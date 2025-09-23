#!/usr/bin/env python3
"""
Vector database for document embeddings and semantic search
Uses ChromaDB for persistent vector storage
"""

import asyncio
import logging
from typing import List, Dict, Any, Optional
import chromadb
from chromadb.config import Settings
from sentence_transformers import SentenceTransformer
import json
import os
from pathlib import Path

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)


class VectorStore:
    """Vector database for document embeddings and semantic search"""

    def __init__(self, persist_directory: str = "./data/chroma_db"):
        self.persist_directory = persist_directory
        self.client = None
        self.collections = {}
        self._sentence_model = None
        self._initialize_client()

    def _initialize_client(self):
        """Initialize ChromaDB client"""
        try:
            # Ensure directory exists
            os.makedirs(self.persist_directory, exist_ok=True)

            self.client = chromadb.PersistentClient(
                path=self.persist_directory,
                settings=Settings(
                    anonymized_telemetry=False,
                    allow_reset=True
                )
            )

            # Initialize collections
            self.collections = {
                "documents": self._get_or_create_collection("documents"),
                "shapes": self._get_or_create_collection("shapes"),
                "presentations": self._get_or_create_collection("presentations")
            }

            logger.info(f"Initialized VectorStore with persistence at {self.persist_directory}")

        except Exception as e:
            logger.error(f"Failed to initialize ChromaDB client: {e}")
            raise

    def _get_or_create_collection(self, name: str):
        """Get or create a collection with specified name"""
        try:
            return self.client.get_or_create_collection(
                name=name,
                metadata={"hnsw:space": "cosine"}
            )
        except Exception as e:
            logger.error(f"Error creating/getting collection {name}: {e}")
            raise

    async def _get_sentence_model(self):
        """Lazy load sentence transformer model"""
        if self._sentence_model is None:
            self._sentence_model = SentenceTransformer('all-MiniLM-L6-v2')
        return self._sentence_model

    async def _generate_embeddings(self, texts: List[str]) -> List[List[float]]:
        """Generate embeddings for texts"""
        try:
            model = await self._get_sentence_model()
            embeddings = model.encode(texts, convert_to_tensor=False)
            return embeddings.tolist() if hasattr(embeddings, 'tolist') else embeddings
        except Exception as e:
            logger.error(f"Error generating embeddings: {e}")
            return []

    async def add_document_chunks(self, doc_id: str, chunks: List[Dict[str, Any]],
                                 metadata: Dict[str, Any] = None) -> bool:
        """
        Add document chunks to vector store

        Args:
            doc_id: Unique document identifier
            chunks: List of processed document chunks
            metadata: Additional metadata for the document

        Returns:
            True if successful, False otherwise
        """
        try:
            if not chunks:
                logger.warning(f"No chunks provided for document {doc_id}")
                return False

            # Extract text content from chunks
            texts = [chunk.get("text", "") for chunk in chunks]

            # Filter out empty texts
            valid_indices = [i for i, text in enumerate(texts) if text.strip()]
            if not valid_indices:
                logger.warning(f"No valid text content in chunks for document {doc_id}")
                return False

            valid_texts = [texts[i] for i in valid_indices]
            valid_chunks = [chunks[i] for i in valid_indices]

            # Generate embeddings
            embeddings = await self._generate_embeddings(valid_texts)
            if not embeddings:
                logger.error(f"Failed to generate embeddings for document {doc_id}")
                return False

            # Prepare metadata for each chunk
            chunk_metadatas = []
            chunk_ids = []

            for i, chunk in enumerate(valid_chunks):
                chunk_metadata = {
                    "doc_id": doc_id,
                    "chunk_id": chunk.get("chunk_id", i),
                    "source_file": chunk.get("source_file", ""),
                    "word_count": chunk.get("word_count", 0),
                    "entities": json.dumps(chunk.get("entities", [])),
                    "key_terms": json.dumps(chunk.get("key_terms", [])),
                    "summary": chunk.get("summary", "")
                }

                # Add document-level metadata
                if metadata:
                    chunk_metadata.update(metadata)

                chunk_metadatas.append(chunk_metadata)
                chunk_ids.append(f"{doc_id}_chunk_{chunk.get('chunk_id', i)}")

            # Add to ChromaDB
            self.collections["documents"].add(
                documents=valid_texts,
                embeddings=embeddings,
                metadatas=chunk_metadatas,
                ids=chunk_ids
            )

            logger.info(f"Added {len(valid_texts)} chunks for document {doc_id}")
            return True

        except Exception as e:
            logger.error(f"Error adding document chunks for {doc_id}: {e}")
            return False

    async def semantic_search(self, query: str, collection: str = "documents",
                             n_results: int = 5, filters: Dict[str, Any] = None) -> List[Dict[str, Any]]:
        """
        Perform semantic search

        Args:
            query: Search query
            collection: Collection to search in
            n_results: Number of results to return
            filters: Metadata filters

        Returns:
            List of search results with documents, metadata, and similarity scores
        """
        try:
            if collection not in self.collections:
                logger.error(f"Collection {collection} not found")
                return []

            # Generate query embedding
            query_embeddings = await self._generate_embeddings([query])
            if not query_embeddings:
                logger.error("Failed to generate query embedding")
                return []

            # Perform search
            search_kwargs = {
                "query_embeddings": query_embeddings,
                "n_results": n_results,
                "include": ["documents", "metadatas", "distances"]
            }

            if filters:
                search_kwargs["where"] = filters

            results = self.collections[collection].query(**search_kwargs)

            # Format results
            formatted_results = []
            if results["documents"] and results["documents"][0]:
                for doc, meta, dist in zip(
                    results["documents"][0],
                    results["metadatas"][0],
                    results["distances"][0]
                ):
                    # Convert distance to similarity (ChromaDB uses cosine distance)
                    similarity = 1 - dist

                    formatted_result = {
                        "document": doc,
                        "metadata": meta,
                        "similarity": similarity,
                        "distance": dist
                    }
                    formatted_results.append(formatted_result)

            logger.info(f"Found {len(formatted_results)} results for query: {query[:50]}...")
            return formatted_results

        except Exception as e:
            logger.error(f"Error performing semantic search: {e}")
            return []

    async def hybrid_search(self, query: str, filters: Dict[str, Any] = None,
                          search_type: str = "hybrid", n_results: int = 10) -> List[Dict[str, Any]]:
        """
        Perform hybrid search combining semantic and keyword search

        Args:
            query: Search query
            filters: Metadata filters
            search_type: Type of search (semantic, keyword, hybrid)
            n_results: Number of results

        Returns:
            Combined search results
        """
        try:
            results = []

            if search_type in ["semantic", "hybrid"]:
                semantic_results = await self.semantic_search(
                    query, n_results=n_results, filters=filters
                )
                for result in semantic_results:
                    result["match_type"] = "semantic"
                results.extend(semantic_results)

            if search_type in ["keyword", "hybrid"]:
                keyword_results = await self._keyword_search(query, filters, n_results)
                for result in keyword_results:
                    result["match_type"] = "keyword"
                results.extend(keyword_results)

            if search_type == "hybrid":
                # Combine and re-rank results
                results = self._combine_search_results(results)

            return results[:n_results]

        except Exception as e:
            logger.error(f"Error performing hybrid search: {e}")
            return []

    async def _keyword_search(self, query: str, filters: Dict[str, Any] = None,
                            n_results: int = 10) -> List[Dict[str, Any]]:
        """Perform keyword-based search"""
        try:
            # Simple keyword matching implementation
            query_terms = set(query.lower().split())

            # Get all documents (this could be optimized for large collections)
            search_kwargs = {"include": ["documents", "metadatas"]}
            if filters:
                search_kwargs["where"] = filters

            all_results = self.collections["documents"].get(**search_kwargs)

            keyword_results = []

            if all_results["documents"]:
                for doc, meta in zip(all_results["documents"], all_results["metadatas"]):
                    doc_terms = set(doc.lower().split())
                    overlap = len(query_terms.intersection(doc_terms))

                    if overlap > 0:
                        # Calculate Jaccard similarity
                        score = overlap / len(query_terms.union(doc_terms))
                        keyword_results.append({
                            "document": doc,
                            "metadata": meta,
                            "similarity": score,
                            "distance": 1 - score
                        })

            # Sort by similarity score
            keyword_results.sort(key=lambda x: x["similarity"], reverse=True)
            return keyword_results[:n_results]

        except Exception as e:
            logger.error(f"Error performing keyword search: {e}")
            return []

    def _combine_search_results(self, results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
        """Combine and re-rank search results from different methods"""
        try:
            # Remove duplicates and combine scores
            doc_scores = {}

            for result in results:
                doc_text = result["document"]
                if doc_text not in doc_scores:
                    doc_scores[doc_text] = {
                        "document": doc_text,
                        "metadata": result["metadata"],
                        "scores": [],
                        "match_types": []
                    }

                doc_scores[doc_text]["scores"].append(result["similarity"])
                doc_scores[doc_text]["match_types"].append(result.get("match_type", "semantic"))

            # Calculate combined scores
            combined_results = []
            for doc_data in doc_scores.values():
                # Weighted average of scores
                semantic_scores = [s for s, t in zip(doc_data["scores"], doc_data["match_types"]) if t == "semantic"]
                keyword_scores = [s for s, t in zip(doc_data["scores"], doc_data["match_types"]) if t == "keyword"]

                combined_score = 0
                if semantic_scores and keyword_scores:
                    combined_score = 0.7 * max(semantic_scores) + 0.3 * max(keyword_scores)
                elif semantic_scores:
                    combined_score = max(semantic_scores)
                elif keyword_scores:
                    combined_score = max(keyword_scores)

                combined_results.append({
                    "document": doc_data["document"],
                    "metadata": doc_data["metadata"],
                    "similarity": combined_score,
                    "distance": 1 - combined_score,
                    "match_types": doc_data["match_types"]
                })

            return sorted(combined_results, key=lambda x: x["similarity"], reverse=True)

        except Exception as e:
            logger.error(f"Error combining search results: {e}")
            return results  # Return original results if combination fails

    async def delete_document(self, doc_id: str) -> bool:
        """Delete all chunks for a document"""
        try:
            # Get all chunk IDs for the document
            results = self.collections["documents"].get(
                where={"doc_id": doc_id},
                include=["metadatas"]
            )

            if not results["ids"]:
                logger.warning(f"No chunks found for document {doc_id}")
                return True

            # Delete all chunks
            self.collections["documents"].delete(ids=results["ids"])
            logger.info(f"Deleted {len(results['ids'])} chunks for document {doc_id}")
            return True

        except Exception as e:
            logger.error(f"Error deleting document {doc_id}: {e}")
            return False

    def get_collection_stats(self) -> Dict[str, Any]:
        """Get statistics about the collections"""
        try:
            stats = {}
            for name, collection in self.collections.items():
                count = collection.count()
                stats[name] = {"document_count": count}

            return stats

        except Exception as e:
            logger.error(f"Error getting collection stats: {e}")
            return {}

    def reset_collections(self) -> bool:
        """Reset all collections (delete all data)"""
        try:
            for name, collection in self.collections.items():
                collection.delete()
                logger.info(f"Reset collection: {name}")
            return True

        except Exception as e:
            logger.error(f"Error resetting collections: {e}")
            return False


async def main():
    """Test the vector store"""
    # Initialize vector store
    vector_store = VectorStore("./test_chroma_db")

    # Test documents
    test_chunks = [
        {
            "text": "PowerPoint presentations are used for business communication and data visualization.",
            "chunk_id": 0,
            "source_file": "test.txt",
            "word_count": 11,
            "entities": ["PowerPoint"],
            "key_terms": ["PowerPoint", "presentations", "business", "communication"],
            "summary": "About PowerPoint presentations"
        },
        {
            "text": "Data analysis and strategic planning require effective presentation tools.",
            "chunk_id": 1,
            "source_file": "test.txt",
            "word_count": 10,
            "entities": [],
            "key_terms": ["data", "analysis", "strategic", "planning", "presentation"],
            "summary": "About data analysis and planning"
        }
    ]

    try:
        # Add test documents
        success = await vector_store.add_document_chunks("test_doc", test_chunks)
        print(f"Added documents: {success}")

        # Test semantic search
        results = await vector_store.semantic_search("business presentations")
        print(f"\nSemantic search results: {len(results)}")
        for i, result in enumerate(results):
            print(f"  {i+1}. Similarity: {result['similarity']:.3f}")
            print(f"     Text: {result['document'][:80]}...")

        # Test hybrid search
        hybrid_results = await vector_store.hybrid_search("data PowerPoint", search_type="hybrid")
        print(f"\nHybrid search results: {len(hybrid_results)}")
        for i, result in enumerate(hybrid_results):
            print(f"  {i+1}. Similarity: {result['similarity']:.3f} ({result.get('match_types', [])})")
            print(f"     Text: {result['document'][:80]}...")

        # Get stats
        stats = vector_store.get_collection_stats()
        print(f"\nCollection stats: {stats}")

    except Exception as e:
        print(f"Error testing vector store: {e}")

    finally:
        # Cleanup
        vector_store.reset_collections()


if __name__ == "__main__":
    asyncio.run(main())