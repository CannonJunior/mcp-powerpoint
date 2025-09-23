#!/usr/bin/env python3
"""
RAG (Retrieval-Augmented Generation) MCP Server
Provides document analysis and context-aware content generation
"""

import asyncio
import json
import logging
import os
import glob
from datetime import datetime
from typing import List, Dict, Any, Optional
from pathlib import Path

from fastmcp import FastMCP
import ollama

from .document_processor import DocumentProcessor
from .vector_store import VectorStore
from .powerpoint_models import DocumentContext, ProcessingMetadata

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize MCP server
mcp = FastMCP("RAG System Server")

# Global instances
document_processor = None
vector_store = None


def get_document_processor():
    """Get or create document processor instance"""
    global document_processor
    if document_processor is None:
        document_processor = DocumentProcessor()
    return document_processor


def get_vector_store():
    """Get or create vector store instance"""
    global vector_store
    if vector_store is None:
        vector_store = VectorStore()
    return vector_store


@mcp.tool()
async def ingest_documents(document_directory: str,
                          file_patterns: List[str] = None) -> str:
    """
    Ingest documents into RAG system for contextual analysis

    Args:
        document_directory: Directory containing documents to ingest
        file_patterns: File patterns to match (default: ["*.txt", "*.md", "*.pdf", "*.docx"])

    Returns:
        Ingestion summary with document count and processing details
    """
    try:
        if not file_patterns:
            file_patterns = ["*.txt", "*.md", "*.pdf", "*.docx"]

        if not os.path.exists(document_directory):
            return json.dumps({
                "error": f"Directory does not exist: {document_directory}",
                "summary": {"total_files_processed": 0, "total_chunks_created": 0, "errors": 1}
            })

        processor = get_document_processor()
        vector_store = get_vector_store()

        ingestion_results = []

        for pattern in file_patterns:
            files = glob.glob(os.path.join(document_directory, pattern))

            for file_path in files:
                try:
                    logger.info(f"Processing file: {file_path}")

                    # Process document
                    chunks = await processor.process_document(file_path)

                    if not chunks:
                        ingestion_results.append({
                            "file": file_path,
                            "status": "warning",
                            "message": "No content extracted",
                            "chunks": 0
                        })
                        continue

                    # Generate metadata
                    metadata = {
                        "source_file": file_path,
                        "file_type": os.path.splitext(file_path)[1],
                        "processed_at": datetime.utcnow().isoformat(),
                        "chunk_count": len(chunks)
                    }

                    # Add to vector store
                    doc_id = os.path.basename(file_path)
                    success = await vector_store.add_document_chunks(doc_id, chunks, metadata)

                    if success:
                        ingestion_results.append({
                            "file": file_path,
                            "status": "success",
                            "chunks": len(chunks)
                        })
                    else:
                        ingestion_results.append({
                            "file": file_path,
                            "status": "error",
                            "error": "Failed to add to vector store",
                            "chunks": 0
                        })

                except Exception as e:
                    logger.error(f"Error processing {file_path}: {e}")
                    ingestion_results.append({
                        "file": file_path,
                        "status": "error",
                        "error": str(e),
                        "chunks": 0
                    })

        # Calculate summary
        successful = [r for r in ingestion_results if r["status"] == "success"]
        errors = [r for r in ingestion_results if r["status"] == "error"]
        total_chunks = sum(r.get("chunks", 0) for r in ingestion_results)

        summary = {
            "total_files_processed": len(successful),
            "total_chunks_created": total_chunks,
            "errors": len(errors),
            "warnings": len([r for r in ingestion_results if r["status"] == "warning"])
        }

        logger.info(f"Ingestion complete: {summary}")

        return json.dumps({
            "summary": summary,
            "results": ingestion_results
        }, indent=2)

    except Exception as e:
        logger.error(f"Error in document ingestion: {e}")
        return json.dumps({
            "error": str(e),
            "summary": {"total_files_processed": 0, "total_chunks_created": 0, "errors": 1}
        })


@mcp.tool()
async def contextual_shape_analysis(json_data: str, query_context: str = "") -> str:
    """
    Analyze shapes using RAG context from ingested documents

    Args:
        json_data: PowerPoint JSON data
        query_context: Additional context for analysis

    Returns:
        Contextual analysis results with document references
    """
    try:
        presentation_data = json.loads(json_data)
        vector_store = get_vector_store()

        analysis_results = {
            "contextual_insights": [],
            "shape_enhancements": [],
            "document_references": [],
            "processing_metadata": {
                "timestamp": datetime.utcnow().isoformat(),
                "query_context": query_context
            }
        }

        # Extract all text content for context analysis
        all_text_content = []
        shape_contexts = []

        for slide_idx, slide in enumerate(presentation_data.get("slides", [])):
            for shape_idx, shape in enumerate(slide.get("shapes", [])):
                if shape.get("text_frame") and shape["text_frame"].get("text"):
                    text_content = shape["text_frame"]["text"]
                    all_text_content.append(text_content)
                    shape_contexts.append({
                        "slide_index": slide_idx,
                        "shape_index": shape_idx,
                        "shape_name": shape.get("name", "Unknown"),
                        "text_content": text_content
                    })

        if not all_text_content:
            return json.dumps({
                "error": "No text content found in presentation",
                "analysis_results": analysis_results
            })

        # Combine all text for global context search
        combined_content = " ".join(all_text_content)
        if query_context:
            combined_content += " " + query_context

        # Perform semantic search for relevant context
        context_results = await vector_store.semantic_search(
            combined_content,
            n_results=10
        )

        analysis_results["document_references"] = [
            {
                "document": result["document"][:200] + "..." if len(result["document"]) > 200 else result["document"],
                "similarity": result["similarity"],
                "source_file": result["metadata"].get("source_file", "Unknown"),
                "summary": result["metadata"].get("summary", "")
            }
            for result in context_results
        ]

        # Enhance individual shapes with context
        for shape_context in shape_contexts:
            shape_text = shape_context["text_content"]

            # Find relevant context for this specific shape
            shape_context_results = await vector_store.semantic_search(
                shape_text, n_results=3
            )

            # Generate enhanced analysis with context
            enhanced_analysis = await _analyze_shape_with_context(
                shape_context, shape_context_results
            )

            analysis_results["shape_enhancements"].append({
                "slide_index": shape_context["slide_index"],
                "shape_index": shape_context["shape_index"],
                "original_name": shape_context["shape_name"],
                "enhanced_analysis": enhanced_analysis
            })

        # Generate contextual insights
        analysis_results["contextual_insights"] = await _generate_contextual_insights(
            presentation_data, context_results
        )

        return json.dumps(analysis_results, indent=2)

    except json.JSONDecodeError as e:
        logger.error(f"Invalid JSON data: {e}")
        return json.dumps({"error": f"Invalid JSON data: {e}"})
    except Exception as e:
        logger.error(f"Error in contextual shape analysis: {e}")
        return json.dumps({"error": str(e)})


@mcp.tool()
async def enhance_shapes_with_documents(json_data: str, documents_dir: str,
                                       strategy: str = "context_aware") -> str:
    """
    Enhanced shape naming using document context - simplified single tool approach

    Args:
        json_data: PowerPoint JSON data
        documents_dir: Directory containing context documents
        strategy: Enhancement strategy (context_aware, semantic, functional)

    Returns:
        Enhanced JSON with improved descriptive names
    """
    try:
        # First, ingest documents if directory provided
        if documents_dir and os.path.exists(documents_dir):
            await ingest_documents(documents_dir)

        # Perform contextual analysis
        analysis_result = await contextual_shape_analysis(json_data)
        analysis_data = json.loads(analysis_result)

        if "error" in analysis_data:
            return json_data  # Return original if analysis failed

        # Parse original presentation data
        presentation_data = json.loads(json_data)

        # Apply enhancements from analysis
        shape_enhancements = analysis_data.get("shape_enhancements", [])

        for enhancement in shape_enhancements:
            slide_idx = enhancement["slide_index"]
            shape_idx = enhancement["shape_index"]

            if (slide_idx < len(presentation_data.get("slides", [])) and
                shape_idx < len(presentation_data["slides"][slide_idx].get("shapes", []))):

                shape = presentation_data["slides"][slide_idx]["shapes"][shape_idx]
                enhanced_analysis = enhancement.get("enhanced_analysis", {})

                # Update shape with enhanced information
                shape["original_name"] = enhancement["original_name"]
                shape["descriptive_name"] = enhanced_analysis.get("suggested_name", shape.get("name"))
                shape["confidence_score"] = enhanced_analysis.get("confidence", 0.5)
                shape["context_analysis"] = enhanced_analysis.get("context_summary", "")
                shape["naming_rationale"] = enhanced_analysis.get("rationale", "")
                shape["semantic_tags"] = enhanced_analysis.get("semantic_tags", [])
                shape["document_references"] = enhanced_analysis.get("document_references", [])

        # Add document context to presentation
        document_context = DocumentContext(
            source_documents=analysis_data.get("document_references", []),
            key_concepts=analysis_data.get("contextual_insights", []),
            document_type="mixed",
            confidence_metrics={"analysis_quality": 0.8}
        )

        processing_metadata = ProcessingMetadata(
            llm_model_used="ollama",
            processing_time_ms=1000,  # Placeholder
            quality_score=0.8
        )

        presentation_data["document_context"] = document_context.dict()
        presentation_data["processing_metadata"] = processing_metadata.dict()

        return json.dumps(presentation_data, indent=2)

    except Exception as e:
        logger.error(f"Error enhancing shapes with documents: {e}")
        return json_data  # Return original data if enhancement fails


async def _analyze_shape_with_context(shape_context: Dict[str, Any],
                                     context_results: List[Dict[str, Any]]) -> Dict[str, Any]:
    """Analyze a shape with document context to generate enhanced naming"""
    try:
        shape_text = shape_context["text_content"]
        shape_name = shape_context["shape_name"]

        # Extract relevant context
        context_summary = ""
        document_refs = []

        for result in context_results[:3]:  # Top 3 results
            context_summary += f"{result['document'][:100]}... "
            doc_ref = result["metadata"].get("source_file", "Unknown")
            if doc_ref not in document_refs:
                document_refs.append(doc_ref)

        # Generate enhanced naming using LLM
        enhanced_name = await _generate_enhanced_name(shape_text, shape_name, context_summary)

        # Extract semantic tags
        semantic_tags = await _extract_semantic_tags(shape_text, context_summary)

        return {
            "suggested_name": enhanced_name,
            "confidence": 0.8,  # Placeholder confidence score
            "context_summary": context_summary[:200] + "..." if len(context_summary) > 200 else context_summary,
            "rationale": f"Based on shape text '{shape_text[:50]}...' and document context",
            "semantic_tags": semantic_tags,
            "document_references": document_refs
        }

    except Exception as e:
        logger.error(f"Error analyzing shape with context: {e}")
        return {
            "suggested_name": shape_context["shape_name"],
            "confidence": 0.3,
            "context_summary": "Error in analysis",
            "rationale": f"Error: {str(e)}",
            "semantic_tags": [],
            "document_references": []
        }


async def _generate_enhanced_name(shape_text: str, original_name: str, context: str) -> str:
    """Generate enhanced shape name using LLM"""
    try:
        prompt = f"""
Given the following information:
- Original shape name: {original_name}
- Shape text content: {shape_text}
- Document context: {context[:300]}

Generate a descriptive, programmatic name for this PowerPoint shape. The name should be:
- Lowercase with underscores (snake_case)
- Descriptive of the content or function
- Maximum 3 words
- Professional and clear

Examples:
- "Rectangle 5" with text "Company Overview" → "company_overview"
- "TextBox 42" with text "Q3 Sales Results" → "q3_sales_results"
- "AutoShape 12" with text "Next Steps" → "next_steps"

Return only the name, nothing else:
"""

        response = ollama.generate(
            model='llama3.2',
            prompt=prompt,
            options={'temperature': 0.3, 'num_predict': 20}
        )

        enhanced_name = response['response'].strip().lower()

        # Clean up the response
        enhanced_name = ''.join(c if c.isalnum() or c == '_' else '_' for c in enhanced_name)
        enhanced_name = '_'.join(part for part in enhanced_name.split('_') if part)

        return enhanced_name if enhanced_name else original_name.lower().replace(' ', '_')

    except Exception as e:
        logger.error(f"Error generating enhanced name: {e}")
        return original_name.lower().replace(' ', '_')


async def _extract_semantic_tags(shape_text: str, context: str) -> List[str]:
    """Extract semantic tags from shape text and context"""
    try:
        # Simple keyword extraction
        combined_text = f"{shape_text} {context}".lower()
        words = combined_text.split()

        # Common business/presentation keywords
        business_keywords = {
            'revenue', 'sales', 'profit', 'growth', 'market', 'strategy', 'goals',
            'objectives', 'results', 'performance', 'analysis', 'data', 'metrics',
            'overview', 'summary', 'conclusion', 'recommendation', 'action', 'plan'
        }

        found_tags = []
        for word in words:
            clean_word = ''.join(c for c in word if c.isalnum())
            if clean_word in business_keywords:
                found_tags.append(clean_word)

        # Remove duplicates and limit to top 5
        return list(set(found_tags))[:5]

    except Exception as e:
        logger.error(f"Error extracting semantic tags: {e}")
        return []


async def _generate_contextual_insights(presentation_data: Dict[str, Any],
                                       context_results: List[Dict[str, Any]]) -> List[str]:
    """Generate contextual insights about the presentation"""
    try:
        insights = []

        # Analyze document types
        doc_types = set()
        for result in context_results:
            file_type = result["metadata"].get("file_type", "unknown")
            doc_types.add(file_type)

        if doc_types:
            insights.append(f"Document context includes: {', '.join(doc_types)}")

        # Analyze content themes
        if context_results:
            high_similarity = [r for r in context_results if r["similarity"] > 0.7]
            if high_similarity:
                insights.append(f"Found {len(high_similarity)} highly relevant document sections")

        # Analyze presentation structure
        slide_count = len(presentation_data.get("slides", []))
        total_shapes = sum(len(slide.get("shapes", [])) for slide in presentation_data.get("slides", []))

        insights.append(f"Presentation structure: {slide_count} slides, {total_shapes} total shapes")

        return insights

    except Exception as e:
        logger.error(f"Error generating contextual insights: {e}")
        return ["Error generating insights"]


@mcp.tool()
async def search_documents(query: str, max_results: int = 5) -> str:
    """
    Search ingested documents for relevant content

    Args:
        query: Search query
        max_results: Maximum number of results to return

    Returns:
        JSON with search results
    """
    try:
        vector_store = get_vector_store()

        results = await vector_store.semantic_search(
            query,
            n_results=max_results
        )

        formatted_results = [
            {
                "text": result["document"][:300] + "..." if len(result["document"]) > 300 else result["document"],
                "similarity": result["similarity"],
                "source_file": result["metadata"].get("source_file", "Unknown"),
                "summary": result["metadata"].get("summary", "")
            }
            for result in results
        ]

        return json.dumps({
            "query": query,
            "results_count": len(formatted_results),
            "results": formatted_results
        }, indent=2)

    except Exception as e:
        logger.error(f"Error searching documents: {e}")
        return json.dumps({"error": str(e)})


@mcp.tool()
async def get_vector_store_stats() -> str:
    """
    Get statistics about the vector store

    Returns:
        JSON with vector store statistics
    """
    try:
        vector_store = get_vector_store()
        stats = vector_store.get_collection_stats()

        return json.dumps({
            "statistics": stats,
            "timestamp": datetime.utcnow().isoformat()
        }, indent=2)

    except Exception as e:
        logger.error(f"Error getting vector store stats: {e}")
        return json.dumps({"error": str(e)})


def main():
    """Main entry point for RAG server"""
    logger.info("Starting RAG MCP Server...")
    mcp.run()


if __name__ == "__main__":
    main()