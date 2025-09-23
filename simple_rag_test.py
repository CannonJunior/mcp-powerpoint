#!/usr/bin/env python3
"""
Simple RAG test to verify the core functionality
"""

import asyncio
import json
import os
import sys
from pathlib import Path

# Add the package to the path
sys.path.insert(0, str(Path(__file__).parent))

from mcp_powerpoint.document_processor import DocumentProcessor
from mcp_powerpoint.vector_store import VectorStore


async def main():
    """Simple test of core RAG functionality"""
    print("🧪 Simple RAG System Test")
    print("=" * 40)

    try:
        # Test 1: Document Processing
        print("\n📄 Testing Document Processing...")
        processor = DocumentProcessor()

        test_text = """
        PowerPoint presentations are essential tools for business communication.
        This project implements intelligent shape naming using machine learning.
        The system processes documents to understand context and improve naming accuracy.
        """

        # Create test file
        test_file = "temp_test.txt"
        with open(test_file, 'w') as f:
            f.write(test_text)

        try:
            chunks = await processor.process_document(test_file)
            print(f"✅ Processed document into {len(chunks)} chunks")

            if chunks:
                chunk = chunks[0]
                print(f"  📝 Chunk text length: {len(chunk['text'])}")
                print(f"  🔤 Entities: {chunk.get('entities', [])}")
                print(f"  🏷️  Key terms: {chunk.get('key_terms', [])}")
                print(f"  📋 Summary: {chunk.get('summary', '')}")
        finally:
            if os.path.exists(test_file):
                os.unlink(test_file)

        # Test 2: Vector Store
        print("\n🔍 Testing Vector Store...")
        vector_store = VectorStore("./simple_test_db")

        # Create test chunks
        test_chunks = [
            {
                "text": "PowerPoint presentations require intelligent shape naming for better organization.",
                "chunk_id": 0,
                "source_file": "test.txt",
                "word_count": 10,
                "entities": ["PowerPoint"],
                "key_terms": ["PowerPoint", "presentations", "shape", "naming"],
                "summary": "About PowerPoint shape naming"
            },
            {
                "text": "Business presentations benefit from automated content analysis and context-aware naming.",
                "chunk_id": 1,
                "source_file": "test.txt",
                "word_count": 11,
                "entities": [],
                "key_terms": ["business", "presentations", "automated", "analysis"],
                "summary": "About business presentation automation"
            }
        ]

        # Add documents to vector store
        success = await vector_store.add_document_chunks("test_doc", test_chunks)
        print(f"✅ Added test chunks to vector store: {success}")

        # Test search
        search_results = await vector_store.semantic_search("PowerPoint shape naming", n_results=3)
        print(f"✅ Search returned {len(search_results)} results")

        for i, result in enumerate(search_results):
            print(f"  {i+1}. Similarity: {result['similarity']:.3f}")
            text_preview = result['document'][:80] + "..." if len(result['document']) > 80 else result['document']
            print(f"     Text: {text_preview}")

        # Test 3: Simple Presentation Enhancement
        print("\n🎯 Testing Presentation Enhancement...")

        # Simple presentation data
        simple_presentation = {
            "slides": [
                {
                    "slide_number": 1,
                    "shapes": [
                        {
                            "shape_id": 1,
                            "name": "Rectangle 1",
                            "shape_type": "TEXT_BOX",
                            "text_frame": {
                                "text": "Project Status Overview"
                            }
                        }
                    ]
                }
            ]
        }

        # Find relevant context for the shape
        shape_text = "Project Status Overview"
        context_results = await vector_store.semantic_search(shape_text, n_results=2)

        if context_results:
            print(f"✅ Found {len(context_results)} relevant context chunks")

            # Simple enhancement simulation
            best_match = context_results[0]
            confidence = best_match['similarity']

            # Generate a simple enhanced name
            enhanced_name = "project_status_overview"

            print(f"  📊 Shape Enhancement:")
            print(f"     Original: {simple_presentation['slides'][0]['shapes'][0]['name']}")
            print(f"     Enhanced: {enhanced_name}")
            print(f"     Confidence: {confidence:.3f}")
            print(f"     Context: {best_match['document'][:100]}...")

        # Clean up
        vector_store.reset_collections()
        print("\n✅ Cleaned up test database")

        print("\n🎉 Simple RAG test completed successfully!")

    except Exception as e:
        print(f"\n❌ Error during test: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    asyncio.run(main())