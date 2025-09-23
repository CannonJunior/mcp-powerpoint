#!/usr/bin/env python3
"""
Test RAG workflow with a simple PowerPoint and documents
Demonstrates the complete RAG-enhanced shape naming process
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
from mcp_powerpoint.rag_server import (
    ingest_documents as ingest_documents_tool,
    enhance_shapes_with_documents as enhance_shapes_tool
)


async def create_simple_presentation_json():
    """Create a simple presentation JSON for testing"""
    simple_presentation = {
        "slide_width": 9144000,
        "slide_height": 6858000,
        "slides": [
            {
                "slide_number": 1,
                "name": "Slide 1",
                "shapes": [
                    {
                        "shape_id": 1,
                        "name": "Title 1",
                        "shape_type": "TEXT_BOX",
                        "left": 1000000,
                        "top": 1000000,
                        "width": 7000000,
                        "height": 1000000,
                        "text_frame": {
                            "text": "Project Status Overview",
                            "paragraphs": [
                                {
                                    "text": "Project Status Overview",
                                    "runs": [
                                        {
                                            "text": "Project Status Overview",
                                            "font": {
                                                "name": "Arial",
                                                "size": 24,
                                                "bold": True
                                            }
                                        }
                                    ]
                                }
                            ]
                        }
                    }
                ]
            }
        ],
        "core_properties": {
            "title": "Test Presentation",
            "author": "MCP PowerPoint Tools",
            "created": "2025-01-20T10:00:00Z"
        }
    }

    return json.dumps(simple_presentation, indent=2)


async def test_rag_workflow():
    """Test the complete RAG workflow"""
    print("🚀 Starting RAG Workflow Test")
    print("=" * 50)

    try:
        # Step 1: Create simple presentation data
        print("\n📄 Step 1: Creating simple test presentation...")
        presentation_json = await create_simple_presentation_json()

        # Save test presentation
        test_presentation_path = "test_case/simple_presentation.json"
        with open(test_presentation_path, 'w') as f:
            f.write(presentation_json)
        print(f"✅ Created test presentation: {test_presentation_path}")

        # Step 2: Test document processing
        print("\n📚 Step 2: Testing document processing...")
        processor = DocumentProcessor()

        documents_dir = "test_case/documents"
        if os.path.exists(documents_dir):
            result = await processor.process_directory(documents_dir)
            print(f"✅ Processed {len(result)} documents")

            for file_path, chunks in result.items():
                print(f"  📄 {os.path.basename(file_path)}: {len(chunks)} chunks")
        else:
            print(f"⚠️  Documents directory not found: {documents_dir}")

        # Step 3: Test vector store
        print("\n🔍 Step 3: Testing vector store...")
        vector_store = VectorStore("./test_case/chroma_db")

        # Test search before adding documents
        search_results = await vector_store.semantic_search("project status")
        print(f"📊 Initial search results: {len(search_results)}")

        # Step 4: Test document ingestion via RAG server
        print("\n🔄 Step 4: Testing document ingestion...")
        ingestion_result = await ingest_documents(documents_dir)
        ingestion_data = json.loads(ingestion_result)

        if "error" not in ingestion_data:
            summary = ingestion_data.get("summary", {})
            print(f"✅ Ingestion complete:")
            print(f"  📄 Files processed: {summary.get('total_files_processed', 0)}")
            print(f"  📝 Chunks created: {summary.get('total_chunks_created', 0)}")
            print(f"  ❌ Errors: {summary.get('errors', 0)}")
        else:
            print(f"❌ Ingestion error: {ingestion_data['error']}")
            return

        # Step 5: Test enhanced shape naming
        print("\n🎯 Step 5: Testing enhanced shape naming...")
        enhanced_json = await enhance_shapes_with_documents(
            presentation_json,
            documents_dir,
            strategy="context_aware"
        )

        # Parse and display results
        enhanced_data = json.loads(enhanced_json)

        # Save enhanced presentation
        enhanced_path = "test_case/enhanced_presentation.json"
        with open(enhanced_path, 'w') as f:
            f.write(enhanced_json)
        print(f"✅ Enhanced presentation saved: {enhanced_path}")

        # Display shape enhancements
        print("\n📊 Shape Enhancement Results:")
        print("-" * 40)

        for slide_idx, slide in enumerate(enhanced_data.get("slides", [])):
            print(f"\n📑 Slide {slide_idx + 1}:")

            for shape_idx, shape in enumerate(slide.get("shapes", [])):
                original_name = shape.get("original_name", shape.get("name", "Unknown"))
                descriptive_name = shape.get("descriptive_name", "Not generated")
                confidence = shape.get("confidence_score", 0.0)

                print(f"  🔤 Shape {shape_idx + 1}:")
                print(f"    Original: {original_name}")
                print(f"    Enhanced: {descriptive_name}")
                print(f"    Confidence: {confidence:.2f}")

                if shape.get("context_analysis"):
                    context = shape["context_analysis"][:100] + "..." if len(shape["context_analysis"]) > 100 else shape["context_analysis"]
                    print(f"    Context: {context}")

                if shape.get("semantic_tags"):
                    print(f"    Tags: {', '.join(shape['semantic_tags'])}")

        # Display document context
        doc_context = enhanced_data.get("document_context")
        if doc_context:
            print("\n📚 Document Context:")
            print(f"  Source documents: {len(doc_context.get('source_documents', []))}")
            print(f"  Key concepts: {len(doc_context.get('key_concepts', []))}")
            print(f"  Document type: {doc_context.get('document_type', 'Unknown')}")

        # Test search functionality
        print("\n🔍 Step 6: Testing search functionality...")
        search_results = await vector_store.semantic_search("project status overview", n_results=3)

        if search_results:
            print(f"✅ Found {len(search_results)} relevant document sections:")
            for i, result in enumerate(search_results):
                print(f"  {i+1}. Similarity: {result['similarity']:.3f}")
                text_preview = result['document'][:100] + "..." if len(result['document']) > 100 else result['document']
                print(f"     Text: {text_preview}")
                print(f"     Source: {result['metadata'].get('source_file', 'Unknown')}")
        else:
            print("⚠️  No search results found")

        print("\n🎉 RAG Workflow Test Complete!")
        print("=" * 50)
        print("\nFiles created:")
        print(f"  📄 {test_presentation_path}")
        print(f"  📄 {enhanced_path}")
        print(f"  📂 ./test_case/chroma_db/ (vector database)")

    except Exception as e:
        print(f"\n❌ Error during RAG workflow test: {e}")
        import traceback
        traceback.print_exc()


async def test_individual_components():
    """Test individual components separately"""
    print("\n🧪 Testing Individual Components")
    print("=" * 40)

    try:
        # Test document processor
        print("\n📄 Testing Document Processor...")
        processor = DocumentProcessor()

        test_doc = "test_case/documents/project_overview.txt"
        if os.path.exists(test_doc):
            chunks = await processor.process_document(test_doc)
            print(f"✅ Processed document into {len(chunks)} chunks")

            if chunks:
                first_chunk = chunks[0]
                print(f"  First chunk:")
                print(f"    Text length: {len(first_chunk['text'])}")
                print(f"    Entities: {first_chunk.get('entities', [])}")
                print(f"    Key terms: {first_chunk.get('key_terms', [])}")

        # Test vector store
        print("\n🔍 Testing Vector Store...")
        vector_store = VectorStore("./test_case/test_chroma_db")

        # Add test data
        test_chunks = [
            {
                "text": "PowerPoint automation enables efficient presentation creation and content management.",
                "chunk_id": 0,
                "source_file": "test.txt",
                "word_count": 10,
                "entities": ["PowerPoint"],
                "key_terms": ["PowerPoint", "automation", "presentation"],
                "summary": "About PowerPoint automation"
            }
        ]

        success = await vector_store.add_document_chunks("test_doc", test_chunks)
        print(f"✅ Added test chunks: {success}")

        # Test search
        results = await vector_store.semantic_search("presentation automation")
        print(f"✅ Search returned {len(results)} results")

        if results:
            print(f"  Top result similarity: {results[0]['similarity']:.3f}")

        # Cleanup test vector store
        vector_store.reset_collections()
        print("✅ Cleaned up test vector store")

    except Exception as e:
        print(f"❌ Error testing individual components: {e}")
        import traceback
        traceback.print_exc()


async def main():
    """Main test function"""
    print("🧪 MCP PowerPoint RAG System Test Suite")
    print("=" * 60)

    # Ensure test directories exist
    os.makedirs("test_case/documents", exist_ok=True)
    os.makedirs("test_case", exist_ok=True)

    # Test individual components first
    await test_individual_components()

    # Test complete workflow
    await test_rag_workflow()

    print("\n✨ All tests completed!")


if __name__ == "__main__":
    asyncio.run(main())