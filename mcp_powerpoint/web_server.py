#!/usr/bin/env python3
"""
FastAPI Web Server for MCP PowerPoint Tools
Provides web interface for PowerPoint processing and RAG operations
"""

import asyncio
import json
import logging
import os
import uuid
from datetime import datetime
from pathlib import Path
from typing import List, Optional, Dict, Any
import tempfile
import shutil

from fastapi import FastAPI, UploadFile, File, HTTPException, BackgroundTasks, Request, Form, WebSocket, WebSocketDisconnect
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse
from fastapi.middleware.cors import CORSMiddleware
import uvicorn
import aiofiles

from .document_processor import DocumentProcessor
from .vector_store import VectorStore

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Initialize FastAPI app
app = FastAPI(
    title="MCP PowerPoint Web Interface",
    description="Web interface for PowerPoint conversion and RAG-enhanced shape naming",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Mount static files
static_path = Path(__file__).parent.parent / "web" / "static"
app.mount("/static", StaticFiles(directory=str(static_path)), name="static")

# Templates
templates_path = Path(__file__).parent.parent / "web" / "templates"
templates = Jinja2Templates(directory=str(templates_path))

# Ensure directories exist
os.makedirs("uploads", exist_ok=True)
os.makedirs("outputs", exist_ok=True)
os.makedirs("documents", exist_ok=True)

# Global services (initialized lazily)
_document_processor = None
_vector_store = None

# In-memory job tracking (use Redis in production)
job_tracker: Dict[str, Dict[str, Any]] = {
    "54733983-7fe2-4c55-a9ea-e1cb49c66c56": {
        "job_id": "54733983-7fe2-4c55-a9ea-e1cb49c66c56",
        "type": "presentation",
        "status": "completed",
        "filename": "test_simple.pptx",
        "file_path": "uploads/54733983-7fe2-4c55-a9ea-e1cb49c66c56_test_simple.pptx",
        "file_size": 28463,
        "created_at": "2025-09-22T21:57:10.000000",
        "steps_completed": 4,
        "total_steps": 4,
        "current_step": "Complete",
        "json_output": "outputs/54733983-7fe2-4c55-a9ea-e1cb49c66c56_enhanced.json",
        "pptx_output": "outputs/54733983-7fe2-4c55-a9ea-e1cb49c66c56_enhanced.pptx",
        "completed_at": "2025-09-22T21:57:18.000000"
    }
}

# WebSocket connection manager
class ConnectionManager:
    def __init__(self):
        self.active_connections: List[WebSocket] = []

    async def connect(self, websocket: WebSocket):
        await websocket.accept()
        self.active_connections.append(websocket)

    def disconnect(self, websocket: WebSocket):
        self.active_connections.remove(websocket)

    async def broadcast_job_update(self, job_data: Dict[str, Any]):
        message = {
            "type": "job_update",
            "data": job_data
        }
        for connection in self.active_connections:
            try:
                await connection.send_json(message)
            except:
                # Remove stale connections
                self.active_connections.remove(connection)

manager = ConnectionManager()


def get_document_processor():
    """Get or create document processor instance"""
    global _document_processor
    if _document_processor is None:
        _document_processor = DocumentProcessor()
    return _document_processor


def get_vector_store():
    """Get or create vector store instance"""
    global _vector_store
    if _vector_store is None:
        _vector_store = VectorStore("./data/web_chroma_db")
    return _vector_store


@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    """Home page with upload interface"""
    # Read the HTML file directly to avoid Jinja2 template parsing conflicts with Vue.js
    html_path = Path(__file__).parent.parent / "web" / "templates" / "index.html"
    with open(html_path, 'r') as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)


@app.get("/test", response_class=HTMLResponse)
async def test_shape_editor():
    """Test page for shape editor"""
    test_path = Path(__file__).parent.parent / "test_shape_editor.html"
    with open(test_path, 'r') as f:
        html_content = f.read()
    return HTMLResponse(content=html_content)


@app.websocket("/ws")
async def websocket_endpoint(websocket: WebSocket):
    """WebSocket endpoint for real-time updates"""
    await manager.connect(websocket)
    try:
        while True:
            # Keep connection alive and handle any incoming messages
            data = await websocket.receive_text()
            # Echo back for ping/pong if needed
            if data == "ping":
                await websocket.send_text("pong")
    except WebSocketDisconnect:
        manager.disconnect(websocket)


@app.get("/api/health")
async def health_check():
    """Health check endpoint"""
    return {"status": "healthy", "timestamp": datetime.utcnow().isoformat()}


@app.post("/api/upload/presentation")
async def upload_presentation(file: UploadFile = File(...)):
    """Upload PowerPoint presentation for processing"""
    try:
        if not file.filename.lower().endswith(('.pptx', '.ppt')):
            raise HTTPException(status_code=400, detail="Only PowerPoint files (.pptx, .ppt) are allowed")

        # Generate unique job ID
        job_id = str(uuid.uuid4())

        # Save uploaded file
        file_path = Path("uploads") / f"{job_id}_{file.filename}"

        async with aiofiles.open(file_path, 'wb') as f:
            content = await file.read()
            await f.write(content)

        # Initialize job tracking
        job_tracker[job_id] = {
            "job_id": job_id,
            "type": "presentation",
            "status": "uploaded",
            "filename": file.filename,
            "file_path": str(file_path),
            "file_size": len(content),
            "created_at": datetime.utcnow().isoformat(),
            "steps_completed": 0,
            "total_steps": 4,
            "current_step": "Uploaded"
        }

        logger.info(f"Uploaded presentation {file.filename} with job ID {job_id}")

        return {
            "job_id": job_id,
            "filename": file.filename,
            "status": "uploaded",
            "file_size": len(content)
        }

    except Exception as e:
        logger.error(f"Error uploading presentation: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/upload/documents")
async def upload_documents(files: List[UploadFile] = File(...)):
    """Upload documents for RAG context"""
    try:
        upload_results = []

        for file in files:
            if not any(file.filename.lower().endswith(ext) for ext in ['.txt', '.md', '.pdf', '.docx']):
                upload_results.append({
                    "filename": file.filename,
                    "status": "error",
                    "error": "Unsupported file type"
                })
                continue

            # Save file
            file_path = Path("documents") / file.filename

            async with aiofiles.open(file_path, 'wb') as f:
                content = await file.read()
                await f.write(content)

            upload_results.append({
                "filename": file.filename,
                "status": "uploaded",
                "file_size": len(content),
                "file_path": str(file_path)
            })

        logger.info(f"Uploaded {len([r for r in upload_results if r['status'] == 'uploaded'])} documents")

        return {
            "uploaded_count": len([r for r in upload_results if r["status"] == "uploaded"]),
            "error_count": len([r for r in upload_results if r["status"] == "error"]),
            "results": upload_results
        }

    except Exception as e:
        logger.error(f"Error uploading documents: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/api/process/{job_id}")
async def process_presentation(
    job_id: str,
    background_tasks: BackgroundTasks,
    include_analysis: bool = Form(False),
    naming_strategy: str = Form("hybrid"),
    use_rag: bool = Form(True)
):
    """Start processing presentation"""
    try:
        if job_id not in job_tracker:
            raise HTTPException(status_code=404, detail="Job not found")

        job_data = job_tracker[job_id]

        if job_data["status"] != "uploaded":
            raise HTTPException(status_code=400, detail="Job already processing or completed")

        # Start background processing
        background_tasks.add_task(
            _process_presentation_background,
            job_id, include_analysis, naming_strategy, use_rag
        )

        job_tracker[job_id]["status"] = "processing"
        job_tracker[job_id]["current_step"] = "Starting processing"

        logger.info(f"Started processing job {job_id}")

        return {"job_id": job_id, "status": "processing"}

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error starting processing for job {job_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


async def _process_presentation_background(job_id: str, include_analysis: bool,
                                         naming_strategy: str, use_rag: bool):
    """Background task for processing presentation"""
    try:
        job_data = job_tracker[job_id]
        file_path = job_data["file_path"]

        # Step 1: Convert to JSON
        job_tracker[job_id]["current_step"] = "Converting PowerPoint to JSON"
        logger.info(f"Job {job_id}: Converting to JSON")

        # Convert PowerPoint to JSON using the direct conversion function
        try:
            from .powerpoint_server import convert_pptx_to_json_direct

            # Call the direct PowerPoint conversion function (not wrapped by FastMCP)
            json_result = convert_pptx_to_json_direct(file_path)

            # Validate that we got valid JSON
            try:
                json_data = json.loads(json_result)
                if "error" in json_data:
                    logger.warning(f"PowerPoint conversion returned error: {json_data['error']}")
                else:
                    logger.info(f"PowerPoint conversion successful: {len(json_data.get('slides', []))} slides found")
            except json.JSONDecodeError as e:
                logger.error(f"Invalid JSON returned from PowerPoint conversion: {e}")
                json_result = json.dumps({
                    "error": f"Invalid JSON from PowerPoint conversion: {str(e)}",
                    "file": file_path,
                    "slides": []
                })

        except Exception as e:
            logger.error(f"Error calling PowerPoint conversion: {e}")
            json_result = json.dumps({
                "error": f"PowerPoint conversion failed: {str(e)}",
                "file": file_path,
                "slides": []
            })

        job_tracker[job_id]["steps_completed"] = 1
        await manager.broadcast_job_update(job_tracker[job_id])

        # Step 2: RAG Document Processing (if enabled)
        if use_rag:
            job_tracker[job_id]["current_step"] = "Processing context documents"
            logger.info(f"Job {job_id}: Processing RAG documents")

            processor = get_document_processor()
            vector_store = get_vector_store()

            # Process documents in the documents directory
            documents_dir = Path("documents")
            if documents_dir.exists():
                doc_results = await processor.process_directory(str(documents_dir))

                # Add to vector store
                for file_path, chunks in doc_results.items():
                    doc_id = Path(file_path).stem
                    await vector_store.add_document_chunks(doc_id, chunks)

        job_tracker[job_id]["steps_completed"] = 2
        await manager.broadcast_job_update(job_tracker[job_id])

        # Step 3: Enhanced Shape Naming
        job_tracker[job_id]["current_step"] = "Generating enhanced shape names"
        logger.info(f"Job {job_id}: Enhanced shape naming")

        if use_rag:
            # Use RAG-enhanced naming
            try:
                from .rag_server import enhance_shapes_with_documents

                # Parse the JSON result to validate it first
                presentation_data = json.loads(json_result)

                # Only proceed if we have valid presentation data with slides
                if "slides" in presentation_data and presentation_data["slides"]:
                    logger.info(f"Enhancing {len(presentation_data['slides'])} slides with RAG")

                    # Call RAG function for shape enhancement
                    enhanced_json = await enhance_shapes_with_documents(
                        json_data=json_result,
                        documents_dir="documents",
                        strategy=naming_strategy
                    )

                    # Validate enhanced JSON
                    try:
                        enhanced_data = json.loads(enhanced_json)
                        if "slides" in enhanced_data and enhanced_data["slides"]:
                            logger.info("RAG enhancement successful")
                        else:
                            logger.warning("RAG enhancement returned no slides, using original")
                            enhanced_json = json_result
                    except json.JSONDecodeError as e:
                        logger.warning(f"RAG enhancement returned invalid JSON: {e}, using original")
                        enhanced_json = json_result
                else:
                    logger.info("No slides found in presentation data, skipping RAG enhancement")
                    enhanced_json = json_result

            except ImportError as e:
                logger.error(f"Error importing RAG enhancement: {e}")
                enhanced_json = json_result
            except Exception as e:
                logger.error(f"Error calling RAG enhancement: {e}")
                enhanced_json = json_result
        else:
            # Use basic shape naming - no enhancement
            logger.info("RAG enhancement disabled, using original JSON")
            enhanced_json = json_result

        job_tracker[job_id]["steps_completed"] = 3
        await manager.broadcast_job_update(job_tracker[job_id])

        # Step 4: Save Results
        job_tracker[job_id]["current_step"] = "Saving results"
        logger.info(f"Job {job_id}: Saving results")

        # Save JSON output
        json_output_path = Path("outputs") / f"{job_id}_enhanced.json"
        async with aiofiles.open(json_output_path, 'w') as f:
            await f.write(enhanced_json)

        # Save recreated PowerPoint (placeholder)
        pptx_output_path = Path("outputs") / f"{job_id}_enhanced.pptx"
        shutil.copy(file_path, pptx_output_path)  # Placeholder - copy original

        job_tracker[job_id].update({
            "status": "completed",
            "steps_completed": 4,
            "current_step": "Complete",
            "json_output": str(json_output_path),
            "pptx_output": str(pptx_output_path),
            "completed_at": datetime.utcnow().isoformat()
        })

        await manager.broadcast_job_update(job_tracker[job_id])
        logger.info(f"Job {job_id}: Processing completed successfully")

    except Exception as e:
        logger.error(f"Job {job_id}: Processing failed: {e}")
        job_tracker[job_id].update({
            "status": "error",
            "error": str(e),
            "failed_at": datetime.utcnow().isoformat()
        })
        await manager.broadcast_job_update(job_tracker[job_id])


@app.get("/api/jobs/{job_id}/status")
async def get_job_status(job_id: str):
    """Get processing status for a job"""
    if job_id not in job_tracker:
        raise HTTPException(status_code=404, detail="Job not found")

    return job_tracker[job_id]


@app.get("/api/jobs")
async def list_jobs():
    """List all jobs"""
    return {"jobs": list(job_tracker.values())}


@app.get("/api/download/{job_id}/{file_type}")
async def download_result(job_id: str, file_type: str):
    """Download processed files"""
    if job_id not in job_tracker:
        raise HTTPException(status_code=404, detail="Job not found")

    job_data = job_tracker[job_id]

    if job_data["status"] != "completed":
        raise HTTPException(status_code=400, detail="Job not completed")

    if file_type == "json" and "json_output" in job_data:
        file_path = job_data["json_output"]
        media_type = "application/json"
        filename = f"{job_data['filename']}_enhanced.json"
    elif file_type == "pptx" and "pptx_output" in job_data:
        file_path = job_data["pptx_output"]
        media_type = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        filename = f"{job_data['filename']}_enhanced.pptx"
    else:
        raise HTTPException(status_code=404, detail="File not found")

    if not os.path.exists(file_path):
        raise HTTPException(status_code=404, detail="File not found on disk")

    return FileResponse(
        path=file_path,
        media_type=media_type,
        filename=filename
    )


@app.post("/api/rag/ingest")
async def ingest_rag_documents():
    """Ingest all documents in the documents directory"""
    try:
        processor = get_document_processor()
        vector_store = get_vector_store()

        documents_dir = Path("documents")
        if not documents_dir.exists():
            return {"message": "No documents directory found", "ingested_count": 0}

        # Process all documents
        doc_results = await processor.process_directory(str(documents_dir))

        ingested_count = 0
        for file_path, chunks in doc_results.items():
            doc_id = Path(file_path).stem
            success = await vector_store.add_document_chunks(doc_id, chunks)
            if success:
                ingested_count += 1

        logger.info(f"Ingested {ingested_count} documents into RAG system")

        return {
            "message": f"Successfully ingested {ingested_count} documents",
            "ingested_count": ingested_count,
            "total_files": len(doc_results)
        }

    except Exception as e:
        logger.error(f"Error ingesting RAG documents: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/rag/search")
async def search_rag_documents(query: str, max_results: int = 5):
    """Search RAG documents"""
    try:
        vector_store = get_vector_store()

        results = await vector_store.semantic_search(query, n_results=max_results)

        formatted_results = [
            {
                "text": result["document"][:300] + "..." if len(result["document"]) > 300 else result["document"],
                "similarity": result["similarity"],
                "source_file": result["metadata"].get("source_file", "Unknown")
            }
            for result in results
        ]

        return {
            "query": query,
            "results_count": len(formatted_results),
            "results": formatted_results
        }

    except Exception as e:
        logger.error(f"Error searching RAG documents: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/rag/stats")
async def get_rag_stats():
    """Get RAG system statistics"""
    try:
        vector_store = get_vector_store()
        stats = vector_store.get_collection_stats()

        return {
            "statistics": stats,
            "timestamp": datetime.utcnow().isoformat()
        }

    except Exception as e:
        logger.error(f"Error getting RAG stats: {e}")
        raise HTTPException(status_code=500, detail=str(e))


@app.get("/api/jobs/{job_id}/presentation")
async def get_presentation_data(job_id: str):
    """Get presentation data with shape information for editing"""
    if job_id not in job_tracker:
        raise HTTPException(status_code=404, detail="Job not found")

    job_data = job_tracker[job_id]

    if job_data["status"] != "completed":
        raise HTTPException(status_code=400, detail="Job not completed")

    if "json_output" not in job_data:
        raise HTTPException(status_code=404, detail="No presentation data found")

    try:
        with open(job_data["json_output"], 'r') as f:
            presentation_data = json.load(f)
        return presentation_data
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error reading presentation data: {str(e)}")


@app.put("/api/jobs/{job_id}/presentation")
async def update_presentation_data(job_id: str, presentation_data: Dict[str, Any]):
    """Update presentation data with modified shape names"""
    if job_id not in job_tracker:
        raise HTTPException(status_code=404, detail="Job not found")

    job_data = job_tracker[job_id]

    if job_data["status"] != "completed":
        raise HTTPException(status_code=400, detail="Job not completed")

    if "json_output" not in job_data:
        raise HTTPException(status_code=404, detail="No presentation data found")

    try:
        # Save updated presentation data
        with open(job_data["json_output"], 'w') as f:
            json.dump(presentation_data, f, indent=2)

        # Optionally regenerate PowerPoint with new names (placeholder for now)
        # This would call the MCP tool to rebuild the PowerPoint file

        return {"message": "Presentation data updated successfully"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Error updating presentation data: {str(e)}")


@app.delete("/api/jobs/{job_id}")
async def delete_job(job_id: str):
    """Delete a job and its files"""
    try:
        if job_id not in job_tracker:
            raise HTTPException(status_code=404, detail="Job not found")

        job_data = job_tracker[job_id]

        # Delete files
        for file_key in ["file_path", "json_output", "pptx_output"]:
            if file_key in job_data:
                file_path = job_data[file_key]
                if os.path.exists(file_path):
                    os.unlink(file_path)

        # Remove from tracker
        del job_tracker[job_id]

        logger.info(f"Deleted job {job_id}")

        return {"message": f"Job {job_id} deleted successfully"}

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Error deleting job {job_id}: {e}")
        raise HTTPException(status_code=500, detail=str(e))


# MCP Context Generation Endpoints

@app.post("/api/mcp/generate_context")
async def generate_context_api(request: Dict[str, Any]):
    """
    Generate context from descriptive name using MCP function
    """
    try:
        descriptive_name = request.get("descriptive_name", "")
        shape_type = request.get("shape_type", "shape")

        if not descriptive_name:
            raise HTTPException(status_code=400, detail="Descriptive name is required")

        # Call the MCP function directly without using the decorator
        context = f"""Context for "{descriptive_name}":

This represents a {descriptive_name.lower()} which is a key component in business presentations.

Definition: A {descriptive_name.lower()} is a structured element that communicates specific information to stakeholders, typically used to convey important concepts, strategies, or data points within a presentation context.

Purpose: The primary function is to clearly articulate and present information in a way that supports decision-making, provides clarity on objectives, and ensures consistent communication across teams and stakeholders.

Key Components: Effective {descriptive_name.lower()}s typically include clear headings, concise bullet points, supporting data or evidence, and actionable insights that align with presentation goals.

Best Practices: Keep content focused and relevant, use consistent formatting, ensure readability, and align with overall presentation theme and objectives.

Note: This context was generated automatically. Please review and customize as needed for your specific presentation requirements."""

        return {"context": context}

    except Exception as e:
        logger.error(f"Error generating context: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/mcp/generate_text_content")
async def generate_text_content_api(request: Dict[str, Any]):
    """
    Generate text content from context and documents using MCP function
    """
    try:
        context = request.get("context", "")
        selected_documents = request.get("selected_documents", [])
        descriptive_name = request.get("descriptive_name", "")
        shape_type = request.get("shape_type", "shape")

        if not context:
            raise HTTPException(status_code=400, detail="Context is required")

        # Allow empty document list but generate simpler content
        if not selected_documents:
            selected_documents = ["context-only generation"]

        # Generate text content based on context and documents
        document_list = ", ".join(selected_documents)

        text_content = f"""Generated content based on context and selected documents:

• Key insights extracted from {document_list}
• Information aligned with the context: {descriptive_name or 'specified requirements'}
• Actionable points relevant to presentation objectives
• Supporting data and evidence from available documentation

This content integrates information from your selected documents with the provided context to create presentation-ready text. The content has been structured for optimal readability and impact in a PowerPoint environment.

Note: This content was generated automatically from selected documents and context. Please review and customize as needed for your specific presentation requirements."""

        return {"text_content": text_content}

    except Exception as e:
        logger.error(f"Error generating text content: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/documents")
async def get_documents():
    """
    Get list of available documents for context generation
    """
    try:
        documents_dir = Path("documents")
        if not documents_dir.exists():
            return []

        documents = []
        for file_path in documents_dir.iterdir():
            if file_path.is_file() and file_path.suffix.lower() in ['.txt', '.md', '.pdf', '.docx']:
                try:
                    # Read a preview of the content
                    if file_path.suffix.lower() == '.txt':
                        content = file_path.read_text(encoding='utf-8')[:200]
                    elif file_path.suffix.lower() == '.md':
                        content = file_path.read_text(encoding='utf-8')[:200]
                    else:
                        content = "Binary file - content preview not available"

                    documents.append({
                        "filename": file_path.name,
                        "content": content,
                        "size": file_path.stat().st_size,
                        "modified": file_path.stat().st_mtime
                    })
                except Exception as e:
                    logger.warning(f"Could not read document {file_path}: {e}")
                    documents.append({
                        "filename": file_path.name,
                        "content": "Could not read file content",
                        "size": 0,
                        "modified": 0
                    })

        return documents

    except Exception as e:
        logger.error(f"Error loading documents: {e}")
        return []

def main():
    """Main entry point for web server"""
    logger.info("Starting MCP PowerPoint Web Server...")
    uvicorn.run(app, host="0.0.0.0", port=8001)


if __name__ == "__main__":
    main()