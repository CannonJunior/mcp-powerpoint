#!/bin/bash

# MCP PowerPoint Web Application Startup Script
# This script launches the FastAPI web server for PowerPoint processing

set -e

echo "🚀 Starting MCP PowerPoint Web Application..."

# Check if we're in the correct directory
if [ ! -f "pyproject.toml" ]; then
    echo "❌ Error: pyproject.toml not found. Please run this script from the project root directory."
    exit 1
fi

# Create required directories
echo "📁 Creating required directories..."
mkdir -p uploads
mkdir -p outputs
mkdir -p documents
mkdir -p data

# Check if uv is available
if ! command -v uv &> /dev/null; then
    echo "❌ Error: uv is not installed. Please install uv first:"
    echo "   curl -LsSf https://astral.sh/uv/install.sh | sh"
    exit 1
fi

# Install dependencies if needed
if [ ! -d ".venv" ]; then
    echo "📦 Installing dependencies..."
    uv sync
fi

# Check for existing processes on port 8001 and kill them
echo "🔍 Checking for existing processes on port 8001..."
EXISTING_PIDS=$(lsof -ti:8001 2>/dev/null || true)
if [ -n "$EXISTING_PIDS" ]; then
    echo "⚠️  Found existing processes on port 8001: $EXISTING_PIDS"
    echo "🛑 Stopping existing processes..."
    kill $EXISTING_PIDS 2>/dev/null || true
    sleep 2
    echo "✅ Existing processes stopped"
else
    echo "✅ Port 8001 is available"
fi

# Start the web server
echo "🌐 Starting FastAPI web server on http://localhost:8001"
echo "   Use Ctrl+C to stop the server"
echo ""

# Run the web server using uv
uv run python -m mcp_powerpoint.web_server

echo "👋 MCP PowerPoint Web Application stopped."