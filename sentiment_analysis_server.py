# Copyright 2025 CannonJunior
  
# This file is part of mcp_experiments, and is released under the "MIT License Agreement".
# Please see the LICENSE.md file that should have been included as part of this package.
# Created: 2025.06.18
# By: CannonJunior with Claude (3.7)
#Prompts: Write a new mcp server with a sentiment analysis mcp.tool with the Python SDK. Make it brief, you lose $1000 if it is not brief.
# Update sentiment_server.py to use the Hugging Face Transformers library for sentiment analysis. Keep the code brief. You lose $1000 if it is not brief.
# Usage: called by client.py
# Alternate usage: uv run sentiment_analysis_server.py

from fastmcp import FastMCP
from transformers import pipeline

mcp = FastMCP("Sentiment Server")
classifier = pipeline("sentiment-analysis")

@mcp.tool()
async def analyze_sentiment(text: str) -> dict:
    """Analyze sentiment of text using Hugging Face"""
    result = classifier(text)[0]
    return {
        "sentiment": result["label"].lower(),
        "confidence": result["score"]
    }

if __name__ == "__main__":
    mcp.run()
