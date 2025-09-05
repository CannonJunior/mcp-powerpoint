# Copyright 2025 CannonJunior

# This file is part of mcp_experiments, and is released under the "MIT License Agreement".
# Please see the LICENSE.md file that should have been included as part of this package.
# Created: 2025.09.05
# By: CannonJunior with Claude (3.7 free version)
# Prompt (amongst others):
# Usage: uv run client.py

import asyncio
import json
import os
import ollama

from fastmcp import Client, FastMCP
from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

async def main():
    config = {
        "mcpServers": {
            "powerpoint": {
                "command": "python",
                "args": ["powerpoint_server.py"]
            }
        }
    }

    os.environ["OPENAI_API_KEY"] = "NA"
    model_name = 'incept5/llama3.1-claude:latest'
 
    async with Client(config) as client:
        # List tools from all servers
        tools = await client.list_tools()
        print(f"Tools: {[t.name for t in tools]}")

        # List resources from all servers
        resources = await client.list_resources()
        print(f"Resources: {[r.name for r in resources]}")

        # List resource templates from all servers
        templates = await client.list_resource_templates()
        print(f"Templates: {[t.name for t in templates]}")
        
        # List prompts
        prompts = await client.list_prompts()
        print(f"Prompts: {[p.name for p in prompts]}")
        
        # Test PowerPoint tools
        print("\n=== PowerPoint Tools Demo ===")
        pptx_file = "MDA-250083-BNB-20250904.v1.RFI.pptx"
        
        # Convert PowerPoint to JSON
        print("Converting PowerPoint to JSON...")
        json_result = await client.call_tool("pptx_to_json", {
            "file_path": pptx_file
        })
        
        # Extract JSON content
        if hasattr(json_result, 'content') and json_result.content:
            json_content = json_result.content[0].text if hasattr(json_result.content[0], 'text') else str(json_result.content[0])
            
            # Save JSON to file
            with open("demo_presentation.json", 'w', encoding='utf-8') as f:
                f.write(json_content)
            
            # Parse and display info
            presentation_data = json.loads(json_content)
            print(f"Successfully extracted {len(presentation_data.get('slides', []))} slides")
            print(f"Slide dimensions: {presentation_data.get('slide_width')} x {presentation_data.get('slide_height')}")
            
            # Convert JSON back to PowerPoint
            print("Converting JSON back to PowerPoint...")
            pptx_result = await client.call_tool("json_to_pptx", {
                "json_data": json_content,
                "output_path": "demo_recreated.pptx"
            })
            print(f"Recreate result: {pptx_result.content[0].text if hasattr(pptx_result.content[0], 'text') else pptx_result}")
        
        # AI delegation (existing code)
        print(f"\n=== AI Delegation ===")
        print(f"resources {resources}")
        resources_list = [r.uri for r in resources]
        print(resources_list)
        if resources_list:
            query = "Get error logs for troubleshooting"
            delegation = ollama.generate(
            model=model_name,
                prompt=f"Pick one from {resources_list} for: {query}. Return only the URI."
            )
            chosen = delegation['response'].strip()
            if chosen in resources_list:
                result = await client.read_resource(chosen)
                print(f"\nAI chose {chosen}:\n{result.contents[0].text[:200]}...")

if __name__ == "__main__":
    asyncio.run(main())
