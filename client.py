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
                "command": "uvx",
                "args": ["--from", "git+https://github.com/CannonJunior/mcp-powerpoint.git", "mcp-powerpoint", "--server", "powerpoint"]
            },
            "shape_naming": {
                "command": "uvx",
                "args": ["--from", "git+https://github.com/CannonJunior/mcp-powerpoint.git", "mcp-powerpoint", "--server", "shape-naming"]
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
        json_result = await client.call_tool("powerpoint_pptx_to_json", {
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
            pptx_result = await client.call_tool("powerpoint_json_to_pptx", {
                "json_data": json_content,
                "output_path": "demo_recreated.pptx"
            })
            print(f"Recreate result: {pptx_result.content[0].text if hasattr(pptx_result.content[0], 'text') else pptx_result}")
            
            # Shape Naming Demo
            print("\n=== Shape Naming Demo ===")
            print("Generating descriptive names for all shapes...")
            naming_result = await client.call_tool("shape_naming_generate_descriptive_names_for_presentation", {
                "json_data": json_content
            })
            
            if hasattr(naming_result, 'content') and naming_result.content:
                naming_content = naming_result.content[0].text if hasattr(naming_result.content[0], 'text') else str(naming_result.content[0])
                
                # Save the updated JSON with descriptive names
                with open("presentation_with_descriptive_names.json", 'w', encoding='utf-8') as f:
                    f.write(naming_content)
                
                # Show some examples
                try:
                    naming_data = json.loads(naming_content)
                    print(f"Updated presentation with descriptive names saved to 'presentation_with_descriptive_names.json'")
                    
                    # Show first few shape renames as examples
                    example_count = 0
                    for slide_idx, slide in enumerate(naming_data.get('slides', [])):
                        for shape in slide.get('shapes', []):
                            if 'descriptive_name' in shape and 'original_name' in shape and example_count < 5:
                                text_preview = ""
                                if shape.get('text_frame') and shape['text_frame'].get('text'):
                                    text_preview = shape['text_frame']['text'][:40] + "..."
                                print(f"  '{shape['original_name']}' -> '{shape['descriptive_name']}' ({text_preview})")
                                example_count += 1
                        if example_count >= 5:
                            break
                    
                    if example_count == 0:
                        print("  No shapes were renamed (this might indicate an issue)")
                except Exception as e:
                    print(f"  Error processing naming results: {e}")
        
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

        print(f"\n=== Advanced Modes Available ===")
        print("For advanced operations, use client_modes.py:")
        print("  python client_modes.py extract -i input.pptx")
        print("  python client_modes.py refine -i recreated.pptx")
        print("  python client_modes.py rename -i presentation.json --content-dir ./docs")
        print("  python client_modes.py populate -i template.pptx --naming-json names.json --content-dir ./content")

if __name__ == "__main__":
    asyncio.run(main())
