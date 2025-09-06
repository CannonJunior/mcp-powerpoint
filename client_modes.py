#!/usr/bin/env python3
"""
PowerPoint MCP Client with Multiple Operation Modes

This client provides different modes for working with PowerPoint presentations:
1. EXTRACT - Extract and recreate PowerPoint presentations 
2. REFINE - Improve recreated PowerPoint presentations
3. RENAME - Enhance shape naming with additional content analysis
4. POPULATE - Generate new presentations from templates using document content
"""

import argparse
import asyncio
import json
import os
import ollama
from pathlib import Path
from typing import List, Dict, Any

from fastmcp import Client

class PowerPointClient:
    """Multi-mode PowerPoint processing client"""
    
    def __init__(self):
        """Initialize the client with MCP server configuration"""
        self.config = {
            "mcpServers": {
                "powerpoint": {
                    "command": "python",
                    "args": ["powerpoint_server.py"]
                },
                "shape_naming": {
                    "command": "python", 
                    "args": ["shape_naming_server.py"]
                }
            }
        }
        # Set up Ollama environment
        os.environ["OPENAI_API_KEY"] = "NA"
        self.model_name = 'incept5/llama3.1-claude:latest'
    
    async def mode_extract(self, input_file: str, output_json: str = None, output_pptx: str = None):
        """
        EXTRACT Mode: Convert PowerPoint to JSON and recreate as new PowerPoint
        
        Args:
            input_file: Input PowerPoint file path
            output_json: Optional output JSON file path (default: auto-generated)
            output_pptx: Optional output PowerPoint file path (default: auto-generated)
        """
        if not output_json:
            output_json = f"{Path(input_file).stem}_extracted.json"
        if not output_pptx:
            output_pptx = f"{Path(input_file).stem}_recreated.pptx"
            
        print(f"🔄 EXTRACT Mode: Processing '{input_file}'")
        
        async with Client(self.config) as client:
            # Extract PowerPoint to JSON
            print("📄 Converting PowerPoint to JSON...")
            json_result = await client.call_tool("powerpoint_pptx_to_json", {
                "file_path": input_file
            })
            
            if hasattr(json_result, 'content') and json_result.content:
                json_content = json_result.content[0].text if hasattr(json_result.content[0], 'text') else str(json_result.content[0])
                
                # Save JSON
                with open(output_json, 'w', encoding='utf-8') as f:
                    f.write(json_content)
                
                presentation_data = json.loads(json_content)
                print(f"✅ Extracted {len(presentation_data.get('slides', []))} slides to '{output_json}'")
                
                # Recreate PowerPoint
                print("🎨 Recreating PowerPoint from JSON...")
                pptx_result = await client.call_tool("powerpoint_json_to_pptx", {
                    "json_data": json_content,
                    "output_path": output_pptx
                })
                
                print(f"✅ Recreated PowerPoint saved as '{output_pptx}'")
                return output_json, output_pptx
            else:
                print("❌ Failed to extract PowerPoint content")
                return None, None
    
    async def mode_refine(self, input_pptx: str, output_pptx: str = None):
        """
        REFINE Mode: Improve an existing recreated PowerPoint presentation
        
        Args:
            input_pptx: Input PowerPoint file to refine
            output_pptx: Optional output file path (default: auto-generated)
        """
        if not output_pptx:
            output_pptx = f"{Path(input_pptx).stem}_refined.pptx"
            
        print(f"✨ REFINE Mode: Improving '{input_pptx}'")
        
        async with Client(self.config) as client:
            # First convert to JSON for analysis
            print("📊 Analyzing current presentation...")
            json_result = await client.call_tool("powerpoint_pptx_to_json", {
                "file_path": input_pptx
            })
            
            if hasattr(json_result, 'content') and json_result.content:
                json_content = json_result.content[0].text if hasattr(json_result.content[0], 'text') else str(json_result.content[0])
                presentation_data = json.loads(json_content)
                
                # Apply improvements (this is where you'd add refinement logic)
                print("🔧 Applying improvements...")
                
                # Example improvements:
                # - Fix font consistency
                # - Standardize colors
                # - Improve spacing
                improvements_applied = []
                
                # Font standardization
                for slide in presentation_data.get('slides', []):
                    for shape in slide.get('shapes', []):
                        if shape.get('text_frame'):
                            for paragraph in shape['text_frame'].get('paragraphs', []):
                                for run in paragraph.get('runs', []):
                                    if run.get('font'):
                                        # Standardize fonts
                                        if run['font'].get('name') in ['Arial', 'Helvetica']:
                                            run['font']['name'] = 'Arial'
                                        improvements_applied.append("Font standardization")
                
                print(f"🎯 Applied {len(set(improvements_applied))} improvement types")
                
                # Recreate improved presentation
                improved_json = json.dumps(presentation_data, indent=2, ensure_ascii=False)
                pptx_result = await client.call_tool("powerpoint_json_to_pptx", {
                    "json_data": improved_json,
                    "output_path": output_pptx
                })
                
                print(f"✅ Refined presentation saved as '{output_pptx}'")
                return output_pptx
            else:
                print("❌ Failed to analyze presentation for refinement")
                return None
    
    async def mode_rename(self, input_json: str, content_dir: str = None, output_json: str = None):
        """
        RENAME Mode: Enhance shape naming using additional content analysis
        
        Args:
            input_json: Input JSON file with basic shape names
            content_dir: Optional directory with additional content for context
            output_json: Optional output JSON file path (default: auto-generated)
        """
        if not output_json:
            output_json = f"{Path(input_json).stem}_enhanced_names.json"
            
        print(f"🏷️  RENAME Mode: Enhancing names in '{input_json}'")
        
        # Load additional context from content directory
        context_info = {}
        if content_dir and os.path.exists(content_dir):
            print(f"📚 Analyzing content in '{content_dir}' for context...")
            for file_path in Path(content_dir).rglob("*.txt"):
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        context_info[file_path.name] = f.read()[:500]  # First 500 chars
                except:
                    pass
            print(f"📖 Found {len(context_info)} context files")
        
        async with Client(self.config) as client:
            # Load the presentation JSON
            with open(input_json, 'r', encoding='utf-8') as f:
                json_content = f.read()
            
            print("🧠 Generating enhanced shape names...")
            
            # If we have no additional context, just use the standard naming
            if not context_info:
                naming_result = await client.call_tool("shape_naming_generate_descriptive_names_for_presentation", {
                    "json_data": json_content
                })
            else:
                # Use context-aware naming (implement custom logic here)
                presentation_data = json.loads(json_content)
                
                # Enhance naming with context
                for slide_idx, slide in enumerate(presentation_data.get('slides', [])):
                    for shape in slide.get('shapes', []):
                        if shape.get('text_frame') and shape['text_frame'].get('text'):
                            text_content = shape['text_frame']['text']
                            
                            # Try to get better suggestions with context
                            suggestions_result = await client.call_tool("shape_naming_get_shape_suggestions", {
                                "shape_text": text_content,
                                "context": f"Business document with context from {len(context_info)} files"
                            })
                            
                            suggestions_content = suggestions_result.content[0].text if hasattr(suggestions_result.content[0], 'text') else str(suggestions_result.content[0])
                            
                            try:
                                suggestions = json.loads(suggestions_content)
                                if suggestions.get('suggestions') and len(suggestions['suggestions']) > 0:
                                    best_suggestion = suggestions['suggestions'][0]
                                    if isinstance(best_suggestion, dict):
                                        shape['descriptive_name'] = best_suggestion.get('name', 'enhanced_shape')
                                        shape['original_name'] = shape.get('name', 'unknown')
                                        shape['naming_rationale'] = best_suggestion.get('rationale', 'Context-enhanced naming')
                            except:
                                pass
                
                naming_content = json.dumps(presentation_data, indent=2, ensure_ascii=False)
            
            if not context_info:
                naming_content = naming_result.content[0].text if hasattr(naming_result.content[0], 'text') else str(naming_result.content[0])
            
            # Save enhanced names
            with open(output_json, 'w', encoding='utf-8') as f:
                f.write(naming_content)
            
            # Show examples
            enhanced_data = json.loads(naming_content)
            example_count = 0
            print("🎯 Enhanced naming examples:")
            for slide in enhanced_data.get('slides', []):
                for shape in slide.get('shapes', []):
                    if 'descriptive_name' in shape and 'original_name' in shape and example_count < 5:
                        print(f"  '{shape['original_name']}' → '{shape['descriptive_name']}'")
                        example_count += 1
                if example_count >= 5:
                    break
            
            print(f"✅ Enhanced names saved to '{output_json}'")
            return output_json
    
    async def mode_populate(self, template_pptx: str, naming_json: str, content_dir: str, output_json: str = None, output_pptx: str = None):
        """
        POPULATE Mode: Generate new presentation using template and document content
        
        Args:
            template_pptx: Template PowerPoint file
            naming_json: JSON file with descriptive shape names
            content_dir: Directory containing documents to extract content from
            output_json: Optional output JSON file path (default: auto-generated)
            output_pptx: Optional output PowerPoint file path (default: auto-generated)
        """
        if not output_json:
            output_json = f"{Path(template_pptx).stem}_populated.json"
        if not output_pptx:
            output_pptx = f"{Path(template_pptx).stem}_populated.pptx"
            
        print(f"🚀 POPULATE Mode: Creating presentation from template '{template_pptx}' with content from '{content_dir}'")
        
        # Load naming template
        print("📋 Loading naming template...")
        with open(naming_json, 'r', encoding='utf-8') as f:
            template_data = json.load(f)
        
        # Analyze document content
        print(f"📚 Analyzing documents in '{content_dir}'...")
        document_content = {}
        content_types = {
            '.txt': 'text',
            '.md': 'markdown', 
            '.json': 'json',
            '.csv': 'csv'
        }
        
        for file_path in Path(content_dir).rglob("*"):
            if file_path.suffix.lower() in content_types:
                try:
                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()
                        document_content[file_path.name] = {
                            'content': content,
                            'type': content_types[file_path.suffix.lower()],
                            'path': str(file_path)
                        }
                except Exception as e:
                    print(f"⚠️  Could not read {file_path}: {e}")
        
        print(f"📖 Loaded {len(document_content)} documents")
        
        async with Client(self.config) as client:
            # Create content mapping using Ollama
            print("🧠 Mapping document content to shape names...")
            
            populated_data = template_data.copy()
            
            # Process each shape to populate with relevant content
            for slide_idx, slide in enumerate(populated_data.get('slides', [])):
                for shape in slide.get('shapes', []):
                    descriptive_name = shape.get('descriptive_name', '')
                    original_text = ""
                    
                    if shape.get('text_frame') and shape['text_frame'].get('text'):
                        original_text = shape['text_frame']['text']
                    
                    if descriptive_name and document_content:
                        # Use Ollama to find the best content match
                        content_summary = "\n".join([
                            f"File: {name} - {info['content'][:100]}..." 
                            for name, info in list(document_content.items())[:5]
                        ])
                        
                        prompt = f"""Based on the shape name "{descriptive_name}" and current text "{original_text}", 
                        select the most relevant content from these documents to populate this PowerPoint shape:

{content_summary}

Provide a brief, appropriate text that would fit in a PowerPoint slide for the shape named "{descriptive_name}".
Keep it concise (1-3 sentences max). If no content fits, return "KEEP_ORIGINAL".

Response:"""

                        try:
                            response = ollama.generate(
                                model=self.model_name,
                                prompt=prompt,
                                options={
                                    'temperature': 0.3,
                                    'num_predict': 100
                                }
                            )
                            
                            suggested_content = response['response'].strip()
                            
                            if suggested_content and suggested_content != "KEEP_ORIGINAL":
                                # Update the shape's text content
                                if shape.get('text_frame'):
                                    shape['text_frame']['text'] = suggested_content
                                    
                                    # Update the first paragraph and run
                                    if shape['text_frame'].get('paragraphs'):
                                        shape['text_frame']['paragraphs'][0]['text'] = suggested_content
                                        if shape['text_frame']['paragraphs'][0].get('runs'):
                                            shape['text_frame']['paragraphs'][0]['runs'][0]['text'] = suggested_content
                                
                                print(f"  📝 Updated '{descriptive_name}': {suggested_content[:50]}...")
                        
                        except Exception as e:
                            print(f"⚠️  Could not generate content for '{descriptive_name}': {e}")
            
            # Save populated JSON
            populated_json = json.dumps(populated_data, indent=2, ensure_ascii=False)
            with open(output_json, 'w', encoding='utf-8') as f:
                f.write(populated_json)
            
            print(f"✅ Populated content saved to '{output_json}'")
            
            # Create final PowerPoint
            print("🎨 Creating final PowerPoint presentation...")
            pptx_result = await client.call_tool("powerpoint_json_to_pptx", {
                "json_data": populated_json,
                "output_path": output_pptx
            })
            
            print(f"✅ Final presentation saved as '{output_pptx}'")
            return output_json, output_pptx

async def main():
    """Main function with argument parsing and mode selection"""
    parser = argparse.ArgumentParser(description="PowerPoint MCP Client with Multiple Operation Modes")
    parser.add_argument("mode", choices=["extract", "refine", "rename", "populate"], 
                       help="Operation mode to run")
    
    # Common arguments
    parser.add_argument("--input", "-i", required=True, help="Input file path")
    parser.add_argument("--output-json", help="Output JSON file path")
    parser.add_argument("--output-pptx", help="Output PowerPoint file path")
    
    # Mode-specific arguments
    parser.add_argument("--template", help="Template file for POPULATE mode")
    parser.add_argument("--naming-json", help="Naming JSON file for POPULATE mode")
    parser.add_argument("--content-dir", help="Content directory for RENAME/POPULATE modes")
    
    args = parser.parse_args()
    
    client = PowerPointClient()
    
    try:
        if args.mode == "extract":
            await client.mode_extract(args.input, args.output_json, args.output_pptx)
            
        elif args.mode == "refine":
            await client.mode_refine(args.input, args.output_pptx)
            
        elif args.mode == "rename":
            await client.mode_rename(args.input, args.content_dir, args.output_json)
            
        elif args.mode == "populate":
            if not args.naming_json or not args.content_dir:
                print("❌ POPULATE mode requires --naming-json and --content-dir arguments")
                return
            await client.mode_populate(args.input, args.naming_json, args.content_dir, 
                                     args.output_json, args.output_pptx)
        
        print("🎉 Operation completed successfully!")
        
    except Exception as e:
        print(f"❌ Error during {args.mode} operation: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    asyncio.run(main())