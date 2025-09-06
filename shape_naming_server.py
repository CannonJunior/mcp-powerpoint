#!/usr/bin/env python3
"""
Shape Naming MCP Server

This server provides tools to generate descriptive names for PowerPoint shapes
based on their text content using Ollama LLM integration.
"""

import json
import ollama
from typing import List, Dict, Any, Optional
from fastmcp import FastMCP

# Initialize MCP server
mcp = FastMCP("Shape Naming Server")

def _analyze_shape_content_internal(shape_name: str, shape_text: str, shape_type: str) -> str:
    """
    Analyze a shape's content and generate a descriptive name based on its text.
    
    Args:
        shape_name: Current generic name of the shape (e.g., "Rectangle 5")
        shape_text: Text content found within the shape
        shape_type: Type of shape (e.g., "AUTO_SHAPE", "TEXT_BOX")
        
    Returns:
        A descriptive name for the shape based on its content
    """
    try:
        # Handle empty or whitespace-only text
        if not shape_text or not shape_text.strip():
            return f"empty_{shape_type.lower()}"
        
        # Clean the text for analysis
        clean_text = shape_text.strip()[:200]  # Limit to first 200 chars for analysis
        
        # Create prompt for Ollama
        prompt = f"""Based on the following text content found in a PowerPoint shape, generate a concise, descriptive name (2-4 words) that captures the main purpose or content of this shape. The name should be:
- Descriptive and specific
- Suitable for programmatic use (lowercase, underscores instead of spaces)
- Related to the actual content/purpose of the text

Text content: "{clean_text}"

Examples of good names:
- "company_header" for "General Atomics - Integrated Intelligence, Inc."
- "project_title" for "250083 – MDA Advanced Capability Concepts"
- "financial_summary" for text about budgets/costs
- "opportunity_description" for business opportunity text
- "contact_info" for contact details
- "deadline_info" for due dates
- "approval_status" for approval information

Generate only the name (lowercase with underscores), nothing else:"""

        # Query Ollama
        response = ollama.generate(
            model='incept5/llama3.1-claude:latest',  # Using the available model
            prompt=prompt,
            options={
                'temperature': 0.3,  # Lower temperature for more consistent naming
                'num_predict': 20,   # Limit response length
                'stop': ['\n', '.', ' ', '-']  # Stop at first word boundary
            }
        )
        
        # Extract and clean the response
        suggested_name = response['response'].strip().lower()
        
        # Clean up the name to ensure it's suitable for programmatic use
        suggested_name = suggested_name.replace(' ', '_').replace('-', '_')
        suggested_name = ''.join(c for c in suggested_name if c.isalnum() or c == '_')
        
        # Ensure name is not empty and has reasonable length
        if not suggested_name or len(suggested_name) < 2:
            # Fallback to content-based naming
            words = clean_text.split()[:3]
            suggested_name = '_'.join(word.lower().strip('.,!?:;') for word in words if word.isalpha())
            if not suggested_name:
                suggested_name = f"text_{shape_type.lower()}"
        
        # Limit name length
        if len(suggested_name) > 30:
            suggested_name = suggested_name[:30]
            
        return suggested_name
        
    except Exception as e:
        # Fallback naming if Ollama fails
        print(f"Warning: Ollama analysis failed: {e}")
        words = shape_text.split()[:2] if shape_text else []
        fallback_name = '_'.join(word.lower().strip('.,!?:;') for word in words if word.isalpha())
        return fallback_name if fallback_name else f"shape_{shape_type.lower()}"

@mcp.tool()
def analyze_shape_content(shape_name: str, shape_text: str, shape_type: str) -> str:
    """
    Analyze a shape's content and generate a descriptive name based on its text.
    
    Args:
        shape_name: Current generic name of the shape (e.g., "Rectangle 5")
        shape_text: Text content found within the shape
        shape_type: Type of shape (e.g., "AUTO_SHAPE", "TEXT_BOX")
        
    Returns:
        A descriptive name for the shape based on its content
    """
    return _analyze_shape_content_internal(shape_name, shape_text, shape_type)

@mcp.tool()
def generate_descriptive_names_for_presentation(json_data: str) -> str:
    """
    Process an entire presentation JSON and generate descriptive names for all shapes.
    
    Args:
        json_data: JSON string containing the presentation data
        
    Returns:
        Updated JSON string with descriptive shape names
    """
    try:
        # Parse the JSON data
        presentation_data = json.loads(json_data)
        
        # Track name usage to avoid duplicates
        used_names = set()
        name_counters = {}
        
        # Process each slide
        for slide_idx, slide in enumerate(presentation_data.get('slides', [])):
            print(f"Processing slide {slide_idx + 1}...")
            
            # Process each shape in the slide
            for shape_idx, shape in enumerate(slide.get('shapes', [])):
                original_name = shape.get('name', f'shape_{shape_idx}')
                shape_type = shape.get('shape_type', 'UNKNOWN')
                
                # Extract text content from the shape
                text_content = ""
                if shape.get('text_frame') and shape['text_frame'].get('text'):
                    text_content = shape['text_frame']['text']
                
                # Generate descriptive name (call the actual function, not the MCP tool wrapper)
                descriptive_name = _analyze_shape_content_internal(original_name, text_content, shape_type)
                
                # Handle duplicate names
                if descriptive_name in used_names:
                    if descriptive_name not in name_counters:
                        name_counters[descriptive_name] = 2
                    else:
                        name_counters[descriptive_name] += 1
                    descriptive_name = f"{descriptive_name}_{name_counters[descriptive_name]}"
                
                used_names.add(descriptive_name)
                
                # Update the shape name
                shape['descriptive_name'] = descriptive_name
                shape['original_name'] = original_name
                
                print(f"  Shape: '{original_name}' -> '{descriptive_name}' (text: '{text_content[:50]}...')")
        
        # Return updated JSON
        return json.dumps(presentation_data, indent=2, ensure_ascii=False)
        
    except Exception as e:
        return json.dumps({
            "error": f"Failed to process presentation: {str(e)}"
        })

@mcp.tool()
def get_shape_suggestions(shape_text: str, context: str = "") -> str:
    """
    Get multiple naming suggestions for a shape based on its text content.
    
    Args:
        shape_text: Text content of the shape
        context: Additional context about the shape's purpose or location
        
    Returns:
        JSON string with multiple naming suggestions and their rationales
    """
    try:
        if not shape_text or not shape_text.strip():
            return json.dumps({
                "suggestions": ["empty_text", "placeholder", "blank_shape"],
                "rationale": "No text content available"
            })
        
        clean_text = shape_text.strip()[:200]
        
        prompt = f"""Analyze this PowerPoint shape text and suggest 3 different descriptive names. Consider the context and purpose.

Text: "{clean_text}"
Context: {context if context else "General PowerPoint shape"}

For each suggestion, provide:
1. A descriptive name (lowercase with underscores)
2. Brief rationale for why this name fits

Format your response as JSON:
{{
  "suggestions": [
    {{"name": "suggestion1", "rationale": "why this name fits"}},
    {{"name": "suggestion2", "rationale": "alternative perspective"}},
    {{"name": "suggestion3", "rationale": "another valid option"}}
  ]
}}

Focus on the content's purpose, not just keywords. Examples:
- Financial data -> "budget_summary", "cost_breakdown", "financial_overview"
- Title text -> "main_title", "section_header", "presentation_title"
- Contact info -> "contact_details", "presenter_info", "company_info"

Generate only valid JSON:"""

        response = ollama.generate(
            model='incept5/llama3.1-claude:latest',
            prompt=prompt,
            options={
                'temperature': 0.5,
                'num_predict': 200
            }
        )
        
        # Try to parse the response as JSON
        try:
            result = json.loads(response['response'])
            return json.dumps(result, indent=2)
        except:
            # Fallback if JSON parsing fails
            return json.dumps({
                "suggestions": [
                    {"name": "content_text", "rationale": "Generic content-based name"},
                    {"name": "shape_text", "rationale": "Simple descriptive name"},
                    {"name": "text_element", "rationale": "Element-based naming"}
                ],
                "note": "Fallback suggestions due to parsing error"
            })
            
    except Exception as e:
        return json.dumps({
            "error": f"Failed to generate suggestions: {str(e)}",
            "suggestions": [
                {"name": "text_shape", "rationale": "Fallback name"},
                {"name": "content_element", "rationale": "Generic fallback"},
                {"name": "shape_content", "rationale": "Safe fallback option"}
            ]
        })

@mcp.tool()
def batch_rename_shapes(json_data: str, naming_rules: str = "") -> str:
    """
    Batch rename shapes in a presentation based on custom rules or automatic analysis.
    
    Args:
        json_data: JSON string containing the presentation data
        naming_rules: Optional JSON string with custom naming rules
        
    Returns:
        JSON string with rename operations and results
    """
    try:
        presentation_data = json.loads(json_data)
        
        # Parse naming rules if provided
        rules = {}
        if naming_rules:
            try:
                rules = json.loads(naming_rules)
            except:
                pass
        
        results = {
            "renamed_shapes": [],
            "total_shapes": 0,
            "successful_renames": 0,
            "failed_renames": 0
        }
        
        # Process all shapes
        for slide_idx, slide in enumerate(presentation_data.get('slides', [])):
            for shape_idx, shape in enumerate(slide.get('shapes', [])):
                results["total_shapes"] += 1
                
                original_name = shape.get('name', f'shape_{shape_idx}')
                text_content = ""
                
                if shape.get('text_frame') and shape['text_frame'].get('text'):
                    text_content = shape['text_frame']['text']
                
                try:
                    # Check if there's a specific rule for this shape
                    new_name = None
                    for pattern, replacement in rules.items():
                        if pattern in original_name or (text_content and pattern.lower() in text_content.lower()):
                            new_name = replacement
                            break
                    
                    # If no rule matched, use automatic naming
                    if not new_name:
                        new_name = _analyze_shape_content_internal(
                            original_name, 
                            text_content, 
                            shape.get('shape_type', 'UNKNOWN')
                        )
                    
                    # Apply the rename
                    shape['descriptive_name'] = new_name
                    shape['original_name'] = original_name
                    
                    results["renamed_shapes"].append({
                        "slide": slide_idx + 1,
                        "shape_index": shape_idx,
                        "original_name": original_name,
                        "new_name": new_name,
                        "text_preview": text_content[:50] if text_content else "No text",
                        "method": "rule_based" if new_name != _analyze_shape_content_internal(original_name, text_content, shape.get('shape_type', 'UNKNOWN')) else "automatic"
                    })
                    
                    results["successful_renames"] += 1
                    
                except Exception as e:
                    results["failed_renames"] += 1
                    print(f"Failed to rename shape {original_name}: {e}")
        
        # Add the updated presentation data to results
        results["updated_presentation"] = presentation_data
        
        return json.dumps(results, indent=2, ensure_ascii=False)
        
    except Exception as e:
        return json.dumps({
            "error": f"Batch rename failed: {str(e)}"
        })

if __name__ == "__main__":
    mcp.run()