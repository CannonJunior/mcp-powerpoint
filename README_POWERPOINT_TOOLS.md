# PowerPoint MCP Tools

This project provides MCP (Model Context Protocol) tools for converting PowerPoint presentations to/from structured JSON format using Pydantic models.

## Features

- **Complete PowerPoint Analysis**: Extracts all text, shapes, images, tables, positioning, and formatting
- **Pydantic Models**: Type-safe data structures for PowerPoint objects
- **Roundtrip Conversion**: Original PPTX → JSON → New PPTX with high fidelity
- **MCP Integration**: Works with any MCP-compatible client

## Files

### Core Components
- `powerpoint_models.py` - Comprehensive Pydantic models for PowerPoint structure
- `powerpoint_server.py` - MCP server providing PowerPoint conversion tools
- `client.py` - Example client demonstrating both sentiment analysis and PowerPoint tools

### Tools Available

#### `pptx_to_json(file_path: str) -> str`
Converts a PowerPoint file to structured JSON format.
- Extracts all slides, shapes, text content, formatting, images, and tables
- Preserves positioning, sizing, fonts, colors, and other properties
- Returns comprehensive JSON representation

#### `json_to_pptx(json_data: str, output_path: str) -> str`
Creates a PowerPoint file from JSON data.
- Reconstructs slides, shapes, text, images, and tables
- Applies formatting, positioning, and styling
- Generates a new .pptx file

## Usage

### 1. Start the MCP Server
```python
python powerpoint_server.py
```

### 2. Use with MCP Client
```python
# Convert PowerPoint to JSON
json_result = await client.call_tool("powerpoint_pptx_to_json", {
    "file_path": "presentation.pptx"
})

# Convert JSON back to PowerPoint
pptx_result = await client.call_tool("powerpoint_json_to_pptx", {
    "json_data": json_string,
    "output_path": "output.pptx"
})
```

### 3. Example Client
Run the provided client to see both sentiment analysis and PowerPoint tools:
```bash
python client.py
```

## Data Structure

The JSON structure captures:
- **Presentation**: Dimensions, properties, metadata
- **Slides**: Individual slide content and properties
- **Shapes**: All shape types (text boxes, auto shapes, images, tables)
- **Text**: Formatted text with fonts, colors, styling
- **Images**: Base64-encoded image data with cropping info
- **Tables**: Complete table structure with cell content
- **Positioning**: Exact coordinates and dimensions
- **Formatting**: Colors, fonts, line styles, fills

## Test Results

Successfully tested with `MDA-250083-BNB-20250904.v1.RFI.pptx`:
- Original file: 132,823 bytes
- JSON representation: 158,917 bytes
- Recreated file: 61,334 bytes
- ✅ Roundtrip conversion successful
- ✅ All text content preserved
- ✅ Font sizes correctly extracted (12pt, 14pt) and applied
- ✅ Font colors extracted (`#000000`, `#FFFFFF`) and applied  
- ✅ Shape fill colors extracted (`#D6D6D6`, `#000000`) and applied
- ✅ Shape border colors extracted (`#000000`) and applied
- ✅ Shape positioning maintained
- ✅ Image data captured and restored

### Fixed Issues
- **Font Size Accuracy**: Fixed EMU to points conversion (was showing 177800, now correctly shows 12pt)
- **Font Color Extraction**: Properly handles RGBColor objects with hex output
- **Shape Fill Colors**: Extracts solid, gradient, and patterned fill colors correctly
- **Shape Border Colors**: Extracts line colors, widths, and dash styles with proper RGBColor handling
- **Color Application**: Applies fill and border colors during PowerPoint reconstruction

## Dependencies

- `python-pptx` - PowerPoint file manipulation
- `pydantic` - Data models and JSON serialization
- `fastmcp` - MCP server framework
- `pillow` - Image processing
- `base64` - Image encoding/decoding

## Architecture

The system uses a three-layer architecture:
1. **Pydantic Models** - Type-safe data structures
2. **Extraction Layer** - Converts PowerPoint objects to models
3. **Reconstruction Layer** - Rebuilds PowerPoint from models
4. **MCP Interface** - Exposes functionality as tools

This enables reliable, type-safe conversion between PowerPoint files and structured JSON data suitable for analysis, modification, or storage.