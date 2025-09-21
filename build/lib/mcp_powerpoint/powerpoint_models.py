#!/usr/bin/env python3
"""
Pydantic models for PowerPoint structure serialization
"""

from pydantic import BaseModel, Field
from typing import List, Optional, Dict, Any, Union
from enum import Enum
import base64

class ShapeType(str, Enum):
    """PowerPoint shape types"""
    AUTO_SHAPE = "AUTO_SHAPE"
    TEXT_BOX = "TEXT_BOX" 
    PICTURE = "PICTURE"
    TABLE = "TABLE"
    PLACEHOLDER = "PLACEHOLDER"
    GROUP = "GROUP"
    CHART = "CHART"
    FREEFORM = "FREEFORM"
    OTHER = "OTHER"

class FontInfo(BaseModel):
    """Font information for text"""
    name: Optional[str] = None
    size: Optional[int] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    underline: Optional[bool] = None
    color_rgb: Optional[str] = None  # Hex color like "#FF0000"

class ParagraphFormat(BaseModel):
    """Paragraph formatting information"""
    alignment: Optional[str] = None  # LEFT, CENTER, RIGHT, JUSTIFY
    space_before: Optional[int] = None
    space_after: Optional[int] = None
    line_spacing: Optional[float] = None
    bullet_font: Optional[FontInfo] = None
    bullet_char: Optional[str] = None
    level: Optional[int] = None

class TextRun(BaseModel):
    """Individual text run with formatting"""
    text: str
    font: Optional[FontInfo] = None

class Paragraph(BaseModel):
    """Paragraph with formatting and runs"""
    text: str
    runs: List[TextRun] = []
    format: Optional[ParagraphFormat] = None

class TextFrame(BaseModel):
    """Text frame containing paragraphs"""
    text: str
    paragraphs: List[Paragraph] = []
    margin_left: Optional[int] = None
    margin_right: Optional[int] = None
    margin_top: Optional[int] = None
    margin_bottom: Optional[int] = None
    word_wrap: Optional[bool] = None
    auto_size: Optional[str] = None

class TableCell(BaseModel):
    """Table cell data"""
    text: str
    text_frame: Optional[TextFrame] = None
    fill_color: Optional[str] = None
    border_color: Optional[str] = None
    border_width: Optional[int] = None

class TableInfo(BaseModel):
    """Table structure and data"""
    rows: int
    columns: int
    cells: List[List[TableCell]] = []  # 2D array of cells

class ImageInfo(BaseModel):
    """Image data and properties"""
    filename: Optional[str] = None
    image_data: Optional[str] = None  # Base64 encoded image data
    content_type: Optional[str] = None  # image/jpeg, image/png, etc.
    crop_left: Optional[float] = None
    crop_top: Optional[float] = None
    crop_right: Optional[float] = None
    crop_bottom: Optional[float] = None

class FillFormat(BaseModel):
    """Fill formatting for shapes"""
    fill_type: Optional[str] = None  # SOLID, GRADIENT, PICTURE, etc.
    fore_color: Optional[str] = None  # RGB hex color
    back_color: Optional[str] = None  # RGB hex color
    transparency: Optional[float] = None

class LineFormat(BaseModel):
    """Line formatting for shapes"""
    color: Optional[str] = None  # RGB hex color
    width: Optional[int] = None
    dash_style: Optional[str] = None
    transparency: Optional[float] = None

class ShadowFormat(BaseModel):
    """Shadow formatting"""
    visible: Optional[bool] = None
    blur_radius: Optional[float] = None
    distance: Optional[float] = None
    direction: Optional[float] = None
    color: Optional[str] = None

class Shape(BaseModel):
    """PowerPoint shape with all properties"""
    shape_id: int
    name: str
    shape_type: ShapeType
    left: int
    top: int
    width: int
    height: int
    rotation: Optional[float] = None
    z_order: Optional[int] = None
    
    # Text content
    text_frame: Optional[TextFrame] = None
    
    # Table content
    table: Optional[TableInfo] = None
    
    # Image content
    image: Optional[ImageInfo] = None
    
    # Formatting
    fill: Optional[FillFormat] = None
    line: Optional[LineFormat] = None
    shadow: Optional[ShadowFormat] = None
    
    # Auto shape properties
    auto_shape_type: Optional[str] = None
    adjustments: Dict[str, float] = {}
    
    # Group properties
    shapes: Optional[List['Shape']] = None  # For grouped shapes
    
    # Additional properties
    click_action: Optional[str] = None
    hyperlink: Optional[str] = None
    
    model_config = {"arbitrary_types_allowed": True}

class SlideLayout(BaseModel):
    """Slide layout information"""
    name: Optional[str] = None
    layout_type: Optional[str] = None

class SlideMaster(BaseModel):
    """Slide master information"""
    name: Optional[str] = None
    
class Slide(BaseModel):
    """PowerPoint slide"""
    slide_number: int
    name: Optional[str] = None
    layout: Optional[SlideLayout] = None
    shapes: List[Shape] = []
    notes: Optional[str] = None
    slide_id: Optional[int] = None
    
    # Background properties
    background_fill: Optional[FillFormat] = None
    
    # Timing and transitions
    advance_time: Optional[float] = None
    transition_type: Optional[str] = None
    transition_duration: Optional[float] = None

class DocumentProperties(BaseModel):
    """Document-level properties"""
    title: Optional[str] = None
    author: Optional[str] = None
    subject: Optional[str] = None
    keywords: Optional[str] = None
    comments: Optional[str] = None
    created: Optional[str] = None  # ISO datetime string
    modified: Optional[str] = None  # ISO datetime string
    last_modified_by: Optional[str] = None

class Presentation(BaseModel):
    """Complete PowerPoint presentation"""
    slide_width: int
    slide_height: int
    slides: List[Slide] = []
    slide_masters: List[SlideMaster] = []
    
    # Document properties
    core_properties: Optional[DocumentProperties] = None
    
    # Theme and styling
    theme_name: Optional[str] = None
    color_scheme: Dict[str, str] = {}  # Named colors to RGB hex
    font_scheme: Dict[str, str] = {}   # Named fonts
    
    # Presentation-level settings
    slide_size_type: Optional[str] = None  # SCREEN_4X3, SCREEN_16X9, etc.
    
    model_config = {"json_encoders": {}}

# Update forward reference
Shape.model_rebuild()