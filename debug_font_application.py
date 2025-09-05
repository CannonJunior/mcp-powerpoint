#!/usr/bin/env python3
"""
Debug script to test font color application
"""

from pptx.dml.color import RGBColor
from pptx.util import Pt

def hex_to_rgb(hex_color: str):
    """Convert hex color string to RGBColor"""
    if not hex_color or not hex_color.startswith('#'):
        return None
    try:
        rgb_int = int(hex_color[1:], 16)
        return RGBColor((rgb_int >> 16) & 0xFF, (rgb_int >> 8) & 0xFF, rgb_int & 0xFF)
    except:
        return None

def test_hex_to_rgb():
    """Test hex color conversion"""
    test_colors = ["#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF"]
    
    print("=== TESTING HEX TO RGB CONVERSION ===")
    for hex_color in test_colors:
        rgb_color = hex_to_rgb(hex_color)
        if rgb_color:
            print(f"  {hex_color} -> {rgb_color} (type: {type(rgb_color)})")
        else:
            print(f"  {hex_color} -> FAILED")

if __name__ == "__main__":
    test_hex_to_rgb()