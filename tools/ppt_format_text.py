#!/usr/bin/env python3
"""
PowerPoint Format Text Tool v3.1.0
Format existing text with accessibility validation and contrast checking

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Font name, size, color, bold, italic formatting
    - WCAG 2.1 AA/AAA color contrast validation
    - Font size accessibility warnings (<12pt)
    - Before/after formatting comparison
    - Detailed validation results and recommendations

Usage:
    uv run tools/ppt_format_text.py --file deck.pptx --slide 0 --shape 0 --font-name "Arial" --font-size 24 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    - Minimum font size: 12pt (14pt recommended for presentations)
    - Color contrast: 4.5:1 for normal text, 3:1 for large text (≥18pt)
    - Tool validates and warns about accessibility issues
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import math
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exception if not available in core
try:
    from core.powerpoint_agent_core import ShapeNotFoundError
except ImportError:
    class ShapeNotFoundError(PowerPointAgentError):
        """Exception raised when shape is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)

# Color helper functions (fallback if not in core)
try:
    from core.powerpoint_agent_core import ColorHelper, RGBColor
except ImportError:
    # Define minimal color helpers locally
    class RGBColor:
        """Simple RGB color class."""
        def __init__(self, r: int, g: int, b: int):
            self.r = r
            self.g = g
            self.b = b
    
    class ColorHelper:
        """Color utilities for accessibility checking."""
        
        @staticmethod
        def from_hex(hex_color: str) -> RGBColor:
            """Convert hex color to RGBColor."""
            hex_color = hex_color.lstrip('#')
            if len(hex_color) != 6:
                raise ValueError(f"Invalid hex color format: #{hex_color}")
            try:
                r = int(hex_color[0:2], 16)
                g = int(hex_color[2:4], 16)
                b = int(hex_color[4:6], 16)
                return RGBColor(r, g, b)
            except ValueError:
                raise ValueError(f"Invalid hex color: #{hex_color}")
        
        @staticmethod
        def _relative_luminance(color: RGBColor) -> float:
            """Calculate relative luminance per WCAG 2.1."""
            def channel_luminance(c: int) -> float:
                c_srgb = c / 255.0
                if c_srgb <= 0.03928:
                    return c_srgb / 12.92
                else:
                    return ((c_srgb + 0.055) / 1.055) ** 2.4
            
            return (0.2126 * channel_luminance(color.r) + 
                    0.7152 * channel_luminance(color.g) + 
                    0.0722 * channel_luminance(color.b))
        
        @staticmethod
        def contrast_ratio(color1: RGBColor, color2: RGBColor) -> float:
            """Calculate contrast ratio between two colors per WCAG 2.1."""
            l1 = ColorHelper._relative_luminance(color1)
            l2 = ColorHelper._relative_luminance(color2)
            
            lighter = max(l1, l2)
            darker = min(l1, l2)
            
            return (lighter + 0.05) / (darker + 0.05)
        
        @staticmethod
        def meets_wcag(text_color: RGBColor, bg_color: RGBColor, is_large_text: bool = False) -> bool:
            """Check if colors meet WCAG AA contrast requirements."""
            ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            required = 3.0 if is_large_text else 4.5
            return ratio >= required


def validate_formatting(
    font_size: Optional[int] = None,
    color: Optional[str] = None,
    current_font_size: Optional[int] = None
) -> Dict[str, Any]:
    """
    Validate formatting parameters against accessibility guidelines.
    
    Args:
        font_size: New font size to validate
        color: New color to validate (hex format)
        current_font_size: Current font size for comparison
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    # Font size validation
    if font_size is not None:
        validation_results["font_size"] = font_size
        validation_results["font_size_ok"] = font_size >= 12
        
        if font_size < 10:
            warnings.append(
                f"Font size {font_size}pt is extremely small. "
                "Minimum recommended: 12pt for handouts, 14pt for presentations."
            )
        elif font_size < 12:
            warnings.append(
                f"Font size {font_size}pt is below minimum recommended 12pt. "
                "Audience may struggle to read."
            )
            recommendations.append("Use 12pt minimum, 14pt+ for projected content")
        elif font_size < 14:
            recommendations.append(
                f"Font size {font_size}pt is acceptable for handouts but consider 14pt+ for projected presentations"
            )
        
        # Check if decreasing size
        if current_font_size and font_size < current_font_size:
            diff = current_font_size - font_size
            recommendations.append(
                f"Decreasing font size by {diff}pt (from {current_font_size}pt to {font_size}pt). "
                "Verify readability on target display."
            )
    
    # Color contrast validation
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            bg_color = RGBColor(255, 255, 255)  # Assume white background
            
            # Determine if large text
            effective_font_size = font_size if font_size else (current_font_size if current_font_size else 18)
            is_large_text = effective_font_size >= 18
            
            contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            wcag_aa = ColorHelper.meets_wcag(text_color, bg_color, is_large_text)
            
            validation_results["color_contrast"] = {
                "color": color,
                "ratio": round(contrast_ratio, 2),
                "wcag_aa": wcag_aa,
                "is_large_text": is_large_text,
                "required_ratio": 3.0 if is_large_text else 4.5
            }
            
            if not wcag_aa:
                required = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color {color} has contrast ratio {contrast_ratio:.2f}:1 "
                    f"(WCAG AA requires {required}:1 for {'large' if is_large_text else 'normal'} text). "
                    "May not meet accessibility standards."
                )
                recommendations.append(
                    "Use high-contrast colors: #000000 (black), #333333 (dark gray), #0070C0 (dark blue)"
                )
            elif contrast_ratio < 7.0:
                recommendations.append(
                    f"Color contrast {contrast_ratio:.2f}:1 meets WCAG AA but not AAA (7:1). "
                    "Consider darker color for maximum accessibility."
                )
        except ValueError as e:
            validation_results["color_error"] = str(e)
            warnings.append(f"Invalid color format: {color}. Use hex format like #FF0000")
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results
    }


def format_text(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    font_name: Optional[str] = None,
    font_size: Optional[int] = None,
    color: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None
) -> Dict[str, Any]:
    """
    Format text in a shape with validation and accessibility checking.
    
    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        shape_index: Shape index (0-based)
        font_name: Optional font name to apply
        font_size: Optional font size in points
        color: Optional text color (hex format, e.g., "#0070C0")
        bold: Optional bold setting (True/False/None for no change)
        italic: Optional italic setting (True/False/None for no change)
        
    Returns:
        Dict containing:
            - status: "success" or "warning" (if accessibility issues)
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the shape
            - before: Original formatting state
            - after: New formatting applied
            - changes_applied: List of changed properties
            - validation: Validation results
            - warnings: Accessibility/formatting warnings (if any)
            - recommendations: Suggested improvements (if any)
            - presentation_version_before: State hash before modification
            - presentation_version_after: State hash after modification
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ShapeNotFoundError: If shape index is out of range
        ValueError: If no formatting options provided or shape has no text
        
    Example:
        >>> result = format_text(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     shape_index=2,
        ...     font_size=18,
        ...     color="#0070C0",
        ...     bold=True
        ... )
        >>> print(result["changes_applied"])
        ['font_size', 'color', 'bold']
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Check that at least one formatting option is provided
    if all(v is None for v in [font_name, font_size, color, bold, italic]):
        raise ValueError(
            "At least one formatting option must be specified. "
            "Use --font-name, --font-size, --color, --bold, or --italic"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE formatting
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Get slide info to validate shape index
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", len(slide_info.get("shapes", [])))
        
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={
                    "requested_index": shape_index,
                    "available_shapes": shape_count
                }
            )
        
        # Check if shape has text
        shapes = slide_info.get("shapes", [])
        shape_info = shapes[shape_index] if shape_index < len(shapes) else {}
        
        if not shape_info.get("has_text", False):
            raise ValueError(
                f"Shape {shape_index} ({shape_info.get('type', 'unknown')}) does not contain text. "
                "Cannot format non-text shape. Use ppt_get_slide_info.py to find text-containing shapes."
            )
        
        # Extract current formatting info
        before_formatting = {
            "shape_type": shape_info.get("type"),
            "shape_name": shape_info.get("name"),
            "has_text": shape_info.get("has_text", False)
        }
        
        # Try to get current font size for validation
        current_font_size = None
        try:
            # Access slide directly for font size extraction
            slide = agent.prs.slides[slide_index]
            shape = slide.shapes[shape_index]
            if hasattr(shape, 'text_frame') and shape.text_frame.paragraphs:
                first_para = shape.text_frame.paragraphs[0]
                if first_para.runs and first_para.runs[0].font.size:
                    current_font_size = int(first_para.runs[0].font.size.pt)
                    before_formatting["font_size"] = current_font_size
                elif first_para.font.size:
                    current_font_size = int(first_para.font.size.pt)
                    before_formatting["font_size"] = current_font_size
        except Exception:
            pass  # Continue without current font size
        
        # Validate formatting parameters
        validation = validate_formatting(font_size, color, current_font_size)
        
        # Apply formatting via core
        agent.format_text(
            slide_index=slide_index,
            shape_index=shape_index,
            font_name=font_name,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=color
        )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER formatting
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Determine status based on warnings
    status = "success" if len(validation["warnings"]) == 0 else "warning"
    
    # Build after formatting dict
    after_formatting: Dict[str, Any] = {}
    if font_name is not None:
        after_formatting["font_name"] = font_name
    if font_size is not None:
        after_formatting["font_size"] = font_size
    if color is not None:
        after_formatting["color"] = color
    if bold is not None:
        after_formatting["bold"] = bold
    if italic is not None:
        after_formatting["italic"] = italic
    
    # Build result
    result: Dict[str, Any] = {
        "status": status,
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "before": before_formatting,
        "after": after_formatting,
        "changes_applied": list(after_formatting.keys()),
        "validation": validation["validation_results"],
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Format text in PowerPoint shape with accessibility validation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Change font and size
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 0 \\
    --font-name "Arial" \\
    --font-size 24 \\
    --json
  
  # Make text bold and colored (with validation)
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --shape 2 \\
    --bold \\
    --color "#0070C0" \\
    --json
  
  # Comprehensive formatting
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --font-name "Calibri" \\
    --font-size 18 \\
    --bold \\
    --color "#000000" \\
    --json
  
  # Fix accessibility issue
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --shape 3 \\
    --font-size 16 \\
    --color "#333333" \\
    --json
  
  # Remove bold/italic
  uv run tools/ppt_format_text.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --shape 1 \\
    --no-bold \\
    --no-italic \\
    --json

Finding Shape Index:
  Use ppt_get_slide_info.py to list shapes and their indices:
  uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 0 --json
  
  Look for shapes with "has_text": true

Common Cross-Platform Fonts:
  - Calibri (default Microsoft Office)
  - Arial (universal)
  - Times New Roman (classic serif)
  - Verdana (screen-optimized)

Accessible Color Palette:
  High Contrast (WCAG AAA - 7:1):
  - #000000 (Black)
  - #333333 (Dark Charcoal)
  - #003366 (Navy Blue)
  
  Good Contrast (WCAG AA - 4.5:1):
  - #595959 (Dark Gray)
  - #0070C0 (Corporate Blue)
  - #006400 (Forest Green)

Accessibility Guidelines:
  - Minimum font size: 12pt (14pt for presentations)
  - Color contrast: 4.5:1 for normal text, 3:1 for large text (≥18pt)
  - Tool automatically validates and warns about issues

Output Format:
  {
    "status": "warning",
    "slide_index": 0,
    "shape_index": 2,
    "before": {"font_size": 24},
    "after": {"font_size": 11, "color": "#CCCCCC"},
    "changes_applied": ["font_size", "color"],
    "validation": {
      "font_size": 11,
      "font_size_ok": false,
      "color_contrast": {"ratio": 2.1, "wcag_aa": false}
    },
    "warnings": ["Font size 11pt is below minimum..."],
    "recommendations": ["Use 12pt minimum..."],
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--shape',
        required=True,
        type=int,
        help='Shape index (0-based, use ppt_get_slide_info.py to find)'
    )
    
    parser.add_argument(
        '--font-name',
        help='Font name (e.g., Arial, Calibri)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        help='Font size in points (minimum recommended: 12pt)'
    )
    
    parser.add_argument(
        '--color',
        help='Text color hex (e.g., #0070C0). Contrast will be validated.'
    )
    
    parser.add_argument(
        '--bold',
        action='store_true',
        dest='bold',
        help='Make text bold'
    )
    
    parser.add_argument(
        '--no-bold',
        action='store_false',
        dest='bold',
        help='Remove bold formatting'
    )
    
    parser.add_argument(
        '--italic',
        action='store_true',
        dest='italic',
        help='Make text italic'
    )
    
    parser.add_argument(
        '--no-italic',
        action='store_false',
        dest='italic',
        help='Remove italic formatting'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.set_defaults(bold=None, italic=None)
    
    args = parser.parse_args()
    
    try:
        result = format_text(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            font_name=args.font_name,
            font_size=args.font_size,
            color=args.color,
            bold=args.bold,
            italic=args.italic
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Provide at least one formatting option and ensure shape has text"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "file": str(args.file) if args.file else None,
            "slide_index": args.slide if hasattr(args, 'slide') else None,
            "shape_index": args.shape if hasattr(args, 'shape') else None,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
