#!/usr/bin/env python3
"""
PowerPoint Add Text Box Tool v3.1.0
Add text box with flexible positioning, comprehensive validation, and accessibility checking.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_text_box.py --file deck.pptx --slide 0 \\
        --text "Revenue: $1.5M" --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"10%"}' --json

Exit Codes:
    0: Success
    1: Error occurred

Position Formats:
  1. Percentage: {"left": "20%", "top": "30%"}
  2. Inches: {"left": 2.0, "top": 3.0}
  3. Anchor: {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
  4. Grid: {"grid_row": 2, "grid_col": 3, "grid_size": 12}
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ColorHelper,
)

__version__ = "3.1.0"

COLOR_PRESETS = {
    "black": "#000000",
    "white": "#FFFFFF",
    "primary": "#0070C0",
    "secondary": "#595959",
    "accent": "#ED7D31",
    "success": "#70AD47",
    "warning": "#FFC000",
    "danger": "#C00000",
    "dark_gray": "#333333",
    "light_gray": "#808080",
    "muted": "#808080",
}

FONT_PRESETS = {
    "default": "Calibri",
    "heading": "Calibri Light",
    "body": "Calibri",
    "code": "Consolas",
    "serif": "Georgia",
    "sans": "Arial",
}


def resolve_color(color: Optional[str]) -> Optional[str]:
    """
    Resolve color value, handling presets and hex formats.
    
    Args:
        color: Color specification (hex, preset name, or None)
        
    Returns:
        Resolved hex color or None
    """
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    if not color.startswith('#') and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass
    
    return color


def resolve_font(font: Optional[str]) -> str:
    """
    Resolve font name, handling presets.
    
    Args:
        font: Font name or preset
        
    Returns:
        Resolved font name
    """
    if font is None:
        return "Calibri"
    
    font_lower = font.lower().strip()
    if font_lower in FONT_PRESETS:
        return FONT_PRESETS[font_lower]
    
    return font


def validate_text_box(
    text: str,
    font_size: int,
    color: Optional[str] = None,
    position: Optional[Dict[str, Any]] = None,
    size: Optional[Dict[str, Any]] = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Validate text box parameters and return warnings/recommendations.
    
    Args:
        text: Text content
        font_size: Font size in points
        color: Text color hex
        position: Position specification
        size: Size specification
        allow_offslide: Allow off-slide positioning
        
    Returns:
        Dict with warnings, recommendations, and validation results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    text_length = len(text)
    line_count = text.count('\n') + 1
    
    validation_results["text_length"] = text_length
    validation_results["line_count"] = line_count
    validation_results["is_multiline"] = line_count > 1
    
    if line_count == 1 and text_length > 100:
        warnings.append(
            f"Text is {text_length} characters for single line (recommended: ≤100). "
            "Long single-line text may be hard to read."
        )
        recommendations.append("Consider breaking into multiple lines or shortening text")
    
    if line_count > 1 and text_length > 500:
        warnings.append(
            f"Multi-line text is {text_length} characters. Very long text blocks reduce readability."
        )
    
    validation_results["font_size"] = font_size
    validation_results["font_size_accessible"] = font_size >= 14
    
    if font_size < 10:
        warnings.append(
            f"Font size {font_size}pt is below minimum (10pt). Text will be very hard to read."
        )
    elif font_size < 12:
        warnings.append(
            f"Font size {font_size}pt is very small. Consider 14pt+ for projected presentations."
        )
        recommendations.append("Use 14pt or larger for projected content")
    elif font_size < 14:
        recommendations.append(
            f"Font size {font_size}pt is below recommended 14pt for projected content"
        )
    
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            from pptx.dml.color import RGBColor
            bg_color = RGBColor(255, 255, 255)
            
            is_large_text = font_size >= 18
            contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
            meets_wcag = ColorHelper.meets_wcag(text_color, bg_color, is_large_text)
            
            validation_results["color_contrast"] = {
                "ratio": round(contrast_ratio, 2),
                "wcag_aa": meets_wcag,
                "required_ratio": 3.0 if is_large_text else 4.5,
                "is_large_text": is_large_text
            }
            
            if not meets_wcag:
                required = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color contrast {contrast_ratio:.2f}:1 may not meet WCAG AA "
                    f"(required: {required}:1). Consider darker color."
                )
                recommendations.append(
                    "Use black (#000000), dark_gray (#333333), or primary (#0070C0) for better contrast"
                )
        except Exception as e:
            validation_results["color_error"] = str(e)
    
    if position:
        try:
            for key in ["left", "top"]:
                if key in position:
                    value_str = str(position[key])
                    if value_str.endswith('%'):
                        pct = float(value_str.rstrip('%'))
                        if not allow_offslide and (pct < 0 or pct > 100):
                            warnings.append(
                                f"Position '{key}' is {pct}% which is outside slide bounds (0-100%). "
                                f"Text box may not be visible."
                            )
        except (ValueError, TypeError):
            pass
    
    if size:
        try:
            if "height" in size:
                height_str = str(size["height"])
                if height_str.endswith('%'):
                    height_pct = float(height_str.rstrip('%'))
                    min_height = font_size * 0.15
                    if height_pct < min_height:
                        warnings.append(
                            f"Height {height_pct}% may be too small for {font_size}pt text. "
                            f"Consider at least {min_height:.1f}%."
                        )
            
            if "width" in size:
                width_str = str(size["width"])
                if width_str.endswith('%'):
                    width_pct = float(width_str.rstrip('%'))
                    if width_pct < 5:
                        warnings.append(
                            f"Width {width_pct}% is very narrow. Text may be excessively wrapped."
                        )
        except (ValueError, TypeError):
            pass
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def add_text_box(
    filepath: Path,
    slide_index: int,
    text: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    font_name: str = "Calibri",
    font_size: int = 18,
    bold: bool = False,
    italic: bool = False,
    color: Optional[str] = None,
    alignment: str = "left",
    vertical_alignment: str = "top",
    word_wrap: bool = True,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Add text box with comprehensive validation and formatting.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Slide index (0-based)
        text: Text content
        position: Position dict (supports %, inches, anchor, grid)
        size: Size dict
        font_name: Font name or preset
        font_size: Font size in points
        bold: Bold text
        italic: Italic text
        color: Text color (hex or preset)
        alignment: Horizontal alignment (left, center, right, justify)
        vertical_alignment: Vertical alignment (top, middle, bottom)
        word_wrap: Enable word wrap
        allow_offslide: Allow off-slide positioning
        
    Returns:
        Result dict with shape_index and validation info
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid
        SlideNotFoundError: If slide index out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    resolved_color = resolve_color(color)
    resolved_font = resolve_font(font_name)
    
    validation = validate_text_box(
        text=text,
        font_size=font_size,
        color=resolved_color,
        position=position,
        size=size,
        allow_offslide=allow_offslide
    )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        version_before = agent.get_presentation_version()
        
        add_result = agent.add_text_box(
            slide_index=slide_index,
            text=text,
            position=position,
            size=size,
            font_name=resolved_font,
            font_size=font_size,
            bold=bold,
            italic=italic,
            color=resolved_color,
            alignment=alignment
        )
        
        agent.save()
        
        version_after = agent.get_presentation_version()
        slide_info = agent.get_slide_info(slide_index)
    
    result = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": add_result.get("shape_index") if isinstance(add_result, dict) else add_result,
        "text": text,
        "text_length": len(text),
        "position": add_result.get("position", position) if isinstance(add_result, dict) else position,
        "size": add_result.get("size", size) if isinstance(add_result, dict) else size,
        "formatting": {
            "font_name": resolved_font,
            "font_size": font_size,
            "bold": bold,
            "italic": italic,
            "color": resolved_color,
            "alignment": alignment,
            "vertical_alignment": vertical_alignment,
            "word_wrap": word_wrap
        },
        "slide_shape_count": slide_info.get("shape_count", 0),
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
        description="Add text box to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
POSITION FORMATS:

  Percentage (recommended):
    {"left": "20%", "top": "30%"}
    
  Absolute inches:
    {"left": 2.0, "top": 3.0}
    
  Anchor-based:
    {"anchor": "center", "offset_x": 0, "offset_y": -1.0}
    Anchors: top_left, top_center, top_right,
             center_left, center, center_right,
             bottom_left, bottom_center, bottom_right
    
  Grid (12-column):
    {"grid_row": 2, "grid_col": 3, "grid_size": 12}

COLOR PRESETS:

  black (#000000)      white (#FFFFFF)      primary (#0070C0)
  secondary (#595959)  accent (#ED7D31)     success (#70AD47)
  warning (#FFC000)    danger (#C00000)     dark_gray (#333333)
  light_gray (#808080) muted (#808080)

FONT PRESETS:

  default (Calibri)    heading (Calibri Light)   body (Calibri)
  code (Consolas)      serif (Georgia)           sans (Arial)

EXAMPLES:

  # Simple text box
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 0 \\
    --text "Revenue: \\$1.5M" \\
    --position '{"left":"20%","top":"30%"}' \\
    --size '{"width":"60%","height":"10%"}' \\
    --font-size 24 --bold --json

  # Centered headline
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 1 \\
    --text "Key Results" \\
    --position '{"anchor":"center","offset_y":-2}' \\
    --size '{"width":"80%","height":"15%"}' \\
    --font-size 48 --bold --color primary --alignment center --json

  # Copyright notice (bottom-right)
  uv run tools/ppt_add_text_box.py \\
    --file deck.pptx --slide 0 \\
    --text "© 2024 Company Inc." \\
    --position '{"anchor":"bottom_right","offset_x":-0.5,"offset_y":-0.3}' \\
    --size '{"width":"20%","height":"5%"}' \\
    --font-size 10 --color muted --json

ACCESSIBILITY GUIDELINES:

  Font Size:
    - Minimum: 10pt (absolute minimum)
    - Recommended: 14pt+ for projected presentations
    - Large text: 18pt+ (relaxed contrast requirements)

  Color Contrast (WCAG 2.1 AA):
    - Normal text (<18pt): 4.5:1 minimum
    - Large text (>=18pt): 3.0:1 minimum
    - Best colors: black, dark_gray, primary
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--text',
        required=True,
        help='Text content (use \\n for line breaks)'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict as JSON'
    )
    
    parser.add_argument(
        '--size',
        type=json.loads,
        help='Size dict as JSON (defaults to 40%% x 20%% if omitted)'
    )
    
    parser.add_argument(
        '--font-name',
        default='Calibri',
        help='Font name or preset (default, heading, body, code, serif, sans)'
    )
    
    parser.add_argument(
        '--font-size',
        type=int,
        default=18,
        help='Font size in points (default: 18, recommended: >=14)'
    )
    
    parser.add_argument(
        '--bold',
        action='store_true',
        help='Make text bold'
    )
    
    parser.add_argument(
        '--italic',
        action='store_true',
        help='Make text italic'
    )
    
    parser.add_argument(
        '--color',
        help='Text color: hex (#0070C0) or preset (primary, danger, etc.)'
    )
    
    parser.add_argument(
        '--alignment',
        choices=['left', 'center', 'right', 'justify'],
        default='left',
        help='Horizontal text alignment (default: left)'
    )
    
    parser.add_argument(
        '--vertical-alignment',
        choices=['top', 'middle', 'bottom'],
        default='top',
        help='Vertical text alignment (default: top)'
    )
    
    parser.add_argument(
        '--no-word-wrap',
        action='store_true',
        help='Disable word wrap'
    )
    
    parser.add_argument(
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        size = args.size if args.size else {}
        position = args.position
        
        if "width" in position and "width" not in size:
            size["width"] = position.get("width")
        if "height" in position and "height" not in size:
            size["height"] = position.get("height")
        
        if "width" not in size:
            size["width"] = "40%"
        if "height" not in size:
            size["height"] = "20%"
        
        result = add_text_box(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            position=position,
            size=size,
            font_name=args.font_name,
            font_size=args.font_size,
            bold=args.bold,
            italic=args.italic,
            color=args.color,
            alignment=args.alignment,
            vertical_alignment=args.vertical_alignment,
            word_wrap=not args.no_word_wrap,
            allow_offslide=args.allow_offslide
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON: {e}",
            "error_type": "JSONDecodeError",
            "suggestion": "Use single quotes around JSON: '{\"left\":\"20%\",\"top\":\"30%\"}'"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check file format is .pptx and parameters are valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check the presentation file is valid and not corrupted."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
