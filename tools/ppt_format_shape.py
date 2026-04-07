#!/usr/bin/env python3
"""
PowerPoint Format Shape Tool v3.1.0
Update styling of existing shapes including fill, line, opacity, and text formatting.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_format_shape.py --file presentation.pptx --slide 0 --shape 1 \\
        --fill-color "#FF0000" --fill-opacity 0.8 --json

Exit Codes:
    0: Success
    1: Error occurred

Note: The --transparency parameter is DEPRECATED. Use --fill-opacity instead.
      Opacity: 0.0 = invisible, 1.0 = opaque (opposite of transparency)
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
    ShapeNotFoundError,
    ColorHelper,
)

__version__ = "3.1.0"

COLOR_PRESETS = {
    "primary": "#0070C0",
    "secondary": "#595959",
    "accent": "#ED7D31",
    "success": "#70AD47",
    "warning": "#FFC000",
    "danger": "#C00000",
    "white": "#FFFFFF",
    "black": "#000000",
    "light_gray": "#D9D9D9",
    "dark_gray": "#404040",
    "transparent": None,
}

OPACITY_PRESETS = {
    "opaque": 1.0,
    "subtle": 0.85,
    "light": 0.7,
    "medium": 0.5,
    "heavy": 0.3,
    "very_light": 0.15,
    "nearly_invisible": 0.05,
}


def resolve_color(color: Optional[str]) -> Optional[str]:
    """Resolve color value, handling presets and hex formats."""
    if color is None:
        return None
    
    color_lower = color.lower().strip()
    
    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]
    
    if color_lower in ("none", "transparent", "clear"):
        return None
    
    if not color.startswith('#') and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass
    
    return color


def resolve_opacity(value: Optional[str], is_transparency: bool = False) -> Optional[float]:
    """
    Resolve opacity value, handling presets and numeric values.
    
    Args:
        value: Opacity specification (float, preset name, or percentage string)
        is_transparency: If True, value is transparency (inverted)
        
    Returns:
        Opacity as float (0.0 = invisible, 1.0 = opaque) or None
    """
    if value is None:
        return None
    
    result: float
    
    if isinstance(value, str):
        value_lower = value.lower().strip()
        
        if value_lower in OPACITY_PRESETS:
            result = OPACITY_PRESETS[value_lower]
        elif value_lower.endswith('%'):
            try:
                pct = float(value_lower[:-1]) / 100.0
                result = pct if not is_transparency else (1.0 - pct)
            except ValueError:
                raise ValueError(f"Invalid opacity value: {value}")
        else:
            try:
                result = float(value_lower)
                if is_transparency:
                    result = 1.0 - result
            except ValueError:
                raise ValueError(f"Invalid opacity value: {value}")
    else:
        result = float(value)
        if is_transparency:
            result = 1.0 - result
    
    return max(0.0, min(1.0, result))


def validate_formatting_params(
    fill_color: Optional[str],
    line_color: Optional[str],
    fill_opacity: Optional[float]
) -> Dict[str, Any]:
    """Validate formatting parameters and generate warnings."""
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    if fill_opacity is not None:
        if fill_opacity < 0.05:
            warnings.append(
                f"Fill opacity {fill_opacity} is very low. Shape may be nearly invisible."
            )
        validation_results["fill_opacity"] = fill_opacity
    
    if fill_color:
        try:
            ColorHelper.from_hex(fill_color)
            validation_results["fill_color_valid"] = True
        except Exception as e:
            validation_results["fill_color_valid"] = False
            validation_results["fill_color_error"] = str(e)
            warnings.append(f"Invalid fill color format: {fill_color}")
    
    if line_color:
        try:
            ColorHelper.from_hex(line_color)
            validation_results["line_color_valid"] = True
        except Exception as e:
            validation_results["line_color_valid"] = False
            warnings.append(f"Invalid line color format: {line_color}")
    
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0
    }


def format_shape(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    line_width: Optional[float] = None,
    fill_opacity: Optional[float] = None,
    line_opacity: Optional[float] = None,
    text_color: Optional[str] = None,
    text_size: Optional[int] = None,
    text_bold: Optional[bool] = None
) -> Dict[str, Any]:
    """
    Format existing shape with comprehensive styling options.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        shape_index: Target shape index (0-based)
        fill_color: Fill color (hex or preset name)
        line_color: Line/border color (hex or preset name)
        line_width: Line width in points
        fill_opacity: Fill opacity (0.0=invisible to 1.0=opaque)
        line_opacity: Line opacity (0.0=invisible to 1.0=opaque)
        text_color: Text color within shape
        text_size: Text size in points
        text_bold: Text bold setting
        
    Returns:
        Result dict with formatting details
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If no formatting options or file format invalid
        SlideNotFoundError: If slide index invalid
        ShapeNotFoundError: If shape index invalid
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    formatting_options = [
        fill_color, line_color, line_width, fill_opacity, line_opacity,
        text_color, text_size, text_bold
    ]
    if all(v is None for v in formatting_options):
        raise ValueError(
            "At least one formatting option required. "
            "Use --fill-color, --line-color, --fill-opacity, etc."
        )
    
    resolved_fill = resolve_color(fill_color)
    resolved_line = resolve_color(line_color)
    resolved_text_color = resolve_color(text_color)
    
    validation = validate_formatting_params(
        fill_color=resolved_fill,
        line_color=resolved_line,
        fill_opacity=fill_opacity
    )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", 0)
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={"requested": shape_index, "available": shape_count}
            )
        
        version_before = agent.get_presentation_version()
        
        format_result = agent.format_shape(
            slide_index=slide_index,
            shape_index=shape_index,
            fill_color=resolved_fill,
            line_color=resolved_line,
            line_width=line_width,
            fill_opacity=fill_opacity,
            line_opacity=line_opacity
        )
        
        text_formatted = False
        if any(v is not None for v in [text_color, text_size, text_bold]):
            try:
                agent.format_text(
                    slide_index=slide_index,
                    shape_index=shape_index,
                    color=resolved_text_color,
                    font_size=text_size,
                    bold=text_bold
                )
                text_formatted = True
            except Exception as e:
                validation["warnings"].append(f"Could not format text: {e}")
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    result: Dict[str, Any] = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "formatting_applied": {
            "fill_color": resolved_fill,
            "fill_opacity": fill_opacity,
            "line_color": resolved_line,
            "line_opacity": line_opacity,
            "line_width": line_width,
            "text_color": resolved_text_color if text_formatted else None,
            "text_size": text_size if text_formatted else None,
            "text_bold": text_bold if text_formatted else None
        },
        "changes_from_core": format_result.get("changes_applied", []) if isinstance(format_result, dict) else [],
        "text_formatted": text_formatted,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if validation["validation_results"]:
        result["validation"] = validation["validation_results"]
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Format existing PowerPoint shape",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
FORMATTING OPTIONS:
  --fill-color     Shape fill color (hex or preset)
  --fill-opacity   Fill opacity: 0.0 (invisible) to 1.0 (opaque)
  --line-color     Border/line color (hex or preset)
  --line-opacity   Line opacity: 0.0 (invisible) to 1.0 (opaque)
  --line-width     Border width in points
  --text-color     Text color within shape
  --text-size      Text size in points
  --text-bold      Make text bold

COLOR PRESETS:
  primary (#0070C0)    secondary (#595959)    accent (#ED7D31)
  success (#70AD47)    warning (#FFC000)      danger (#C00000)
  white (#FFFFFF)      black (#000000)        transparent (none)

OPACITY PRESETS:
  opaque (1.0)         subtle (0.85)          light (0.7)
  medium (0.5)         heavy (0.3)            very_light (0.15)

EXAMPLES:

  # Change fill color
  uv run tools/ppt_format_shape.py --file deck.pptx --slide 0 --shape 1 \\
    --fill-color "#FF0000" --json

  # Semi-transparent overlay
  uv run tools/ppt_format_shape.py --file deck.pptx --slide 1 --shape 0 \\
    --fill-color black --fill-opacity 0.5 --json

  # Format text within shape
  uv run tools/ppt_format_shape.py --file deck.pptx --slide 0 --shape 3 \\
    --fill-color primary --text-color white --text-size 24 --text-bold --json

FINDING SHAPE INDEX:
  Use ppt_get_slide_info.py to find shape indices:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

DEPRECATED:
  --transparency is deprecated. Use --fill-opacity instead.
  transparency = 1.0 - fill_opacity (values are inverted)
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--fill-color', help='Fill color: hex or preset')
    parser.add_argument('--fill-opacity', help='Fill opacity: 0.0 (invisible) to 1.0 (opaque)')
    parser.add_argument('--line-color', help='Line/border color')
    parser.add_argument('--line-opacity', help='Line opacity: 0.0 to 1.0')
    parser.add_argument('--line-width', type=float, help='Line width in points')
    parser.add_argument('--transparency', help='DEPRECATED: Use --fill-opacity. Transparency: 0.0 (opaque) to 1.0 (invisible)')
    parser.add_argument('--text-color', help='Text color within shape')
    parser.add_argument('--text-size', type=int, help='Text size in points')
    parser.add_argument('--text-bold', action='store_true', help='Make text bold')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        fill_opacity: Optional[float] = None
        deprecation_warning: Optional[str] = None
        
        if args.fill_opacity is not None:
            fill_opacity = resolve_opacity(args.fill_opacity, is_transparency=False)
        elif args.transparency is not None:
            fill_opacity = resolve_opacity(args.transparency, is_transparency=True)
            deprecation_warning = (
                "--transparency is deprecated. Use --fill-opacity instead. "
                f"Converted transparency {args.transparency} to fill_opacity {fill_opacity}"
            )
        
        line_opacity: Optional[float] = None
        if args.line_opacity is not None:
            line_opacity = resolve_opacity(args.line_opacity, is_transparency=False)
        
        result = format_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            fill_color=args.fill_color,
            line_color=args.line_color,
            line_width=args.line_width,
            fill_opacity=fill_opacity,
            line_opacity=line_opacity,
            text_color=args.text_color,
            text_size=args.text_size,
            text_bold=args.text_bold if args.text_bold else None
        )
        
        if deprecation_warning:
            if "warnings" not in result:
                result["warnings"] = []
            result["warnings"].insert(0, deprecation_warning)
            result["status"] = "warning"
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
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
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Provide at least one formatting option and check opacity values."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check the presentation file is valid."
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
