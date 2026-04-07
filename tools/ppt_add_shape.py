#!/usr/bin/env python3
"""
PowerPoint Add Shape Tool v3.1.0
Add shapes (rectangle, circle, arrow, etc.) to slides with comprehensive styling options.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"20%","top":"30%"}' \\
        --size '{"width":"60%","height":"40%"}' --fill-color "#0070C0" --json

    # Overlay with opacity
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --position '{"left":"0%","top":"0%"}' \\
        --size '{"width":"100%","height":"100%"}' \\
        --fill-color "#000000" --fill-opacity 0.15 --json

    # Quick overlay preset
    uv run tools/ppt_add_shape.py --file presentation.pptx --slide 0 \\
        --shape rectangle --overlay --fill-color "#FFFFFF" --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, "w")

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    ColorHelper,
)

__version__ = "3.1.0"

AVAILABLE_SHAPES = [
    "rectangle",
    "rounded_rectangle",
    "ellipse",
    "oval",
    "triangle",
    "arrow_right",
    "arrow_left",
    "arrow_up",
    "arrow_down",
    "diamond",
    "pentagon",
    "hexagon",
    "star",
    "heart",
    "lightning",
    "sun",
    "moon",
    "cloud",
]

SHAPE_ALIASES = {
    "rect": "rectangle",
    "round_rect": "rounded_rectangle",
    "circle": "ellipse",
    "arrow": "arrow_right",
    "right_arrow": "arrow_right",
    "left_arrow": "arrow_left",
    "up_arrow": "arrow_up",
    "down_arrow": "arrow_down",
    "5_point_star": "star",
}

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
}

OVERLAY_DEFAULTS = {
    "position": {"left": "0%", "top": "0%"},
    "size": {"width": "100%", "height": "100%"},
    "fill_opacity": 0.15,
    "z_order_action": "send_to_back",
}


def resolve_shape_type(shape_type: str) -> str:
    """Resolve shape type, handling aliases."""
    shape_lower = shape_type.lower().strip()

    if shape_lower in SHAPE_ALIASES:
        return SHAPE_ALIASES[shape_lower]

    if shape_lower in AVAILABLE_SHAPES:
        return shape_lower

    for available in AVAILABLE_SHAPES:
        if shape_lower in available or available in shape_lower:
            return available

    return shape_lower


def resolve_color(color: Optional[str]) -> Optional[str]:
    """Resolve color, handling presets and validation."""
    if color is None:
        return None

    color_lower = color.lower().strip()

    if color_lower in COLOR_PRESETS:
        return COLOR_PRESETS[color_lower]

    if not color.startswith("#") and len(color) == 6:
        try:
            int(color, 16)
            return f"#{color}"
        except ValueError:
            pass

    return color


def validate_opacity(
    fill_opacity: float, line_opacity: float
) -> Tuple[List[str], List[str]]:
    """Validate opacity values and return warnings/recommendations."""
    warnings: List[str] = []
    recommendations: List[str] = []

    if not 0.0 <= fill_opacity <= 1.0:
        raise ValueError(
            f"fill_opacity must be between 0.0 and 1.0, got {fill_opacity}"
        )

    if not 0.0 <= line_opacity <= 1.0:
        raise ValueError(
            f"line_opacity must be between 0.0 and 1.0, got {line_opacity}"
        )

    if fill_opacity == 0.0:
        warnings.append(
            "Fill opacity is 0.0 (fully transparent). Shape fill will be invisible."
        )
    elif fill_opacity < 0.05:
        warnings.append(
            f"Fill opacity {fill_opacity} is extremely low (<5%). Shape may be nearly invisible."
        )

    if line_opacity == 0.0 and fill_opacity == 0.0:
        warnings.append(
            "Both fill and line opacity are 0.0. Shape will be completely invisible."
        )

    if 0.1 <= fill_opacity <= 0.3:
        recommendations.append(
            f"Opacity {fill_opacity} is appropriate for overlay backgrounds. "
            "Remember to use ppt_set_z_order.py --action send_to_back after adding."
        )

    return warnings, recommendations


def validate_shape_params(
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    fill_opacity: float = 1.0,
    line_color: Optional[str] = None,
    line_opacity: float = 1.0,
    text: Optional[str] = None,
    allow_offslide: bool = False,
    is_overlay: bool = False,
) -> Dict[str, Any]:
    """Validate shape parameters and return warnings/recommendations."""
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}

    opacity_warnings, opacity_recommendations = validate_opacity(
        fill_opacity, line_opacity
    )
    warnings.extend(opacity_warnings)
    recommendations.extend(opacity_recommendations)

    validation_results["fill_opacity"] = fill_opacity
    validation_results["line_opacity"] = line_opacity
    validation_results["effective_fill_transparency"] = round(1.0 - fill_opacity, 2)

    if position:
        try:
            for key in ["left", "top"]:
                if key in position:
                    value_str = str(position[key])
                    if value_str.endswith("%"):
                        pct = float(value_str.rstrip("%"))
                        if not allow_offslide and (pct < 0 or pct > 100):
                            warnings.append(
                                f"Position '{key}' is {pct}% which is outside slide bounds (0-100%)."
                            )
        except (ValueError, TypeError):
            pass

    if size:
        try:
            for key in ["width", "height"]:
                if key in size:
                    value_str = str(size[key])
                    if value_str.endswith("%"):
                        pct = float(value_str.rstrip("%"))
                        if pct <= 0:
                            warnings.append(
                                f"Size '{key}' is {pct}% which is invalid (must be > 0%)."
                            )
                        elif pct < 1:
                            warnings.append(
                                f"Size '{key}' is {pct}% which is extremely small (<1%)."
                            )
        except (ValueError, TypeError):
            pass

    if fill_color:
        try:
            from pptx.dml.color import RGBColor

            shape_rgb = ColorHelper.from_hex(fill_color)
            validation_results["fill_color_hex"] = fill_color

            if fill_opacity < 1.0:
                # RGBColor is a tuple-like (r, g, b)
                r, g, b = shape_rgb[0], shape_rgb[1], shape_rgb[2]
                effective_r = int(fill_opacity * r + (1 - fill_opacity) * 255)
                effective_g = int(fill_opacity * g + (1 - fill_opacity) * 255)
                effective_b = int(fill_opacity * b + (1 - fill_opacity) * 255)
                validation_results["effective_color_on_white"] = {
                    "hex": f"#{effective_r:02X}{effective_g:02X}{effective_b:02X}"
                }
        except Exception as e:
            validation_results["color_validation_error"] = str(e)

    if text and fill_opacity < 0.5:
        warnings.append(
            f"Shape has text but fill opacity is only {fill_opacity}. "
            "Text may be hard to read against varied backgrounds."
        )

    if is_overlay:
        validation_results["is_overlay"] = True

        if fill_opacity > 0.3:
            warnings.append(
                f"Overlay opacity {fill_opacity} is relatively high (>30%). "
                "System prompt recommends 0.15 for subtle overlays."
            )

        recommendations.append(
            "IMPORTANT: After adding this overlay, run ppt_set_z_order.py "
            "--action send_to_back to ensure overlay appears behind content."
        )

    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results,
        "has_warnings": len(warnings) > 0,
    }


def add_shape(
    filepath: Path,
    slide_index: int,
    shape_type: str,
    position: Dict[str, Any],
    size: Dict[str, Any],
    fill_color: Optional[str] = None,
    fill_opacity: float = 1.0,
    line_color: Optional[str] = None,
    line_opacity: float = 1.0,
    line_width: float = 1.0,
    text: Optional[str] = None,
    allow_offslide: bool = False,
    is_overlay: bool = False,
) -> Dict[str, Any]:
    """
    Add shape to slide with comprehensive validation and opacity support.

    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        shape_type: Type of shape to add
        position: Position specification dict
        size: Size specification dict
        fill_color: Fill color (hex or preset name)
        fill_opacity: Fill opacity (0.0=transparent to 1.0=opaque)
        line_color: Line/border color (hex or preset name)
        line_opacity: Line/border opacity (0.0=transparent to 1.0=opaque)
        line_width: Line width in points
        text: Optional text to add inside shape
        allow_offslide: Allow positioning outside slide bounds
        is_overlay: Whether this is an overlay shape

    Returns:
        Result dict with shape details and validation info

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid or opacity out of range
        SlideNotFoundError: If slide index is invalid
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    if filepath.suffix.lower() != ".pptx":
        raise ValueError("Only .pptx files are supported")

    resolved_shape = resolve_shape_type(shape_type)
    resolved_fill = resolve_color(fill_color)
    resolved_line = resolve_color(line_color)

    if is_overlay:
        if not position:
            position = OVERLAY_DEFAULTS["position"].copy()
        if not size:
            size = OVERLAY_DEFAULTS["size"].copy()
        if fill_opacity == 1.0:
            fill_opacity = OVERLAY_DEFAULTS["fill_opacity"]

    validation = validate_shape_params(
        position=position,
        size=size,
        fill_color=resolved_fill,
        fill_opacity=fill_opacity,
        line_color=resolved_line,
        line_opacity=line_opacity,
        text=text,
        allow_offslide=allow_offslide,
        is_overlay=is_overlay,
    )

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)

        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides},
            )

        version_before = agent.get_presentation_version()

        add_result = agent.add_shape(
            slide_index=slide_index,
            shape_type=resolved_shape,
            position=position,
            size=size,
            fill_color=resolved_fill,
            fill_opacity=fill_opacity,
            line_color=resolved_line,
            line_opacity=line_opacity,
            line_width=line_width,
            text=text,
        )

        agent.save()

        version_after = agent.get_presentation_version()

    shape_index = (
        add_result.get("shape_index") if isinstance(add_result, dict) else add_result
    )

    result: Dict[str, Any] = {
        "status": "success" if not validation["has_warnings"] else "warning",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_type": resolved_shape,
        "shape_type_requested": shape_type,
        "shape_index": shape_index,
        "position": add_result.get("position", position)
        if isinstance(add_result, dict)
        else position,
        "size": add_result.get("size", size) if isinstance(add_result, dict) else size,
        "styling": {
            "fill_color": resolved_fill,
            "fill_opacity": fill_opacity,
            "fill_transparency": round(1.0 - fill_opacity, 2),
            "line_color": resolved_line,
            "line_opacity": line_opacity,
            "line_width": line_width,
        },
        "text": text,
        "is_overlay": is_overlay,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__,
    }

    if validation["validation_results"]:
        result["validation"] = validation["validation_results"]

    if validation["warnings"]:
        result["warnings"] = validation["warnings"]

    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]

    notes = [
        "Shape added to top of z-order (in front of existing shapes).",
        f"Shape index {shape_index} may change if other shapes are added/removed.",
    ]

    if is_overlay or fill_opacity < 1.0:
        notes.insert(
            1,
            "Use ppt_set_z_order.py --action send_to_back to move overlay behind content.",
        )

    result["notes"] = notes

    if is_overlay:
        result["next_step"] = {
            "command": "ppt_set_z_order.py",
            "args": {
                "--file": str(filepath.resolve()),
                "--slide": slide_index,
                "--shape": shape_index,
                "--action": "send_to_back",
            },
            "description": "Send overlay to back so it appears behind content",
        }

    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add shape to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
AVAILABLE SHAPES:
  Basic:        rectangle, rounded_rectangle, ellipse/oval, triangle, diamond
  Arrows:       arrow_right, arrow_left, arrow_up, arrow_down
  Polygons:     pentagon, hexagon
  Decorative:   star, heart, lightning, sun, moon, cloud

SHAPE ALIASES:
  rect -> rectangle, circle -> ellipse, arrow -> arrow_right

OPACITY/TRANSPARENCY:
  --fill-opacity 1.0    Fully opaque (default)
  --fill-opacity 0.5    50% transparent
  --fill-opacity 0.15   85% transparent (subtle overlay, recommended)
  --fill-opacity 0.0    Fully transparent (invisible)

OVERLAY MODE (--overlay):
  Quick preset for creating background overlays:
  - Full-slide position and size
  - 15% opacity (subtle, non-competing)
  - Reminder to use ppt_set_z_order.py after

COLOR PRESETS:
  primary (#0070C0)    secondary (#595959)    accent (#ED7D31)
  success (#70AD47)    warning (#FFC000)      danger (#C00000)
  white (#FFFFFF)      black (#000000)

EXAMPLES:

  # Semi-transparent callout box
  uv run tools/ppt_add_shape.py --file deck.pptx --slide 0 --shape rounded_rectangle \\
    --position '{"left":"10%","top":"15%"}' --size '{"width":"30%","height":"15%"}' \\
    --fill-color primary --fill-opacity 0.8 --text "Key Point" --json

  # Subtle white overlay for text readability
  uv run tools/ppt_add_shape.py --file deck.pptx --slide 2 --shape rectangle \\
    --position '{"left":"0%","top":"0%"}' --size '{"width":"100%","height":"100%"}' \\
    --fill-color "#FFFFFF" --fill-opacity 0.15 --json

  # Quick overlay using --overlay preset
  uv run tools/ppt_add_shape.py --file deck.pptx --slide 3 --shape rectangle \\
    --overlay --fill-color black --json

Z-ORDER (LAYERING):
  Shapes are added on TOP of existing shapes by default.
  For overlays, you MUST send them to back:
    1. Add the overlay shape
    2. Note the shape_index from the output
    3. Run: ppt_set_z_order.py --file FILE --slide N --shape INDEX --action send_to_back
        """,
    )

    parser.add_argument(
        "--file", required=True, type=Path, help="PowerPoint file path (.pptx)"
    )
    parser.add_argument(
        "--slide", required=True, type=int, help="Slide index (0-based)"
    )
    parser.add_argument("--shape", required=True, help="Shape type")
    parser.add_argument(
        "--position", type=json.loads, default={}, help="Position dict as JSON"
    )
    parser.add_argument("--size", type=json.loads, help="Size dict as JSON")
    parser.add_argument("--fill-color", help="Fill color: hex or preset name")
    parser.add_argument(
        "--fill-opacity", type=float, default=1.0, help="Fill opacity (0.0-1.0)"
    )
    parser.add_argument("--line-color", help="Line/border color")
    parser.add_argument(
        "--line-opacity", type=float, default=1.0, help="Line opacity (0.0-1.0)"
    )
    parser.add_argument(
        "--line-width", type=float, default=1.0, help="Line width in points"
    )
    parser.add_argument("--text", help="Text to add inside shape")
    parser.add_argument(
        "--overlay", action="store_true", help="Overlay preset: full-slide, 15% opacity"
    )
    parser.add_argument(
        "--allow-offslide", action="store_true", help="Allow off-slide positioning"
    )
    parser.add_argument(
        "--json", action="store_true", default=True, help="Output JSON (default: true)"
    )

    args = parser.parse_args()

    try:
        size = args.size if args.size else {}
        position = args.position if args.position else {}

        if "width" in position and "width" not in size:
            size["width"] = position.pop("width")
        if "height" in position and "height" not in size:
            size["height"] = position.pop("height")

        if args.overlay:
            if "left" not in position:
                position["left"] = "0%"
            if "top" not in position:
                position["top"] = "0%"
            if "width" not in size:
                size["width"] = "100%"
            if "height" not in size:
                size["height"] = "100%"
        else:
            if "width" not in size:
                size["width"] = "20%"
            if "height" not in size:
                size["height"] = "20%"

        result = add_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_type=args.shape,
            position=position,
            size=size,
            fill_color=args.fill_color,
            fill_opacity=args.fill_opacity,
            line_color=args.line_color,
            line_opacity=args.line_opacity,
            line_width=args.line_width,
            text=args.text,
            allow_offslide=args.allow_offslide,
            is_overlay=args.overlay,
        )

        print(json.dumps(result, indent=2))
        sys.exit(0)

    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible.",
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, "details", {}),
            "suggestion": "Use ppt_get_info.py to check available slides.",
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check opacity values are between 0.0 and 1.0, and file is .pptx format.",
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON: {e}",
            "error_type": "JSONDecodeError",
            "suggestion": "Ensure --position and --size are valid JSON strings.",
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check the presentation file is valid.",
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__,
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)


if __name__ == "__main__":
    main()
