#!/usr/bin/env python3
"""
PowerPoint Reposition Shape Tool v3.1.1
Move and/or resize a shape on a slide.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_reposition_shape.py --file deck.pptx --slide 0 --shape 2 --position '{"left":1.0,"top":2.0}' --json
    uv run tools/ppt_reposition_shape.py --file deck.pptx --slide 0 --shape 2 --size '{"width":8.0,"height":4.0}' --json
    uv run tools/ppt_reposition_shape.py --file deck.pptx --slide 0 --shape 2 --position '{"left":1.0,"top":2.0}' --size '{"width":8.0,"height":4.0}' --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

# --- HYGIENE BLOCK START ---
sys.stderr = open(os.devnull, "w")
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
)

__version__ = "3.1.1"

EMU_PER_INCH = 914400


def reposition_shape(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    position: Optional[Dict[str, Any]] = None,
    size: Optional[Dict[str, Any]] = None,
) -> Dict[str, Any]:
    """
    Reposition and/or resize a shape on a slide.

    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        shape_index: Shape index (0-based)
        position: Dict with 'left' and/or 'top' in inches
        size: Dict with 'width' and/or 'height' in inches

    Returns:
        Dict with shape details after repositioning
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)

        version_before = agent.get_presentation_version()
        total_slides = agent.get_slide_count()

        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides},
            )

        slide = agent.prs.slides[slide_index]
        shape_count = len(slide.shapes)

        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={"requested": shape_index, "available": shape_count},
            )

        shape = slide.shapes[shape_index]
        old_left = shape.left / EMU_PER_INCH
        old_top = shape.top / EMU_PER_INCH
        old_width = shape.width / EMU_PER_INCH
        old_height = shape.height / EMU_PER_INCH

        if position:
            if "left" in position:
                shape.left = int(position["left"] * EMU_PER_INCH)
            if "top" in position:
                shape.top = int(position["top"] * EMU_PER_INCH)

        if size:
            if "width" in size:
                shape.width = int(size["width"] * EMU_PER_INCH)
            if "height" in size:
                shape.height = int(size["height"] * EMU_PER_INCH)

        agent.save()
        version_after = agent.get_presentation_version()

    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "shape_name": shape.name,
        "before": {
            "left": round(old_left, 2),
            "top": round(old_top, 2),
            "width": round(old_width, 2),
            "height": round(old_height, 2),
        },
        "after": {
            "left": round(shape.left / EMU_PER_INCH, 2),
            "top": round(shape.top / EMU_PER_INCH, 2),
            "width": round(shape.width / EMU_PER_INCH, 2),
            "height": round(shape.height / EMU_PER_INCH, 2),
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Reposition and/or resize a shape on a slide",
        epilog="""
Position and size use inches. Examples:

  # Move shape to (1.0", 2.0")
  uv run tools/ppt_reposition_shape.py --file deck.pptx --slide 0 --shape 2 --position '{"left":1.0,"top":2.0}' --json

  # Resize shape to 8"x4"
  uv run tools/ppt_reposition_shape.py --file deck.pptx --slide 0 --shape 2 --size '{"width":8.0,"height":4.0}' --json

  # Move and resize
  uv run tools/ppt_reposition_shape.py --file deck.pptx --slide 0 --shape 2 --position '{"left":1.0,"top":2.0}' --size '{"width":8.0,"height":4.0}' --json
        """,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )

    parser.add_argument(
        "--file", required=True, type=Path, help="PowerPoint file path (.pptx)"
    )
    parser.add_argument(
        "--slide", required=True, type=int, help="Slide index (0-based)"
    )
    parser.add_argument(
        "--shape", required=True, type=int, help="Shape index (0-based)"
    )
    parser.add_argument(
        "--position",
        type=json.loads,
        help='Position dict: {"left": inches, "top": inches}',
    )
    parser.add_argument(
        "--size", type=json.loads, help='Size dict: {"width": inches, "height": inches}'
    )
    parser.add_argument(
        "--json", action="store_true", default=True, help="Output JSON (default: true)"
    )

    args = parser.parse_args()

    if not args.position and not args.size:
        error_result = {
            "status": "error",
            "error": "Must specify at least one of --position or --size",
            "error_type": "ValueError",
            "suggestion": 'Provide --position \'{"left":1.0,"top":2.0}\' and/or --size \'{"width":8.0,"height":4.0}\'',
            "tool_version": __version__,
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(2)

    try:
        result = reposition_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            position=args.position,
            size=args.size,
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)

    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file path exists and is accessible.",
            "tool_version": __version__,
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except (SlideNotFoundError, ShapeNotFoundError) as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, "details", {}),
            "suggestion": "Use ppt_get_slide_info.py to refresh shape indices before targeting shapes.",
            "tool_version": __version__,
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, "details", {}),
            "suggestion": "Check source file integrity.",
            "tool_version": __version__,
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information.",
            "tool_version": __version__,
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(5)


if __name__ == "__main__":
    main()
