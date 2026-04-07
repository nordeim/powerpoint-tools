#!/usr/bin/env python3
"""
PowerPoint Set Shape Text Tool v3.1.1
Update text content of an existing shape or text box.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_set_shape_text.py --file deck.pptx --slide 0 --shape 2 --text "New content" --json
    uv run tools/ppt_set_shape_text.py --file deck.pptx --slide 0 --shape 2 --text "Line 1\nLine 2" --json

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
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
)

__version__ = "3.1.1"


def set_shape_text(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    text: str,
) -> Dict[str, Any]:
    """
    Replace all text in a shape with new content.

    Args:
        filepath: Path to PowerPoint file
        slide_index: Slide index (0-based)
        shape_index: Shape index (0-based)
        text: New text content (supports \\n for line breaks)

    Returns:
        Dict with shape details after text update
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

        if not shape.has_text_frame:
            raise PowerPointAgentError(
                f"Shape '{shape.name}' does not have a text frame",
                details={"shape_type": shape.shape_type},
            )

        # Clear existing paragraphs and set new text
        text_frame = shape.text_frame
        text_frame.clear()
        text_frame.word_wrap = True

        lines = text.split("\n")
        for i, line in enumerate(lines):
            if i == 0:
                text_frame.text = line
            else:
                p = text_frame.add_paragraph()
                p.text = line

        agent.save()
        version_after = agent.get_presentation_version()

    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "shape_name": shape.name,
        "text_preview": text[:100] + ("..." if len(text) > 100 else ""),
        "text_length": len(text),
        "line_count": len(lines),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__,
    }


def main():
    parser = argparse.ArgumentParser(
        description="Update text content of an existing shape or text box",
        epilog="""
Examples:

  # Set simple text
  uv run tools/ppt_set_shape_text.py --file deck.pptx --slide 0 --shape 2 --text "New content" --json

  # Set multi-line text
  uv run tools/ppt_set_shape_text.py --file deck.pptx --slide 0 --shape 2 --text "Line 1\nLine 2\nLine 3" --json
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
        "--text", required=True, help="New text content (use \\n for line breaks)"
    )
    parser.add_argument(
        "--json", action="store_true", default=True, help="Output JSON (default: true)"
    )

    args = parser.parse_args()

    try:
        result = set_shape_text(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            text=args.text,
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
            "suggestion": "Use ppt_get_slide_info.py to refresh shape indices.",
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
            "suggestion": "Ensure the target shape has a text frame (text box, placeholder, or shape with text).",
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
