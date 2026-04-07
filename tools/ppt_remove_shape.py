#!/usr/bin/env python3
"""
PowerPoint Remove Shape Tool v3.1.0
Safely remove shapes from slides with comprehensive safety controls.

⚠️  DESTRUCTIVE OPERATION WARNING ⚠️
- Shape removal CANNOT be undone
- Shape indices WILL shift after removal
- Always CLONE the presentation first
- Always use --dry-run to preview

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    # Preview removal (RECOMMENDED FIRST)
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 2 --dry-run --json

    # Execute removal
    uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 2 --json

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
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    ApprovalTokenError,
)

__version__ = "3.1.0"


def get_shape_details(
    agent: PowerPointAgent, slide_index: int, shape_index: int
) -> Dict[str, Any]:
    """Get detailed information about a shape before removal."""
    try:
        slide_info = agent.get_slide_info(slide_index)
        shapes = slide_info.get("shapes", [])

        if 0 <= shape_index < len(shapes):
            shape = shapes[shape_index]
            return {
                "index": shape_index,
                "type": shape.get("type", "unknown"),
                "name": shape.get("name", ""),
                "has_text": shape.get("has_text", False),
                "text_preview": (shape.get("text", "")[:100] + "...")
                if len(shape.get("text", "")) > 100
                else shape.get("text", ""),
                "position": shape.get("position", {}),
                "size": shape.get("size", {}),
            }
    except Exception as e:
        return {"index": shape_index, "error": str(e)}

    return {"index": shape_index, "type": "unknown"}


def find_shape_by_name(
    agent: PowerPointAgent, slide_index: int, name: str
) -> Optional[int]:
    """Find shape index by name (partial match)."""
    try:
        slide_info = agent.get_slide_info(slide_index)
        shapes = slide_info.get("shapes", [])

        for idx, shape in enumerate(shapes):
            if shape.get("name", "") == name:
                return idx

        name_lower = name.lower()
        for idx, shape in enumerate(shapes):
            shape_name = shape.get("name", "").lower()
            if name_lower in shape_name or shape_name in name_lower:
                return idx

        return None
    except Exception:
        return None


def remove_shape(
    filepath: Path,
    slide_index: int,
    shape_index: Optional[int] = None,
    shape_name: Optional[str] = None,
    dry_run: bool = False,
    approval_token: Optional[str] = None,
) -> Dict[str, Any]:
    """
    Remove shape from slide with safety controls.

    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        shape_index: Shape index to remove (0-based)
        shape_name: Shape name to remove (alternative to index)
        dry_run: If True, preview only without actual removal

    Returns:
        Result dict with removal details

    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If invalid parameters
        SlideNotFoundError: If slide index invalid
        ShapeNotFoundError: If shape not found
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")

    if filepath.suffix.lower() != ".pptx":
        raise ValueError("Only .pptx files are supported")

    if shape_index is None and shape_name is None:
        raise ValueError("Must specify either --shape (index) or --name (shape name)")

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)

        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides},
            )

        slide_info_before = agent.get_slide_info(slide_index)
        shape_count_before = slide_info_before.get("shape_count", 0)

        resolved_index = shape_index
        if shape_name is not None:
            resolved_index = find_shape_by_name(agent, slide_index, shape_name)
            if resolved_index is None:
                raise ShapeNotFoundError(
                    f"Shape with name '{shape_name}' not found on slide {slide_index}",
                    details={
                        "slide_index": slide_index,
                        "shape_name": shape_name,
                        "available_shapes": [
                            s.get("name") for s in slide_info_before.get("shapes", [])
                        ],
                    },
                )

        if not 0 <= resolved_index < shape_count_before:
            raise ShapeNotFoundError(
                f"Shape index {resolved_index} out of range (0-{shape_count_before - 1})",
                details={"requested": resolved_index, "available": shape_count_before},
            )

        shape_details = get_shape_details(agent, slide_index, resolved_index)
        version_before = agent.get_presentation_version()

        result: Dict[str, Any] = {
            "file": str(filepath.resolve()),
            "slide_index": slide_index,
            "shape_index": resolved_index,
            "shape_details": shape_details,
            "shape_count_before": shape_count_before,
            "dry_run": dry_run,
            "presentation_version_before": version_before,
            "tool_version": __version__,
        }

        if dry_run:
            result["status"] = "preview"
            result["message"] = (
                "DRY RUN: Shape would be removed. Run without --dry-run to execute."
            )
            result["shape_count_after"] = shape_count_before - 1
            shapes_affected = shape_count_before - resolved_index - 1
            result["index_shift_info"] = {
                "shapes_affected": shapes_affected,
                "message": f"Shapes at indices {resolved_index + 1} to {shape_count_before - 1} would shift down by 1"
                if shapes_affected > 0
                else "No other shapes would be affected",
            }
        else:
            agent.remove_shape(
                slide_index=slide_index,
                shape_index=resolved_index,
                approval_token=approval_token,
            )
            agent.save()

            version_after = agent.get_presentation_version()
            slide_info_after = agent.get_slide_info(slide_index)
            shape_count_after = slide_info_after.get("shape_count", 0)

            result["status"] = "success"
            result["message"] = "Shape removed successfully"
            result["shape_count_after"] = shape_count_after
            result["presentation_version_after"] = version_after

            shapes_shifted = shape_count_before - resolved_index - 1
            if shapes_shifted > 0:
                result["index_shift_info"] = {
                    "shapes_shifted": shapes_shifted,
                    "warning": f"⚠️ {shapes_shifted} shape(s) have new indices. Re-query before further operations.",
                    "refresh_command": f"uv run tools/ppt_get_slide_info.py --file {filepath} --slide {slide_index} --json",
                }

            result["rollback_guidance"] = (
                "This operation cannot be undone. Restore from backup clone."
            )

    return result


def main():
    parser = argparse.ArgumentParser(
        description="Remove shape from PowerPoint slide ⚠️ DESTRUCTIVE",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️  DESTRUCTIVE OPERATION - READ CAREFULLY ⚠️

This tool PERMANENTLY REMOVES shapes from presentations.
- Shape removal CANNOT be undone
- Shape indices WILL shift after removal
- Always CLONE the presentation first
- Always use --dry-run to preview

SAFE REMOVAL PROTOCOL:

  1. CLONE: ppt_clone_presentation.py --source original.pptx --output work.pptx
  2. INSPECT: ppt_get_slide_info.py --file work.pptx --slide 0 --json
  3. PREVIEW: ppt_remove_shape.py --file work.pptx --slide 0 --shape 2 --dry-run --json
  4. EXECUTE: ppt_remove_shape.py --file work.pptx --slide 0 --shape 2 --json
  5. REFRESH: ppt_get_slide_info.py --file work.pptx --slide 0 --json

EXAMPLES:

  # Preview removal (ALWAYS DO FIRST)
  uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 3 --dry-run --json

  # Remove by index
  uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 3 --json

  # Remove by name
  uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --name "Rectangle 1" --json
        """,
    )

    parser.add_argument(
        "--file", required=True, type=Path, help="PowerPoint file path (.pptx)"
    )
    parser.add_argument(
        "--slide", required=True, type=int, help="Slide index (0-based)"
    )

    shape_group = parser.add_mutually_exclusive_group(required=True)
    shape_group.add_argument(
        "--shape", type=int, help="Shape index to remove (0-based)"
    )
    shape_group.add_argument("--name", help="Shape name to remove")

    parser.add_argument(
        "--dry-run", action="store_true", help="Preview without executing"
    )
    parser.add_argument(
        "--json", action="store_true", default=True, help="Output JSON (default: true)"
    )
    parser.add_argument(
        "--approval-token",
        type=str,
        default=None,
        help="Approval token for destructive operation",
    )

    args = parser.parse_args()

    try:
        result = remove_shape(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            shape_name=args.name,
            dry_run=args.dry_run,
            approval_token=args.approval_token,
        )

        print(json.dumps(result, indent=2))
        sys.exit(0)

    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file path exists and is accessible.",
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

    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, "details", {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shapes.",
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Specify --shape INDEX or --name NAME, and ensure .pptx format.",
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

    except ApprovalTokenError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ApprovalTokenError",
            "details": getattr(e, "details", {}),
            "suggestion": "Generate approval token with scope 'shape:remove:<slide>:<shape>' and retry",
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(4)

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
