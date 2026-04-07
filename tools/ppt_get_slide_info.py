#!/usr/bin/env python3
"""
PowerPoint Get Slide Info Tool v3.1.0
Get detailed information about slide content (shapes, images, text, positions)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Full text content (no truncation)
    - Position information (inches and percentages)
    - Size information (inches and percentages)
    - Human-readable placeholder type names
    - Notes detection

Usage:
    uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 0 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Finding shape indices for ppt_format_text.py
    - Locating images for ppt_replace_image.py
    - Debugging positioning issues
    - Auditing slide content
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
from pathlib import Path
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"


def get_slide_info(
    filepath: Path,
    slide_index: int
) -> Dict[str, Any]:
    """
    Get detailed slide information including full text and positioning.
    
    This is a read-only operation that does not modify the file.
    It acquires no lock, allowing concurrent reads.
    
    Args:
        filepath: Path to the PowerPoint file
        slide_index: Slide index (0-based)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to file
            - slide_index: Index of the slide
            - layout: Layout name
            - shape_count: Total number of shapes
            - shapes: List of shape information dicts
            - has_notes: Whether slide has speaker notes
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Each shape dict contains:
        - index: Shape index (for targeting with other tools)
        - type: Shape type (with human-readable placeholder names)
        - name: Shape name
        - has_text: Boolean
        - text: Full text content (no truncation)
        - text_length: Character count
        - text_preview: First 100 chars (if text > 100 chars)
        - position: Dict with inches and percentages
        - size: Dict with inches and percentages
        - is_placeholder: Boolean
        - placeholder_type: Human-readable type (if placeholder)
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        
    Example:
        >>> result = get_slide_info(Path("presentation.pptx"), 0)
        >>> print(result["shape_count"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        # Open without acquiring lock (read-only operation)
        agent.open(filepath, acquire_lock=False)
        
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
        
        # Get enhanced slide info from core
        slide_info = agent.get_slide_info(slide_index)
        
        # Get presentation version
        prs_info = agent.get_presentation_info()
        presentation_version = prs_info.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_info.get("slide_index", slide_index),
        "layout": slide_info.get("layout", "Unknown"),
        "shape_count": slide_info.get("shape_count", 0),
        "shapes": slide_info.get("shapes", []),
        "has_notes": slide_info.get("has_notes", False),
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint slide information with full text and positioning",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get info for first slide
  uv run tools/ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --json
  
  # Get info for specific slide
  uv run tools/ppt_get_slide_info.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --json
  
  # Find text shapes
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json | \\
    jq '.shapes[] | select(.has_text == true)'
  
  # Find footer elements (shapes at bottom)
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json | \\
    jq '.shapes[] | select(.type | contains("FOOTER"))'

Output Information:
  - Slide layout name
  - Total shape count
  - Presentation version (for change tracking)
  - List of all shapes with:
    - Shape index (for targeting with other tools)
    - Shape type with human-readable placeholder names
    - Shape name
    - Whether it contains text
    - FULL text content (no truncation)
    - Position in inches and percentages
    - Size in inches and percentages

Use Cases:
  - Find shape indices for ppt_format_text.py
  - Locate images for ppt_replace_image.py
  - Inspect slide layout and structure
  - Audit slide content
  - Debug positioning issues
  - Verify footer/header presence

Finding Shape Indices:
  Use this tool before:
  - ppt_format_text.py (needs shape index)
  - ppt_replace_image.py (needs image name)
  - ppt_format_shape.py (needs shape index)
  - ppt_set_image_properties.py (needs shape index)
  - ppt_crop_image.py (needs shape index)

Example Output:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "layout": "Title Slide",
    "shape_count": 5,
    "shapes": [
      {
        "index": 0,
        "type": "PLACEHOLDER (TITLE)",
        "name": "Title 1",
        "has_text": true,
        "text": "My Presentation Title",
        "text_length": 21,
        "position": {
          "left_inches": 0.5,
          "top_inches": 1.0,
          "left_percent": "5.0%",
          "top_percent": "13.3%"
        },
        "size": {
          "width_inches": 9.0,
          "height_inches": 1.5,
          "width_percent": "90.0%",
          "height_percent": "20.0%"
        }
      }
    ],
    "has_notes": false,
    "presentation_version": "a1b2c3d4e5f6g7h8",
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
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = get_slide_info(
            filepath=args.file,
            slide_index=args.slide
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
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
