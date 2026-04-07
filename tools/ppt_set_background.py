#!/usr/bin/env python3
"""
PowerPoint Set Background Tool v3.1.0
Set slide background to a solid color or image.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --color "#FFFFFF" --json
    uv run tools/ppt_set_background.py --file deck.pptx --all-slides --color "#F5F5F5" --json
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --image background.jpg --json

Exit Codes:
    0: Success
    1: Error occurred
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ColorHelper,
)

__version__ = "3.1.0"


def set_background(
    filepath: Path,
    color: Optional[str] = None,
    image: Optional[Path] = None,
    slide_index: Optional[int] = None,
    all_slides: bool = False
) -> Dict[str, Any]:
    """
    Set slide background to a solid color or image.
    
    Args:
        filepath: Path to PowerPoint file (.pptx only)
        color: Hex color code (e.g., "#FFFFFF")
        image: Path to background image file
        slide_index: Specific slide index (0-based), or None
        all_slides: If True, apply to all slides
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute path to file
            - slides_affected: Number of slides modified
            - slide_indices: List of modified slide indices
            - type: 'color' or 'image'
            - value: The color code or image path used
            - presentation_version_before: Version hash before changes
            - presentation_version_after: Version hash after changes
            - tool_version: Tool version string
            - deprecated_default_used: True if defaulted to all slides (backward compat)
            
    Raises:
        FileNotFoundError: If file or image doesn't exist
        ValueError: If parameters are invalid or mutually exclusive
        SlideNotFoundError: If slide index is out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError(
            f"Invalid file format '{filepath.suffix}'. Only .pptx files are supported."
        )
    
    if not color and not image:
        raise ValueError("Must specify either --color or --image")
    
    if color and image:
        raise ValueError("Cannot specify both --color and --image; choose one")
    
    if slide_index is not None and all_slides:
        raise ValueError("Cannot specify both --slide and --all-slides; choose one")
    
    if color:
        color_clean = color.strip()
        if not color_clean.startswith('#'):
            color_clean = '#' + color_clean
        if len(color_clean) != 7:
            raise ValueError(
                f"Invalid color format '{color}'. Use hex format: #RRGGBB (e.g., #FFFFFF)"
            )
        try:
            int(color_clean[1:], 16)
        except ValueError:
            raise ValueError(
                f"Invalid color format '{color}'. Must contain valid hex characters."
            )
    
    if image and not image.exists():
        raise FileNotFoundError(f"Image file not found: {image}")
    
    deprecated_default_used = False
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        slide_count = agent.get_slide_count()
        
        if slide_count == 0:
            raise PowerPointAgentError("Presentation has no slides")
        
        if slide_index is not None:
            if not 0 <= slide_index < slide_count:
                raise SlideNotFoundError(
                    f"Slide index {slide_index} out of range (0-{slide_count - 1})",
                    details={"requested": slide_index, "available": slide_count}
                )
            target_indices = [slide_index]
        elif all_slides:
            target_indices = list(range(slide_count))
        else:
            target_indices = list(range(slide_count))
            deprecated_default_used = True
        
        for idx in target_indices:
            slide = agent.prs.slides[idx]
            bg = slide.background
            fill = bg.fill
            
            if color:
                fill.solid()
                fill.fore_color.rgb = ColorHelper.from_hex(color)
            elif image:
                fill.user_picture(str(image.resolve()))
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    result = {
        "status": "success",
        "file": str(filepath.resolve()),
        "slides_affected": len(target_indices),
        "slide_indices": target_indices,
        "type": "color" if color else "image",
        "value": color if color else str(image.resolve()),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if deprecated_default_used:
        result["deprecated_default_used"] = True
        result["deprecation_warning"] = (
            "Defaulting to all slides is deprecated. "
            "Future versions will require explicit --slide N or --all-slides flag."
        )
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Set PowerPoint slide background to a solid color or image",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Set single slide background to white
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --color "#FFFFFF" --json
    
    # Set all slides to light gray (explicit)
    uv run tools/ppt_set_background.py --file deck.pptx --all-slides --color "#F5F5F5" --json
    
    # Set background image on single slide
    uv run tools/ppt_set_background.py --file deck.pptx --slide 0 --image bg.jpg --json
    
    # Set background image on all slides
    uv run tools/ppt_set_background.py --file deck.pptx --all-slides --image pattern.png --json

Color Format:
    Use hex color codes: #RRGGBB
    Examples: #FFFFFF (white), #000000 (black), #0070C0 (blue)

Supported Image Formats:
    PNG, JPEG, GIF, BMP, TIFF

Notes:
    - Use --slide N for a single slide (0-based index)
    - Use --all-slides to apply to entire presentation
    - If neither is specified, defaults to all slides (deprecated behavior)
    - Cannot use both --color and --image simultaneously
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='Path to PowerPoint file (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        type=int,
        dest='slide_index',
        help='Slide index to modify (0-based)'
    )
    
    parser.add_argument(
        '--all-slides',
        action='store_true',
        dest='all_slides',
        help='Apply background to all slides'
    )
    
    parser.add_argument(
        '--color',
        type=str,
        help='Hex color code (e.g., #FFFFFF)'
    )
    
    parser.add_argument(
        '--image',
        type=Path,
        help='Path to background image file'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output as JSON (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_background(
            filepath=args.file,
            color=args.color,
            image=args.image,
            slide_index=args.slide_index,
            all_slides=args.all_slides
        )
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file and image paths exist and are accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide count."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check color format (#RRGGBB), ensure only one of --color or --image is used."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Verify the file is not corrupted and has at least one slide."
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
