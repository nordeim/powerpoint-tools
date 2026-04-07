#!/usr/bin/env python3
"""
PowerPoint Create New Tool v3.1.1
Create a new PowerPoint presentation with specified slides.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_create_new.py --output presentation.pptx --slides 5 --layout "Title and Content" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Note:
    For creating presentations from existing templates with branding,
    consider using ppt_create_from_template.py instead.

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
    - Added get_available_layouts() fallback for compatibility
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
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError,
    LayoutNotFoundError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# MAIN LOGIC
# ============================================================================

def create_new_presentation(
    output: Path,
    slides: int,
    template: Optional[Path] = None,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation with specified number of slides.
    
    Creates a blank presentation (or from optional template) and populates
    it with the requested number of slides. The first slide uses "Title Slide"
    layout if available, subsequent slides use the specified layout.
    
    Args:
        output: Path where the new presentation will be saved
        slides: Number of slides to create (1-100)
        template: Optional path to template .pptx file (default: None for blank)
        layout: Layout name for slides after the first (default: "Title and Content")
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to created file
            - slides_created: Number of slides created
            - slide_indices: List of slide indices
            - file_size_bytes: Size of created file
            - slide_dimensions: Width, height, and aspect ratio
            - available_layouts: List of all available layouts
            - layout_used: Layout name used for non-title slides
            - template_used: Path to template if used, else None
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Raises:
        ValueError: If slide count is invalid (not 1-100)
        FileNotFoundError: If template specified but not found
        
    Example:
        >>> result = create_new_presentation(
        ...     output=Path("pitch_deck.pptx"),
        ...     slides=10,
        ...     layout="Title and Content"
        ... )
        >>> print(result["slides_created"])
        10
    """
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
    if template is not None:
        if not template.exists():
            raise FileNotFoundError(f"Template file not found: {template}")
        if not template.suffix.lower() == '.pptx':
            raise ValueError(f"Template must be .pptx file, got: {template.suffix}")
    
    with PowerPointAgent() as agent:
        agent.create_new(template=template)
        
        try:
            available_layouts = agent.get_available_layouts()
        except AttributeError:
            info = agent.get_presentation_info()
            available_layouts = info.get("layouts", [])
        
        resolved_layout = layout
        if layout not in available_layouts:
            layout_lower = layout.lower()
            matched = False
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    resolved_layout = avail
                    matched = True
                    break
            
            if not matched:
                resolved_layout = available_layouts[0] if available_layouts else "Title Slide"
        
        slide_indices: List[int] = []
        
        for i in range(slides):
            if i == 0 and "Title Slide" in available_layouts:
                slide_layout = "Title Slide"
            else:
                slide_layout = resolved_layout
            
            result = agent.add_slide(layout_name=slide_layout)
            if isinstance(result, dict):
                idx = result.get("slide_index", result.get("index", i))
            else:
                idx = result
            slide_indices.append(idx)
        
        agent.save(output)
        
        info = agent.get_presentation_info()
        presentation_version = info.get("presentation_version", None)
    
    file_size = output.stat().st_size if output.exists() else 0
    
    return {
        "status": "success",
        "file": str(output.resolve()),
        "slides_created": slides,
        "slide_indices": slide_indices,
        "file_size_bytes": file_size,
        "slide_dimensions": {
            "width_inches": info.get("slide_width_inches", 13.333),
            "height_inches": info.get("slide_height_inches", 7.5),
            "aspect_ratio": info.get("aspect_ratio", "16:9")
        },
        "available_layouts": info.get("layouts", available_layouts),
        "layout_used": resolved_layout,
        "template_used": str(template.resolve()) if template else None,
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Create new PowerPoint presentation with specified slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create presentation with 5 blank slides
  uv run tools/ppt_create_new.py --output presentation.pptx --slides 5 --json
  
  # Create with specific layout
  uv run tools/ppt_create_new.py --output pitch_deck.pptx --slides 10 --layout "Title and Content" --json
  
  # Create from template (for simple cases; use ppt_create_from_template.py for advanced)
  uv run tools/ppt_create_new.py --output new_deck.pptx --slides 3 --template corporate_template.pptx --json
  
  # Create single title slide
  uv run tools/ppt_create_new.py --output title.pptx --slides 1 --layout "Title Slide" --json

Available Layouts (typical):
  - Title Slide
  - Title and Content
  - Section Header
  - Two Content
  - Comparison
  - Title Only
  - Blank
  - Content with Caption
  - Picture with Caption

First Slide Behavior:
  The first slide automatically uses "Title Slide" layout if available,
  regardless of the --layout parameter. Subsequent slides use --layout.

For Template-Based Creation:
  If you need to preserve template content or work with branded templates,
  use ppt_create_from_template.py instead. This tool is optimized for
  creating presentations from scratch.

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slides_created": 5,
    "slide_indices": [0, 1, 2, 3, 4],
    "file_size_bytes": 28432,
    "slide_dimensions": {
      "width_inches": 13.333,
      "height_inches": 7.5,
      "aspect_ratio": "16:9"
    },
    "available_layouts": ["Title Slide", "Title and Content", ...],
    "layout_used": "Title and Content",
    "presentation_version": "a1b2c3d4...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Number of slides to create (default: 1)'
    )
    
    parser.add_argument(
        '--template',
        type=Path,
        default=None,
        help='Optional template file to use (.pptx)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout to use for slides after the first (default: "Title and Content")'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        output_path = args.output
        if not output_path.suffix.lower() == '.pptx':
            output_path = output_path.with_suffix('.pptx')
        
        result = create_new_presentation(
            output=output_path.resolve(),
            slides=args.slides,
            template=args.template.resolve() if args.template else None,
            layout=args.layout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            sys.stdout.write(f"Created presentation: {result['file']}\n")
            sys.stdout.write(f"  Slides: {result['slides_created']}\n")
            sys.stdout.write(f"  Layout: {result['layout_used']}\n")
            sys.stdout.write(f"  Dimensions: {result['slide_dimensions']['aspect_ratio']}\n")
            if args.template:
                sys.stdout.write(f"  Template: {result['template_used']}\n")
            sys.stdout.flush()
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the template file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check slide count (1-100) and template file extension (.pptx)",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except LayoutNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "LayoutNotFoundError",
            "suggestion": "Use ppt_capability_probe.py to discover available layouts",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Check file permissions and template compatibility",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
