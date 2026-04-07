#!/usr/bin/env python3
"""
PowerPoint Create From Template Tool v3.1.1
Create new presentation from existing .pptx template.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_create_from_template.py --template corporate_template.pptx --output new_presentation.pptx --slides 10 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

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
from typing import Dict, Any, List

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

def create_from_template(
    template: Path,
    output: Path,
    slides: int = 1,
    layout: str = "Title and Content"
) -> Dict[str, Any]:
    """
    Create a new PowerPoint presentation from an existing template.
    
    This tool copies the template (including its theme, master slides, and
    any existing content) and optionally adds additional slides using the
    specified layout.
    
    Args:
        template: Path to the source template .pptx file
        output: Path where the new presentation will be saved
        slides: Total number of slides desired in the output (default: 1)
        layout: Layout name for additional slides (default: "Title and Content")
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to created file
            - template_used: Path to source template
            - total_slides: Final slide count
            - slides_requested: Number of slides requested
            - template_slides: Number of slides in original template
            - slides_added: Number of slides added
            - layout_used: Layout name used for added slides
            - available_layouts: List of all available layouts
            - file_size_bytes: Size of created file
            - slide_dimensions: Width, height, and aspect ratio
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If template file does not exist
        ValueError: If template is not .pptx or slide count invalid
        LayoutNotFoundError: If specified layout not found (falls back)
        
    Example:
        >>> result = create_from_template(
        ...     template=Path("templates/corporate.pptx"),
        ...     output=Path("q4_report.pptx"),
        ...     slides=15,
        ...     layout="Title and Content"
        ... )
        >>> print(result["total_slides"])
        15
    """
    if not template.exists():
        raise FileNotFoundError(f"Template file not found: {template}")
    
    if not template.suffix.lower() == '.pptx':
        raise ValueError(f"Template must be .pptx file, got: {template.suffix}")
    
    if slides < 1:
        raise ValueError("Must create at least 1 slide")
    
    if slides > 100:
        raise ValueError("Maximum 100 slides per creation (performance limit)")
    
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
        
        current_slides = agent.get_slide_count()
        
        slides_to_add = max(0, slides - current_slides)
        
        slide_indices: List[int] = list(range(current_slides))
        
        for i in range(slides_to_add):
            result = agent.add_slide(layout_name=resolved_layout)
            if isinstance(result, dict):
                idx = result.get("slide_index", result.get("index", len(slide_indices)))
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
        "template_used": str(template.resolve()),
        "total_slides": info["slide_count"],
        "slides_requested": slides,
        "template_slides": current_slides,
        "slides_added": slides_to_add,
        "layout_used": resolved_layout,
        "available_layouts": info.get("layouts", available_layouts),
        "file_size_bytes": file_size,
        "slide_dimensions": {
            "width_inches": info.get("slide_width_inches", 13.333),
            "height_inches": info.get("slide_height_inches", 7.5),
            "aspect_ratio": info.get("aspect_ratio", "16:9")
        },
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Create PowerPoint presentation from template",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Create from corporate template with 15 slides
  uv run tools/ppt_create_from_template.py \\
    --template templates/corporate.pptx \\
    --output q4_report.pptx \\
    --slides 15 \\
    --json
  
  # Create presentation using specific layout
  uv run tools/ppt_create_from_template.py \\
    --template templates/minimal.pptx \\
    --output demo.pptx \\
    --slides 5 \\
    --layout "Section Header" \\
    --json
  
  # Quick presentation from template (uses template's existing slides)
  uv run tools/ppt_create_from_template.py \\
    --template templates/branded.pptx \\
    --output quick_deck.pptx \\
    --json

Use Cases:
  - Corporate presentations with consistent branding
  - Team presentations with shared theme
  - Pre-formatted layouts (fonts, colors, logos)
  - Department-specific templates
  - Client-specific branded decks

Template Benefits:
  - Consistent branding across organization
  - Pre-configured master slides
  - Corporate colors and fonts
  - Logo placements
  - Standard layouts
  - Accessibility features built-in

Creating Templates:
  1. Design in PowerPoint with desired theme
  2. Configure master slides
  3. Set up color scheme
  4. Define standard layouts
  5. Save as .pptx template
  6. Use with this tool

Best Practices:
  - Maintain template library for different purposes
  - Version control templates
  - Document template usage guidelines
  - Test templates before distribution
  - Include variety of layouts in template

Output Format:
  {
    "status": "success",
    "file": "/path/to/output.pptx",
    "template_used": "/path/to/template.pptx",
    "total_slides": 15,
    "template_slides": 1,
    "slides_added": 14,
    "layout_used": "Title and Content",
    "presentation_version": "a1b2c3d4...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--template',
        required=True,
        type=Path,
        help='Path to template .pptx file'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output presentation path'
    )
    
    parser.add_argument(
        '--slides',
        type=int,
        default=1,
        help='Total number of slides desired (default: 1)'
    )
    
    parser.add_argument(
        '--layout',
        default='Title and Content',
        help='Layout for additional slides (default: "Title and Content")'
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
        
        result = create_from_template(
            template=args.template.resolve(),
            output=output_path.resolve(),
            slides=args.slides,
            layout=args.layout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            sys.stdout.write(f"Created presentation from template: {result['file']}\n")
            sys.stdout.write(f"  Template: {result['template_used']}\n")
            sys.stdout.write(f"  Total slides: {result['total_slides']}\n")
            sys.stdout.write(f"  Template had: {result['template_slides']} slides\n")
            sys.stdout.write(f"  Added: {result['slides_added']} slides\n")
            sys.stdout.write(f"  Layout: {result['layout_used']}\n")
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
            "suggestion": "Check that template is .pptx and slide count is 1-100",
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
            "suggestion": "Check template file integrity and available layouts",
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
