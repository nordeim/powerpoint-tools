#!/usr/bin/env python3
"""
PowerPoint Set Slide Layout Tool v3.1.0
Change the layout of an existing slide with safety warnings

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

⚠️ IMPORTANT WARNING:
    Changing slide layouts can cause CONTENT LOSS!
    - Text in removed placeholders may disappear
    - Shapes may be repositioned
    - This is a python-pptx limitation
    
    ALWAYS backup your presentation before changing layouts!

Usage:
    uv run tools/ppt_set_slide_layout.py --file presentation.pptx --slide 2 --layout "Title Only" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Safety:
    The --force flag is required for layouts that may cause content loss
    (e.g., "Blank", "Title Only")
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
from difflib import get_close_matches

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exception
try:
    from core.powerpoint_agent_core import LayoutNotFoundError
except ImportError:
    class LayoutNotFoundError(PowerPointAgentError):
        """Exception raised when layout is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)

# Layouts known to potentially cause content loss
DESTRUCTIVE_LAYOUTS = ["Blank", "Title Only"]


def set_slide_layout(
    filepath: Path,
    slide_index: int,
    layout_name: str,
    force: bool = False
) -> Dict[str, Any]:
    """
    Change slide layout with safety warnings.
    
    ⚠️ WARNING: Changing layouts can cause content loss due to python-pptx
    limitations. Always backup presentations before layout changes.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Slide index (0-based)
        layout_name: Target layout name (fuzzy matching supported)
        force: Acknowledge content loss risk (required for destructive layouts)
        
    Returns:
        Dict containing:
            - status: "success" or "warning"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - old_layout: Previous layout name
            - new_layout: New layout name
            - layout_changed: Whether layout actually changed
            - placeholders: Before/after/change counts
            - available_layouts: All available layouts
            - warnings: Content loss warnings (if any)
            - recommendations: Suggested actions (if any)
            - presentation_version_before: State hash before change
            - presentation_version_after: State hash after change
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        LayoutNotFoundError: If layout is not found
        PowerPointAgentError: If force required but not provided
        
    Example:
        >>> result = set_slide_layout(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=2,
        ...     layout_name="Section Header"
        ... )
        >>> print(result["new_layout"])
        'Section Header'
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    warnings: List[str] = []
    recommendations: List[str] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE change
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
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
        
        # Get available layouts
        available_layouts = agent.get_available_layouts()
        
        # Get current slide info
        slide_info_before = agent.get_slide_info(slide_index)
        old_layout = slide_info_before.get("layout", "Unknown")
        placeholders_before = sum(
            1 for shape in slide_info_before.get("shapes", [])
            if "PLACEHOLDER" in shape.get("type", "")
        )
        
        # Layout name matching with fuzzy search
        matched_layout: Optional[str] = None
        
        # Exact match (case-insensitive)
        for layout in available_layouts:
            if layout.lower() == layout_name.lower():
                matched_layout = layout
                break
        
        # Substring match if no exact match
        if not matched_layout:
            for layout in available_layouts:
                if layout_name.lower() in layout.lower():
                    matched_layout = layout
                    warnings.append(
                        f"Matched '{layout_name}' to layout '{layout}' (substring match)"
                    )
                    break
        
        # Fuzzy match using difflib
        if not matched_layout:
            close_matches = get_close_matches(
                layout_name, available_layouts, n=3, cutoff=0.6
            )
            if close_matches:
                raise LayoutNotFoundError(
                    f"Layout '{layout_name}' not found. Did you mean one of these?\n" +
                    "\n".join(f"  - {match}" for match in close_matches) +
                    f"\n\nAll available layouts:\n" +
                    "\n".join(f"  - {layout}" for layout in available_layouts),
                    details={
                        "requested_layout": layout_name,
                        "suggestions": close_matches,
                        "available_layouts": available_layouts
                    }
                )
            else:
                raise LayoutNotFoundError(
                    f"Layout '{layout_name}' not found.\n\n" +
                    f"Available layouts:\n" +
                    "\n".join(f"  - {layout}" for layout in available_layouts),
                    details={
                        "requested_layout": layout_name,
                        "available_layouts": available_layouts
                    }
                )
        
        # Safety warnings for destructive layouts
        if matched_layout in DESTRUCTIVE_LAYOUTS and placeholders_before > 0:
            warnings.append(
                f"⚠️ CONTENT LOSS RISK: Changing from '{old_layout}' to '{matched_layout}' "
                f"may remove {placeholders_before} placeholder(s) and their content!"
            )
            
            if not force:
                raise PowerPointAgentError(
                    f"Layout change from '{old_layout}' to '{matched_layout}' requires --force flag.\n"
                    f"This change may cause content loss ({placeholders_before} placeholders affected).\n\n"
                    "To proceed, add --force flag:\n"
                    f"  --layout \"{matched_layout}\" --force\n\n"
                    "RECOMMENDATION: Backup your presentation first!"
                )
        
        # Warn about same layout
        if matched_layout == old_layout:
            recommendations.append(
                f"Slide already uses '{old_layout}' layout. No change needed."
            )
        
        # Apply layout change
        agent.set_slide_layout(slide_index, matched_layout)
        
        # Get slide info after change
        slide_info_after = agent.get_slide_info(slide_index)
        placeholders_after = sum(
            1 for shape in slide_info_after.get("shapes", [])
            if "PLACEHOLDER" in shape.get("type", "")
        )
        
        # Detect content loss
        if placeholders_after < placeholders_before:
            lost_count = placeholders_before - placeholders_after
            warnings.append(
                f"Content loss detected: {lost_count} placeholder(s) removed during layout change."
            )
            recommendations.append(
                "Review slide content and restore any lost text using ppt_add_text_box.py"
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER change
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Build response
    status = "success" if len(warnings) == 0 else "warning"
    
    result: Dict[str, Any] = {
        "status": status,
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "old_layout": old_layout,
        "new_layout": matched_layout,
        "layout_changed": (old_layout != matched_layout),
        "placeholders": {
            "before": placeholders_before,
            "after": placeholders_after,
            "change": placeholders_after - placeholders_before
        },
        "available_layouts": available_layouts,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if warnings:
        result["warnings"] = warnings
    
    if recommendations:
        result["recommendations"] = recommendations
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Change PowerPoint slide layout with safety warnings",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ IMPORTANT WARNING ⚠️
    Changing slide layouts can cause CONTENT LOSS!
    - Text in removed placeholders may disappear
    - Shapes may be repositioned
    
    ALWAYS backup your presentation before changing layouts!

Examples:
  # List available layouts first
  uv run tools/ppt_get_info.py --file presentation.pptx --json | jq '.layouts'
  
  # Change to Title Only layout (low risk)
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --layout "Title Only" \\
    --json
  
  # Change to Blank layout (HIGH RISK - requires --force)
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --layout "Blank" \\
    --force \\
    --json
  
  # Fuzzy matching (will match "Title and Content")
  uv run tools/ppt_set_slide_layout.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --layout "title content" \\
    --json

Common Layouts:
  Low Risk (preserve most content):
  - "Title and Content" - Most versatile
  - "Two Content" - Side-by-side content
  - "Section Header" - Section dividers
  
  Medium Risk:
  - "Title Only" - Removes content placeholders
  - "Content with Caption" - Repositions content
  
  High Risk (requires --force):
  - "Blank" - Removes ALL placeholders!

Layout Matching:
  This tool supports flexible matching:
  - Exact: "Title and Content" matches "Title and Content"
  - Case-insensitive: "title slide" matches "Title Slide"
  - Substring: "content" matches "Title and Content"
  - Fuzzy: "tile slide" suggests "Title Slide"

Safety Features:
  - Warns about content loss risk
  - Requires --force for destructive layouts
  - Reports placeholder count changes
  - Suggests recovery actions

Output Format:
  {
    "status": "warning",
    "slide_index": 2,
    "old_layout": "Title and Content",
    "new_layout": "Title Only",
    "layout_changed": true,
    "placeholders": {
      "before": 2,
      "after": 1,
      "change": -1
    },
    "warnings": ["Content loss detected..."],
    "recommendations": ["Review slide content..."],
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.0"
  }

Recovery from Content Loss:
  If content was lost during layout change:
  1. Restore from backup (you did backup, right?)
  2. Use ppt_get_slide_info.py to inspect current state
  3. Restore text with ppt_add_text_box.py

Related Tools:
  - ppt_get_info.py: List all available layouts
  - ppt_get_slide_info.py: Inspect current slide layout
  - ppt_add_text_box.py: Restore lost content
  - ppt_clone_presentation.py: Create backup before changes
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
        '--layout',
        required=True,
        help='New layout name (fuzzy matching supported)'
    )
    
    parser.add_argument(
        '--force',
        action='store_true',
        help='Force destructive layout change (acknowledges content loss risk)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_slide_layout(
            filepath=args.file,
            slide_index=args.slide,
            layout_name=args.layout,
            force=args.force
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
            "suggestion": "Use ppt_get_info.py to check available slides"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except LayoutNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "LayoutNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to list available layouts"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Add --force flag if you accept the content loss risk"
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
