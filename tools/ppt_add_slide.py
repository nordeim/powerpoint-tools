#!/usr/bin/env python3
"""
PowerPoint Add Slide Tool v3.1.0
Add new slide to existing presentation with specific layout

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Compatible with PowerPoint Agent Core v3.1.0 (Dictionary Returns)

Usage:
    uv run tools/ppt_add_slide.py --file presentation.pptx --layout "Title and Content" --index 2 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
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
from typing import Dict, Any, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError
)

__version__ = "3.1.0"

# Define fallback exceptions if not in core
try:
    from core.powerpoint_agent_core import LayoutNotFoundError
except ImportError:
    class LayoutNotFoundError(PowerPointAgentError):
        """Exception raised when layout is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def add_slide(
    filepath: Path,
    layout: str,
    index: Optional[int] = None,
    set_title: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add a new slide to a presentation.
    
    Handles the v3.1.0 Core API where add_slide returns a dictionary
    with slide_index and version information.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        layout: Layout name for the new slide (fuzzy matching supported)
        index: Position to insert slide (0-based, default: end of presentation)
        set_title: Optional title text to set on the new slide
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the new slide
            - layout: Actual layout name used
            - title_set: Title text if provided
            - title_set_success: Whether title was set successfully
            - total_slides: Total slide count after addition
            - slide_info: Shape count and notes info
            - presentation_version_before: State hash before addition
            - presentation_version_after: State hash after addition
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        LayoutNotFoundError: If layout is not found
        
    Example:
        >>> result = add_slide(
        ...     filepath=Path("presentation.pptx"),
        ...     layout="Title and Content",
        ...     set_title="Q4 Results"
        ... )
        >>> print(result["slide_index"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Get available layouts for validation
        available_layouts = agent.get_available_layouts()
        
        # Validate layout with fuzzy matching
        matched_layout = layout
        if layout not in available_layouts:
            layout_lower = layout.lower()
            match_found = False
            for avail in available_layouts:
                if layout_lower in avail.lower():
                    matched_layout = avail
                    match_found = True
                    break
            
            if not match_found:
                raise LayoutNotFoundError(
                    f"Layout '{layout}' not found. Available layouts: {available_layouts}",
                    details={
                        "requested_layout": layout,
                        "available_layouts": available_layouts
                    }
                )
        
        # Add slide (Core v3.1.0 returns a dict)
        add_result = agent.add_slide(layout_name=matched_layout, index=index)
        
        # Extract the integer index from the returned dictionary
        # Core v3.1.0 returns dict, older versions may return int
        if isinstance(add_result, dict):
            slide_index = add_result["slide_index"]
            version_before = add_result.get("presentation_version_before")
        else:
            slide_index = add_result
            version_before = None
        
        # Set title if provided
        title_set_result = None
        title_set_success = False
        if set_title:
            try:
                title_set_result = agent.set_title(slide_index, set_title)
                if isinstance(title_set_result, dict):
                    title_set_success = title_set_result.get("title_set", False)
                else:
                    title_set_success = True
            except Exception:
                title_set_success = False
        
        # Get slide info before saving (for verification)
        slide_info = agent.get_slide_info(slide_index)
        
        # Save the file
        agent.save()
        
        # Get updated presentation info (includes final version hash)
        prs_info = agent.get_presentation_info()
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "layout": matched_layout,
        "title_set": set_title,
        "title_set_success": title_set_success,
        "total_slides": prs_info["slide_count"],
        "slide_info": {
            "shape_count": slide_info.get("shape_count", 0),
            "has_notes": slide_info.get("has_notes", False)
        },
        "presentation_version_before": version_before,
        "presentation_version_after": prs_info.get("presentation_version"),
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add new slide to PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Add slide at end
  uv run tools/ppt_add_slide.py \\
    --file presentation.pptx \\
    --layout "Title and Content" \\
    --json
  
  # Add slide at specific position
  uv run tools/ppt_add_slide.py \\
    --file deck.pptx \\
    --layout "Section Header" \\
    --index 2 \\
    --json
  
  # Add slide with title
  uv run tools/ppt_add_slide.py \\
    --file presentation.pptx \\
    --layout "Title Slide" \\
    --title "Q4 Results" \\
    --json

Common Layouts:
  - Title Slide
  - Title and Content
  - Section Header
  - Two Content
  - Comparison
  - Title Only
  - Blank

Layout Matching:
  The tool supports fuzzy matching:
  - Exact match first
  - Then substring match (case-insensitive)
  
  Example: "content" will match "Title and Content"

Finding Available Layouts:
  Use ppt_get_info.py to list layouts:
  uv run tools/ppt_get_info.py --file presentation.pptx --json | jq '.layouts'

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 5,
    "layout": "Title and Content",
    "title_set": "Q4 Results",
    "title_set_success": true,
    "total_slides": 6,
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
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
        '--layout',
        required=True,
        help='Layout name for new slide (fuzzy matching supported)'
    )
    
    parser.add_argument(
        '--index',
        type=int,
        help='Position to insert slide (0-based, default: end)'
    )
    
    parser.add_argument(
        '--title',
        help='Optional title text to set on new slide'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_slide(
            filepath=args.file,
            layout=args.layout,
            index=args.index,
            set_title=args.title
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
            "details": getattr(e, 'details', {})
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)


if __name__ == "__main__":
    main()
