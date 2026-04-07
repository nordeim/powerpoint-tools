#!/usr/bin/env python3
"""
PowerPoint Duplicate Slide Tool v3.1.1
Clone an existing slide within the presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_duplicate_slide.py --file presentation.pptx --index 0 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Changelog v3.1.1:
    - Added sys.stdout.flush() for pipeline safety
    - Added suggestion field to all error handlers
    - Added tool_version to all error responses
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

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"


# ============================================================================
# MAIN LOGIC
# ============================================================================

def duplicate_slide(
    filepath: Path, 
    index: int
) -> Dict[str, Any]:
    """
    Duplicate a slide at the specified index.
    
    Creates a deep copy of the slide including all shapes, text runs,
    formatting, and styles. The duplicated slide is inserted immediately
    after the source slide.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        index: Index of the slide to duplicate (0-based)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - source_index: Index of the original slide
            - new_slide_index: Index of the newly created duplicate
            - total_slides: Total slide count after duplication
            - layout: Layout name of the duplicated slide
            - presentation_version_before: State hash before duplication
            - presentation_version_after: State hash after duplication
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        
    Example:
        >>> result = duplicate_slide(
        ...     filepath=Path("presentation.pptx"),
        ...     index=0
        ... )
        >>> print(result["new_slide_index"])
        1
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        total = agent.get_slide_count()
        if not 0 <= index < total:
            raise SlideNotFoundError(
                f"Slide index {index} out of range (0-{total - 1})",
                details={
                    "requested_index": index,
                    "available_slides": total,
                    "valid_range": f"0 to {total - 1}"
                }
            )
        
        result = agent.duplicate_slide(index)
        
        if isinstance(result, dict):
            new_index = result.get("slide_index", result.get("new_slide_index", index + 1))
        else:
            new_index = result
        
        agent.save()
        
        slide_info = agent.get_slide_info(new_index)
        layout_name = slide_info.get("layout", "Unknown")
        
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
        final_count = info_after["slide_count"]
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "source_index": index,
        "new_slide_index": new_index,
        "total_slides": final_count,
        "layout": layout_name,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Duplicate a PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Duplicate the first slide
  uv run tools/ppt_duplicate_slide.py --file presentation.pptx --index 0 --json
  
  # Duplicate slide at index 5
  uv run tools/ppt_duplicate_slide.py --file deck.pptx --index 5 --json

Behavior:
  - Creates a deep copy of the slide at the specified index
  - The duplicate is inserted immediately after the source slide
  - All shapes, text, formatting, and styles are preserved
  - Returns the index of the newly created slide

Use Cases:
  - Creating similar slides with slight variations
  - Building slide sequences from a template slide
  - Backing up a slide before major changes

Important Notes:
  - Shape indices on the duplicated slide start fresh
  - The new slide gets the next available index
  - Use ppt_get_slide_info.py to inspect the new slide

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "source_index": 0,
    "new_slide_index": 1,
    "total_slides": 6,
    "layout": "Title and Content",
    "presentation_version_before": "a1b2c3d4...",
    "presentation_version_after": "e5f6g7h8...",
    "tool_version": "3.1.1"
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
        '--index', 
        required=True, 
        type=int, 
        help='Source slide index to duplicate (0-based)'
    )
    
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = duplicate_slide(
            filepath=args.file.resolve(),
            index=args.index
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide indices",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
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
            "suggestion": "Check file integrity and slide index validity",
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
