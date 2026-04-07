#!/usr/bin/env python3
"""
PowerPoint Get Info Tool v3.1.0
Get presentation metadata (slide count, dimensions, file size, version)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_get_info.py --file presentation.pptx --json

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
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError
)

__version__ = "3.1.0"


def get_info(filepath: Path) -> Dict[str, Any]:
    """
    Get comprehensive information about a PowerPoint presentation.
    
    This is a read-only operation that does not modify the file.
    It acquires no lock, allowing concurrent reads.
    
    Args:
        filepath: Path to the PowerPoint file to inspect
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to the file
            - slide_count: Number of slides
            - file_size_bytes: File size in bytes
            - file_size_mb: File size in megabytes (rounded to 2 decimals)
            - slide_dimensions: Width, height (inches), and aspect ratio
            - layouts: List of available layout names
            - layout_count: Number of available layouts
            - modified: Last modification timestamp (if available)
            - presentation_version: State hash for change tracking
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        PowerPointAgentError: If the file cannot be read
        
    Example:
        >>> result = get_info(Path("presentation.pptx"))
        >>> print(result["slide_count"])
        15
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    with PowerPointAgent(filepath) as agent:
        # Open without acquiring lock (read-only operation)
        agent.open(filepath, acquire_lock=False)
        
        # Get comprehensive presentation info
        info = agent.get_presentation_info()
    
    # Build response with all available information
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_count": info.get("slide_count", 0),
        "file_size_bytes": info.get("file_size_bytes", 0),
        "file_size_mb": round(info.get("file_size_bytes", 0) / (1024 * 1024), 2),
        "slide_dimensions": {
            "width_inches": info.get("slide_width_inches", 13.333),
            "height_inches": info.get("slide_height_inches", 7.5),
            "aspect_ratio": info.get("aspect_ratio", "16:9")
        },
        "layouts": info.get("layouts", []),
        "layout_count": len(info.get("layouts", [])),
        "modified": info.get("modified"),
        "presentation_version": info.get("presentation_version"),
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Get PowerPoint presentation information",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Get presentation info
  uv run tools/ppt_get_info.py --file presentation.pptx --json
  
  # Check before making modifications
  uv run tools/ppt_get_info.py --file deck.pptx --json | jq '.slide_count'

Output Information:
  - file: Absolute path to the file
  - slide_count: Total number of slides
  - file_size_bytes/mb: File size
  - slide_dimensions: Width, height (inches), and aspect ratio
  - layouts: List of available layout names
  - layout_count: Number of available layouts
  - modified: Last modification timestamp
  - presentation_version: State hash for change tracking

Use Cases:
  - Verify presentation structure before editing
  - Check aspect ratio for compatibility
  - List available layouts for slide creation
  - Track presentation state via version hash
  - Validate file size limits

Aspect Ratios:
  - 16:9 (Widescreen): Most common, modern standard
  - 4:3 (Standard): Traditional, older format
  - 16:10: Some displays, between 16:9 and 4:3

Layout Information:
  The layouts list shows all slide layouts available in the presentation.
  Use these exact names with other tools:
  - ppt_create_new.py --layout "Title Slide"
  - ppt_add_slide.py --layout "Title and Content"
  - ppt_set_slide_layout.py --layout "Section Header"

Version Tracking:
  The presentation_version field is a hash of the presentation state
  including slide count, layouts, shape geometry, and text content.
  Use this to detect changes between operations.

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_count": 15,
    "file_size_bytes": 2568192,
    "file_size_mb": 2.45,
    "slide_dimensions": {
      "width_inches": 13.333,
      "height_inches": 7.5,
      "aspect_ratio": "16:9"
    },
    "layouts": [
      "Title Slide",
      "Title and Content",
      "Section Header",
      "Two Content",
      "Comparison",
      "Title Only",
      "Blank"
    ],
    "layout_count": 7,
    "modified": "2024-01-15T10:30:00",
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
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = get_info(filepath=args.file)
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
