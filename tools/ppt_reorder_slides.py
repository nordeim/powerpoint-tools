#!/usr/bin/env python3
"""
PowerPoint Reorder Slides Tool v3.1.0
Move a slide from one position to another

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_reorder_slides.py --file presentation.pptx --from-index 3 --to-index 1 --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Notes:
    - Indices are 0-based
    - Moving a slide shifts other slides accordingly
    - Original content is preserved during move
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


def reorder_slides(
    filepath: Path, 
    from_index: int, 
    to_index: int
) -> Dict[str, Any]:
    """
    Move a slide from one position to another.
    
    The slide at from_index is moved to to_index. Other slides
    shift accordingly to accommodate the move.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        from_index: Current position of the slide (0-based)
        to_index: Target position for the slide (0-based)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - moved_from: Original slide position
            - moved_to: New slide position
            - total_slides: Total slide count
            - presentation_version_before: State hash before reorder
            - presentation_version_after: State hash after reorder
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If from_index or to_index is out of range
        
    Example:
        >>> result = reorder_slides(
        ...     filepath=Path("presentation.pptx"),
        ...     from_index=5,
        ...     to_index=1
        ... )
        >>> print(result["moved_to"])
        1
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate indices are different
    if from_index == to_index:
        # Not an error, but no operation needed
        with PowerPointAgent(filepath) as agent:
            agent.open(filepath, acquire_lock=False)
            total = agent.get_slide_count()
            prs_info = agent.get_presentation_info()
        
        return {
            "status": "success",
            "file": str(filepath.resolve()),
            "moved_from": from_index,
            "moved_to": to_index,
            "total_slides": total,
            "note": "Source and target indices are the same. No change made.",
            "presentation_version_before": prs_info.get("presentation_version"),
            "presentation_version_after": prs_info.get("presentation_version"),
            "tool_version": __version__
        }
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE reorder
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate indices
        total = agent.get_slide_count()
        
        if not 0 <= from_index < total:
            raise SlideNotFoundError(
                f"Source index {from_index} out of range (0-{total - 1})",
                details={
                    "requested_index": from_index,
                    "available_slides": total,
                    "parameter": "from_index"
                }
            )
        
        if not 0 <= to_index < total:
            raise SlideNotFoundError(
                f"Target index {to_index} out of range (0-{total - 1})",
                details={
                    "requested_index": to_index,
                    "available_slides": total,
                    "parameter": "to_index"
                }
            )
        
        # Perform reorder
        agent.reorder_slides(from_index, to_index)
        
        # Save changes
        agent.save()
        
        # Capture version AFTER reorder
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "moved_from": from_index,
        "moved_to": to_index,
        "total_slides": total,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Reorder PowerPoint slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Move slide from position 3 to position 1
  uv run tools/ppt_reorder_slides.py \\
    --file presentation.pptx \\
    --from-index 3 \\
    --to-index 1 \\
    --json
  
  # Move last slide to beginning
  uv run tools/ppt_reorder_slides.py \\
    --file deck.pptx \\
    --from-index 9 \\
    --to-index 0 \\
    --json
  
  # Move first slide to end
  uv run tools/ppt_reorder_slides.py \\
    --file deck.pptx \\
    --from-index 0 \\
    --to-index 9 \\
    --json

Behavior:
  - Slide at from_index is moved to to_index
  - Other slides shift to accommodate the move
  - All slide content is preserved
  - Indices are 0-based

Finding Slide Count:
  Use ppt_get_info.py to check slide count:
  uv run tools/ppt_get_info.py --file presentation.pptx --json | jq '.slide_count'

Use Cases:
  - Reorganizing presentation flow
  - Moving section headers
  - Reordering topic sequences
  - Placing summary slides

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "moved_from": 3,
    "moved_to": 1,
    "total_slides": 10,
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
        '--from-index', 
        required=True, 
        type=int, 
        help='Current slide index (0-based)'
    )
    parser.add_argument(
        '--to-index', 
        required=True, 
        type=int, 
        help='Target slide index (0-based)'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = reorder_slides(
            filepath=args.file, 
            from_index=args.from_index, 
            to_index=args.to_index
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
            "suggestion": "Use ppt_get_info.py to check slide count"
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
