#!/usr/bin/env python3
"""
PowerPoint Delete Slide Tool v3.1.1
Remove a slide from the presentation.

⚠️ DESTRUCTIVE OPERATION - Requires approval token with scope 'delete:slide'

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_delete_slide.py --file presentation.pptx --index 1 --approval-token "HMAC-SHA256:..." --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)
    4: Permission error (missing or invalid approval token)

Security:
    This tool performs a destructive operation and requires a valid approval
    token with scope 'delete:slide'. Generate tokens using the approval token
    system described in the governance documentation.

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
from typing import Dict, Any, Optional

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
# EXCEPTION FALLBACK
# ============================================================================

try:
    from core.powerpoint_agent_core import ApprovalTokenError
except ImportError:
    class ApprovalTokenError(PowerPointAgentError):
        """Exception raised when approval token is missing or invalid."""
        def __init__(self, message: str, details: Optional[Dict] = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)
        
        def __str__(self):
            return self.message


# ============================================================================
# MAIN LOGIC
# ============================================================================

def delete_slide(
    filepath: Path, 
    index: int,
    approval_token: str
) -> Dict[str, Any]:
    """
    Delete a slide at the specified index.
    
    ⚠️ DESTRUCTIVE OPERATION - This permanently removes the slide.
    
    This operation requires a valid approval token with scope 'delete:slide'
    to prevent accidental data loss. Always clone the presentation first
    using ppt_clone_presentation.py before performing destructive operations.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        index: Slide index to delete (0-based)
        approval_token: HMAC-SHA256 approval token with scope 'delete:slide'
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - deleted_index: Index of the deleted slide
            - remaining_slides: Number of slides after deletion
            - presentation_version_before: State hash before deletion
            - presentation_version_after: State hash after deletion
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        ApprovalTokenError: If approval token is missing or invalid
        
    Example:
        >>> result = delete_slide(
        ...     filepath=Path("presentation.pptx"),
        ...     index=2,
        ...     approval_token="HMAC-SHA256:eyJ..."
        ... )
        >>> print(result["remaining_slides"])
        9
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not approval_token:
        raise ApprovalTokenError(
            "Approval token required for slide deletion",
            details={
                "operation": "delete_slide",
                "slide_index": index,
                "required_scope": "delete:slide",
                "file": str(filepath)
            }
        )
    
    if not approval_token.startswith("HMAC-SHA256:"):
        raise ApprovalTokenError(
            "Invalid approval token format. Expected 'HMAC-SHA256:...'",
            details={
                "operation": "delete_slide",
                "slide_index": index,
                "required_scope": "delete:slide",
                "token_prefix_received": approval_token[:20] + "..." if len(approval_token) > 20 else approval_token
            }
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        total_slides = agent.get_slide_count()
        if not 0 <= index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": index,
                    "available_slides": total_slides,
                    "valid_range": f"0 to {total_slides - 1}"
                }
            )
        
        try:
            agent.delete_slide(index, approval_token=approval_token)
        except TypeError:
            agent.delete_slide(index)
        
        agent.save()
        
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
        new_count = info_after["slide_count"]
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "deleted_index": index,
        "remaining_slides": new_count,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Delete PowerPoint slide (⚠️ DESTRUCTIVE - requires approval token)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ DESTRUCTIVE OPERATION ⚠️

This tool permanently removes a slide from the presentation.
An approval token with scope 'delete:slide' is REQUIRED.

Examples:
  # Delete slide at index 2 (third slide)
  uv run tools/ppt_delete_slide.py \\
    --file presentation.pptx \\
    --index 2 \\
    --approval-token "HMAC-SHA256:eyJzY29wZSI6ImRlbGV0ZTpzbGlkZSIsLi4ufQ==.abc123..." \\
    --json

Safety Workflow:
  1. CLONE the presentation first:
     uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx
  
  2. VERIFY slide count and content:
     uv run tools/ppt_get_info.py --file work.pptx --json
     uv run tools/ppt_get_slide_info.py --file work.pptx --slide 2 --json
  
  3. GENERATE approval token with scope 'delete:slide'
  
  4. DELETE the slide:
     uv run tools/ppt_delete_slide.py --file work.pptx --index 2 --approval-token "..." --json

Token Generation:
  Approval tokens must be generated by a trusted service using HMAC-SHA256.
  The token must have scope 'delete:slide' and not be expired.
  See governance documentation for token generation details.

Exit Codes:
  0: Success - slide deleted
  1: Error - check error_type in JSON output
  4: Permission Error - missing or invalid approval token

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "deleted_index": 2,
    "remaining_slides": 9,
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
        help='Slide index to delete (0-based)'
    )
    
    parser.add_argument(
        '--approval-token',
        required=True,
        type=str,
        help='Approval token with scope "delete:slide" (REQUIRED for this destructive operation)'
    )
    
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = delete_slide(
            filepath=args.file.resolve(), 
            index=args.index,
            approval_token=args.approval_token
        )
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
    except ApprovalTokenError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ApprovalTokenError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Generate a valid approval token with scope 'delete:slide'",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(4)
        
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
