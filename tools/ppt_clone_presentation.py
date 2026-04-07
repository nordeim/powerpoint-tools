#!/usr/bin/env python3
"""
PowerPoint Clone Presentation Tool v3.1.0
Create an exact copy of a presentation for safe editing

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

⚠️ GOVERNANCE FOUNDATION - Clone-Before-Edit Principle

This tool implements the foundational safety principle: NEVER modify source
files directly. Always create a working copy first using this tool.

Usage:
    uv run tools/ppt_clone_presentation.py --source original.pptx --output work_copy.pptx --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Safety Workflow:
    1. Clone: ppt_clone_presentation.py --source original.pptx --output work.pptx
    2. Edit: Use other tools on work.pptx
    3. Validate: ppt_validate_presentation.py --file work.pptx
    4. Deliver: Rename/move work.pptx when approved
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


def clone_presentation(
    source: Path, 
    output: Path
) -> Dict[str, Any]:
    """
    Create an exact copy of a PowerPoint presentation.
    
    This is the foundational tool for the Clone-Before-Edit governance
    principle. Always use this before modifying any presentation to:
    
    1. Protect source files from accidental modification
    2. Enable rollback to original if needed
    3. Create audit-safe work copies
    4. Allow parallel editing without conflicts
    
    Args:
        source: Path to the source presentation to clone
        output: Path where the clone will be saved
        
    Returns:
        Dict containing:
            - status: "success"
            - source: Absolute path to source file
            - output: Absolute path to cloned file
            - source_size_bytes: Size of source file
            - output_size_bytes: Size of cloned file
            - slide_count: Number of slides in presentation
            - presentation_version: State hash of the cloned presentation
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If source file doesn't exist
        PermissionError: If output location is not writable
        
    Example:
        >>> result = clone_presentation(
        ...     source=Path("template.pptx"),
        ...     output=Path("work/project.pptx")
        ... )
        >>> print(result["presentation_version"])
        'a1b2c3d4e5f6g7h8'
    """
    # Validate source exists
    if not source.exists():
        raise FileNotFoundError(f"Source file not found: {source}")
    
    # Validate source is a PowerPoint file
    if source.suffix.lower() not in {'.pptx', '.pptm', '.potx'}:
        raise ValueError(
            f"Source must be a PowerPoint file (.pptx, .pptm, .potx), got: {source.suffix}"
        )
    
    # Ensure output has correct extension
    if not output.suffix.lower() == '.pptx':
        output = output.with_suffix('.pptx')
    
    # Create output directory if needed
    output.parent.mkdir(parents=True, exist_ok=True)
    
    # Get source file size
    source_size = source.stat().st_size
    
    # Open source (read-only, no lock) and save to output
    with PowerPointAgent(source) as agent:
        agent.open(source, acquire_lock=False)  # Read-only, don't lock source
        
        # Get presentation info before saving
        info = agent.get_presentation_info()
        
        # Save to new location (creates the clone)
        agent.save(output)
        
        # Get the cloned presentation's version
        presentation_version = info.get("presentation_version")
        slide_count = info.get("slide_count", 0)
    
    # Get output file size (should match source)
    output_size = output.stat().st_size
    
    return {
        "status": "success",
        "source": str(source.resolve()),
        "output": str(output.resolve()),
        "source_size_bytes": source_size,
        "output_size_bytes": output_size,
        "slide_count": slide_count,
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Clone PowerPoint presentation (⚠️ GOVERNANCE FOUNDATION)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
⚠️ GOVERNANCE FOUNDATION: Clone-Before-Edit Principle

This tool implements the first and most important safety rule:
NEVER modify source files directly. Always create a working copy.

Examples:
  # Basic clone
  uv run tools/ppt_clone_presentation.py \\
    --source original.pptx \\
    --output work_copy.pptx \\
    --json
  
  # Clone to work directory
  uv run tools/ppt_clone_presentation.py \\
    --source templates/corporate.pptx \\
    --output work/q4_report.pptx \\
    --json
  
  # Clone for parallel editing
  uv run tools/ppt_clone_presentation.py \\
    --source shared/presentation.pptx \\
    --output my_edits/presentation_v2.pptx \\
    --json

Safety Workflow:
  1. CLONE the source file:
     uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx
  
  2. PROBE the clone:
     uv run tools/ppt_capability_probe.py --file work.pptx --deep --json
  
  3. EDIT the clone (not the original!):
     uv run tools/ppt_add_slide.py --file work.pptx --layout "Title Slide" --json
  
  4. VALIDATE before delivery:
     uv run tools/ppt_validate_presentation.py --file work.pptx --json
     uv run tools/ppt_check_accessibility.py --file work.pptx --json
  
  5. DELIVER when approved:
     mv work.pptx final_presentation.pptx

Why Clone-Before-Edit?
  - Protects original files from accidental modification
  - Enables rollback if edits go wrong
  - Creates audit trail (original preserved)
  - Allows concurrent work without conflicts
  - Required by governance framework

Version Tracking:
  The presentation_version in the output is a state hash that can be used
  to track changes. After editing the clone, the version will change.
  Compare versions to detect modifications.

Output Format:
  {
    "status": "success",
    "source": "/path/to/original.pptx",
    "output": "/path/to/work_copy.pptx",
    "source_size_bytes": 1234567,
    "output_size_bytes": 1234567,
    "slide_count": 15,
    "presentation_version": "a1b2c3d4e5f6g7h8",
    "tool_version": "3.1.0"
  }
        """
    )
    
    parser.add_argument(
        '--source', 
        required=True, 
        type=Path, 
        help='Source PowerPoint file to clone'
    )
    parser.add_argument(
        '--output', 
        required=True, 
        type=Path, 
        help='Destination path for the cloned file'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = clone_presentation(source=args.source, output=args.output)
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the source file path exists and is accessible"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Source must be a PowerPoint file (.pptx, .pptm, .potx)"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except PermissionError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PermissionError",
            "suggestion": "Check write permissions for the output directory"
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
