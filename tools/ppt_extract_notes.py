#!/usr/bin/env python3
"""
PowerPoint Extract Notes Tool v3.1.0
Extract speaker notes from all slides in a presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json

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
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import PowerPointAgent

__version__ = "3.1.0"


def extract_notes(filepath: Path) -> Dict[str, Any]:
    """
    Extract speaker notes from all slides in a presentation.
    
    This is a read-only operation that does not modify the file.
    
    Args:
        filepath: Path to PowerPoint file (.pptx only)
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute path to file
            - presentation_version: Version hash of presentation
            - total_slides: Total number of slides
            - notes_found: Count of slides that have notes content
            - notes: Dict mapping slide index (as string) to notes text
            - tool_version: Tool version string
            
    Raises:
        FileNotFoundError: If PowerPoint file doesn't exist
        ValueError: If file format is not .pptx
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError(
            f"Invalid file format '{filepath.suffix}'. Only .pptx files are supported."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        presentation_version = agent.get_presentation_version()
        notes = agent.extract_notes()
        total_slides = agent.get_slide_count()
    
    notes_with_content = sum(1 for text in notes.values() if text and text.strip())
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "presentation_version": presentation_version,
        "total_slides": total_slides,
        "notes_found": notes_with_content,
        "notes": notes,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Extract speaker notes from all slides in a PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Extract notes from presentation
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json
    
    # Save notes to file
    uv run tools/ppt_extract_notes.py --file presentation.pptx --json > notes.json

Output Format:
    {
      "status": "success",
      "file": "/path/to/presentation.pptx",
      "presentation_version": "a1b2c3d4e5f6g7h8",
      "total_slides": 10,
      "notes_found": 5,
      "notes": {
        "0": "Speaker notes for slide 1...",
        "1": "",
        "2": "Important talking points...",
        ...
      },
      "tool_version": "3.1.0"
    }

Use Cases:
    - Export notes for speaker preparation
    - Backup presentation scripts
    - Convert notes to other formats
    - Accessibility: extract text alternatives
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='Path to PowerPoint file (.pptx)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output as JSON (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = extract_notes(filepath=args.file)
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Ensure file has .pptx extension."
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
