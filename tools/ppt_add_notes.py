#!/usr/bin/env python3
"""
PowerPoint Add Speaker Notes Tool v3.1.0
Add, append, or overwrite speaker notes for a specific slide.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "Key talking point" --json
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "New script" --mode overwrite --json
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 --text "IMPORTANT:" --mode prepend --json

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

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
)

__version__ = "3.1.0"


def add_notes(
    filepath: Path,
    slide_index: int,
    text: str,
    mode: str = "append"
) -> Dict[str, Any]:
    """
    Add speaker notes to a slide.
    
    Args:
        filepath: Path to PowerPoint file (.pptx only)
        slide_index: Index of slide to modify (0-based)
        text: Text content to add to speaker notes
        mode: Insertion mode - 'append' (default), 'prepend', or 'overwrite'
        
    Returns:
        Dict containing:
            - status: 'success'
            - file: Absolute path to file
            - slide_index: Target slide index
            - mode: Mode that was used
            - original_length: Character count of original notes
            - new_length: Character count of final notes
            - preview: First 100 characters of final notes
            - presentation_version_before: Version hash before changes
            - presentation_version_after: Version hash after changes
            - tool_version: Tool version string
            
    Raises:
        FileNotFoundError: If PowerPoint file doesn't exist
        ValueError: If file format is invalid, text is empty, or mode is invalid
        SlideNotFoundError: If slide index is out of range
        PowerPointAgentError: If notes slide cannot be accessed
    """
    if filepath.suffix.lower() != '.pptx':
        raise ValueError(
            f"Invalid file format '{filepath.suffix}'. Only .pptx files are supported."
        )

    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not text or not text.strip():
        raise ValueError("Notes text cannot be empty")
    
    if mode not in ('append', 'prepend', 'overwrite'):
        raise ValueError(
            f"Invalid mode '{mode}'. Must be 'append', 'prepend', or 'overwrite'."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        slide_count = agent.get_slide_count()
        
        if slide_count == 0:
            raise PowerPointAgentError("Presentation has no slides")
        
        if not 0 <= slide_index < slide_count:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{slide_count - 1})",
                details={"requested": slide_index, "available": slide_count}
            )
            
        slide = agent.prs.slides[slide_index]
        
        try:
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
        except Exception as e:
            raise PowerPointAgentError(f"Failed to access notes slide: {str(e)}")
        
        original_text = text_frame.text if text_frame.text else ""
        
        if mode == "overwrite":
            final_text = text
        elif mode == "append":
            if original_text and original_text.strip():
                final_text = original_text + "\n" + text
            else:
                final_text = text
        elif mode == "prepend":
            if original_text and original_text.strip():
                final_text = text + "\n" + original_text
            else:
                final_text = text
        
        text_frame.text = final_text
                
        agent.save()
        
        version_after = agent.get_presentation_version()
        
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "mode": mode,
        "original_length": len(original_text),
        "new_length": len(final_text),
        "preview": final_text[:100] + "..." if len(final_text) > 100 else final_text,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add speaker notes to a PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Append notes (default mode)
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 \\
        --text "Key talking point: Emphasize Q4 growth." --json
    
    # Overwrite existing notes
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 \\
        --text "Complete new script for this slide." --mode overwrite --json
    
    # Prepend notes (add before existing)
    uv run tools/ppt_add_notes.py --file deck.pptx --slide 0 \\
        --text "IMPORTANT: Start with customer story." --mode prepend --json

Modes:
    append    - Add text after existing notes (default)
    prepend   - Add text before existing notes
    overwrite - Replace all existing notes with new text

Use Cases:
    - Presentation scripting and speaker preparation
    - Accessibility: text alternatives for complex visuals
    - Documentation: embedding context for future editors
    - Training: detailed explanations not shown on slides
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='Path to PowerPoint file (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--text',
        required=True,
        help='Notes content to add'
    )
    
    parser.add_argument(
        '--mode',
        choices=['append', 'prepend', 'overwrite'],
        default='append',
        help='Insertion mode: append (default), prepend, or overwrite'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output as JSON (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_notes(
            filepath=args.file,
            slide_index=args.slide,
            text=args.text,
            mode=args.mode
        )
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
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slide count."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check that file is .pptx format, text is not empty, and mode is valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Verify the file is not corrupted and the slide structure is valid."
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
