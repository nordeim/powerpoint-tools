#!/usr/bin/env python3
"""
PowerPoint Search Content Tool v3.1.1
Search for text content across all slides in a presentation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_search_content.py --file presentation.pptx --query "Revenue" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

This tool searches for text content across slides, including:
- Text in shapes and text boxes
- Slide titles and subtitles
- Speaker notes
- Table cell contents

Use this tool to locate content before using ppt_replace_text.py or to
navigate large presentations efficiently.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Pattern

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
# TYPE DEFINITIONS
# ============================================================================

Match = Dict[str, Any]


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def compile_pattern(
    query: str,
    is_regex: bool = False,
    case_sensitive: bool = False
) -> Pattern:
    """
    Compile search pattern from query string.
    
    Args:
        query: Search query (plain text or regex)
        is_regex: If True, treat query as regular expression
        case_sensitive: If True, perform case-sensitive search
        
    Returns:
        Compiled regex pattern
        
    Raises:
        ValueError: If regex is invalid
    """
    flags = 0 if case_sensitive else re.IGNORECASE
    
    if is_regex:
        try:
            return re.compile(query, flags)
        except re.error as e:
            raise ValueError(f"Invalid regex pattern: {e}")
    else:
        escaped = re.escape(query)
        return re.compile(escaped, flags)


def extract_context(text: str, match_start: int, match_end: int, context_chars: int = 50) -> str:
    """
    Extract text context around a match.
    
    Args:
        text: Full text
        match_start: Start position of match
        match_end: End position of match
        context_chars: Characters to include before/after
        
    Returns:
        Context string with match highlighted
    """
    start = max(0, match_start - context_chars)
    end = min(len(text), match_end + context_chars)
    
    prefix = "..." if start > 0 else ""
    suffix = "..." if end < len(text) else ""
    
    context = text[start:end]
    
    return f"{prefix}{context}{suffix}"


def search_text_frame(
    text_frame,
    pattern: Pattern,
    slide_index: int,
    shape_index: int,
    shape_name: str,
    shape_type: str,
    location: str = "text"
) -> List[Match]:
    """
    Search within a text frame.
    
    Args:
        text_frame: TextFrame object
        pattern: Compiled search pattern
        slide_index: Parent slide index
        shape_index: Parent shape index
        shape_name: Shape name
        shape_type: Shape type string
        location: Location identifier ("text" or "notes")
        
    Returns:
        List of match dictionaries
    """
    matches = []
    
    try:
        full_text = ""
        for paragraph in text_frame.paragraphs:
            for run in paragraph.runs:
                full_text += run.text
            full_text += "\n"
        
        full_text = full_text.strip()
        
        if not full_text:
            return matches
        
        for match in pattern.finditer(full_text):
            matches.append({
                "slide_index": slide_index,
                "shape_index": shape_index,
                "shape_name": shape_name,
                "shape_type": shape_type,
                "location": location,
                "match_text": match.group(),
                "match_start": match.start(),
                "match_end": match.end(),
                "context": extract_context(full_text, match.start(), match.end())
            })
    except Exception:
        pass
    
    return matches


def search_table(
    table,
    pattern: Pattern,
    slide_index: int,
    shape_index: int,
    shape_name: str
) -> List[Match]:
    """
    Search within a table.
    
    Args:
        table: Table object
        pattern: Compiled search pattern
        slide_index: Parent slide index
        shape_index: Parent shape index
        shape_name: Shape name
        
    Returns:
        List of match dictionaries
    """
    matches = []
    
    try:
        for row_idx, row in enumerate(table.rows):
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text_frame.text if cell.text_frame else ""
                
                if not cell_text:
                    continue
                
                for match in pattern.finditer(cell_text):
                    matches.append({
                        "slide_index": slide_index,
                        "shape_index": shape_index,
                        "shape_name": shape_name,
                        "shape_type": "TABLE_CELL",
                        "location": "table",
                        "cell_row": row_idx,
                        "cell_col": col_idx,
                        "match_text": match.group(),
                        "match_start": match.start(),
                        "match_end": match.end(),
                        "context": extract_context(cell_text, match.start(), match.end())
                    })
    except Exception:
        pass
    
    return matches


# ============================================================================
# MAIN LOGIC
# ============================================================================

def search_content(
    filepath: Path,
    query: str,
    is_regex: bool = False,
    case_sensitive: bool = False,
    scope: str = "all",
    slide_index: Optional[int] = None
) -> Dict[str, Any]:
    """
    Search for content across a PowerPoint presentation.
    
    Args:
        filepath: Path to the PowerPoint file
        query: Search query (text or regex pattern)
        is_regex: If True, treat query as regular expression
        case_sensitive: If True, perform case-sensitive search
        scope: Search scope - "text", "notes", "tables", or "all"
        slide_index: Optional specific slide to search (None = all slides)
        
    Returns:
        Dict with search results
        
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If specified slide index is invalid
        ValueError: If regex pattern is invalid
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    pattern = compile_pattern(query, is_regex, case_sensitive)
    
    all_matches: List[Match] = []
    slides_searched: List[int] = []
    slides_with_matches: List[int] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)
        
        presentation_version = agent.get_presentation_version()
        total_slides = agent.get_slide_count()
        
        if slide_index is not None:
            if not 0 <= slide_index < total_slides:
                raise SlideNotFoundError(
                    f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                    details={
                        "requested_index": slide_index,
                        "available_slides": total_slides
                    }
                )
            slides_to_search = [slide_index]
        else:
            slides_to_search = list(range(total_slides))
        
        for slide_idx in slides_to_search:
            slides_searched.append(slide_idx)
            slide = agent.prs.slides[slide_idx]
            slide_matches: List[Match] = []
            
            if scope in ["text", "all"]:
                for shape_idx, shape in enumerate(slide.shapes):
                    shape_name = getattr(shape, 'name', f'Shape_{shape_idx}')
                    shape_type = str(shape.shape_type).replace('MSO_SHAPE_TYPE.', '')
                    
                    if hasattr(shape, 'text_frame') and shape.has_text_frame:
                        matches = search_text_frame(
                            shape.text_frame,
                            pattern,
                            slide_idx,
                            shape_idx,
                            shape_name,
                            shape_type,
                            "text"
                        )
                        slide_matches.extend(matches)
                    
                    if scope in ["tables", "all"] and hasattr(shape, 'table') and shape.has_table:
                        matches = search_table(
                            shape.table,
                            pattern,
                            slide_idx,
                            shape_idx,
                            shape_name
                        )
                        slide_matches.extend(matches)
            
            if scope in ["notes", "all"]:
                try:
                    notes_slide = slide.notes_slide
                    if notes_slide and notes_slide.notes_text_frame:
                        notes_text = notes_slide.notes_text_frame.text
                        if notes_text:
                            for match in pattern.finditer(notes_text):
                                slide_matches.append({
                                    "slide_index": slide_idx,
                                    "shape_index": None,
                                    "shape_name": "Speaker Notes",
                                    "shape_type": "NOTES",
                                    "location": "notes",
                                    "match_text": match.group(),
                                    "match_start": match.start(),
                                    "match_end": match.end(),
                                    "context": extract_context(notes_text, match.start(), match.end())
                                })
                except Exception:
                    pass
            
            if slide_matches:
                slides_with_matches.append(slide_idx)
                all_matches.extend(slide_matches)
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "query": query,
        "options": {
            "regex": is_regex,
            "case_sensitive": case_sensitive,
            "scope": scope
        },
        "total_matches": len(all_matches),
        "slides_searched": len(slides_searched),
        "slides_with_matches": slides_with_matches,
        "matches": all_matches,
        "presentation_version": presentation_version,
        "tool_version": __version__
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Search for content across PowerPoint slides",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Simple text search
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "Revenue" --json

  # Case-sensitive search
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "Q4" --case-sensitive --json

  # Regex search for dates
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "\\d{4}-\\d{2}-\\d{2}" --regex --json

  # Search only in speaker notes
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "TODO" --scope notes --json

  # Search specific slide
  uv run tools/ppt_search_content.py \\
    --file presentation.pptx --query "Summary" --slide 5 --json

Scope Options:
  all    - Search everywhere (default)
  text   - Search only in text shapes
  notes  - Search only in speaker notes
  tables - Search only in table cells

Use Cases:
  1. Find slides before using ppt_replace_text.py
  2. Locate placeholder text to update
  3. Audit presentations for sensitive content
  4. Navigate large presentations efficiently
  5. Verify content updates were applied

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "query": "Revenue",
    "total_matches": 5,
    "slides_with_matches": [0, 2, 7],
    "matches": [
      {
        "slide_index": 0,
        "shape_index": 3,
        "shape_name": "Title 1",
        "shape_type": "PLACEHOLDER",
        "location": "text",
        "match_text": "Revenue",
        "context": "...Q4 Revenue Growth..."
      }
    ],
    "presentation_version": "a1b2c3...",
    "tool_version": "3.1.1"
  }
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to search'
    )
    
    parser.add_argument(
        '--query',
        required=True,
        type=str,
        help='Search query (text or regex pattern)'
    )
    
    parser.add_argument(
        '--regex',
        action='store_true',
        help='Treat query as regular expression'
    )
    
    parser.add_argument(
        '--case-sensitive',
        action='store_true',
        help='Perform case-sensitive search'
    )
    
    parser.add_argument(
        '--scope',
        choices=['all', 'text', 'notes', 'tables'],
        default='all',
        help='Search scope (default: all)'
    )
    
    parser.add_argument(
        '--slide',
        type=int,
        default=None,
        help='Limit search to specific slide index'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = search_content(
            filepath=args.file.resolve(),
            query=args.query,
            is_regex=args.regex,
            case_sensitive=args.case_sensitive,
            scope=args.scope,
            slide_index=args.slide
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(0)
        
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
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check regex syntax if using --regex flag",
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
            "suggestion": "Check file integrity and format",
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
