#!/usr/bin/env python3
"""
PowerPoint Replace Text Tool v3.1.0
Find and replace text across presentation or in specific targets

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Features:
    - Global replacement (entire presentation)
    - Targeted replacement (specific slide)
    - Surgical replacement (specific shape)
    - Dry-run mode (preview without changes)
    - Case-sensitive matching option
    - Formatting-preserving replacement (run-level)
    - Location reporting

Usage:
    uv run tools/ppt_replace_text.py --file deck.pptx --find "Old" --replace "New" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Safety:
    Always use --dry-run first to preview changes before applying.
    For mass replacements, consider cloning the presentation first.
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null to prevent library noise from corrupting JSON output
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import re
import json
import argparse
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"


def perform_replacement_on_shape(
    shape, 
    find: str, 
    replace: str, 
    match_case: bool
) -> int:
    """
    Perform text replacement in a single shape.
    
    Uses a two-strategy approach:
    1. Run-level replacement (preserves formatting)
    2. Shape-level fallback (for text split across runs)
    
    Args:
        shape: PowerPoint shape object with text_frame
        find: Text to find
        replace: Replacement text
        match_case: Whether to match case
        
    Returns:
        Number of replacements made
    """
    if not hasattr(shape, 'text_frame'):
        return 0
    
    count = 0
    text_frame = shape.text_frame
    
    # Strategy 1: Replace in runs (preserves formatting)
    for paragraph in text_frame.paragraphs:
        for run in paragraph.runs:
            if match_case:
                if find in run.text:
                    run.text = run.text.replace(find, replace)
                    count += 1
            else:
                if find.lower() in run.text.lower():
                    pattern = re.compile(re.escape(find), re.IGNORECASE)
                    if pattern.search(run.text):
                        run.text = pattern.sub(replace, run.text)
                        count += 1
    
    if count > 0:
        return count
    
    # Strategy 2: Shape-level replacement (if runs didn't catch it due to splitting)
    try:
        full_text = shape.text
        should_replace = False
        
        if match_case:
            if find in full_text:
                should_replace = True
        else:
            if find.lower() in full_text.lower():
                should_replace = True
        
        if should_replace:
            if match_case:
                new_text = full_text.replace(find, replace)
            else:
                pattern = re.compile(re.escape(find), re.IGNORECASE)
                new_text = pattern.sub(replace, full_text)
            
            # Only apply if text actually changed
            if new_text != full_text:
                shape.text = new_text
                count += 1
    except Exception:
        pass  # Continue without shape-level replacement
    
    return count


def replace_text(
    filepath: Path,
    find: str,
    replace: str,
    slide_index: Optional[int] = None,
    shape_index: Optional[int] = None,
    match_case: bool = False,
    dry_run: bool = False
) -> Dict[str, Any]:
    """
    Find and replace text with optional targeting.
    
    Supports three scopes:
    1. Global: All slides, all shapes (default)
    2. Slide-specific: Single slide, all shapes (--slide N)
    3. Shape-specific: Single shape (--slide N --shape M)
    
    Args:
        filepath: Path to PowerPoint file
        find: Text to find
        replace: Replacement text
        slide_index: Optional specific slide index (0-based)
        shape_index: Optional specific shape index (requires slide_index)
        match_case: Whether to match case (default: False)
        dry_run: Preview without making changes (default: False)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Path to file
            - action: "dry_run" or "replace"
            - find/replace: Search parameters
            - scope: Target scope information
            - total_matches/replacements_made: Count
            - locations: List of affected locations
            - presentation_version_before: State hash before (if not dry_run)
            - presentation_version_after: State hash after (if not dry_run)
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If find is empty or invalid parameters
        SlideNotFoundError: If slide index is out of range
        
    Example:
        >>> result = replace_text(
        ...     filepath=Path("presentation.pptx"),
        ...     find="Old Company",
        ...     replace="New Company",
        ...     dry_run=True
        ... )
        >>> print(result["total_matches"])
        15
    """
    # Validate file extension
    valid_extensions = {'.pptx', '.pptm', '.potx'}
    if filepath.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Invalid PowerPoint file format: {filepath.suffix}. "
            f"Supported formats: {', '.join(valid_extensions)}"
        )
    
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate find text
    if not find:
        raise ValueError("Find text cannot be empty")
    
    # Validate parameters
    if shape_index is not None and slide_index is None:
        raise ValueError(
            "If --shape is specified, --slide must also be specified. "
            "Shape indices are slide-specific."
        )
    
    action = "dry_run" if dry_run else "replace"
    total_count = 0
    locations: List[Dict[str, Any]] = []
    version_before = None
    version_after = None
    
    with PowerPointAgent(filepath) as agent:
        # Open with appropriate locking
        agent.open(filepath, acquire_lock=not dry_run)
        
        # Capture version BEFORE (only for actual replacements)
        if not dry_run:
            info_before = agent.get_presentation_info()
            version_before = info_before.get("presentation_version")
        
        slide_count = agent.get_slide_count()
        
        # Include performance note in response for large presentations
        large_presentation = slide_count > 50
        
        # Determine target slides
        target_slides: List[tuple] = []
        
        if slide_index is not None:
            # Single slide scope
            if not 0 <= slide_index < slide_count:
                raise SlideNotFoundError(
                    f"Slide index {slide_index} out of range (0-{slide_count - 1})",
                    details={
                        "requested_index": slide_index,
                        "available_slides": slide_count
                    }
                )
            # NOTE: Direct prs access required for shape-level text manipulation
            target_slides = [(slide_index, agent.prs.slides[slide_index])]
        else:
            # Global scope
            target_slides = [(i, slide) for i, slide in enumerate(agent.prs.slides)]
        
        # Process each target slide
        for s_idx, slide in target_slides:
            # Determine target shapes
            target_shapes: List[tuple] = []
            
            if shape_index is not None:
                # Single shape scope
                if not 0 <= shape_index < len(slide.shapes):
                    raise ValueError(
                        f"Shape index {shape_index} out of range (0-{len(slide.shapes) - 1}) on slide {s_idx}"
                    )
                target_shapes = [(shape_index, slide.shapes[shape_index])]
            else:
                # All shapes on slide
                target_shapes = [(i, shape) for i, shape in enumerate(slide.shapes)]
            
            # Process each target shape
            for sh_idx, shape in target_shapes:
                if not hasattr(shape, 'text_frame'):
                    continue
                
                if dry_run:
                    # Count occurrences without modifying
                    text = shape.text_frame.text
                    occurrences = 0
                    
                    if match_case:
                        occurrences = text.count(find)
                    else:
                        occurrences = text.lower().count(find.lower())
                    
                    if occurrences > 0:
                        total_count += occurrences
                        preview = text[:100] + "..." if len(text) > 100 else text
                        locations.append({
                            "slide": s_idx,
                            "shape": sh_idx,
                            "occurrences": occurrences,
                            "preview": preview
                        })
                else:
                    # Perform actual replacement
                    replacements = perform_replacement_on_shape(
                        shape, find, replace, match_case
                    )
                    
                    if replacements > 0:
                        total_count += replacements
                        locations.append({
                            "slide": s_idx,
                            "shape": sh_idx,
                            "replacements": replacements
                        })
        
        # Save changes (only for actual replacements)
        if not dry_run:
            agent.save()
            
            # Capture version AFTER
            info_after = agent.get_presentation_info()
            version_after = info_after.get("presentation_version")
    
    # Build result
    result: Dict[str, Any] = {
        "status": "success",
        "file": str(filepath.resolve()),
        "action": action,
        "find": find,
        "replace": replace,
        "match_case": match_case,
        "scope": {
            "slide": slide_index if slide_index is not None else "all",
            "shape": shape_index if shape_index is not None else "all"
        },
        "locations": locations,
        "tool_version": __version__
    }
    
    # Add appropriate count field
    if dry_run:
        result["total_matches"] = total_count
    else:
        result["replacements_made"] = total_count
        result["presentation_version_before"] = version_before
        result["presentation_version_after"] = version_after
    
    # Add performance note for large presentations
    if large_presentation:
        result["note"] = f"Large presentation ({slide_count} slides) processed"
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Find and replace text in PowerPoint",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Preview global replacement (ALWAYS do this first!)
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "Old Company" \\
    --replace "New Company" \\
    --dry-run \\
    --json
  
  # Execute global replacement
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "Old Company" \\
    --replace "New Company" \\
    --json
  
  # Targeted replacement (specific slide)
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --find "Draft" \\
    --replace "Final" \\
    --json
  
  # Surgical replacement (specific shape)
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --find "2024" \\
    --replace "2025" \\
    --json
  
  # Case-sensitive replacement
  uv run tools/ppt_replace_text.py \\
    --file presentation.pptx \\
    --find "API" \\
    --replace "REST API" \\
    --match-case \\
    --json

Scope Options:
  Global (default):     All slides, all shapes
  Slide-specific:       --slide N (single slide, all shapes)
  Shape-specific:       --slide N --shape M (single shape)

Safety Recommendations:
  1. ALWAYS use --dry-run first to preview changes
  2. Clone the presentation before mass replacements:
     uv run tools/ppt_clone_presentation.py --source original.pptx --output work.pptx
  3. Check dry-run output for unexpected matches
  4. Use --slide/--shape to limit scope when appropriate

Replacement Strategy:
  The tool uses a two-tier approach:
  1. Run-level replacement (preserves formatting)
  2. Shape-level fallback (for text split across runs)
  
  This ensures text is replaced even when PowerPoint splits it
  across multiple text runs, while preserving formatting when possible.

Finding Shape Indices:
  Use ppt_get_slide_info.py to identify shapes:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Output Format (dry-run):
  {
    "status": "success",
    "action": "dry_run",
    "find": "Old Company",
    "replace": "New Company",
    "scope": {"slide": "all", "shape": "all"},
    "total_matches": 15,
    "locations": [
      {"slide": 0, "shape": 1, "occurrences": 2, "preview": "Welcome to Old Company..."},
      {"slide": 3, "shape": 4, "occurrences": 1, "preview": "Old Company was founded..."}
    ],
    "tool_version": "3.1.0"
  }

Output Format (replace):
  {
    "status": "success",
    "action": "replace",
    "replacements_made": 15,
    "locations": [...],
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
        '--find', 
        required=True, 
        help='Text to find'
    )
    parser.add_argument(
        '--replace', 
        required=True, 
        help='Replacement text'
    )
    parser.add_argument(
        '--slide', 
        type=int, 
        help='Target specific slide index (0-based)'
    )
    parser.add_argument(
        '--shape', 
        type=int, 
        help='Target specific shape index (requires --slide)'
    )
    parser.add_argument(
        '--match-case', 
        action='store_true', 
        help='Case-sensitive matching'
    )
    parser.add_argument(
        '--dry-run', 
        action='store_true', 
        help='Preview changes without modifying (RECOMMENDED first step)'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = replace_text(
            filepath=args.file,
            find=args.find,
            replace=args.replace,
            slide_index=args.slide,
            shape_index=args.shape,
            match_case=args.match_case,
            dry_run=args.dry_run
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
            "suggestion": "Use ppt_get_info.py to check available slide indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check file format and parameter values"
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
