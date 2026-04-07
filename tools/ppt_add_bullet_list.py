#!/usr/bin/env python3
"""
PowerPoint Add Bullet List Tool v3.1.0
Add bullet or numbered list with 6×6 rule validation and accessibility checks.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 1 \\
        --items "Point 1,Point 2,Point 3" \\
        --position '{"left":"10%","top":"25%"}' \\
        --size '{"width":"80%","height":"60%"}' --json

Exit Codes:
    0: Success
    1: Error occurred

6×6 Rule (Best Practice):
    - Maximum 6 bullet points per slide
    - Maximum 6 words per line (~60 characters)
    - Ensures readability and audience engagement
"""

import sys
import os

sys.stderr = open(os.devnull, 'w')

import json
import argparse
from pathlib import Path
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
    ColorHelper,
)
from pptx.dml.color import RGBColor

__version__ = "3.1.0"


def calculate_readability_score(items: List[str]) -> Dict[str, Any]:
    """Calculate readability metrics for bullet list."""
    total_chars = sum(len(item) for item in items)
    avg_chars = total_chars / len(items) if items else 0
    max_chars = max(len(item) for item in items) if items else 0
    
    total_words = sum(len(item.split()) for item in items)
    avg_words = total_words / len(items) if items else 0
    max_words = max(len(item.split()) for item in items) if items else 0
    
    score = 100
    issues = []
    
    if len(items) > 6:
        score -= (len(items) - 6) * 10
        issues.append(f"Exceeds 6×6 rule: {len(items)} items (recommended: ≤6)")
    
    if avg_chars > 60:
        score -= 20
        issues.append(f"Items too long: {avg_chars:.0f} chars average (recommended: ≤60)")
    
    if max_chars > 100:
        score -= 10
        issues.append(f"Longest item: {max_chars} chars (consider splitting)")
    
    if max_words > 12:
        score -= 15
        issues.append(f"Too many words per item: {max_words} max (recommended: ≤10)")
    
    score = max(0, score)
    
    return {
        "score": score,
        "grade": "A" if score >= 90 else "B" if score >= 75 else "C" if score >= 60 else "D" if score >= 50 else "F",
        "metrics": {
            "item_count": len(items),
            "avg_characters": round(avg_chars, 1),
            "max_characters": max_chars,
            "avg_words": round(avg_words, 1),
            "max_words": max_words
        },
        "issues": issues
    }


def add_bullet_list(
    filepath: Path,
    slide_index: int,
    items: List[str],
    position: Dict[str, Any],
    size: Dict[str, Any],
    bullet_style: str = "bullet",
    font_size: int = 18,
    font_name: str = "Calibri",
    color: str = None,
    line_spacing: float = 1.0,
    ignore_rules: bool = False
) -> Dict[str, Any]:
    """
    Add bullet or numbered list with validation.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Slide index (0-based)
        items: List of bullet items
        position: Position dict
        size: Size dict
        bullet_style: "bullet", "numbered", or "none"
        font_size: Font size in points
        font_name: Font name
        color: Optional text color (hex)
        line_spacing: Line spacing multiplier
        ignore_rules: Override 6×6 rule validation
        
    Returns:
        Dict with results and validation info
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If invalid parameters
        SlideNotFoundError: If slide index out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    if not items:
        raise ValueError("At least one item required")
    
    warnings = []
    recommendations = []
    
    readability = calculate_readability_score(items)
    
    if len(items) > 6 and not ignore_rules:
        warnings.append(
            f"6×6 Rule violation: {len(items)} items exceeds recommended 6 per slide. "
            "This reduces readability and audience engagement."
        )
        recommendations.append(
            "Consider splitting into multiple slides or using --ignore-rules to override"
        )
    
    if len(items) > 10 and not ignore_rules:
        raise ValueError(
            f"Too many items: {len(items)} exceeds hard limit of 10 per slide. "
            "Split into multiple slides or use --ignore-rules to override."
        )
    
    for idx, item in enumerate(items):
        if len(item) > 100:
            warnings.append(
                f"Item {idx + 1} is {len(item)} characters (very long). "
                "Consider breaking into multiple bullets."
            )
    
    if font_size < 14:
        warnings.append(
            f"Font size {font_size}pt is below recommended minimum of 14pt."
        )
    
    if color:
        try:
            text_color = ColorHelper.from_hex(color)
            bg_color = RGBColor(255, 255, 255)
            is_large_text = font_size >= 18
            
            if not ColorHelper.meets_wcag(text_color, bg_color, is_large_text):
                contrast_ratio = ColorHelper.contrast_ratio(text_color, bg_color)
                required_ratio = 3.0 if is_large_text else 4.5
                warnings.append(
                    f"Color contrast {contrast_ratio:.2f}:1 may not meet WCAG accessibility "
                    f"(required: {required_ratio}:1)."
                )
        except Exception:
            pass
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        agent.add_bullet_list(
            slide_index=slide_index,
            items=items,
            position=position,
            size=size,
            bullet_style=bullet_style,
            font_size=font_size
        )
        
        slide_info = agent.get_slide_info(slide_index)
        last_shape_idx = slide_info["shape_count"] - 1
        
        if color:
            try:
                agent.format_text(
                    slide_index=slide_index,
                    shape_index=last_shape_idx,
                    color=color
                )
            except Exception as e:
                warnings.append(f"Could not apply color: {str(e)}")
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    if readability["score"] < 75:
        recommendations.append(
            f"Readability score is {readability['grade']} ({readability['score']}/100). "
            "Consider simplifying content."
        )
    
    status = "success"
    if warnings:
        status = "warning"
    
    result: Dict[str, Any] = {
        "status": status,
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "items_added": len(items),
        "items": items,
        "bullet_style": bullet_style,
        "formatting": {
            "font_size": font_size,
            "font_name": font_name,
            "color": color,
            "line_spacing": line_spacing
        },
        "readability": readability,
        "validation": {
            "six_six_rule": {
                "compliant": len(items) <= 6 and readability["metrics"]["max_words"] <= 10,
                "item_count_ok": len(items) <= 6,
                "word_count_ok": readability["metrics"]["max_words"] <= 10
            },
            "accessibility": {
                "font_size_ok": font_size >= 14,
                "color_contrast_checked": color is not None
            }
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if warnings:
        result["warnings"] = warnings
    
    if recommendations:
        result["recommendations"] = recommendations
    
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add bullet/numbered list with 6×6 rule validation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
6×6 Rule (Best Practice):
  - Maximum 6 bullet points per slide
  - Maximum 6 words per line (~60 characters)
  - Ensures readability and audience engagement

Examples:
  # Simple bullet list
  uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 1 \\
    --items "Revenue up 45%,Customer growth 60%,Market share increased" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' --json

  # Numbered list
  uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 2 \\
    --items "Define objectives,Analyze market,Execute plan" \\
    --bullet-style numbered --font-size 20 --color "#0070C0" \\
    --position '{"left":"15%","top":"30%"}' \\
    --size '{"width":"70%","height":"50%"}' --json

  # From JSON file
  echo '["First point", "Second point"]' > items.json
  uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 3 \\
    --items-file items.json --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' --json
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--items', help='Comma-separated list items')
    parser.add_argument('--items-file', type=Path, help='JSON file with array of items')
    parser.add_argument('--position', required=True, type=json.loads, help='Position dict (JSON)')
    parser.add_argument('--size', type=json.loads, help='Size dict (JSON)')
    parser.add_argument('--bullet-style', choices=['bullet', 'numbered', 'none'], default='bullet')
    parser.add_argument('--font-size', type=int, default=18, help='Font size (default: 18)')
    parser.add_argument('--font-name', default='Calibri', help='Font name')
    parser.add_argument('--color', help='Text color hex (e.g., #0070C0)')
    parser.add_argument('--line-spacing', type=float, default=1.0, help='Line spacing')
    parser.add_argument('--ignore-rules', action='store_true', help='Override 6×6 validation')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        if args.items_file:
            if not args.items_file.exists():
                raise FileNotFoundError(f"Items file not found: {args.items_file}")
            with open(args.items_file, 'r', encoding='utf-8') as f:
                items = json.load(f)
            if not isinstance(items, list):
                raise ValueError("Items file must contain JSON array")
        elif args.items:
            if '\\n' in args.items:
                items = args.items.split('\\n')
            else:
                items = args.items.split(',')
            items = [item.strip() for item in items if item.strip()]
        else:
            raise ValueError("Either --items or --items-file required")
        
        size = args.size if args.size else {}
        position = args.position
        
        if "width" not in size:
            size["width"] = position.get("width", "80%")
        if "height" not in size:
            size["height"] = position.get("height", "50%")
        
        result = add_bullet_list(
            filepath=args.file,
            slide_index=args.slide,
            items=items,
            position=position,
            size=size,
            bullet_style=args.bullet_style,
            font_size=args.font_size,
            font_name=args.font_name,
            color=args.color,
            line_spacing=args.line_spacing,
            ignore_rules=args.ignore_rules
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "SlideNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check items format and file extension (.pptx required)."
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
