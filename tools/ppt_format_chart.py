#!/usr/bin/env python3
"""
PowerPoint Format Chart Tool v3.1.0
Format existing chart (title, legend position)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_format_chart.py --file presentation.pptx --slide 1 --chart 0 --title "Revenue Growth" --legend bottom --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Limitations:
    python-pptx has limited chart formatting support. This tool handles:
    - Chart title text
    - Legend position
    
    Not supported (requires PowerPoint):
    - Individual series colors
    - Axis formatting
    - Data labels
    - Chart styles/templates
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

__version__ = "3.1.0"

# Legend positions
LEGEND_POSITIONS = ['bottom', 'left', 'right', 'top', 'none']


def format_chart(
    filepath: Path,
    slide_index: int,
    chart_index: int = 0,
    title: Optional[str] = None,
    legend_position: Optional[str] = None
) -> Dict[str, Any]:
    """
    Format an existing chart on a slide.
    
    Updates chart title and/or legend position. At least one
    formatting option must be specified.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the chart (0-based)
        chart_index: Index of the chart on the slide (0-based, default: 0)
        title: New chart title text (optional)
        legend_position: Legend position - 'bottom', 'left', 'right', 'top', 'none'
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - chart_index: Index of the chart
            - formatting_applied: Dict with applied formatting
            - presentation_version_before: State hash before formatting
            - presentation_version_after: State hash after formatting
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ValueError: If no formatting options specified or chart not found
        
    Example:
        >>> result = format_chart(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=1,
        ...     chart_index=0,
        ...     title="Revenue Growth Trend",
        ...     legend_position="bottom"
        ... )
        >>> print(result["formatting_applied"]["title"])
        'Revenue Growth Trend'
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate at least one option is specified
    if title is None and legend_position is None:
        raise ValueError(
            "At least one formatting option (--title or --legend) must be specified"
        )
    
    # Validate legend position if provided
    if legend_position is not None and legend_position not in LEGEND_POSITIONS:
        raise ValueError(
            f"Invalid legend position: {legend_position}. "
            f"Valid options: {', '.join(LEGEND_POSITIONS)}"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE formatting
        info_before = agent.get_presentation_info()
        version_before = info_before.get("presentation_version")
        
        # Validate slide index
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={
                    "requested_index": slide_index,
                    "available_slides": total_slides
                }
            )
        
        # Get slide info to check for charts
        slide_info = agent.get_slide_info(slide_index)
        
        # Count charts on slide (check shapes for chart type)
        chart_count = 0
        for shape in slide_info.get("shapes", []):
            if shape.get("type") == "CHART" or shape.get("has_chart", False):
                chart_count += 1
        
        # If we couldn't detect charts from slide_info, try the operation anyway
        # The core method will raise if chart doesn't exist
        if chart_count == 0:
            # Could be that slide_info doesn't expose chart detection
            # Let the core method handle validation
            pass
        elif not 0 <= chart_index < chart_count:
            raise ValueError(
                f"Chart index {chart_index} out of range. "
                f"Slide has {chart_count} chart(s) (indices 0-{chart_count - 1})."
            )
        
        # Format chart
        agent.format_chart(
            slide_index=slide_index,
            chart_index=chart_index,
            title=title,
            legend_position=legend_position
        )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER formatting
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Build formatting applied dict
    formatting_applied: Dict[str, Any] = {}
    if title is not None:
        formatting_applied["title"] = title
    if legend_position is not None:
        formatting_applied["legend_position"] = legend_position
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "chart_index": chart_index,
        "formatting_applied": formatting_applied,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Format PowerPoint chart (title, legend)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set chart title
  uv run tools/ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart 0 \\
    --title "Revenue Growth Trend" \\
    --json
  
  # Position legend at bottom
  uv run tools/ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --chart 0 \\
    --legend bottom \\
    --json
  
  # Set both title and legend
  uv run tools/ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --chart 0 \\
    --title "Q4 Performance" \\
    --legend right \\
    --json
  
  # Hide legend
  uv run tools/ppt_format_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart 0 \\
    --legend none \\
    --json

Legend Positions:
  bottom  Below chart (common for wide charts)
  right   Right side of chart (default)
  top     Above chart
  left    Left side of chart
  none    Hide legend entirely

Finding Charts:
  Charts are indexed in order they appear on the slide (0, 1, 2...).
  Use ppt_get_slide_info.py to find charts:
  uv run tools/ppt_get_slide_info.py --file presentation.pptx --slide 1 --json

Best Practices:
  - Keep titles concise and descriptive
  - Use 'bottom' legend for wide charts
  - Use 'right' legend for tall charts
  - Hide legend if only one series
  - Match title to chart type (e.g., "Trend" for line charts)

⚠️ Formatting Limitations:
  python-pptx has limited chart formatting support.
  
  Supported by this tool:
  ✓ Chart title text
  ✓ Legend position
  
  Not supported (use PowerPoint directly):
  ✗ Individual series colors
  ✗ Axis formatting (labels, scale)
  ✗ Data labels
  ✗ Chart styles/templates
  ✗ Gridlines

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 1,
    "chart_index": 0,
    "formatting_applied": {
      "title": "Revenue Growth Trend",
      "legend_position": "bottom"
    },
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
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--chart',
        type=int,
        default=0,
        help='Chart index on slide (default: 0)'
    )
    
    parser.add_argument(
        '--title',
        help='Chart title text'
    )
    
    parser.add_argument(
        '--legend',
        choices=LEGEND_POSITIONS,
        help='Legend position (bottom, left, right, top, none)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = format_chart(
            filepath=args.file,
            slide_index=args.slide,
            chart_index=args.chart,
            title=args.title,
            legend_position=args.legend
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
            "suggestion": "Specify --title and/or --legend, and verify chart index"
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
