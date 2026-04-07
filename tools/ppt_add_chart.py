#!/usr/bin/env python3
"""
PowerPoint Add Chart Tool v3.1.0
Add data visualization chart to slide

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_chart.py --file presentation.pptx --slide 1 --chart-type column --data chart_data.json --position '{"left":"10%","top":"20%"}' --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Supported Chart Types:
    column, column_stacked, bar, bar_stacked, line, line_markers,
    pie, area, scatter, doughnut
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
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Supported chart types
CHART_TYPES = [
    'column', 'column_stacked', 'bar', 'bar_stacked',
    'line', 'line_markers', 'pie', 'area', 'scatter', 'doughnut'
]


def add_chart(
    filepath: Path,
    slide_index: int,
    chart_type: str,
    data: Dict[str, Any],
    position: Dict[str, Any],
    size: Dict[str, Any],
    chart_title: Optional[str] = None
) -> Dict[str, Any]:
    """
    Add a data visualization chart to a slide.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the target slide (0-based)
        chart_type: Type of chart (column, bar, line, pie, etc.)
        data: Chart data dict with 'categories' and 'series' keys
        position: Position specification dict
        size: Size specification dict
        chart_title: Optional chart title
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the added chart shape
            - chart_type: Type of chart added
            - chart_title: Title if provided
            - categories: Number of categories
            - series: Number of data series
            - data_points: Total number of data points
            - presentation_version_before: State hash before addition
            - presentation_version_after: State hash after addition
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ValueError: If data format is invalid
        
    Example:
        >>> data = {
        ...     "categories": ["Q1", "Q2", "Q3", "Q4"],
        ...     "series": [{"name": "Revenue", "values": [100, 120, 140, 160]}]
        ... }
        >>> result = add_chart(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=1,
        ...     chart_type="column",
        ...     data=data,
        ...     position={"left": "10%", "top": "20%"},
        ...     size={"width": "80%", "height": "60%"},
        ...     chart_title="Revenue Growth"
        ... )
        >>> print(result["shape_index"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate chart type
    if chart_type not in CHART_TYPES:
        raise ValueError(
            f"Invalid chart type: {chart_type}. "
            f"Supported types: {', '.join(CHART_TYPES)}"
        )
    
    # Validate data structure
    if "categories" not in data:
        raise ValueError(
            "Data must contain 'categories' key. "
            "Example: {\"categories\": [\"Q1\", \"Q2\"], \"series\": [...]}"
        )
    
    if "series" not in data or not data["series"]:
        raise ValueError(
            "Data must contain at least one series. "
            "Example: {\"series\": [{\"name\": \"Sales\", \"values\": [10, 20]}]}"
        )
    
    # Validate all series have same length as categories
    cat_len = len(data["categories"])
    for i, series in enumerate(data["series"]):
        if "values" not in series:
            raise ValueError(f"Series {i} missing 'values' key")
        if len(series.get("values", [])) != cat_len:
            raise ValueError(
                f"Series '{series.get('name', f'[{i}]')}' has {len(series['values'])} values, "
                f"but there are {cat_len} categories. Counts must match."
            )
    
    # Validate pie chart has only one series
    if chart_type in ['pie', 'doughnut'] and len(data["series"]) > 1:
        raise ValueError(
            f"{chart_type.capitalize()} charts support only one data series. "
            f"Found {len(data['series'])} series."
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE addition
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
        
        # Add chart
        result = agent.add_chart(
            slide_index=slide_index,
            chart_type=chart_type,
            data=data,
            position=position,
            size=size,
            chart_title=chart_title
        )
        
        # Extract shape index from result (handle v3.0.x and v3.1.x)
        if isinstance(result, dict):
            shape_index = result.get("shape_index")
        else:
            # Fallback: get last shape index
            slide_info = agent.get_slide_info(slide_index)
            shape_index = slide_info.get("shape_count", 1) - 1
        
        # Save changes
        agent.save()
        
        # Capture version AFTER addition
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "chart_type": chart_type,
        "chart_title": chart_title,
        "categories": len(data["categories"]),
        "series": len(data["series"]),
        "data_points": sum(len(s["values"]) for s in data["series"]),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add data visualization chart to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Chart Types:
  column          Vertical bars (compare across categories)
  column_stacked  Stacked vertical bars (show composition)
  bar             Horizontal bars (compare items)
  bar_stacked     Stacked horizontal bars
  line            Line chart (show trends over time)
  line_markers    Line with data point markers
  pie             Pie chart (show proportions, single series only)
  doughnut        Doughnut chart (pie with hole, single series only)
  area            Area chart (emphasize magnitude of change)
  scatter         Scatter plot (show relationships)

Data Format (JSON file or inline):
{
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [
    {"name": "Revenue", "values": [100, 120, 140, 160]},
    {"name": "Costs", "values": [80, 90, 100, 110]}
  ]
}

Examples:
  # Revenue growth chart from JSON file
  uv run tools/ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --chart-type column \\
    --data revenue_data.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --title "Revenue Growth Trajectory" \\
    --json
  
  # Inline data (short example)
  uv run tools/ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --chart-type column \\
    --data-string '{"categories":["A","B","C"],"series":[{"name":"Sales","values":[10,20,15]}]}' \\
    --position '{"left":"20%","top":"25%"}' \\
    --size '{"width":"60%","height":"50%"}' \\
    --json
  
  # Pie chart (single series)
  uv run tools/ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --chart-type pie \\
    --data-string '{"categories":["Us","Competitor A","Others"],"series":[{"name":"Share","values":[35,40,25]}]}' \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"60%","height":"60%"}' \\
    --title "Market Share" \\
    --json
  
  # Line chart for trends
  uv run tools/ppt_add_chart.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --chart-type line_markers \\
    --data trend_data.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"65%"}' \\
    --title "Monthly Trends" \\
    --json

Chart Selection Guide:
  Compare values across categories  → column or bar
  Show trends over time             → line or line_markers
  Show proportions/percentages      → pie or doughnut
  Show composition over time        → column_stacked or area
  Show correlation between values   → scatter

Best Practices:
  - Use column charts for most comparisons
  - Limit pie charts to 5-7 slices maximum
  - Use line charts for time series data
  - Keep series count to 3-5 for readability
  - Always include a descriptive title
  - Round numbers for better readability

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 1,
    "shape_index": 5,
    "chart_type": "column",
    "chart_title": "Revenue Growth",
    "categories": 4,
    "series": 2,
    "data_points": 8,
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
        '--chart-type',
        required=True,
        choices=CHART_TYPES,
        help='Chart type'
    )
    
    parser.add_argument(
        '--data',
        type=Path,
        help='JSON file with chart data'
    )
    
    parser.add_argument(
        '--data-string',
        help='Inline JSON data string'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=str,
        help='Position dict as JSON string'
    )
    
    parser.add_argument(
        '--size',
        type=str,
        help='Size dict as JSON string'
    )
    
    parser.add_argument(
        '--title',
        help='Chart title'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse position JSON
        try:
            position = json.loads(args.position)
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON in --position: {e}")
        
        # Load chart data
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r') as f:
                data = json.load(f)
        elif args.data_string:
            try:
                data = json.loads(args.data_string)
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON in --data-string: {e}")
        else:
            raise ValueError("Either --data or --data-string is required")
        
        # Parse size JSON or set defaults
        size: Dict[str, Any] = {}
        if args.size:
            try:
                size = json.loads(args.size)
            except json.JSONDecodeError as e:
                raise ValueError(f"Invalid JSON in --size: {e}")
        
        # Handle size from position if not specified
        if "width" in position and "width" not in size:
            size["width"] = position["width"]
        if "height" in position and "height" not in size:
            size["height"] = position["height"]
        
        # Apply defaults
        if "width" not in size:
            size["width"] = "50%"
        if "height" not in size:
            size["height"] = "50%"
        
        result = add_chart(
            filepath=args.file,
            slide_index=args.slide,
            chart_type=args.chart_type,
            data=data,
            position=position,
            size=size,
            chart_title=args.title
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file paths exist and are accessible"
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
            "suggestion": "Check data format and JSON syntax"
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
