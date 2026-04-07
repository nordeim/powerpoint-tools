#!/usr/bin/env python3
"""
PowerPoint Add Table Tool v3.1.0
Add data table to slide with comprehensive validation.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_table.py --file presentation.pptx --slide 1 --rows 5 --cols 3 \\
        --data table_data.json --position '{"left":"10%","top":"25%"}' \\
        --size '{"width":"80%","height":"50%"}' --json

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
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
    SlideNotFoundError,
)

__version__ = "3.1.0"


def validate_table_params(
    rows: int,
    cols: int,
    position: Dict[str, Any],
    size: Dict[str, Any],
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Validate table parameters and return warnings/recommendations.
    
    Args:
        rows: Number of rows
        cols: Number of columns
        position: Position specification dict
        size: Size specification dict
        allow_offslide: Whether to allow off-slide positioning
        
    Returns:
        Dict with warnings, recommendations, and validation_results
    """
    warnings: List[str] = []
    recommendations: List[str] = []
    validation_results: Dict[str, Any] = {}
    
    if position:
        try:
            if "left" in position:
                left_str = str(position["left"])
                if left_str.endswith('%'):
                    left_pct = float(left_str.rstrip('%'))
                    if (left_pct < 0 or left_pct > 100) and not allow_offslide:
                        warnings.append(
                            f"Left position {left_pct}% is outside slide bounds (0-100%). "
                            "Table may not be visible. Use --allow-offslide if intentional."
                        )
            
            if "top" in position:
                top_str = str(position["top"])
                if top_str.endswith('%'):
                    top_pct = float(top_str.rstrip('%'))
                    if (top_pct < 0 or top_pct > 100) and not allow_offslide:
                        warnings.append(
                            f"Top position {top_pct}% is outside slide bounds (0-100%). "
                            "Table may not be visible. Use --allow-offslide if intentional."
                        )
        except (ValueError, TypeError):
            pass
    
    if size:
        try:
            if "height" in size:
                height_str = str(size["height"])
                if height_str.endswith('%'):
                    height_pct = float(height_str.rstrip('%'))
                    min_height = rows * 2
                    if height_pct < min_height:
                        warnings.append(
                            f"Table height {height_pct}% is very small for {rows} rows "
                            f"(recommended: >{min_height}%). Text may be unreadable."
                        )
            
            if "width" in size:
                width_str = str(size["width"])
                if width_str.endswith('%'):
                    width_pct = float(width_str.rstrip('%'))
                    min_width = cols * 5
                    if width_pct < min_width:
                        warnings.append(
                            f"Table width {width_pct}% is very small for {cols} columns "
                            f"(recommended: >{min_width}%). Text may be unreadable."
                        )
        except (ValueError, TypeError):
            pass
            
    return {
        "warnings": warnings,
        "recommendations": recommendations,
        "validation_results": validation_results
    }


def add_table(
    filepath: Path,
    slide_index: int,
    rows: int,
    cols: int,
    position: Dict[str, Any],
    size: Dict[str, Any],
    data: Optional[List[List[Any]]] = None,
    headers: Optional[List[str]] = None,
    allow_offslide: bool = False
) -> Dict[str, Any]:
    """
    Add table to slide with validation.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        rows: Number of rows (including header row if headers provided)
        cols: Number of columns
        position: Position specification dict
        size: Size specification dict
        data: Optional 2D list of cell values
        headers: Optional list of header strings
        allow_offslide: Allow positioning outside slide bounds
        
    Returns:
        Dict with operation results, validation info, and version tracking
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If parameters are invalid
        SlideNotFoundError: If slide index out of range
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    validation = validate_table_params(rows, cols, position, size, allow_offslide)
    
    if rows < 1 or cols < 1:
        raise ValueError("Table must have at least 1 row and 1 column")
    
    if rows > 50 or cols > 20:
        raise ValueError("Maximum table size: 50 rows × 20 columns (readability limit)")
    
    table_data: List[List[Any]] = []
    
    if headers:
        if len(headers) != cols:
            raise ValueError(f"Headers count ({len(headers)}) must match columns ({cols})")
        table_data.append(headers)
        data_rows = rows - 1
    else:
        data_rows = rows
    
    if data:
        if len(data) > data_rows:
            raise ValueError(
                f"Too many data rows ({len(data)}) for table size ({data_rows} data rows)"
            )
        
        for row_idx, row in enumerate(data):
            if len(row) != cols:
                raise ValueError(
                    f"Data row {row_idx} has {len(row)} items, expected {cols}"
                )
            table_data.append([str(cell) for cell in row])
        
        while len(table_data) < rows:
            table_data.append([""] * cols)
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        total_slides = agent.get_slide_count()
        if not 0 <= slide_index < total_slides:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{total_slides - 1})",
                details={"requested": slide_index, "available": total_slides}
            )
        
        agent.add_table(
            slide_index=slide_index,
            rows=rows,
            cols=cols,
            position=position,
            size=size,
            data=table_data if table_data else None
        )
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    result: Dict[str, Any] = {
        "status": "success" if not validation["warnings"] else "warning",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "rows": rows,
        "cols": cols,
        "has_headers": headers is not None,
        "data_rows_filled": len(data) if data else 0,
        "total_cells": rows * cols,
        "validation": validation["validation_results"],
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }
    
    if validation["warnings"]:
        result["warnings"] = validation["warnings"]
        
    if validation["recommendations"]:
        result["recommendations"] = validation["recommendations"]
        
    return result


def main():
    parser = argparse.ArgumentParser(
        description="Add data table to PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Data Format (JSON):
  - 2D array: [["A1","B1","C1"], ["A2","B2","C2"]]
  - CSV file: converted to 2D array
  - Pandas DataFrame: exported to JSON array

Examples:
  # Simple pricing table
  cat > pricing.json << 'EOF'
[
  ["Starter", "$9/mo", "Basic features"],
  ["Pro", "$29/mo", "Advanced features"],
  ["Enterprise", "$99/mo", "All features + support"]
]
EOF
  
  uv run tools/ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 3 \\
    --rows 4 \\
    --cols 3 \\
    --headers "Plan,Price,Features" \\
    --data pricing.json \\
    --position '{"left":"15%","top":"25%"}' \\
    --size '{"width":"70%","height":"50%"}' \\
    --json
  
  # Quarterly results table
  cat > results.json << 'EOF'
[
  ["Q1", "10.5", "8.2", "2.3"],
  ["Q2", "12.8", "9.1", "3.7"],
  ["Q3", "15.2", "10.5", "4.7"],
  ["Q4", "18.6", "12.1", "6.5"]
]
EOF
  
  uv run tools/ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 4 \\
    --rows 5 \\
    --cols 4 \\
    --headers "Quarter,Revenue,Costs,Profit" \\
    --data results.json \\
    --position '{"left":"10%","top":"20%"}' \\
    --size '{"width":"80%","height":"55%"}' \\
    --json
  
  # Comparison table (centered)
  cat > comparison.json << 'EOF'
[
  ["Speed", "Fast", "Very Fast"],
  ["Security", "Standard", "Enterprise"],
  ["Support", "Email", "24/7 Phone"]
]
EOF
  
  uv run tools/ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --rows 4 \\
    --cols 3 \\
    --headers "Feature,Basic,Premium" \\
    --data comparison.json \\
    --position '{"anchor":"center"}' \\
    --size '{"width":"60%","height":"40%"}' \\
    --json
  
  # Empty table (for manual filling)
  uv run tools/ppt_add_table.py \\
    --file presentation.pptx \\
    --slide 6 \\
    --rows 6 \\
    --cols 4 \\
    --headers "Name,Role,Department,Email" \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"60%"}' \\
    --json

Best Practices:
  - Keep tables under 10 rows for readability
  - Use headers for all tables
  - Align numbers right, text left
  - Use consistent decimal places
  - Highlight key values with color
  - Leave white space around table
  - Use alternating row colors for large tables

Table Size Guidelines:
  - 3-5 columns: Optimal for most presentations
  - 6-10 rows: Maximum for comfortable reading
  - Font size: 12-16pt for body, 14-18pt for headers
  - Cell padding: Leave breathing room

When to Use Tables vs Charts:
  - Use tables: Exact values matter, detailed data
  - Use charts: Show trends, comparisons, patterns
  - Use both: Table with summary chart
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path (.pptx)'
    )
    
    parser.add_argument(
        '--slide',
        required=True,
        type=int,
        help='Slide index (0-based)'
    )
    
    parser.add_argument(
        '--rows',
        required=True,
        type=int,
        help='Number of rows (including header if present)'
    )
    
    parser.add_argument(
        '--cols',
        required=True,
        type=int,
        help='Number of columns'
    )
    
    parser.add_argument(
        '--position',
        required=True,
        type=json.loads,
        help='Position dict (JSON string)'
    )
    
    parser.add_argument(
        '--size',
        required=True,
        type=json.loads,
        help='Size dict (JSON string)'
    )
    
    parser.add_argument(
        '--data',
        type=Path,
        help='JSON file with 2D array of cell values'
    )
    
    parser.add_argument(
        '--data-string',
        help='Inline JSON 2D array string'
    )
    
    parser.add_argument(
        '--headers',
        help='Comma-separated header row (will be row 0)'
    )
    
    parser.add_argument(
        '--allow-offslide',
        action='store_true',
        help='Allow positioning outside slide bounds'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        headers = None
        if args.headers:
            headers = [h.strip() for h in args.headers.split(',')]
        
        data = None
        if args.data:
            if not args.data.exists():
                raise FileNotFoundError(f"Data file not found: {args.data}")
            with open(args.data, 'r', encoding='utf-8') as f:
                data = json.load(f)
        elif args.data_string:
            data = json.loads(args.data_string)
        
        if data is not None:
            if not isinstance(data, list):
                raise ValueError("Data must be a 2D array (list of lists)")
            if data and not isinstance(data[0], list):
                raise ValueError("Data must be a 2D array (list of lists)")
        
        result = add_table(
            filepath=args.file,
            slide_index=args.slide,
            rows=args.rows,
            cols=args.cols,
            position=args.position,
            size=args.size,
            data=data,
            headers=headers,
            allow_offslide=args.allow_offslide
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
            "suggestion": "Use ppt_get_info.py to check available slides."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check table dimensions, data format, and position/size JSON."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON: {str(e)}",
            "error_type": "JSONDecodeError",
            "suggestion": "Validate JSON syntax. Use single quotes around JSON strings."
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
