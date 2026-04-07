#!/usr/bin/env python3
"""
PowerPoint Create From Structure Tool v3.1.0
Create a complete presentation from a JSON structure definition.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_create_from_structure.py --structure deck.json --output presentation.pptx --json

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
from typing import Dict, Any, List

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError,
)

__version__ = "3.1.0"


def validate_structure(structure: Dict[str, Any]) -> None:
    """
    Validate the JSON structure schema.
    
    Args:
        structure: Dictionary containing presentation structure
        
    Raises:
        ValueError: If structure is invalid
    """
    if "slides" not in structure:
        raise ValueError("Structure must contain 'slides' array")
    
    if not isinstance(structure["slides"], list):
        raise ValueError("'slides' must be an array")
    
    if len(structure["slides"]) == 0:
        raise ValueError("Must have at least one slide")
    
    if len(structure["slides"]) > 100:
        raise ValueError("Maximum 100 slides supported (performance limit)")


def create_from_structure(
    structure: Dict[str, Any],
    output: Path
) -> Dict[str, Any]:
    """
    Create a PowerPoint presentation from a JSON structure definition.
    
    Args:
        structure: Dictionary defining presentation structure with slides and content
        output: Output path for the created presentation (.pptx)
        
    Returns:
        Dict containing:
            - status: 'success' or 'success_with_errors'
            - file: Absolute path to created file
            - presentation_version: Version hash of created presentation
            - slides_created: Number of slides created
            - content_added: Dict with counts per content type
            - errors: List of error messages encountered
            - error_count: Total number of errors
            - file_size_bytes: Size of created file
            - tool_version: Tool version string
            
    Raises:
        ValueError: If structure is invalid
        PowerPointAgentError: If presentation creation fails
    """
    validate_structure(structure)
    
    stats = {
        "slides_created": 0,
        "text_boxes_added": 0,
        "images_inserted": 0,
        "charts_added": 0,
        "tables_added": 0,
        "shapes_added": 0,
        "errors": []
    }
    
    with PowerPointAgent() as agent:
        template = structure.get("template")
        if template and Path(template).exists():
            agent.create_new(template=Path(template))
        else:
            agent.create_new()
        
        for slide_idx, slide_def in enumerate(structure["slides"]):
            try:
                layout = slide_def.get("layout", "Title and Content")
                agent.add_slide(layout_name=layout)
                stats["slides_created"] += 1
                
                if "title" in slide_def:
                    agent.set_title(
                        slide_index=slide_idx,
                        title=slide_def["title"],
                        subtitle=slide_def.get("subtitle")
                    )
                
                for item in slide_def.get("content", []):
                    try:
                        item_type = item.get("type")
                        
                        if item_type == "text_box":
                            agent.add_text_box(
                                slide_index=slide_idx,
                                text=item["text"],
                                position=item["position"],
                                size=item["size"],
                                font_name=item.get("font_name", "Calibri"),
                                font_size=item.get("font_size", 18),
                                bold=item.get("bold", False),
                                italic=item.get("italic", False),
                                color=item.get("color"),
                                alignment=item.get("alignment", "left")
                            )
                            stats["text_boxes_added"] += 1
                        
                        elif item_type == "image":
                            image_path = Path(item["path"])
                            if image_path.exists():
                                agent.insert_image(
                                    slide_index=slide_idx,
                                    image_path=image_path,
                                    position=item["position"],
                                    size=item.get("size"),
                                    compress=item.get("compress", False)
                                )
                                stats["images_inserted"] += 1
                            else:
                                stats["errors"].append(f"Image not found: {item['path']}")
                        
                        elif item_type == "chart":
                            agent.add_chart(
                                slide_index=slide_idx,
                                chart_type=item["chart_type"],
                                data=item["data"],
                                position=item["position"],
                                size=item["size"],
                                chart_title=item.get("title")
                            )
                            stats["charts_added"] += 1
                        
                        elif item_type == "table":
                            agent.add_table(
                                slide_index=slide_idx,
                                rows=item["rows"],
                                cols=item["cols"],
                                position=item["position"],
                                size=item["size"],
                                data=item.get("data")
                            )
                            stats["tables_added"] += 1
                        
                        elif item_type == "shape":
                            agent.add_shape(
                                slide_index=slide_idx,
                                shape_type=item["shape_type"],
                                position=item["position"],
                                size=item["size"],
                                fill_color=item.get("fill_color"),
                                line_color=item.get("line_color"),
                                line_width=item.get("line_width", 1.0)
                            )
                            stats["shapes_added"] += 1
                        
                        elif item_type == "bullet_list":
                            agent.add_bullet_list(
                                slide_index=slide_idx,
                                items=item["items"],
                                position=item["position"],
                                size=item["size"],
                                bullet_style=item.get("bullet_style", "bullet"),
                                font_size=item.get("font_size", 18)
                            )
                            stats["text_boxes_added"] += 1
                        
                        else:
                            stats["errors"].append(f"Unknown content type: {item_type}")
                    
                    except Exception as e:
                        stats["errors"].append(f"Error adding {item.get('type', 'unknown')}: {str(e)}")
            
            except Exception as e:
                stats["errors"].append(f"Error processing slide {slide_idx}: {str(e)}")
        
        agent.save(output)
        
        presentation_version = agent.get_presentation_version()
    
    file_size = output.stat().st_size if output.exists() else 0
    
    return {
        "status": "success" if len(stats["errors"]) == 0 else "success_with_errors",
        "file": str(output.resolve()),
        "presentation_version": presentation_version,
        "slides_created": stats["slides_created"],
        "content_added": {
            "text_boxes": stats["text_boxes_added"],
            "images": stats["images_inserted"],
            "charts": stats["charts_added"],
            "tables": stats["tables_added"],
            "shapes": stats["shapes_added"]
        },
        "errors": stats["errors"],
        "error_count": len(stats["errors"]),
        "file_size_bytes": file_size,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Create PowerPoint presentation from JSON structure definition",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
JSON Structure Schema:
{
  "template": "optional_template.pptx",
  "slides": [
    {
      "layout": "Title Slide",
      "title": "Presentation Title",
      "subtitle": "Subtitle",
      "content": [
        {
          "type": "text_box",
          "text": "Content here",
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "10%"},
          "font_size": 18,
          "color": "#000000"
        },
        {
          "type": "image",
          "path": "image.png",
          "position": {"left": "20%", "top": "30%"},
          "size": {"width": "60%", "height": "auto"}
        },
        {
          "type": "chart",
          "chart_type": "column",
          "data": {
            "categories": ["Q1", "Q2", "Q3"],
            "series": [{"name": "Revenue", "values": [100, 120, 140]}]
          },
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "60%"}
        },
        {
          "type": "table",
          "rows": 3,
          "cols": 3,
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "50%"},
          "data": [["A", "B", "C"], ["1", "2", "3"]]
        },
        {
          "type": "shape",
          "shape_type": "rectangle",
          "position": {"left": "10%", "top": "10%"},
          "size": {"width": "30%", "height": "15%"},
          "fill_color": "#0070C0"
        },
        {
          "type": "bullet_list",
          "items": ["Item 1", "Item 2", "Item 3"],
          "position": {"left": "10%", "top": "25%"},
          "size": {"width": "80%", "height": "60%"}
        }
      ]
    }
  ]
}

Examples:
    # Create simple presentation
    cat > structure.json << 'EOF'
{
  "slides": [
    {
      "layout": "Title Slide",
      "title": "My Presentation",
      "subtitle": "Created from Structure"
    },
    {
      "layout": "Title and Content",
      "title": "Agenda",
      "content": [
        {
          "type": "bullet_list",
          "items": ["Introduction", "Main Content", "Conclusion"],
          "position": {"left": "10%", "top": "25%"},
          "size": {"width": "80%", "height": "60%"}
        }
      ]
    }
  ]
}
EOF
    
    uv run tools/ppt_create_from_structure.py \\
        --structure structure.json \\
        --output presentation.pptx \\
        --json

    # Create presentation with charts
    cat > complex.json << 'EOF'
{
  "slides": [
    {
      "layout": "Title Slide",
      "title": "Q4 Report"
    },
    {
      "layout": "Title and Content",
      "title": "Revenue Growth",
      "content": [
        {
          "type": "chart",
          "chart_type": "column",
          "data": {
            "categories": ["Q1", "Q2", "Q3", "Q4"],
            "series": [
              {"name": "2023", "values": [100, 110, 120, 130]},
              {"name": "2024", "values": [120, 135, 145, 160]}
            ]
          },
          "position": {"left": "10%", "top": "20%"},
          "size": {"width": "80%", "height": "65%"},
          "title": "Year over Year Comparison"
        }
      ]
    }
  ]
}
EOF
    
    uv run tools/ppt_create_from_structure.py \\
        --structure complex.json \\
        --output q4_report.pptx \\
        --json

Content Types:
    text_box    - Text container with formatting options
    image       - Image from file path
    chart       - Data visualization (column, bar, line, pie)
    table       - Grid of cells
    shape       - Geometric shape (rectangle, arrow, etc.)
    bullet_list - Bulleted or numbered list

Use Cases:
    - Automated report generation
    - Template-based presentations from data
    - Batch presentation creation
    - AI-generated presentations
    - Programmatic deck building
        """
    )
    
    parser.add_argument(
        '--structure',
        required=True,
        type=Path,
        help='Path to JSON structure file'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output path for created presentation (.pptx)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output as JSON (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        if not args.structure.exists():
            raise FileNotFoundError(f"Structure file not found: {args.structure}")
        
        with open(args.structure, 'r', encoding='utf-8') as f:
            structure = json.load(f)
        
        output_path = args.output
        if output_path.suffix.lower() != '.pptx':
            output_path = output_path.with_suffix('.pptx')
        
        result = create_from_structure(
            structure=structure,
            output=output_path
        )
        
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the structure file path exists and is accessible."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except json.JSONDecodeError as e:
        error_result = {
            "status": "error",
            "error": f"Invalid JSON in structure file: {str(e)}",
            "error_type": "JSONDecodeError",
            "suggestion": "Validate JSON syntax using a JSON linter."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check structure file matches the required schema."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Check template file exists and is valid if specified."
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
