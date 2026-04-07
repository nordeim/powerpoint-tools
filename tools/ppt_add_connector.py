#!/usr/bin/env python3
"""
PowerPoint Add Connector Tool v3.1.0
Draw a line/connector between two shapes on a slide

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_add_connector.py --file deck.pptx --slide 0 --from-shape 0 --to-shape 1 --type straight --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Flowcharts and process diagrams
    - Org charts
    - Network diagrams
    - Relationship mapping
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

# Supported connector types
CONNECTOR_TYPES = ['straight', 'elbow', 'curve']

# Define fallback exception
try:
    from core.powerpoint_agent_core import ShapeNotFoundError
except ImportError:
    class ShapeNotFoundError(PowerPointAgentError):
        """Exception raised when shape is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def add_connector(
    filepath: Path,
    slide_index: int,
    from_shape: int,
    to_shape: int,
    connector_type: str = "straight",
    line_color: Optional[str] = None,
    line_width: Optional[float] = None
) -> Dict[str, Any]:
    """
    Add a connector line between two shapes on a slide.
    
    Creates a line that visually connects two shapes, useful for
    flowcharts, org charts, and process diagrams.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the shapes (0-based)
        from_shape: Index of the starting shape (0-based)
        to_shape: Index of the ending shape (0-based)
        connector_type: Type of connector ('straight', 'elbow', 'curve')
        line_color: Optional line color in hex format (e.g., "#000000")
        line_width: Optional line width in points
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the new connector shape
            - connection: Dict with from, to, and type info
            - presentation_version_before: State hash before addition
            - presentation_version_after: State hash after addition
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ShapeNotFoundError: If from_shape or to_shape index is invalid
        ValueError: If connector type is invalid
        
    Example:
        >>> result = add_connector(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     from_shape=0,
        ...     to_shape=1,
        ...     connector_type="straight"
        ... )
        >>> print(result["shape_index"])
        5
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Validate connector type
    if connector_type not in CONNECTOR_TYPES:
        raise ValueError(
            f"Invalid connector type: {connector_type}. "
            f"Supported types: {', '.join(CONNECTOR_TYPES)}"
        )
    
    # Validate from and to are different
    if from_shape == to_shape:
        raise ValueError(
            "Cannot connect a shape to itself. "
            "from_shape and to_shape must be different."
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
        
        # Get slide info to validate shape indices
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", 0)
        
        # Validate from_shape
        if not 0 <= from_shape < shape_count:
            raise ShapeNotFoundError(
                f"from_shape index {from_shape} out of range (0-{shape_count - 1})",
                details={
                    "requested_index": from_shape,
                    "available_shapes": shape_count,
                    "parameter": "from_shape"
                }
            )
        
        # Validate to_shape
        if not 0 <= to_shape < shape_count:
            raise ShapeNotFoundError(
                f"to_shape index {to_shape} out of range (0-{shape_count - 1})",
                details={
                    "requested_index": to_shape,
                    "available_shapes": shape_count,
                    "parameter": "to_shape"
                }
            )
        
        # Add connector
        result = agent.add_connector(
            slide_index=slide_index,
            from_shape=from_shape,
            to_shape=to_shape,
            connector_type=connector_type,
            line_color=line_color,
            line_width=line_width
        )
        
        # Extract shape index from result
        if isinstance(result, dict):
            connector_index = result.get("shape_index", result.get("connector_index"))
        else:
            # Fallback: new shape is at end
            updated_info = agent.get_slide_info(slide_index)
            connector_index = updated_info.get("shape_count", 1) - 1
        
        # Save changes
        agent.save()
        
        # Capture version AFTER addition
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": connector_index,
        "connection": {
            "from_shape": from_shape,
            "to_shape": to_shape,
            "type": connector_type,
            "line_color": line_color,
            "line_width": line_width
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Add connector line between shapes in PowerPoint",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Connector Types:
  straight  Direct line between shapes (default)
  elbow     Right-angle connector (90-degree bends)
  curve     Curved/bezier connector

Examples:
  # Simple straight connector
  uv run tools/ppt_add_connector.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --from-shape 0 \\
    --to-shape 1 \\
    --json
  
  # Elbow connector with styling
  uv run tools/ppt_add_connector.py \\
    --file flowchart.pptx \\
    --slide 2 \\
    --from-shape 3 \\
    --to-shape 5 \\
    --type elbow \\
    --color "#0070C0" \\
    --width 2.0 \\
    --json
  
  # Curved connector
  uv run tools/ppt_add_connector.py \\
    --file diagram.pptx \\
    --slide 1 \\
    --from-shape 0 \\
    --to-shape 2 \\
    --type curve \\
    --json

Finding Shape Indices:
  Use ppt_get_slide_info.py to identify shape indices:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Use Cases:
  - Flowcharts: Connect process steps
  - Org charts: Connect hierarchy levels
  - Network diagrams: Show connections
  - Mind maps: Connect ideas
  - Process flows: Show sequence

Best Practices:
  - Use straight connectors for simple diagrams
  - Use elbow connectors for flowcharts (cleaner appearance)
  - Use curved connectors for org charts
  - Keep connector colors consistent with theme
  - Add shapes before connecting them

⚠️ Shape Index Warning:
  After adding a connector, shape indices may change.
  Always refresh shape indices using ppt_get_slide_info.py
  before performing additional operations.

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "shape_index": 5,
    "connection": {
      "from_shape": 0,
      "to_shape": 1,
      "type": "straight"
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
        '--from-shape', 
        required=True, 
        type=int, 
        help='Starting shape index (0-based)'
    )
    parser.add_argument(
        '--to-shape', 
        required=True, 
        type=int, 
        help='Ending shape index (0-based)'
    )
    parser.add_argument(
        '--type', 
        choices=CONNECTOR_TYPES,
        default='straight', 
        help='Connector type (default: straight)'
    )
    parser.add_argument(
        '--color',
        help='Line color in hex format (e.g., "#000000")'
    )
    parser.add_argument(
        '--width',
        type=float,
        help='Line width in points'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = add_connector(
            filepath=args.file,
            slide_index=args.slide,
            from_shape=args.from_shape,
            to_shape=args.to_shape,
            connector_type=args.type,
            line_color=args.color,
            line_width=args.width
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
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shape indices"
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": f"Check connector type (supported: {', '.join(CONNECTOR_TYPES)})"
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
