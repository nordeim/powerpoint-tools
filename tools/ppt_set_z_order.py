#!/usr/bin/env python3
"""
PowerPoint Set Z-Order Tool v3.1.0
Manage shape layering (Bring to Front, Send to Back, etc.).

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_z_order.py --file deck.pptx --slide 0 --shape 1 --action bring_to_front --json

Exit Codes:
    0: Success
    1: Error occurred

⚠️  IMPORTANT: Shape indices change after z-order operations!
    Always refresh indices with ppt_get_slide_info.py before targeting shapes.
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
    ShapeNotFoundError,
)

__version__ = "3.1.0"


def _validate_xml_structure(sp_tree) -> bool:
    """Validate XML tree integrity after manipulation."""
    return all(child is not None for child in sp_tree)


def set_z_order(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    action: str
) -> Dict[str, Any]:
    """
    Change the Z-order (stacking order) of a shape.
    
    Args:
        filepath: Path to PowerPoint file (.pptx)
        slide_index: Target slide index (0-based)
        shape_index: Target shape index (0-based)
        action: One of 'bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'
        
    Returns:
        Result dict with z-order change details
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ValueError: If file format invalid or invalid action
        SlideNotFoundError: If slide index invalid
        ShapeNotFoundError: If shape index invalid
        PowerPointAgentError: If XML manipulation fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if filepath.suffix.lower() != '.pptx':
        raise ValueError("Only .pptx files are supported")
    
    valid_actions = ['bring_to_front', 'send_to_back', 'bring_forward', 'send_backward']
    if action not in valid_actions:
        raise ValueError(f"Invalid action '{action}'. Must be one of: {valid_actions}")
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        version_before = agent.get_presentation_version()
        
        slide_count = agent.get_slide_count()
        if not 0 <= slide_index < slide_count:
            raise SlideNotFoundError(
                f"Slide index {slide_index} out of range (0-{slide_count - 1})",
                details={"requested": slide_index, "available": slide_count}
            )
        
        slide = agent.prs.slides[slide_index]
        shape_count = len(slide.shapes)
        
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={"requested": shape_index, "available": shape_count}
            )
        
        shape = slide.shapes[shape_index]
        
        # XML Manipulation for Z-Order
        sp_tree = slide.shapes._spTree
        element = shape.element
        
        # Find current position in XML tree
        current_index = -1
        for i, child in enumerate(sp_tree):
            if child == element:
                current_index = i
                break
        
        if current_index == -1:
            raise PowerPointAgentError("Could not locate shape in XML tree")
        
        new_index = current_index
        max_index = len(sp_tree) - 1
        
        # Execute Z-Order Action
        if action == 'bring_to_front':
            sp_tree.remove(element)
            sp_tree.append(element)
            new_index = max_index
            
        elif action == 'send_to_back':
            sp_tree.remove(element)
            sp_tree.insert(0, element)
            new_index = 0
            
        elif action == 'bring_forward':
            if current_index < max_index:
                sp_tree.remove(element)
                sp_tree.insert(current_index + 1, element)
                new_index = current_index + 1
                
        elif action == 'send_backward':
            if current_index > 0:
                sp_tree.remove(element)
                sp_tree.insert(current_index - 1, element)
                new_index = current_index - 1
        
        # Validate XML structure after manipulation
        if not _validate_xml_structure(sp_tree):
            raise PowerPointAgentError("XML structure corrupted during Z-order operation")
        
        agent.save()
        
        version_after = agent.get_presentation_version()
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index_target": shape_index,
        "action": action,
        "z_order_change": {
            "from": current_index,
            "to": new_index
        },
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__,
        "warning": "⚠️ Shape indices may have changed. Use ppt_get_slide_info.py to refresh before further operations.",
        "refresh_command": f"uv run tools/ppt_get_slide_info.py --file {filepath} --slide {slide_index} --json"
    }


def main():
    parser = argparse.ArgumentParser(
        description="Set shape Z-Order (layering)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Actions:
  bring_to_front  - Move shape to top layer (in front of all)
  send_to_back    - Move shape to bottom layer (behind all)
  bring_forward   - Move shape up one layer
  send_backward   - Move shape down one layer

Examples:
  # Send overlay to back (for readability overlays)
  uv run tools/ppt_set_z_order.py --file deck.pptx --slide 0 --shape 5 \\
    --action send_to_back --json

  # Bring logo to front
  uv run tools/ppt_set_z_order.py --file deck.pptx --slide 0 --shape 2 \\
    --action bring_to_front --json

⚠️  IMPORTANT: Shape indices change after z-order operations!
    Always run ppt_get_slide_info.py to refresh indices before targeting shapes.
        """
    )
    
    parser.add_argument('--file', required=True, type=Path, help='PowerPoint file path (.pptx)')
    parser.add_argument('--slide', required=True, type=int, help='Slide index (0-based)')
    parser.add_argument('--shape', required=True, type=int, help='Shape index (0-based)')
    parser.add_argument('--action', required=True,
                        choices=['bring_to_front', 'send_to_back', 'bring_forward', 'send_backward'],
                        help='Layering action')
    parser.add_argument('--json', action='store_true', default=True, help='Output JSON (default: true)')
    
    args = parser.parse_args()
    
    try:
        result = set_z_order(
            filepath=args.file,
            slide_index=args.slide,
            shape_index=args.shape,
            action=args.action
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
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ShapeNotFoundError",
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to check available shapes."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check file format (.pptx) and action is valid."
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "XML manipulation failed. File may be corrupted."
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
