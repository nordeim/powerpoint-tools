#!/usr/bin/env python3
"""
PowerPoint Set Image Properties Tool v3.1.0
Set alt text and opacity for image shapes (accessibility support)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_set_image_properties.py --file deck.pptx --slide 0 --shape 1 --alt-text "Company Logo" --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    Alt text is required for WCAG 2.1 compliance. All images should have
    descriptive alternative text that conveys the image's content and purpose.
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
import warnings

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exception if not available in core
try:
    from core.powerpoint_agent_core import ShapeNotFoundError
except ImportError:
    class ShapeNotFoundError(PowerPointAgentError):
        """Exception raised when shape is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def set_image_properties(
    filepath: Path,
    slide_index: int,
    shape_index: int,
    alt_text: Optional[str] = None,
    opacity: Optional[float] = None,
    transparency: Optional[float] = None  # Deprecated, for backward compat
) -> Dict[str, Any]:
    """
    Set properties on an image shape.
    
    Supports setting alternative text for accessibility and opacity
    for visual effects. At least one property must be specified.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the shape (0-based)
        shape_index: Index of the image shape (0-based)
        alt_text: Alternative text for accessibility (recommended for all images)
        opacity: Image opacity from 0.0 (invisible) to 1.0 (opaque)
        transparency: DEPRECATED - use opacity instead. If provided, converted to opacity.
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the shape
            - properties_set: Dict of properties that were set
            - presentation_version_before: State hash before modification
            - presentation_version_after: State hash after modification
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        ShapeNotFoundError: If the shape index is out of range
        ValueError: If no properties specified or invalid values
        
    Example:
        >>> result = set_image_properties(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     shape_index=1,
        ...     alt_text="Company Logo - Blue and white design"
        ... )
        >>> print(result["properties_set"]["alt_text"])
        'Company Logo - Blue and white design'
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Handle deprecated transparency parameter
    effective_opacity = opacity
    transparency_converted = False
    
    if transparency is not None:
        if opacity is not None:
            raise ValueError(
                "Cannot specify both 'opacity' and 'transparency'. "
                "Use 'opacity' (transparency is deprecated)."
            )
        # Convert transparency to opacity (inverse relationship)
        effective_opacity = 1.0 - transparency
        transparency_converted = True
    
    # Validate at least one property is being set
    if alt_text is None and effective_opacity is None:
        raise ValueError(
            "At least one property must be set (--alt-text or --opacity)"
        )
    
    # Validate opacity range
    if effective_opacity is not None:
        if not (0.0 <= effective_opacity <= 1.0):
            raise ValueError(
                f"Opacity must be between 0.0 and 1.0, got: {effective_opacity}"
            )

    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE modification
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
        
        # Get slide info to validate shape index
        slide_info = agent.get_slide_info(slide_index)
        shape_count = slide_info.get("shape_count", 0)
        
        if not 0 <= shape_index < shape_count:
            raise ShapeNotFoundError(
                f"Shape index {shape_index} out of range (0-{shape_count - 1})",
                details={
                    "requested_index": shape_index,
                    "available_shapes": shape_count
                }
            )
        
        # Set image properties
        # Note: Core method may use different parameter names
        try:
            agent.set_image_properties(
                slide_index=slide_index,
                shape_index=shape_index,
                alt_text=alt_text,
                # Pass opacity as fill_opacity if core supports it
                fill_opacity=effective_opacity
            )
        except TypeError:
            # Fallback if core uses different signature
            agent.set_image_properties(
                slide_index=slide_index,
                shape_index=shape_index,
                alt_text=alt_text,
                transparency=1.0 - effective_opacity if effective_opacity is not None else None
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER modification
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Build properties dict
    properties_set = {}
    if alt_text is not None:
        properties_set["alt_text"] = alt_text
    if effective_opacity is not None:
        properties_set["opacity"] = effective_opacity
        if transparency_converted:
            properties_set["transparency_converted"] = True
            properties_set["original_transparency"] = transparency
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "properties_set": properties_set,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Set image properties (alt text, opacity) in PowerPoint",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Set alt text for accessibility
  uv run tools/ppt_set_image_properties.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 1 \\
    --alt-text "Company Logo - Blue and white circular design" \\
    --json
  
  # Set opacity for watermark effect
  uv run tools/ppt_set_image_properties.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --shape 3 \\
    --opacity 0.3 \\
    --json
  
  # Set both properties
  uv run tools/ppt_set_image_properties.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --shape 0 \\
    --alt-text "Background watermark" \\
    --opacity 0.15 \\
    --json

Finding Shape Indices:
  Use ppt_get_slide_info.py to identify shape indices:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Alt Text Guidelines (WCAG 2.1):
  - Describe image content and purpose
  - For logos: "Company Name Logo"
  - For charts: Include key data points
  - For photos: Describe what's shown
  - For decorative images: Use empty string ""
  - Keep under 125 characters when possible

Opacity Values:
  - 0.0 = Fully transparent (invisible)
  - 0.5 = 50% visible
  - 1.0 = Fully opaque (default)
  
  Use Cases:
  - Watermarks: 0.1-0.2
  - Background images: 0.3-0.5
  - Subtle overlays: 0.15-0.25

Deprecation Notice:
  --transparency is deprecated. Use --opacity instead.
  transparency = 1.0 - opacity

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "shape_index": 1,
    "properties_set": {
      "alt_text": "Company Logo",
      "opacity": 1.0
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
        '--shape', 
        required=True, 
        type=int, 
        help='Shape index (0-based)'
    )
    parser.add_argument(
        '--alt-text', 
        help='Alternative text for accessibility'
    )
    parser.add_argument(
        '--opacity', 
        type=float, 
        help='Opacity from 0.0 (invisible) to 1.0 (opaque)'
    )
    parser.add_argument(
        '--transparency', 
        type=float, 
        help='DEPRECATED: Use --opacity instead. Transparency from 0.0 to 1.0'
    )
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = set_image_properties(
            filepath=args.file, 
            slide_index=args.slide, 
            shape_index=args.shape,
            alt_text=args.alt_text,
            opacity=args.opacity,
            transparency=args.transparency
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
            "suggestion": "Specify at least --alt-text or --opacity (0.0-1.0)"
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
