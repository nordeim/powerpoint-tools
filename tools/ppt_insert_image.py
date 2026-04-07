#!/usr/bin/env python3
"""
PowerPoint Insert Image Tool v3.1.0
Insert image into slide with automatic aspect ratio handling

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_insert_image.py --file presentation.pptx --slide 0 --image logo.png --position '{"left":"10%","top":"10%"}' --size '{"width":"20%","height":"auto"}' --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Accessibility:
    Always use --alt-text to provide alternative text for screen readers.
    This is required for WCAG 2.1 compliance.
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

# Define fallback exceptions if not available in core
try:
    from core.powerpoint_agent_core import ImageNotFoundError
except ImportError:
    class ImageNotFoundError(PowerPointAgentError):
        """Exception raised when image file is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)

try:
    from core.powerpoint_agent_core import InvalidPositionError
except ImportError:
    class InvalidPositionError(PowerPointAgentError):
        """Exception raised when position specification is invalid."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def insert_image(
    filepath: Path,
    slide_index: int,
    image_path: Path,
    position: Dict[str, Any],
    size: Optional[Dict[str, Any]] = None,
    compress: bool = False,
    alt_text: Optional[str] = None
) -> Dict[str, Any]:
    """
    Insert an image into a PowerPoint slide.
    
    Supports automatic aspect ratio preservation when using "auto" for
    width or height. Optionally compresses large images and sets
    accessibility alt text.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the target slide (0-based)
        image_path: Path to the image file to insert
        position: Position specification dict (percentage, anchor, or grid-based)
        size: Size specification dict (optional, defaults to 50% width with auto height)
        compress: Whether to compress the image before insertion (default: False)
        alt_text: Alternative text for accessibility (highly recommended)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - shape_index: Index of the inserted image shape
            - image_file: Path to the source image
            - image_size_bytes: Original image file size
            - image_size_mb: Original image size in MB
            - position: Applied position
            - size: Applied size
            - compressed: Whether compression was applied
            - alt_text: Applied alt text (or None)
            - presentation_version_before: State hash before insertion
            - presentation_version_after: State hash after insertion
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If the PowerPoint or image file doesn't exist
        SlideNotFoundError: If the slide index is out of range
        ValueError: If image format is unsupported
        InvalidPositionError: If position specification is invalid
        
    Example:
        >>> result = insert_image(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     image_path=Path("logo.png"),
        ...     position={"left": "10%", "top": "10%"},
        ...     size={"width": "20%", "height": "auto"},
        ...     alt_text="Company Logo"
        ... )
        >>> print(result["shape_index"])
        5
    """
    # Validate presentation file exists
    if not filepath.exists():
        raise FileNotFoundError(f"Presentation file not found: {filepath}")
    
    # Validate image file exists
    if not image_path.exists():
        raise ImageNotFoundError(
            f"Image file not found: {image_path}",
            details={"image_path": str(image_path)}
        )
    
    # Validate image format
    valid_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'}
    if image_path.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Unsupported image format: {image_path.suffix}. "
            f"Supported formats: {', '.join(sorted(valid_extensions))}"
        )
    
    # Default size if not provided
    if size is None:
        size = {"width": "50%", "height": "auto"}
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE insertion
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
        
        # Insert image
        result = agent.insert_image(
            slide_index=slide_index,
            image_path=image_path,
            position=position,
            size=size,
            compress=compress
        )
        
        # Extract shape index from result (handle both v3.0.x and v3.1.x)
        if isinstance(result, dict):
            shape_index = result.get("shape_index")
        else:
            # Fallback: get last shape index from slide info
            slide_info = agent.get_slide_info(slide_index)
            shape_index = slide_info["shape_count"] - 1
        
        # Set alt text if provided
        if alt_text and shape_index is not None:
            agent.set_image_properties(
                slide_index=slide_index,
                shape_index=shape_index,
                alt_text=alt_text
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER insertion
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
        
        # Get final slide info
        final_slide_info = agent.get_slide_info(slide_index)
    
    # Get image file info
    image_size = image_path.stat().st_size
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "shape_index": shape_index,
        "image_file": str(image_path.resolve()),
        "image_size_bytes": image_size,
        "image_size_mb": round(image_size / (1024 * 1024), 2),
        "position": position,
        "size": size,
        "compressed": compress,
        "alt_text": alt_text,
        "slide_shape_count": final_slide_info.get("shape_count", 0),
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Insert image into PowerPoint slide",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Insert logo with alt text (accessibility)
  uv run tools/ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --image company_logo.png \\
    --position '{"left":"5%","top":"5%"}' \\
    --size '{"width":"15%","height":"auto"}' \\
    --alt-text "Company Logo" \\
    --json
  
  # Insert centered hero image with compression
  uv run tools/ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 1 \\
    --image product_photo.jpg \\
    --position '{"anchor":"center","offset_x":0,"offset_y":0}' \\
    --size '{"width":"80%","height":"auto"}' \\
    --compress \\
    --alt-text "Product photograph showing new design" \\
    --json
  
  # Insert chart with grid positioning
  uv run tools/ppt_insert_image.py \\
    --file presentation.pptx \\
    --slide 2 \\
    --image revenue_chart.png \\
    --position '{"left":"10%","top":"25%"}' \\
    --size '{"width":"80%","height":"auto"}' \\
    --alt-text "Revenue growth chart: Q1 $100K, Q2 $150K, Q3 $200K, Q4 $250K" \\
    --json

Size Options:
  {"width": "50%", "height": "auto"}  - Auto-calculate height (recommended)
  {"width": "auto", "height": "40%"}  - Auto-calculate width
  {"width": "30%", "height": "20%"}   - Fixed dimensions
  {"width": 3.0, "height": 2.0}       - Absolute inches

Position Options:
  {"left": "10%", "top": "20%"}       - Percentage of slide
  {"anchor": "center"}                - Anchor-based
  {"left": 1.5, "top": 2.0}           - Absolute inches

Supported Formats:
  - PNG (recommended for logos, diagrams, transparency)
  - JPG/JPEG (recommended for photos)
  - GIF (first frame only, animation not supported)
  - BMP, TIFF (not recommended, large file size)

Compression (--compress):
  - Resizes to max 1920px width
  - Converts RGBA to RGB
  - JPEG quality 85%
  - Typically reduces size 50-70%

Accessibility (--alt-text):
  - REQUIRED for WCAG 2.1 compliance
  - Describe the image content and purpose
  - For charts/data: include key data points
  - For decorative images: use empty string ""

Best Practices:
  - Always use --alt-text for accessibility
  - Use "auto" for height OR width to maintain aspect ratio
  - Use --compress for images > 1MB
  - Recommended max resolution: 1920x1080
  - Use PNG for transparency, JPG for photos

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "shape_index": 5,
    "image_file": "/path/to/logo.png",
    "image_size_mb": 0.25,
    "alt_text": "Company Logo",
    "presentation_version_after": "a1b2c3d4...",
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
        '--image',
        required=True,
        type=Path,
        help='Image file path'
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
        default=None,
        help='Size dict as JSON string (default: 50%% width with auto height)'
    )
    
    parser.add_argument(
        '--compress',
        action='store_true',
        help='Compress image before inserting (recommended for large images)'
    )
    
    parser.add_argument(
        '--alt-text',
        help='Alternative text for accessibility (highly recommended)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        # Parse JSON arguments
        try:
            position = json.loads(args.position)
        except json.JSONDecodeError as e:
            raise ValueError(
                f"Invalid JSON in --position: {e}. "
                "Use single quotes around JSON: '{\"left\":\"10%\",\"top\":\"20%\"}'"
            )
        
        size = None
        if args.size:
            try:
                size = json.loads(args.size)
            except json.JSONDecodeError as e:
                raise ValueError(
                    f"Invalid JSON in --size: {e}. "
                    "Use single quotes around JSON: '{\"width\":\"50%\",\"height\":\"auto\"}'"
                )
        
        result = insert_image(
            filepath=args.file,
            slide_index=args.slide,
            image_path=args.image,
            position=position,
            size=size,
            compress=args.compress,
            alt_text=args.alt_text
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except (FileNotFoundError, ImageNotFoundError) as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
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
            "suggestion": "Check JSON format and image file format"
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
