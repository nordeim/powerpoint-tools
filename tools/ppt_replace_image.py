#!/usr/bin/env python3
"""
PowerPoint Replace Image Tool v3.1.0
Replace an existing image with a new one (preserves position and size)

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_replace_image.py --file presentation.pptx --slide 0 --old-image "logo" --new-image new_logo.png --json

Exit Codes:
    0: Success
    1: Error occurred (check error_type in JSON for details)

Use Cases:
    - Logo updates during rebranding
    - Product photo updates
    - Chart/diagram refreshes
    - Team photo updates
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
from typing import Dict, Any

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError
)

__version__ = "3.1.0"

# Define fallback exception if not available in core
try:
    from core.powerpoint_agent_core import ImageNotFoundError
except ImportError:
    class ImageNotFoundError(PowerPointAgentError):
        """Exception raised when image is not found."""
        def __init__(self, message: str, details: Dict = None):
            self.message = message
            self.details = details or {}
            super().__init__(message)


def replace_image(
    filepath: Path,
    slide_index: int,
    old_image: str,
    new_image: Path,
    compress: bool = False
) -> Dict[str, Any]:
    """
    Replace an existing image with a new one.
    
    Searches for an image by name (exact or partial match) and replaces
    it with the new image while preserving the original position and size.
    
    Args:
        filepath: Path to the PowerPoint file to modify
        slide_index: Index of the slide containing the image (0-based)
        old_image: Name or partial name of the image to replace
        new_image: Path to the new image file
        compress: Whether to compress the new image (default: False)
        
    Returns:
        Dict containing:
            - status: "success"
            - file: Absolute path to modified file
            - slide_index: Index of the slide
            - old_image: Name/pattern that was searched
            - new_image: Path to the new image
            - new_image_size_bytes: Size of new image file
            - new_image_size_mb: Size in MB
            - compressed: Whether compression was applied
            - replaced: True if replacement succeeded
            - presentation_version_before: State hash before replacement
            - presentation_version_after: State hash after replacement
            - tool_version: Version of this tool
            
    Raises:
        FileNotFoundError: If PowerPoint or new image file doesn't exist
        SlideNotFoundError: If slide index is out of range
        ImageNotFoundError: If old image is not found on the slide
        
    Example:
        >>> result = replace_image(
        ...     filepath=Path("presentation.pptx"),
        ...     slide_index=0,
        ...     old_image="company_logo",
        ...     new_image=Path("new_logo.png")
        ... )
        >>> print(result["replaced"])
        True
    """
    # Validate presentation file exists
    if not filepath.exists():
        raise FileNotFoundError(f"Presentation file not found: {filepath}")
    
    # Validate new image file exists
    if not new_image.exists():
        raise ImageNotFoundError(
            f"New image file not found: {new_image}",
            details={"new_image_path": str(new_image)}
        )
    
    # Validate image format
    valid_extensions = {'.png', '.jpg', '.jpeg', '.gif', '.bmp', '.tiff', '.tif'}
    if new_image.suffix.lower() not in valid_extensions:
        raise ValueError(
            f"Unsupported image format: {new_image.suffix}. "
            f"Supported formats: {', '.join(sorted(valid_extensions))}"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture version BEFORE replacement
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
        
        # Attempt replacement
        replaced = agent.replace_image(
            slide_index=slide_index,
            old_image_name=old_image,
            new_image_path=new_image,
            compress=compress
        )
        
        if not replaced:
            raise ImageNotFoundError(
                f"Image matching '{old_image}' not found on slide {slide_index}. "
                "Use ppt_get_slide_info.py to list available images.",
                details={
                    "search_pattern": old_image,
                    "slide_index": slide_index
                }
            )
        
        # Save changes
        agent.save()
        
        # Capture version AFTER replacement
        info_after = agent.get_presentation_info()
        version_after = info_after.get("presentation_version")
    
    # Get new image size
    new_size = new_image.stat().st_size
    
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "slide_index": slide_index,
        "old_image": old_image,
        "new_image": str(new_image.resolve()),
        "new_image_size_bytes": new_size,
        "new_image_size_mb": round(new_size / (1024 * 1024), 2),
        "compressed": compress,
        "replaced": True,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "tool_version": __version__
    }


def main():
    parser = argparse.ArgumentParser(
        description="Replace image in PowerPoint presentation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Replace logo by name
  uv run tools/ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --old-image "company_logo" \\
    --new-image new_logo.png \\
    --json
  
  # Replace with compression
  uv run tools/ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 5 \\
    --old-image "product_photo" \\
    --new-image updated_photo.jpg \\
    --compress \\
    --json
  
  # Partial name match
  uv run tools/ppt_replace_image.py \\
    --file presentation.pptx \\
    --slide 0 \\
    --old-image "logo" \\
    --new-image rebrand_logo.png \\
    --json

Finding Images:
  Use ppt_get_slide_info.py to list images on a slide:
  uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json

Search Strategy:
  The tool searches for images by:
  1. Exact name match
  2. Partial name match (contains)
  3. First match is replaced

Compression (--compress):
  - Resizes to max 1920px width
  - Converts to JPEG at 85% quality
  - Typically reduces size 50-70%
  - Recommended for images > 1MB

Best Practices:
  - Use descriptive image names in PowerPoint
  - Keep new image dimensions similar to original
  - Use --compress for large replacement images
  - Test on a cloned copy first
  - Verify aspect ratios match

Output Format:
  {
    "status": "success",
    "file": "/path/to/presentation.pptx",
    "slide_index": 0,
    "old_image": "company_logo",
    "new_image": "/path/to/new_logo.png",
    "new_image_size_mb": 0.15,
    "compressed": false,
    "replaced": true,
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
        '--old-image',
        required=True,
        help='Name or partial name of image to replace'
    )
    
    parser.add_argument(
        '--new-image',
        required=True,
        type=Path,
        help='Path to new image file'
    )
    
    parser.add_argument(
        '--compress',
        action='store_true',
        help='Compress new image before inserting'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = replace_image(
            filepath=args.file,
            slide_index=args.slide,
            old_image=args.old_image,
            new_image=args.new_image,
            compress=args.compress
        )
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
        
    except (FileNotFoundError, ImageNotFoundError) as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "details": getattr(e, 'details', {}),
            "suggestion": "Use ppt_get_slide_info.py to list available images on the slide"
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
            "suggestion": "Check image file format (PNG, JPG, GIF, BMP supported)"
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
