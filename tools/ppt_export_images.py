#!/usr/bin/env python3
"""
PowerPoint Export Images Tool v3.1.1
Export each slide as PNG or JPG image using LibreOffice.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_export_images.py --file presentation.pptx --output-dir output/ --format png --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for image export:
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/

Changelog v3.1.1:
    - Added hygiene block for JSON pipeline safety
    - Added presentation_version tracking
    - Added tool_version to output
    - Added --timeout argument
    - Fixed error response format with suggestions
    - Removed stderr print statements
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that JSON parsers only see valid JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import subprocess
import shutil
from pathlib import Path
from typing import Dict, Any, List, Optional

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent,
    PowerPointAgentError
)

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"

# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def check_libreoffice() -> bool:
    """Check if LibreOffice is installed and accessible."""
    return shutil.which('soffice') is not None or shutil.which('libreoffice') is not None


def get_libreoffice_command() -> str:
    """Get the LibreOffice command for the current system."""
    if shutil.which('soffice'):
        return 'soffice'
    return 'libreoffice'


# ============================================================================
# MAIN LOGIC
# ============================================================================

def export_images(
    filepath: Path,
    output_dir: Path,
    image_format: str = "png",
    prefix: str = "slide_",
    timeout: int = 120
) -> Dict[str, Any]:
    """
    Export PowerPoint slides as images.
    
    Args:
        filepath: Path to PowerPoint file (must be .pptx)
        output_dir: Directory for output images
        image_format: Image format ('png', 'jpg', 'jpeg')
        prefix: Filename prefix for output images
        timeout: Subprocess timeout in seconds
        
    Returns:
        Dict with export results including file list and sizes
        
    Raises:
        FileNotFoundError: If input file doesn't exist
        ValueError: If invalid format or file type
        RuntimeError: If LibreOffice not installed
        PowerPointAgentError: If export process fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.suffix.lower() == '.pptx':
        raise ValueError(f"Input must be .pptx file, got: {filepath.suffix}")
    
    if image_format.lower() not in ['png', 'jpg', 'jpeg']:
        raise ValueError(f"Format must be png or jpg, got: {image_format}")
    
    format_ext = 'png' if image_format.lower() == 'png' else 'jpg'
    
    if not check_libreoffice():
        raise RuntimeError(
            "LibreOffice not found. Image export requires LibreOffice.\n"
            "Install:\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    warnings_collected: List[str] = []
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only operation, no lock needed
        slide_count = agent.get_slide_count()
        presentation_version = agent.get_presentation_version()
    
    base_name = filepath.stem
    pdf_path = output_dir / f"{base_name}.pdf"
    lo_command = get_libreoffice_command()
    
    cmd_pdf = [
        lo_command,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output_dir),
        str(filepath)
    ]
    
    try:
        result_pdf = subprocess.run(
            cmd_pdf,
            capture_output=True,
            text=True,
            timeout=timeout
        )
    except subprocess.TimeoutExpired:
        raise PowerPointAgentError(
            f"PDF export timed out after {timeout} seconds. "
            f"Try increasing --timeout for large presentations."
        )
    
    if result_pdf.returncode != 0:
        raise PowerPointAgentError(
            f"PDF export failed: {result_pdf.stderr}\n"
            f"Command: {' '.join(cmd_pdf)}"
        )
    
    use_pdftoppm = shutil.which('pdftoppm') is not None
    
    if use_pdftoppm and pdf_path.exists():
        cmd_img = [
            'pdftoppm',
            f"-{format_ext}",
            '-r', '150',
            str(pdf_path),
            str(output_dir / base_name)
        ]
        
        try:
            result_img = subprocess.run(
                cmd_img,
                capture_output=True,
                text=True,
                timeout=timeout
            )
            
            if result_img.returncode != 0:
                warnings_collected.append(
                    f"pdftoppm failed, using LibreOffice direct export"
                )
                _export_direct(filepath, output_dir, format_ext, lo_command, timeout)
                
        except subprocess.TimeoutExpired:
            warnings_collected.append(
                f"pdftoppm timed out, using LibreOffice direct export"
            )
            _export_direct(filepath, output_dir, format_ext, lo_command, timeout)
        
        if pdf_path.exists():
            pdf_path.unlink()
    else:
        if not use_pdftoppm:
            warnings_collected.append(
                "pdftoppm not found, using LibreOffice direct export (may be incomplete)"
            )
        _export_direct(filepath, output_dir, format_ext, lo_command, timeout)
    
    result = _scan_and_process_results(filepath, output_dir, format_ext, prefix)
    
    result["presentation_version"] = presentation_version
    result["slide_count_source"] = slide_count
    result["tool_version"] = __version__
    
    if warnings_collected:
        result["warnings"] = warnings_collected
    
    return result


def _export_direct(
    filepath: Path,
    output_dir: Path,
    format_ext: str,
    lo_command: str,
    timeout: int
) -> None:
    """
    Direct export using LibreOffice (fallback method).
    
    Args:
        filepath: Input PowerPoint file
        output_dir: Output directory
        format_ext: Image format extension
        lo_command: LibreOffice command to use
        timeout: Subprocess timeout in seconds
    """
    cmd = [
        lo_command,
        '--headless',
        '--convert-to', format_ext,
        '--outdir', str(output_dir),
        str(filepath)
    ]
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=timeout
        )
    except subprocess.TimeoutExpired:
        raise PowerPointAgentError(
            f"Direct image export timed out after {timeout} seconds"
        )
    
    if result.returncode != 0:
        raise PowerPointAgentError(
            f"Image export failed: {result.stderr}\n"
            f"Command: {' '.join(cmd)}"
        )


def _scan_and_process_results(
    filepath: Path,
    output_dir: Path,
    format_ext: str,
    prefix: str
) -> Dict[str, Any]:
    """
    Find, rename, and report exported images.
    
    Args:
        filepath: Original input file (for base name)
        output_dir: Directory containing exported images
        format_ext: Image format extension
        prefix: Prefix for renamed files
        
    Returns:
        Dict with export statistics and file list
    """
    base_name = filepath.stem
    
    candidates = sorted(output_dir.glob(f"{base_name}*.{format_ext}"))
    
    if not candidates:
        candidates = sorted(output_dir.glob(f"*.{format_ext}"))
    
    exported_files: List[Path] = []
    
    for i, old_file in enumerate(candidates):
        new_file = output_dir / f"{prefix}{i+1:03d}.{format_ext}"
        
        if old_file != new_file:
            if new_file.exists():
                new_file.unlink()
            old_file.rename(new_file)
            exported_files.append(new_file)
        else:
            exported_files.append(old_file)
    
    if len(exported_files) == 0:
        raise PowerPointAgentError(
            f"Export completed but no image files found in: {output_dir}"
        )
    
    total_size = sum(f.stat().st_size for f in exported_files)
    
    return {
        "status": "success",
        "input_file": str(filepath.resolve()),
        "output_dir": str(output_dir.resolve()),
        "format": format_ext,
        "slides_exported": len(exported_files),
        "files": [str(f.resolve()) for f in exported_files],
        "total_size_bytes": total_size,
        "total_size_mb": round(total_size / (1024 * 1024), 2),
        "average_size_mb": round(total_size / (1024 * 1024) / len(exported_files), 2) if exported_files else 0
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Export PowerPoint slides as images",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Export as PNG
  uv run tools/ppt_export_images.py \\
    --file presentation.pptx \\
    --output-dir slides/ \\
    --format png \\
    --json
  
  # Export as JPG with custom prefix and timeout
  uv run tools/ppt_export_images.py \\
    --file presentation.pptx \\
    --output-dir images/ \\
    --format jpg \\
    --prefix deck_ \\
    --timeout 300 \\
    --json

Output Files:
  Files are named: <prefix><number>.<format>
  Examples: slide_001.png, slide_002.png, deck_001.jpg

Requirements:
  LibreOffice must be installed:
  - Linux: sudo apt install libreoffice-impress
  - macOS: brew install --cask libreoffice
  - Windows: https://www.libreoffice.org/download/

Format Comparison:
  PNG: Lossless, better for text/diagrams, larger files
  JPG: Lossy, better for photos, smaller files
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to export'
    )
    
    parser.add_argument(
        '--output-dir',
        required=True,
        type=Path,
        help='Output directory for images'
    )
    
    parser.add_argument(
        '--format',
        choices=['png', 'jpg', 'jpeg'],
        default='png',
        help='Image format (default: png)'
    )
    
    parser.add_argument(
        '--prefix',
        default='slide_',
        help='Filename prefix (default: slide_)'
    )
    
    parser.add_argument(
        '--timeout',
        type=int,
        default=120,
        help='Subprocess timeout in seconds (default: 120)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        result = export_images(
            filepath=args.file.resolve(),
            output_dir=args.output_dir.resolve(),
            image_format=args.format,
            prefix=args.prefix,
            timeout=args.timeout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            print(f"✅ Exported {result['slides_exported']} slides to {result['output_dir']}")
            print(f"   Format: {result['format'].upper()}")
            print(f"   Total size: {result['total_size_mb']} MB")
            print(f"   Average: {result['average_size_mb']} MB per slide")
        
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify the file path exists and is accessible",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except ValueError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ValueError",
            "suggestion": "Check input file format (.pptx) and image format (png/jpg)",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except RuntimeError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "RuntimeError",
            "suggestion": "Install LibreOffice: sudo apt install libreoffice-impress",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Check LibreOffice installation and file integrity",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)
        
    except Exception as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "suggestion": "Check logs for detailed error information",
            "tool_version": __version__
        }
        sys.stdout.write(json.dumps(error_result, indent=2) + "\n")
        sys.stdout.flush()
        sys.exit(1)


if __name__ == "__main__":
    main()
