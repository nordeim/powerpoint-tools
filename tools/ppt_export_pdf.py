#!/usr/bin/env python3
"""
PowerPoint Export PDF Tool v3.1.1
Export presentation to PDF format using LibreOffice.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_export_pdf.py --file presentation.pptx --output presentation.pdf --json

Exit Codes:
    0: Success
    1: Error occurred

Requirements:
    LibreOffice must be installed for PDF export:
    - Linux: sudo apt install libreoffice-impress
    - macOS: brew install --cask libreoffice
    - Windows: Download from https://www.libreoffice.org/

Changelog v3.1.1:
    - Added hygiene block for JSON pipeline safety
    - Added presentation_version tracking via PowerPointAgent
    - Added tool_version and slide_count to output
    - Added --timeout argument (default: 300s for large presentations)
    - Fixed cross-filesystem rename with shutil.move
    - Fixed error response format with suggestions
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
from typing import Dict, Any

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

def export_pdf(
    filepath: Path,
    output: Path,
    timeout: int = 300
) -> Dict[str, Any]:
    """
    Export PowerPoint presentation to PDF.
    
    Args:
        filepath: Path to PowerPoint file (must be .pptx)
        output: Output PDF file path
        timeout: Subprocess timeout in seconds (default: 300)
        
    Returns:
        Dict with export results including file sizes and version info
        
    Raises:
        FileNotFoundError: If input file doesn't exist
        ValueError: If input is not a .pptx file
        RuntimeError: If LibreOffice not installed
        PowerPointAgentError: If export process fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    if not filepath.suffix.lower() == '.pptx':
        raise ValueError(f"Input must be .pptx file, got: {filepath.suffix}")
    
    if not check_libreoffice():
        raise RuntimeError(
            "LibreOffice not found. PDF export requires LibreOffice.\n"
            "Install:\n"
            "  Linux: sudo apt install libreoffice-impress\n"
            "  macOS: brew install --cask libreoffice\n"
            "  Windows: https://www.libreoffice.org/download/"
        )
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only operation, no lock needed
        presentation_version = agent.get_presentation_version()
        slide_count = agent.get_slide_count()
        presentation_info = agent.get_presentation_info()
    
    output.parent.mkdir(parents=True, exist_ok=True)
    
    lo_command = get_libreoffice_command()
    
    cmd = [
        lo_command,
        '--headless',
        '--convert-to', 'pdf',
        '--outdir', str(output.parent.resolve()),
        str(filepath.resolve())
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
            f"PDF export timed out after {timeout} seconds. "
            f"Try increasing --timeout for large presentations (100+ slides may need 5+ minutes)."
        )
    
    if result.returncode != 0:
        raise PowerPointAgentError(
            f"PDF export failed: {result.stderr}\n"
            f"Command: {' '.join(cmd)}"
        )
    
    expected_pdf = output.parent / f"{filepath.stem}.pdf"
    
    if expected_pdf != output:
        if expected_pdf.exists():
            if output.exists():
                output.unlink()
            shutil.move(str(expected_pdf), str(output))
    
    if not output.exists():
        if expected_pdf.exists():
            shutil.move(str(expected_pdf), str(output))
    
    if not output.exists():
        raise PowerPointAgentError(
            f"PDF export completed but output file not found. "
            f"Expected at: {output}"
        )
    
    input_size = filepath.stat().st_size
    output_size = output.stat().st_size
    
    return {
        "status": "success",
        "tool_version": __version__,
        "input_file": str(filepath.resolve()),
        "output_file": str(output.resolve()),
        "presentation_version": presentation_version,
        "slide_count": slide_count,
        "input_size_bytes": input_size,
        "input_size_mb": round(input_size / (1024 * 1024), 2),
        "output_size_bytes": output_size,
        "output_size_mb": round(output_size / (1024 * 1024), 2),
        "size_ratio": round(output_size / input_size, 2) if input_size > 0 else 0,
        "compression_percent": round((1 - output_size / input_size) * 100, 1) if input_size > 0 else 0
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Export PowerPoint presentation to PDF",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic export
  uv run tools/ppt_export_pdf.py \\
    --file presentation.pptx \\
    --output presentation.pdf \\
    --json
  
  # Large presentation with extended timeout
  uv run tools/ppt_export_pdf.py \\
    --file large_deck.pptx \\
    --output reports/output.pdf \\
    --timeout 600 \\
    --json

Requirements:
  LibreOffice must be installed:
  - Linux: sudo apt install libreoffice-impress
  - macOS: brew install --cask libreoffice
  - Windows: https://www.libreoffice.org/download/

Performance Notes:
  - Small decks (<20 slides): ~10-30 seconds
  - Medium decks (20-50 slides): ~1-2 minutes
  - Large decks (100+ slides): ~3-5 minutes
  - Adjust --timeout accordingly

PDF Benefits:
  - Universal compatibility
  - Prevents editing
  - Smaller file size (typically 30-50% of .pptx)
  - Better for printing

Limitations:
  - Animations not preserved
  - Embedded videos become static
  - Speaker notes not included
  - Transitions removed
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to export'
    )
    
    parser.add_argument(
        '--output',
        required=True,
        type=Path,
        help='Output PDF file path'
    )
    
    parser.add_argument(
        '--timeout',
        type=int,
        default=300,
        help='Export timeout in seconds (default: 300)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    args = parser.parse_args()
    
    try:
        output_path = args.output
        if not output_path.suffix.lower() == '.pdf':
            output_path = output_path.with_suffix('.pdf')
        
        result = export_pdf(
            filepath=args.file.resolve(),
            output=output_path.resolve(),
            timeout=args.timeout
        )
        
        if args.json:
            sys.stdout.write(json.dumps(result, indent=2) + "\n")
            sys.stdout.flush()
        else:
            print(f"✅ Exported to PDF: {result['output_file']}")
            print(f"   Slides: {result['slide_count']}")
            print(f"   Input: {result['input_size_mb']} MB")
            print(f"   Output: {result['output_size_mb']} MB")
            print(f"   Compression: {result['compression_percent']}%")
        
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
            "suggestion": "Ensure input file is a .pptx PowerPoint file",
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
            "suggestion": "Check LibreOffice installation and try increasing --timeout",
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
