#!/usr/bin/env python3
"""
PowerPoint Check Accessibility Tool v3.1.0
Run WCAG 2.1 accessibility checks on presentation.

This tool performs comprehensive accessibility validation including:
- Alt text presence for images
- Color contrast ratios
- Reading order verification
- Font size compliance

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.0

Usage:
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json

Exit Codes:
    0: Success (check completed, see 'passed' field for result)
    1: Error occurred (file not found, crash)

Design Principles:
    - Read-only operation (acquire_lock=False)
    - JSON-first output with consistent contract
    - Strict output hygiene (stderr suppressed)
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately.
# This prevents libraries (pptx, warnings) from printing non-JSON text
# which corrupts pipelines that capture 2>&1.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import logging
import time
import uuid
from pathlib import Path
from typing import Dict, Any
from datetime import datetime

# Configure logging to null handler
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.0"

# ============================================================================
# IMPORTS WITH ERROR HANDLING
# ============================================================================

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError
    )
    CORE_AVAILABLE = True
except ImportError as e:
    CORE_AVAILABLE = False
    IMPORT_ERROR = str(e)


# ============================================================================
# MAIN LOGIC
# ============================================================================

def check_accessibility(filepath: Path) -> Dict[str, Any]:
    """
    Run accessibility checks on a PowerPoint presentation.
    
    Args:
        filepath: Path to PowerPoint file
        
    Returns:
        Dict with accessibility check results including:
        - status: "success"
        - passed: bool indicating if all checks passed
        - issues: dict of issue categories
        - summary: counts by category
        
    Raises:
        FileNotFoundError: If file doesn't exist
        ImportError: If core module not available
        PowerPointAgentError: If presentation cannot be opened
    """
    if not CORE_AVAILABLE:
        raise ImportError(f"Core module not available: {IMPORT_ERROR}")
    
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    start_time = time.perf_counter()
    operation_id = str(uuid.uuid4())
    
    with PowerPointAgent(filepath) as agent:
        # acquire_lock=False because validation is read-only
        agent.open(filepath, acquire_lock=False)
        
        # Capture presentation version for audit trail
        try:
            presentation_version = agent.get_presentation_version()
        except AttributeError:
            presentation_version = "unknown"
        
        # Get presentation info for context
        try:
            pres_info = agent.get_presentation_info()
            slide_count = pres_info.get("slide_count", 0)
        except Exception:
            slide_count = 0
        
        # Run accessibility check
        result = agent.check_accessibility()
    
    duration_ms = int((time.perf_counter() - start_time) * 1000)
    
    # Extract issues for summary
    issues = result.get("issues", {})
    missing_alt_text = issues.get("missing_alt_text", [])
    low_contrast = issues.get("low_contrast", [])
    reading_order = issues.get("reading_order_issues", [])
    small_fonts = issues.get("small_fonts", [])
    
    total_issues = (
        len(missing_alt_text) + 
        len(low_contrast) + 
        len(reading_order) + 
        len(small_fonts)
    )
    
    # Determine pass/fail
    passed = result.get("passed", total_issues == 0)
    
    return {
        "status": "success",
        "passed": passed,
        "file": str(filepath.resolve()),
        "tool_version": __version__,
        "presentation_version": presentation_version,
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "operation_id": operation_id,
        "duration_ms": duration_ms,
        "summary": {
            "slide_count": slide_count,
            "total_issues": total_issues,
            "missing_alt_text_count": len(missing_alt_text),
            "low_contrast_count": len(low_contrast),
            "reading_order_issues_count": len(reading_order),
            "small_fonts_count": len(small_fonts)
        },
        "issues": {
            "missing_alt_text": missing_alt_text,
            "low_contrast": low_contrast,
            "reading_order_issues": reading_order,
            "small_fonts": small_fonts
        },
        "wcag_level": "AA",
        "recommendations": _generate_recommendations(issues)
    }


def _generate_recommendations(issues: Dict[str, Any]) -> list:
    """
    Generate actionable recommendations based on issues found.
    
    Args:
        issues: Dict of issue categories
        
    Returns:
        List of recommendation dicts
    """
    recommendations = []
    
    missing_alt = issues.get("missing_alt_text", [])
    if missing_alt:
        recommendations.append({
            "priority": "high",
            "category": "accessibility",
            "action": f"Add alt text to {len(missing_alt)} image(s)",
            "fix_command": "ppt_set_image_properties.py --alt-text"
        })
    
    low_contrast = issues.get("low_contrast", [])
    if low_contrast:
        recommendations.append({
            "priority": "high",
            "category": "accessibility",
            "action": f"Fix contrast on {len(low_contrast)} element(s)",
            "fix_command": "ppt_format_text.py --color"
        })
    
    small_fonts = issues.get("small_fonts", [])
    if small_fonts:
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "action": f"Increase font size on {len(small_fonts)} element(s)",
            "fix_command": "ppt_format_text.py --font-size"
        })
    
    reading_order = issues.get("reading_order_issues", [])
    if reading_order:
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "action": f"Review reading order on {len(reading_order)} slide(s)",
            "fix_command": "Manual review required"
        })
    
    return recommendations


# ============================================================================
# CLI ENTRY POINT
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description=f"Check PowerPoint accessibility (WCAG 2.1) - v{__version__}",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    # Check accessibility
    uv run tools/ppt_check_accessibility.py --file presentation.pptx --json
    
Output includes:
    - passed: Boolean indicating if accessibility checks passed
    - summary: Counts of issues by category
    - issues: Detailed list of accessibility violations
    - recommendations: Suggested fixes with commands

Version: """ + __version__
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file path'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default)'
    )
    
    args = parser.parse_args()
    
    try:
        result = check_accessibility(filepath=args.file)
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
        
    except ImportError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "ImportError",
            "suggestion": "Ensure core module is properly installed"
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
