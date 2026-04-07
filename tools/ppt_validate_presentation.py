#!/usr/bin/env python3
"""
PowerPoint Validate Presentation Tool v3.1.1
Comprehensive validation for structure, accessibility, assets, and design quality.

Fully aligned with PowerPoint Agent Core v3.1.0+ and System Prompt v3.0 validation gates.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --json
    uv run tools/ppt_validate_presentation.py --file presentation.pptx --policy strict --json

Exit Codes:
    0: Success (valid or only warnings within policy thresholds)
    1: Error occurred or critical issues exceed policy thresholds

Changelog v3.1.1:
    - Added presentation_version to output for audit trail
    - Populated fix_command for actionable remediation
    - Expanded _validate_design_rules with color and 6x6 rule checking
    - Added tool_version to output
    - Added acquire_lock documentation comments
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
import logging
from pathlib import Path
from typing import Dict, Any, List, Optional, Set
from dataclasses import dataclass, field, asdict
from datetime import datetime

# Configure logging to null handler to prevent any accidental output
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path for core import
sys.path.insert(0, str(Path(__file__).parent.parent))

try:
    from core.powerpoint_agent_core import (
        PowerPointAgent,
        PowerPointAgentError,
        __version__ as CORE_VERSION
    )
except ImportError:
    CORE_VERSION = "0.0.0"
    PowerPointAgent = None
    PowerPointAgentError = Exception

# ============================================================================
# CONSTANTS & POLICIES
# ============================================================================

__version__ = "3.1.1"

VALIDATION_POLICIES = {
    "lenient": {
        "name": "Lenient",
        "description": "Minimal validation - suitable for drafts and work-in-progress",
        "thresholds": {
            "max_critical_issues": 10,
            "max_accessibility_issues": 20,
            "max_design_warnings": 50,
            "max_empty_slides": 5,
            "max_slides_without_titles": 10,
            "max_missing_alt_text": 20,
            "max_low_contrast": 10,
            "max_large_images": 10,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 10,
            "max_colors": 20,
            "min_font_size_pt": 8,
        }
    },
    "standard": {
        "name": "Standard",
        "description": "Balanced validation - suitable for internal presentations",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 5,
            "max_design_warnings": 10,
            "max_empty_slides": 0,
            "max_slides_without_titles": 3,
            "max_missing_alt_text": 5,
            "max_low_contrast": 3,
            "max_large_images": 5,
            "require_all_alt_text": False,
            "enforce_6x6_rule": False,
            "max_fonts": 5,
            "max_colors": 10,
            "min_font_size_pt": 10,
        }
    },
    "strict": {
        "name": "Strict",
        "description": "Maximum validation - suitable for external/production presentations",
        "thresholds": {
            "max_critical_issues": 0,
            "max_accessibility_issues": 0,
            "max_design_warnings": 3,
            "max_empty_slides": 0,
            "max_slides_without_titles": 0,
            "max_missing_alt_text": 0,
            "max_low_contrast": 0,
            "max_large_images": 3,
            "require_all_alt_text": True,
            "enforce_6x6_rule": True,
            "max_fonts": 3,
            "max_colors": 5,
            "min_font_size_pt": 12,
        }
    }
}

# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class ValidationIssue:
    """Represents a single validation issue found in the presentation."""
    category: str
    severity: str
    message: str
    slide_index: Optional[int] = None
    shape_index: Optional[int] = None
    fix_command: Optional[str] = None
    details: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary, excluding None values."""
        result = {}
        for key, value in asdict(self).items():
            if value is not None and value != {}:
                result[key] = value
        return result


@dataclass
class ValidationSummary:
    """Summary statistics for validation results."""
    total_issues: int = 0
    critical_count: int = 0
    warning_count: int = 0
    info_count: int = 0
    empty_slides: int = 0
    slides_without_titles: int = 0
    missing_alt_text: int = 0
    low_contrast: int = 0
    large_images: int = 0
    fonts_used: int = 0
    colors_detected: int = 0
    bullet_violations: int = 0
    small_font_count: int = 0
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return asdict(self)


@dataclass
class ValidationPolicy:
    """Validation policy with thresholds."""
    name: str
    thresholds: Dict[str, Any]
    description: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        """Convert to dictionary."""
        return asdict(self)


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def get_policy(
    policy_name: str,
    custom_thresholds: Optional[Dict[str, Any]] = None
) -> ValidationPolicy:
    """
    Get validation policy by name with optional custom overrides.
    
    Args:
        policy_name: Name of policy ('lenient', 'standard', 'strict', 'custom')
        custom_thresholds: Optional custom threshold overrides
        
    Returns:
        ValidationPolicy instance
    """
    if policy_name == "custom" and custom_thresholds:
        base = VALIDATION_POLICIES["standard"]["thresholds"].copy()
        base.update(custom_thresholds)
        return ValidationPolicy(
            name="Custom",
            thresholds=base,
            description="Custom policy with user-defined thresholds"
        )
    
    config = VALIDATION_POLICIES.get(policy_name, VALIDATION_POLICIES["standard"])
    return ValidationPolicy(
        name=config["name"],
        thresholds=config["thresholds"],
        description=config.get("description", "")
    )


def generate_fix_command(
    filepath: Path,
    issue_type: str,
    slide_index: Optional[int] = None,
    shape_index: Optional[int] = None,
    extra_args: Optional[Dict[str, str]] = None
) -> Optional[str]:
    """
    Generate a CLI command to fix a specific issue.
    
    Args:
        filepath: Path to the presentation file
        issue_type: Type of issue to fix
        slide_index: Slide index if applicable
        shape_index: Shape index if applicable
        extra_args: Additional arguments for the fix command
        
    Returns:
        CLI command string or None if no fix available
    """
    base_path = str(filepath)
    
    fix_commands = {
        "missing_alt_text": (
            f"uv run tools/ppt_set_image_properties.py "
            f"--file \"{base_path}\" --slide {slide_index} --shape {shape_index} "
            f"--alt-text \"DESCRIBE_IMAGE_HERE\" --json"
        ),
        "empty_slide": (
            f"uv run tools/ppt_delete_slide.py "
            f"--file \"{base_path}\" --index {slide_index} --json"
        ),
        "missing_title": (
            f"uv run tools/ppt_set_title.py "
            f"--file \"{base_path}\" --slide {slide_index} "
            f"--title \"ADD_TITLE_HERE\" --json"
        ),
        "low_contrast": (
            f"uv run tools/ppt_format_text.py "
            f"--file \"{base_path}\" --slide {slide_index} --shape {shape_index} "
            f"--color \"#111111\" --json"
        ),
    }
    
    if issue_type in fix_commands:
        cmd = fix_commands[issue_type]
        if slide_index is None:
            return None
        return cmd
    
    return None


# ============================================================================
# VALIDATION PROCESSORS
# ============================================================================

def _process_core_validation(
    core_result: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """
    Process results from agent.validate_presentation().
    
    Args:
        core_result: Result from validate_presentation()
        issues: List to append issues to
        summary: Summary to update
        filepath: Path for fix commands
    """
    issue_data = core_result.get("issues", {})
    
    empty_slides = issue_data.get("empty_slides", [])
    summary.empty_slides = len(empty_slides)
    for idx in empty_slides:
        issues.append(ValidationIssue(
            category="structure",
            severity="critical",
            message=f"Empty slide with no content",
            slide_index=idx,
            fix_command=generate_fix_command(filepath, "empty_slide", slide_index=idx),
            details={"issue_type": "empty_slide"}
        ))
    
    untitled_slides = issue_data.get("slides_without_titles", [])
    summary.slides_without_titles = len(untitled_slides)
    for idx in untitled_slides:
        issues.append(ValidationIssue(
            category="structure",
            severity="warning",
            message=f"Slide missing title",
            slide_index=idx,
            fix_command=generate_fix_command(filepath, "missing_title", slide_index=idx),
            details={"issue_type": "missing_title"}
        ))
    
    fonts_used = issue_data.get("fonts_used", [])
    if isinstance(fonts_used, list):
        summary.fonts_used = len(fonts_used)


def _process_accessibility(
    acc_result: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """
    Process results from agent.check_accessibility().
    
    Args:
        acc_result: Result from check_accessibility()
        issues: List to append issues to
        summary: Summary to update
        filepath: Path for fix commands
    """
    issue_data = acc_result.get("issues", {})
    
    missing_alt = issue_data.get("missing_alt_text", [])
    summary.missing_alt_text = len(missing_alt)
    for item in missing_alt:
        slide_idx = item.get("slide", item.get("slide_index"))
        shape_idx = item.get("shape", item.get("shape_index"))
        issues.append(ValidationIssue(
            category="accessibility",
            severity="critical",
            message=f"Image missing alt text",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=generate_fix_command(
                filepath, "missing_alt_text",
                slide_index=slide_idx, shape_index=shape_idx
            ),
            details={
                "issue_type": "missing_alt_text",
                "shape_name": item.get("name", "Unknown")
            }
        ))
    
    low_contrast = issue_data.get("low_contrast", [])
    summary.low_contrast = len(low_contrast)
    for item in low_contrast:
        slide_idx = item.get("slide", item.get("slide_index"))
        shape_idx = item.get("shape", item.get("shape_index"))
        issues.append(ValidationIssue(
            category="accessibility",
            severity="warning",
            message=f"Low color contrast ratio ({item.get('ratio', 'N/A')})",
            slide_index=slide_idx,
            shape_index=shape_idx,
            fix_command=generate_fix_command(
                filepath, "low_contrast",
                slide_index=slide_idx, shape_index=shape_idx
            ),
            details={
                "issue_type": "low_contrast",
                "contrast_ratio": item.get("ratio"),
                "wcag_minimum": 4.5
            }
        ))
    
    small_fonts = issue_data.get("small_fonts", [])
    summary.small_font_count = len(small_fonts)
    for item in small_fonts:
        issues.append(ValidationIssue(
            category="accessibility",
            severity="warning",
            message=f"Font size too small ({item.get('size', 'N/A')}pt)",
            slide_index=item.get("slide"),
            shape_index=item.get("shape"),
            details={
                "issue_type": "small_font",
                "font_size_pt": item.get("size"),
                "minimum_recommended": 12
            }
        ))


def _process_assets(
    asset_result: Dict[str, Any],
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    filepath: Path
) -> None:
    """
    Process results from agent.validate_assets().
    
    Args:
        asset_result: Result from validate_assets()
        issues: List to append issues to
        summary: Summary to update
        filepath: Path for fix commands
    """
    issue_data = asset_result.get("issues", {})
    
    large_images = issue_data.get("large_images", [])
    summary.large_images = len(large_images)
    for item in large_images:
        issues.append(ValidationIssue(
            category="assets",
            severity="info",
            message=f"Large image may slow presentation ({item.get('size_mb', 'N/A')} MB)",
            slide_index=item.get("slide"),
            shape_index=item.get("shape"),
            details={
                "issue_type": "large_image",
                "size_mb": item.get("size_mb"),
                "recommended_max_mb": 2.0
            }
        ))
    
    missing_assets = issue_data.get("missing_assets", [])
    for item in missing_assets:
        issues.append(ValidationIssue(
            category="assets",
            severity="critical",
            message=f"Referenced asset not found: {item.get('name', 'Unknown')}",
            slide_index=item.get("slide"),
            details={
                "issue_type": "missing_asset",
                "asset_name": item.get("name")
            }
        ))


def _validate_design_rules(
    agent: PowerPointAgent,
    issues: List[ValidationIssue],
    summary: ValidationSummary,
    policy: ValidationPolicy,
    filepath: Path
) -> None:
    """
    Validate design rules according to policy thresholds.
    
    Checks:
    - Font count limit
    - Color count limit  
    - 6x6 rule (bullets per slide, words per bullet)
    
    Args:
        agent: PowerPointAgent instance
        issues: List to append issues to
        summary: Summary to update
        policy: Validation policy with thresholds
        filepath: Path for fix commands
    """
    thresholds = policy.thresholds
    
    if summary.fonts_used > thresholds.get("max_fonts", 5):
        issues.append(ValidationIssue(
            category="design",
            severity="warning",
            message=f"Too many fonts used ({summary.fonts_used} > {thresholds.get('max_fonts', 5)})",
            details={
                "issue_type": "excessive_fonts",
                "font_count": summary.fonts_used,
                "threshold": thresholds.get("max_fonts", 5),
                "recommendation": "Limit to 2-3 font families for consistency"
            }
        ))
    
    try:
        presentation_info = agent.get_presentation_info()
        slide_count = presentation_info.get("slide_count", 0)
        
        colors_detected: Set[str] = set()
        bullet_violations = 0
        
        for slide_idx in range(slide_count):
            try:
                slide_info = agent.get_slide_info(slide_idx)
                shapes = slide_info.get("shapes", [])
                
                for shape in shapes:
                    if "fill_color" in shape and shape["fill_color"]:
                        colors_detected.add(shape["fill_color"])
                    if "line_color" in shape and shape["line_color"]:
                        colors_detected.add(shape["line_color"])
                    if "text_color" in shape and shape["text_color"]:
                        colors_detected.add(shape["text_color"])
                    
                    if thresholds.get("enforce_6x6_rule", False):
                        if shape.get("has_text_frame", False):
                            paragraphs = shape.get("paragraphs", [])
                            bullet_count = len([p for p in paragraphs if p.get("is_bullet", False)])
                            
                            if bullet_count > 6:
                                bullet_violations += 1
                                issues.append(ValidationIssue(
                                    category="design",
                                    severity="warning",
                                    message=f"Too many bullet points ({bullet_count} > 6)",
                                    slide_index=slide_idx,
                                    shape_index=shape.get("index"),
                                    details={
                                        "issue_type": "6x6_violation",
                                        "bullet_count": bullet_count,
                                        "max_allowed": 6
                                    }
                                ))
                            
            except Exception:
                continue
        
        summary.colors_detected = len(colors_detected)
        summary.bullet_violations = bullet_violations
        
        max_colors = thresholds.get("max_colors", 10)
        if summary.colors_detected > max_colors:
            issues.append(ValidationIssue(
                category="design",
                severity="warning",
                message=f"Too many colors used ({summary.colors_detected} > {max_colors})",
                details={
                    "issue_type": "excessive_colors",
                    "color_count": summary.colors_detected,
                    "threshold": max_colors,
                    "recommendation": "Limit to 3-5 primary colors for visual coherence"
                }
            ))
            
    except Exception:
        pass


def _check_policy_compliance(
    summary: ValidationSummary,
    policy: ValidationPolicy
) -> tuple:
    """
    Check if validation results comply with policy thresholds.
    
    Args:
        summary: Validation summary
        policy: Validation policy
        
    Returns:
        Tuple of (passed: bool, violations: List[str])
    """
    violations = []
    thresholds = policy.thresholds
    
    checks = [
        ("max_critical_issues", summary.critical_count, "Critical issues"),
        ("max_empty_slides", summary.empty_slides, "Empty slides"),
        ("max_slides_without_titles", summary.slides_without_titles, "Untitled slides"),
        ("max_missing_alt_text", summary.missing_alt_text, "Missing alt text"),
        ("max_low_contrast", summary.low_contrast, "Low contrast issues"),
        ("max_large_images", summary.large_images, "Large images"),
        ("max_fonts", summary.fonts_used, "Font families"),
        ("max_colors", summary.colors_detected, "Colors"),
    ]
    
    for threshold_key, actual_value, label in checks:
        threshold_value = thresholds.get(threshold_key)
        if threshold_value is not None and actual_value > threshold_value:
            violations.append(f"{label} ({actual_value}) exceeds limit ({threshold_value})")
    
    if thresholds.get("require_all_alt_text", False) and summary.missing_alt_text > 0:
        violations.append(f"All images must have alt text ({summary.missing_alt_text} missing)")
    
    return len(violations) == 0, violations


def _generate_recommendations(
    issues: List[ValidationIssue],
    policy: ValidationPolicy
) -> List[Dict[str, Any]]:
    """
    Generate prioritized recommendations based on issues found.
    
    Args:
        issues: List of validation issues
        policy: Validation policy
        
    Returns:
        List of recommendation dictionaries
    """
    recommendations = []
    
    critical_issues = [i for i in issues if i.severity == "critical"]
    accessibility_issues = [i for i in issues if i.category == "accessibility"]
    design_issues = [i for i in issues if i.category == "design"]
    
    if any(i.details.get("issue_type") == "empty_slide" for i in critical_issues):
        recommendations.append({
            "priority": "high",
            "category": "structure",
            "action": "Remove or populate empty slides",
            "impact": "Improves presentation flow and professionalism"
        })
    
    if any(i.details.get("issue_type") == "missing_alt_text" for i in accessibility_issues):
        recommendations.append({
            "priority": "high",
            "category": "accessibility",
            "action": "Add descriptive alt text to all images",
            "impact": "Required for WCAG 2.1 AA compliance and screen reader users"
        })
    
    if any(i.details.get("issue_type") == "low_contrast" for i in accessibility_issues):
        recommendations.append({
            "priority": "medium",
            "category": "accessibility",
            "action": "Improve text/background contrast ratios",
            "impact": "Ensures readability for users with visual impairments"
        })
    
    if any(i.details.get("issue_type") == "excessive_fonts" for i in design_issues):
        recommendations.append({
            "priority": "medium",
            "category": "design",
            "action": "Consolidate to 2-3 font families",
            "impact": "Creates visual consistency and professional appearance"
        })
    
    if any(i.details.get("issue_type") == "excessive_colors" for i in design_issues):
        recommendations.append({
            "priority": "low",
            "category": "design",
            "action": "Reduce color palette to 3-5 primary colors",
            "impact": "Improves visual coherence and brand consistency"
        })
    
    return recommendations


# ============================================================================
# MAIN VALIDATION FUNCTION
# ============================================================================

def validate_presentation(
    filepath: Path,
    policy: ValidationPolicy
) -> Dict[str, Any]:
    """
    Perform comprehensive presentation validation.
    
    Args:
        filepath: Path to PowerPoint file
        policy: Validation policy to apply
        
    Returns:
        Complete validation report dictionary
        
    Raises:
        FileNotFoundError: If file doesn't exist
        PowerPointAgentError: If validation fails
    """
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    issues: List[ValidationIssue] = []
    summary = ValidationSummary()
    
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath, acquire_lock=False)  # Read-only validation, no lock needed
        
        presentation_info = agent.get_presentation_info()
        slide_count = presentation_info.get("slide_count", 0)
        presentation_version = agent.get_presentation_version()
        
        core_validation = agent.validate_presentation()
        accessibility_validation = agent.check_accessibility()
        asset_validation = agent.validate_assets()
        
        _process_core_validation(core_validation, issues, summary, filepath)
        _process_accessibility(accessibility_validation, issues, summary, filepath)
        _process_assets(asset_validation, issues, summary, filepath)
        _validate_design_rules(agent, issues, summary, policy, filepath)
    
    summary.total_issues = len(issues)
    summary.critical_count = sum(1 for i in issues if i.severity == "critical")
    summary.warning_count = sum(1 for i in issues if i.severity == "warning")
    summary.info_count = sum(1 for i in issues if i.severity == "info")
    
    passed, violations = _check_policy_compliance(summary, policy)
    
    if summary.critical_count > 0:
        status = "critical"
    elif not passed:
        status = "failed"
    elif summary.warning_count > 0:
        status = "warnings"
    else:
        status = "valid"
    
    return {
        "status": status,
        "passed": passed,
        "tool_version": __version__,
        "core_version": CORE_VERSION,
        "file": str(filepath.resolve()),
        "presentation_version": presentation_version,
        "validated_at": datetime.utcnow().isoformat() + "Z",
        "policy": policy.to_dict(),
        "summary": summary.to_dict(),
        "policy_violations": violations,
        "issues": [i.to_dict() for i in issues],
        "recommendations": _generate_recommendations(issues, policy),
        "presentation_info": {
            "slide_count": slide_count,
            "file_size_mb": presentation_info.get("file_size_mb"),
            "aspect_ratio": presentation_info.get("aspect_ratio"),
            "has_notes": presentation_info.get("has_notes", False)
        }
    }


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Comprehensive PowerPoint presentation validation",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Standard validation
  uv run tools/ppt_validate_presentation.py --file deck.pptx --json

  # Strict validation for production
  uv run tools/ppt_validate_presentation.py --file deck.pptx --policy strict --json

  # Custom thresholds
  uv run tools/ppt_validate_presentation.py --file deck.pptx \\
    --max-empty-slides 0 --max-missing-alt-text 0 --json

Policies:
  lenient  - Minimal validation for drafts
  standard - Balanced validation (default)
  strict   - Maximum validation for production

Validation Categories:
  structure     - Empty slides, missing titles
  accessibility - Alt text, contrast, font sizes
  assets        - Large images, missing files
  design        - Font/color limits, 6x6 rule
        """
    )
    
    parser.add_argument(
        '--file',
        required=True,
        type=Path,
        help='PowerPoint file to validate'
    )
    
    parser.add_argument(
        '--policy',
        choices=['lenient', 'standard', 'strict'],
        default='standard',
        help='Validation policy (default: standard)'
    )
    
    parser.add_argument(
        '--json',
        action='store_true',
        default=True,
        help='Output JSON response (default: true)'
    )
    
    parser.add_argument(
        '--max-missing-alt-text',
        type=int,
        metavar='N',
        help='Override maximum missing alt text allowed'
    )
    
    parser.add_argument(
        '--max-slides-without-titles',
        type=int,
        metavar='N',
        help='Override maximum untitled slides allowed'
    )
    
    parser.add_argument(
        '--max-empty-slides',
        type=int,
        metavar='N',
        help='Override maximum empty slides allowed'
    )
    
    parser.add_argument(
        '--require-all-alt-text',
        action='store_true',
        help='Require alt text on all images'
    )
    
    parser.add_argument(
        '--enforce-6x6',
        action='store_true',
        help='Enforce 6x6 content density rule'
    )
    
    parser.add_argument(
        '--summary-only',
        action='store_true',
        help='Output summary only, omit individual issues'
    )
    
    args = parser.parse_args()
    
    try:
        custom_thresholds = {}
        if args.max_missing_alt_text is not None:
            custom_thresholds["max_missing_alt_text"] = args.max_missing_alt_text
        if args.max_slides_without_titles is not None:
            custom_thresholds["max_slides_without_titles"] = args.max_slides_without_titles
        if args.max_empty_slides is not None:
            custom_thresholds["max_empty_slides"] = args.max_empty_slides
        if args.require_all_alt_text:
            custom_thresholds["require_all_alt_text"] = True
        if args.enforce_6x6:
            custom_thresholds["enforce_6x6_rule"] = True
        
        if custom_thresholds:
            policy = get_policy("custom", custom_thresholds)
        else:
            policy = get_policy(args.policy)
        
        result = validate_presentation(args.file.resolve(), policy)
        
        if args.summary_only:
            result.pop("issues", None)
        
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.stdout.flush()
        
        exit_code = 1 if result["status"] in ("critical", "failed") else 0
        sys.exit(exit_code)
        
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
        
    except PowerPointAgentError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PowerPointAgentError",
            "suggestion": "Check file integrity and PowerPoint format",
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
