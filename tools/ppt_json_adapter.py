#!/usr/bin/env python3
"""
PowerPoint JSON Adapter Tool v3.1.1
Validates and normalizes JSON outputs from presentation CLI tools.

Author: PowerPoint Agent Team
License: MIT
Version: 3.1.1

Usage:
    uv run tools/ppt_json_adapter.py --schema ppt_get_info.schema.json --input raw.json

Behavior:
    - Validates input JSON against provided schema
    - Maps common alias keys to canonical keys
    - Emits normalized JSON to stdout
    - On validation failure, emits structured error JSON and exits non-zero

Exit Codes:
    0: Success (valid and normalized)
    2: Validation Error (schema validation failed)
    3: Input Load Error (could not read input file)
    5: Schema Load Error (could not read schema file)

Changelog v3.1.1:
    - Added hygiene block for JSON pipeline safety
    - Fixed ERROR_TEMPLATE bug causing duplicate keys
    - Added status wrapper to success output
    - Added tool_version to all outputs
    - Improved error response format
"""

import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that JSON parsers only see valid JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import argparse
import json
import hashlib
from typing import Dict, Any, Optional, List, Union
from pathlib import Path

try:
    from jsonschema import validate, ValidationError
except ImportError:
    validate = None
    ValidationError = Exception

# ============================================================================
# CONSTANTS
# ============================================================================

__version__ = "3.1.1"

# Alias mapping table for common drifted/variant keys across tool versions
ALIAS_MAP = {
    # Slide count variants
    "slides_count": "slide_count",
    "slidesTotal": "slide_count",
    "num_slides": "slide_count",
    "total_slides": "slide_count",
    
    # Slides list variants
    "slides_list": "slides",
    "slidesList": "slides",
    
    # Probe variants
    "probe_time": "probe_timestamp",
    "probeTime": "probe_timestamp",
    "probed_at": "probe_timestamp",
    
    # Permission variants
    "canWrite": "can_write",
    "writeable": "can_write",
    "canRead": "can_read",
    "readable": "can_read",
    
    # Size variants
    "maxImageSizeMB": "max_image_size_mb",
    "max_image_size": "max_image_size_mb",
    
    # Version variants
    "version": "presentation_version",
    "pres_version": "presentation_version",
    "file_version": "presentation_version"
}


# ============================================================================
# HELPER FUNCTIONS
# ============================================================================

def emit_error(
    error_code: str,
    message: str,
    details: Optional[Any] = None,
    retryable: bool = False
) -> None:
    """
    Emit a standardized error response to stdout.
    
    Args:
        error_code: Machine-readable error code
        message: Human-readable error message
        details: Additional error details
        retryable: Whether the operation can be retried
    """
    error_response = {
        "status": "error",
        "tool_version": __version__,
        "error": {
            "error_code": error_code,
            "message": message,
            "details": details,
            "retryable": retryable
        }
    }
    sys.stdout.write(json.dumps(error_response, indent=2) + "\n")
    sys.stdout.flush()


def load_json(path: Path) -> Dict[str, Any]:
    """
    Load JSON from a file path.
    
    Args:
        path: Path to JSON file
        
    Returns:
        Parsed JSON as dictionary
        
    Raises:
        FileNotFoundError: If file doesn't exist
        json.JSONDecodeError: If file contains invalid JSON
    """
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def map_aliases(obj: Any) -> Any:
    """
    Recursively map aliased keys to their canonical forms.
    
    Args:
        obj: Object to process (dict, list, or primitive)
        
    Returns:
        Object with aliased keys replaced by canonical keys
    """
    if isinstance(obj, dict):
        new_dict = {}
        for key, value in obj.items():
            canonical_key = ALIAS_MAP.get(key, key)
            if isinstance(value, dict):
                new_dict[canonical_key] = map_aliases(value)
            elif isinstance(value, list):
                new_dict[canonical_key] = [map_aliases(item) for item in value]
            else:
                new_dict[canonical_key] = value
        return new_dict
    elif isinstance(obj, list):
        return [map_aliases(item) for item in obj]
    else:
        return obj


def compute_presentation_version(info_obj: Dict[str, Any]) -> Optional[str]:
    """
    Compute a best-effort presentation_version if missing.
    
    This is a fallback approximation when the actual version from
    PowerPointAgent is unavailable. It uses available metadata to
    produce a deterministic hash.
    
    NOTE: This does NOT include shape geometry (left:top:width:height)
    as specified in the Core Handbook. It is only used when actual
    version tracking data is missing from the input.
    
    Args:
        info_obj: Presentation info dictionary
        
    Returns:
        SHA-256 hash string (first 16 chars) or None if computation fails
    """
    try:
        slides = info_obj.get("slides", [])
        
        slide_identifiers = []
        for slide in slides:
            slide_id = slide.get("id", slide.get("index", slide.get("slide_index", "")))
            slide_identifiers.append(str(slide_id))
        
        slide_ids_str = ",".join(slide_identifiers)
        
        file_path = info_obj.get("file", info_obj.get("filepath", ""))
        slide_count = info_obj.get("slide_count", len(slides))
        
        hash_input = f"{file_path}-{slide_count}-{slide_ids_str}"
        
        full_hash = hashlib.sha256(hash_input.encode("utf-8")).hexdigest()
        return full_hash[:16]
        
    except Exception:
        return None


def should_compute_version(schema: Dict[str, Any]) -> bool:
    """
    Determine if this schema type should have a computed version.
    
    Args:
        schema: JSON Schema dictionary
        
    Returns:
        True if presentation_version should be computed when missing
    """
    schema_id = schema.get("$id", "")
    schema_title = schema.get("title", "").lower()
    
    version_relevant_patterns = [
        "ppt_get_info",
        "get_info",
        "presentation_info",
        "ppt_capability_probe",
        "capability_probe"
    ]
    
    for pattern in version_relevant_patterns:
        if pattern in schema_id.lower() or pattern in schema_title:
            return True
    
    required_fields = schema.get("required", [])
    if "presentation_version" in required_fields:
        return True
    
    return False


# ============================================================================
# MAIN LOGIC
# ============================================================================

def adapt_json(
    schema_path: Path,
    input_path: Path
) -> Dict[str, Any]:
    """
    Validate and normalize JSON input against schema.
    
    Args:
        schema_path: Path to JSON Schema file
        input_path: Path to input JSON file
        
    Returns:
        Normalized and validated JSON wrapped in success response
    """
    if validate is None:
        emit_error(
            "DEPENDENCY_ERROR",
            "jsonschema library not installed",
            details={"required_package": "jsonschema"},
            retryable=False
        )
        sys.exit(5)
    
    try:
        schema = load_json(schema_path)
    except FileNotFoundError:
        emit_error(
            "SCHEMA_NOT_FOUND",
            f"Schema file not found: {schema_path}",
            details={"path": str(schema_path)},
            retryable=False
        )
        sys.exit(5)
    except json.JSONDecodeError as e:
        emit_error(
            "SCHEMA_PARSE_ERROR",
            f"Invalid JSON in schema file: {e.msg}",
            details={"path": str(schema_path), "line": e.lineno, "column": e.colno},
            retryable=False
        )
        sys.exit(5)
    except Exception as e:
        emit_error(
            "SCHEMA_LOAD_ERROR",
            str(e),
            details={"path": str(schema_path)},
            retryable=False
        )
        sys.exit(5)
    
    try:
        raw_input = load_json(input_path)
    except FileNotFoundError:
        emit_error(
            "INPUT_NOT_FOUND",
            f"Input file not found: {input_path}",
            details={"path": str(input_path)},
            retryable=True
        )
        sys.exit(3)
    except json.JSONDecodeError as e:
        emit_error(
            "INPUT_PARSE_ERROR",
            f"Invalid JSON in input file: {e.msg}",
            details={"path": str(input_path), "line": e.lineno, "column": e.colno},
            retryable=True
        )
        sys.exit(3)
    except Exception as e:
        emit_error(
            "INPUT_LOAD_ERROR",
            str(e),
            details={"path": str(input_path)},
            retryable=True
        )
        sys.exit(3)
    
    normalized = map_aliases(raw_input)
    
    if "presentation_version" not in normalized:
        if should_compute_version(schema):
            computed_version = compute_presentation_version(normalized)
            if computed_version:
                normalized["presentation_version"] = computed_version
                normalized["_version_computed"] = True
    
    try:
        validate(instance=normalized, schema=schema)
    except ValidationError as ve:
        schema_path_str = list(ve.schema_path) if ve.schema_path else None
        emit_error(
            "SCHEMA_VALIDATION_ERROR",
            ve.message,
            details={
                "schema_path": schema_path_str,
                "validator": ve.validator,
                "validator_value": str(ve.validator_value) if ve.validator_value else None,
                "instance_path": list(ve.absolute_path) if ve.absolute_path else None
            },
            retryable=False
        )
        sys.exit(2)
    
    return {
        "status": "success",
        "tool_version": __version__,
        "schema_used": str(schema_path),
        "input_file": str(input_path),
        "aliases_mapped": _count_mapped_aliases(raw_input, normalized),
        "data": normalized
    }


def _count_mapped_aliases(original: Any, normalized: Any) -> int:
    """
    Count how many aliases were mapped during normalization.
    
    Args:
        original: Original input object
        normalized: Normalized object
        
    Returns:
        Number of keys that were remapped
    """
    count = 0
    
    if isinstance(original, dict) and isinstance(normalized, dict):
        for key in original:
            if key in ALIAS_MAP:
                count += 1
            if key in original and isinstance(original[key], (dict, list)):
                canonical = ALIAS_MAP.get(key, key)
                if canonical in normalized:
                    count += _count_mapped_aliases(original[key], normalized[canonical])
    elif isinstance(original, list) and isinstance(normalized, list):
        for orig_item, norm_item in zip(original, normalized):
            count += _count_mapped_aliases(orig_item, norm_item)
    
    return count


# ============================================================================
# CLI INTERFACE
# ============================================================================

def main():
    parser = argparse.ArgumentParser(
        description="Validate and normalize JSON outputs from presentation CLI tools",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Validate and normalize tool output
  uv run tools/ppt_json_adapter.py \\
    --schema schemas/ppt_get_info.schema.json \\
    --input raw_output.json

  # Pipeline usage
  uv run tools/ppt_get_info.py --file deck.pptx --json > raw.json
  uv run tools/ppt_json_adapter.py --schema schemas/ppt_get_info.schema.json --input raw.json

Exit Codes:
  0: Success - valid JSON emitted
  2: Validation Error - input doesn't match schema
  3: Input Load Error - couldn't read input file
  5: Schema Load Error - couldn't read schema file

Alias Mapping:
  The adapter normalizes common key variations:
  - slides_count -> slide_count
  - slidesTotal -> slide_count
  - probe_time -> probe_timestamp
  - canWrite -> can_write
  etc.
        """
    )
    
    parser.add_argument(
        "--schema",
        required=True,
        type=Path,
        help="Path to JSON Schema file"
    )
    
    parser.add_argument(
        "--input",
        required=True,
        type=Path,
        help="Path to raw JSON input file"
    )
    
    args = parser.parse_args()
    
    result = adapt_json(
        schema_path=args.schema,
        input_path=args.input
    )
    
    sys.stdout.write(json.dumps(result, indent=2) + "\n")
    sys.stdout.flush()
    sys.exit(0)


if __name__ == "__main__":
    main()
