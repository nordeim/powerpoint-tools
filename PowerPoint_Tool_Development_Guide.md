# CONTRIBUTING_TOOLS.md

This handbook is the definitive reference for creating new CLI tools for the PowerPoint Agent suite. It is **self-sufficient**, containing all necessary templates, API references, and patterns required to build a compliant tool without examining the source code of other files.

---

## 1. The Design Contract

All tools must strictly adhere to these 4 principles to ensure compatibility with the AI Agent:

1.  **Atomic & Stateless:** Tools must open a file, perform one specific action, save, and exit. Do not assume previous state.
2.  **CLI Interface:** Use `argparse`. Complex data (positions, lists) must be passed as **JSON strings**.
3.  **JSON Output:**
    *   **STDOUT:** Must contain *only* the final JSON response.
    *   **STDERR:** Use for logging/debugging.
    *   **Exit Codes:** `0` for Success, `1` for Error.
4.  **Path Safety:** All file paths must be validated using `pathlib.Path` before execution.

---

## 2. Governance Principles (NEW SECTION)
### 2.1 Clone-Before-Edit Principle
**MANDATORY**: Always work on cloned copies, never source files. This is the first, non-negotiable rule.

```python
# ✅ CORRECT: Clone first, then operate
from core.powerpoint_agent_core import PowerPointAgent

# Open the source presentation, then clone it
with PowerPointAgent(Path("/source/template.pptx")) as agent:
    agent.open(Path("/source/template.pptx"))
    new_agent = agent.clone_presentation(Path("/work/modified.pptx"))
    new_agent.save()
    new_agent.close()
    
# Now operate on the work copy
with PowerPointAgent(Path("/work/modified.pptx")) as agent:
    agent.open(Path("/work/modified.pptx"))
    # ... operations ...
    agent.save()
```

### 2.2 Presentation Versioning Protocol
Tools must track presentation versions to prevent race conditions and conflicts:

```python
# ✅ CORRECT: Version tracking pattern
with PowerPointAgent(filepath) as agent:
    agent.open(filepath)
    
    # Capture initial version
    info_before = agent.get_presentation_info()
    initial_version = info_before["presentation_version"]
    
    # Perform operations
    result = agent.some_operation()
    
    # Capture new version
    info_after = agent.get_presentation_info()
    new_version = info_after["presentation_version"]
    
    # Return version tracking in response
    return {
        "status": "success",
        "file": str(filepath),
        "presentation_version_before": initial_version,
        "presentation_version_after": new_version,
        "changes_made": result
    }
```

**Version Format**: SHA-256 hex string (first 16 characters for brevity)
**Version Computation**: Hash of file path + slide count + slide IDs + modification timestamp

### 2.3 Approval Token System
**CRITICAL OPERATIONS REQUIRE APPROVAL**. The following operations require approval tokens (mandated by System Prompt v3.0 for all new destructive tools):
- `ppt_delete_slide.py` 🔒 **Actively enforced (exit code 4)**
- `ppt_remove_shape.py` 🔒 **Actively enforced (exit code 4)**
- Mass text replacements without dry-run
- Background replacements on all slides
- Any operation marked `critical: true` in manifest

**Token Structure**:
```json
{
    "token_id": "apt-YYYYMMDD-NNN",
    "manifest_id": "manifest-xxx",
    "user": "user@domain.com",
    "issued": "ISO8601",
    "expiry": "ISO8601", 
    "scope": ["delete:slide", "replace:all", "remove:shape"],
    "single_use": true,
    "signature": "HMAC-SHA256:base64.signature"
}
```

**Enforcement Protocol**:
```python
def validate_approval_token(token: str, required_scope: str) -> bool:
    """
    Validate approval token for destructive operations.
    
    Args:
        token: Base64-encoded token string
        required_scope: Required scope (e.g., "delete:slide")
        
    Returns:
        bool: True if token is valid and has required scope
        
    Raises:
        PermissionError: If token is invalid, expired, or lacks scope
    """
    if not token:
        raise PermissionError(f"Approval token required for {required_scope} operation")
    
    # Token validation logic here
    # ...
    
    if required_scope not in decoded_token["scope"]:
        raise PermissionError(f"Token lacks required scope: {required_scope}")
    
    return True
```

### 2.4 Shape Index Management Best Practices
**CRITICAL**: Shape indices shift after structural operations. Tools must handle this correctly.

**Operations That Invalidate Indices**:
| Operation | Effect |
|-----------|--------|
| `add_shape()` | Adds new index at end |
| `remove_shape()` | Shifts subsequent indices down |
| `set_z_order()` | Reorders indices (requires immediate refresh) |
| `delete_slide()` | Invalidates all indices on slide |
| `add_slide()` | Invalidates slide indices |

**Best Practices**:
```python
# ❌ WRONG - indices become stale after structural changes
result1 = agent.add_shape(slide_index=0, ...)  # Returns shape_index: 5
result2 = agent.add_shape(slide_index=0, ...)  # Returns shape_index: 6
agent.remove_shape(slide_index=0, shape_index=5)
agent.format_shape(slide_index=0, shape_index=6, ...)  # ❌ Now index 5!

# ✅ CORRECT - re-query after structural changes
result1 = agent.add_shape(slide_index=0, ...)
result2 = agent.add_shape(slide_index=0, ...)
agent.remove_shape(slide_index=0, shape_index=result1["shape_index"])

# IMMEDIATELY refresh indices
slide_info = agent.get_slide_info(slide_index=0)

# Find target shape by characteristics
for shape in slide_info["shapes"]:
    if shape["name"] == "target_shape":
        agent.format_shape(slide_index=0, shape_index=shape["index"], ...)
```

**Rule**: After any operation that affects shape indices, tools must call `get_slide_info()` and use the refreshed indices for subsequent operations.

## 3. The Master Template

Copy this code to start a new tool (e.g., `tools/ppt_new_feature.py`).

```python
#!/usr/bin/env python3
"""
[Tool Name]
[Short Description of what the tool does]

Usage:
    uv python ppt_new_feature.py --file deck.pptx --param value --json

Exit Codes:
    0: Success
    1: Error (Standard)
    2-5: Advanced Error Codes (Optional - see Error Handling Standards)
"""

import sys
import json
import argparse
from pathlib import Path
from typing import Dict, Any, Optional

# 1. PATH SETUP: Allow importing 'core' without installation
sys.path.insert(0, str(Path(__file__).parent.parent))

# 2. IMPORTS: Bring in the Core Agent and Exceptions
from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError, 
    SlideNotFoundError,
    ShapeNotFoundError,
    LayoutNotFoundError,
    ApprovalTokenError
)
from core.strict_validator import validate_against_schema, ValidationError

def logic_function(filepath: Path, param: str) -> Dict[str, Any]:
    """
    The main logic handler.
    1. Validate Inputs
    2. Open Agent
    3. Execute Core Method
    4. Save & Return Info
    """
    # Validate file exists
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    # Context Manager handles Open/Lock/Close automatically
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # Capture presentation version before changes
        info_before = agent.get_presentation_info()
        version_before = info_before["presentation_version"]
        
        # --- CORE LOGIC HERE ---
        # Example: Validating a slide index
        total_slides = agent.get_slide_count()
        if total_slides == 0:
            raise PowerPointAgentError("Presentation is empty")
            
        # Example: Calling a core method
        # result_data = agent.some_core_method(param)
        
        # Save changes
        agent.save()
        
        # Get fresh info for response including new version
        info_after = agent.get_presentation_info()
        version_after = info_after["presentation_version"]
        
    # Return standardized success dictionary with version tracking
    return {
        "status": "success",
        "file": str(filepath.resolve()),
        "action_performed": "new_feature",
        "presentation_version_before": version_before,
        "presentation_version_after": version_after,
        "details": {
            "param_used": param,
            "total_slides": info_after["slide_count"]
        }
    }

def main():
    # 3. ARGUMENT PARSING
    parser = argparse.ArgumentParser(description="Tool Description")
    
    # Standard Argument: File Path
    parser.add_argument(
        '--file', 
        required=True, 
        type=Path, 
        help='PowerPoint file path (absolute path required)'
    )
    
    # Custom Arguments
    parser.add_argument(
        '--param', 
        required=True, 
        help='Description of parameter'
    )
    
    # Standard Argument: JSON Flag (Required convention)
    parser.add_argument(
        '--json', 
        action='store_true', 
        default=True, 
        help='Output JSON response'
    )
    
    # Governance: Approval token for destructive operations
    parser.add_argument(
        '--approval-token',
        type=str,
        help='Approval token for destructive operations (required for delete/remove operations)'
    )
    
    args = parser.parse_args()
    
    # 4. VALIDATION & GOVERNANCE CHECKS
    # Clone-before-edit check (for tools that modify files)
    if not str(args.file).startswith('/work/') and not str(args.file).startswith('work_'):
        print(json.dumps({
            "status": "error",
            "error": "Safety violation: Direct editing of source files prohibited",
            "error_type": "SafetyViolationError",
            "suggestion": "Always clone files first using ppt_clone_presentation.py before editing"
        }, indent=2))
        sys.exit(1) # Standard error code
    
    # 5. ERROR HANDLING & OUTPUT
    try:
        # Validate approval token if required
        if hasattr(args, 'requires_approval') and args.requires_approval:
            if not args.approval_token:
                raise PermissionError("Approval token required for destructive operation")
        
        result = logic_function(filepath=args.file, param=args.param)
        print(json.dumps(result, indent=2))
        sys.exit(0)
        
    except FileNotFoundError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "FileNotFoundError",
            "suggestion": "Verify file path exists and is accessible"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except SlideNotFoundError as e:
        error_result = {
            "status": "error", 
            "error": e.message,
            "error_type": "SlideNotFoundError",
            "details": e.details,
            "suggestion": "Use ppt_get_info.py to check available slides"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except ShapeNotFoundError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "ShapeNotFoundError", 
            "details": e.details,
            "suggestion": "Use ppt_get_slide_info.py to refresh shape indices"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)
        
    except PermissionError as e:
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": "PermissionError",
            "suggestion": "Obtain approval token for destructive operation"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1) # Use 1 for standard permission error, 4 is advanced
        
    except ValidationError as e:
        error_result = {
            "status": "error",
            "error": e.message,
            "error_type": "ValidationError",
            "details": e.details,
            "suggestion": "Fix input data to match schema requirements"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1) # Use 1 for standard validation error, 2 is advanced
        
    except Exception as e:
        # Standard catch-all
        error_result = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__,
            "hint": "Check logs for detailed error information"
        }
        print(json.dumps(error_result, indent=2))
        sys.exit(1)

if __name__ == "__main__":
    main()
```

---

## 3. Data Structures Reference

When passing complex arguments to `PowerPointAgent` methods, use these dictionary schemas.

### **Position Dictionary** (`Dict[str, Any]`)
Used in: `add_text_box`, `insert_image`, `add_chart`, `add_shape`

*   **Percentage (Recommended):** `{"left": "10%", "top": "20%"}`
*   **Absolute (Inches):** `{"left": 1.5, "top": 2.0}`
*   **Anchor:** `{"anchor": "center", "offset_x": 0, "offset_y": -0.5}`
    *   *Anchors:* `top_left`, `top_center`, `top_right`, `center_left`, `center`, `center_right`, `bottom_left`, `bottom_center`, `bottom_right`
*   **Grid:** `{"grid_row": 2, "grid_col": 2, "grid_size": 12}`

### **Size Dictionary** (`Dict[str, Any]`)
Used in: `add_text_box`, `insert_image`, `add_chart`, `add_shape`

*   **Percentage:** `{"width": "50%", "height": "50%"}`
*   **Absolute:** `{"width": 5.0, "height": 3.0}`
*   **Auto (Aspect Ratio):** `{"width": "50%", "height": "auto"}`

### **Colors**
*   **Format:** Hex String `"#FF0000"` or `"#0070C0"`.

---

## 4. Core API Cheatsheet

You do not need to check `powerpoint_agent_core.py`. Use this reference for available methods on the `agent` instance.

### **File & Info**
| Method | Args | Returns |
| :--- | :--- | :--- |
| `create_new()` | `template: Path=None` | `None` |
| `open()` | `filepath: Path` | `None` |
| `save()` | `filepath: Path=None` | `None` |
| `get_slide_count()` | *None* | `int` |
| `get_presentation_info()` | `None` | `Dict` (metadata with `presentation_version`) |
| `get_slide_info()` | `slide_index: int` | `Dict` (shapes/text) |

### **Slide Manipulation**
| Method | Args | Returns |
| :--- | :--- | :--- |
| `add_slide()` | `layout_name: str, index: int=None` | `Dict[str, Any]` (slide_index, layout_name, total_slides, presentation_version_before/after) |
| `delete_slide()` | `index: int, approval_token: str=None` | `Dict[str, Any]` (deleted_index, previous_count, new_count, presentation_version_before/after) ⚠️ **Requires approval token** |
| `duplicate_slide()` | `index: int` | `Dict[str, Any]` (new_slide_index, total_slides, presentation_version_before/after) |
| `reorder_slides()` | `from_index: int, to_index: int` | `Dict[str, Any]` (from_index, to_index, total_slides, presentation_version_before/after) |
| `set_slide_layout()` | `slide_index: int, layout_name: str` | `None` |

### **Content Creation**
| Method | Args | Notes |
| :--- | :--- | :--- |
| `add_text_box()` | `slide_index, text, position, size, font_name=None, font_size=18, bold=False, italic=False, color=None, alignment="left"` | See Data Structures |
| `add_bullet_list()` | `slide_index, items: List[str], position, size, bullet_style="bullet", font_size=18, font_name=None` | Styles: `bullet`, `numbered`, `none` |
| `set_title()` | `slide_index, title: str, subtitle: str=None` | Uses layout placeholders |
| `insert_image()` | `slide_index, image_path, position, size=None, alt_text=None, compress=False` | Handles `auto` size. alt_text for accessibility |
| `add_shape()` | `slide_index, shape_type, position, size, fill_color=None, fill_opacity=1.0, line_color=None, line_opacity=1.0, line_width=1.0, text=None` | Types: `rectangle`, `arrow`, etc. **Opacity range: 0.0-1.0** |
| `replace_image()` | `slide_index, old_image_name: str, new_image_path, compress=False` | Replace by name or partial match |
| `add_chart()` | `slide_index, chart_type, data: Dict, position, size, title=None` | Data: `{"categories":[], "series":[]}` |
| `add_table()` | `slide_index, rows, cols, position, size, data: List[List]=None, header_row=True` | Data is 2D array. header_row for styling hint |

### **Formatting & Editing**
| Method | Args | Notes |
| :--- | :--- | :--- |
| `format_text()` | `slide_index, shape_index, font_name=None, font_size=None, bold=None, italic=None, color=None` | Update text formatting |
| `format_shape()` | `slide_index, shape_index, fill_color=None, fill_opacity=None, line_color=None, line_opacity=None, line_width=None` | **Opacity range: 0.0-1.0** ⚠️ `transparency` parameter **DEPRECATED** - use `fill_opacity` instead |
| `replace_text()` | `find: str, replace: str, match_case: bool=False` | Global text replacement |
| `remove_shape()` | `slide_index, shape_index` | Remove shape from slide ⚠️ **Requires approval token** |
| `set_z_order()` | `slide_index, shape_index, action` | Actions: `bring_to_front`, `send_to_back`, `bring_forward`, `send_backward` ⚠️ **Refresh indices after** |
| `reposition_shape()` | `slide_index, shape_index, position=None, size=None` | Move and/or resize shapes by absolute inches |
| `set_shape_text()` | `slide_index, shape_index, text` | Update text content of existing shapes |
| `add_connector()` | `slide_index, connector_type, start_shape_index, end_shape_index` | Types: `straight`, `elbow`, `curve` |
| `crop_image()` | `slide_index, shape_index, crop_box: Dict` | crop_box: `{"left": %, "top": %, "right": %, "bottom": %}` |
| `set_image_properties()` | `slide_index, shape_index, alt_text=None` | Set accessibility |

### **Validation**
| Method | Returns |
| :--- | :--- |
| `check_accessibility()` | `Dict` (WCAG issues) |
| `validate_presentation()` | `Dict` (Empty slides, missing assets) |

### **Chart & Presentation Operations**
| Method | Args | Notes |
| :--- | :--- | :--- |
| `update_chart_data()` | `slide_index, chart_index, data: Dict` | Update existing chart data |
| `format_chart()` | `slide_index, chart_index, title=None, legend_position=None` | Modify chart appearance |
| `add_notes()` | `slide_index, text, mode="append"` | Modes: `append`, `prepend`, `overwrite` (v3.1.0+) |
| `extract_notes()` | `None` | Returns `Dict[int, str]` of all notes by slide |
| `set_footer()` | `slide_index=None, text=None, show_number=False, show_date=False` | Configure slide footer |
| `set_background()` | `slide_index=None, color=None, image_path=None` | Set slide or presentation background |

## 6. Error Handling Standards (NEW SECTION)
### Exit Code Matrix
| Code | Category | Meaning | Retryable | Action |
|------|----------|---------|-----------|--------|
| 0 | Success | Operation completed | N/A | Proceed |
| 1 | Usage Error | Invalid arguments | No | Fix arguments |
| 2 | Validation Error | Schema/content invalid | No | Fix input |
| 3 | Transient Error | Timeout, I/O, network | Yes | Retry with backoff |
| 4 | Permission Error | Approval token missing/invalid | No | Obtain token |
| 5 | Internal Error | Unexpected failure | Maybe | Investigate |

### Standard Error Response Format
```json
{
    "status": "error",
    "error": {
        "error_code": "SCHEMA_VALIDATION_ERROR",
        "message": "Human-readable description",
        "details": {"path": "$.slides[0].layout"},
        "retryable": false,
        "hint": "Check that layout name matches available layouts from probe"
    }
}
```

### 6.1 Platform-Independent Paths (NEW SUBSECTION)
**CRITICAL**: Always use `pathlib.Path` for file operations to ensure compatibility across Linux, macOS, and Windows.

```python
# ❌ WRONG - String manipulation is fragile
file_path = "/tmp/" + filename
if "\\" in file_path: ...

# ✅ CORRECT - Use pathlib
from pathlib import Path
file_path = Path("/tmp") / filename
if not file_path.exists(): ...
```

### Tool-Specific Error Examples

**Permission Error (Exit Code 4):**
```json
{
    "status": "error",
    "error": "Approval token required for slide deletion",
    "error_type": "PermissionError",
    "details": {
        "operation": "delete_slide",
        "slide_index": 5
    },
    "suggestion": "Generate approval token with scope 'delete:slide' and retry"
}
```

**Shape Index Error (Exit Code 1):**
```json
{
    "status": "error",
    "error": "Shape index 10 out of range (0-8)",
    "error_type": "ShapeNotFoundError",
    "details": {
        "requested": 10,
        "available": 9
    },
    "suggestion": "Refresh shape indices using ppt_get_slide_info.py before targeting shapes"
}
```

**Version Mismatch Error (Exit Code 1):**
```json
{
    "status": "error", 
    "error": "Presentation version mismatch - file was modified externally",
    "error_type": "VersionConflictError",
    "details": {
        "expected": "a1b2c3d4",
        "actual": "e5f6g7h8"
    },
    "suggestion": "Re-probe presentation and update manifest with current version"
}
```

## 7. Opacity & Transparency (v3.1.0+)
The toolkit supports semi-transparent shapes and fills for enhanced visual effects (v3.1.0+):

### Opacity Parameters:
- `fill_opacity`: Float from 0.0 (invisible) to 1.0 (opaque). Default: 1.0
- `line_opacity`: Float from 0.0 (invisible) to 1.0 (opaque). Default: 1.0
- `transparency`: **DEPRECATED** - Use `fill_opacity` instead. Inverse relationship: `opacity = 1.0 - transparency`

### Common Use Case - Text Readability Overlay:
```python
# ✅ MODERN (preferred - v3.1.0+)
agent.add_shape(
    slide_index=0,
    shape_type="rectangle",
    position={"left": "0%", "top": "0%"},
    size={"width": "100%", "height": "100%"},
    fill_color="#FFFFFF",
    fill_opacity=0.15  # Subtle, non-competing overlay
)

# ⚠️ DEPRECATED (backward compatible but logs warning)
agent.format_shape(
    slide_index=0,
    shape_index=5,
    transparency=0.85  # Converts to fill_opacity=0.15 with warning
)
```
**WARNING:** Tools should use `fill_opacity` parameter exclusively. The `transparency` parameter is deprecated and will be removed in v4.0.

### Methods Supporting Opacity:
- `add_shape()` - `fill_opacity` and `line_opacity` parameters
- `format_shape()` - `fill_opacity` and `line_opacity` parameters
- `set_background()` - `fill_opacity` parameter for image backgrounds

## 8. Workflow Context (NEW SECTION)
### The 5-Phase Workflow
Tools are designed to work within a structured 5-phase workflow. Each tool should document which phase(s) it belongs to:

| Phase | Purpose | Tool Examples | Key Requirements |
|-------|---------|---------------|-------------------|
| **DISCOVER** | Deep inspection and capability probing | `ppt_capability_probe.py`, `ppt_get_info.py`, `ppt_get_slide_info.py` | Timeout handling, fallback probes, comprehensive metadata |
| **PLAN** | Manifest creation and design decisions | `ppt_create_from_structure.py`, `ppt_validate_manifest.py` | Schema validation, design rationale documentation |
| **CREATE** | Actual content creation and modification | `ppt_add_shape.py`, `ppt_add_slide.py`, `ppt_replace_text.py` | Version tracking, approval token enforcement, index freshness |
| **VALIDATE** | Quality assurance and compliance checking | `ppt_validate_presentation.py`, `ppt_check_accessibility.py` | WCAG 2.1 compliance, structural validation, contrast checking |
| **DELIVER** | Production handoff and documentation | `ppt_export_pdf.py`, `ppt_extract_notes.py`, `ppt_generate_manifest.py` | Complete audit trails, rollback commands, delivery packages |

### 8.1 Probe Resilience Pattern (NEW SUBSECTION)
**CRITICAL**: Discovery tools must be resilient to large files and timeouts. Implement the Timeout + Transient Slides + Graceful Degradation pattern.

**Core Pattern: Timeout + Transient Slides + Warnings**

The resilience pattern has 3 mandatory layers:

**Layer 1: Timeout Detection (Interval Checks)**
```python
import time

def detect_layouts(prs, timeout_seconds=30):
    """Detect layouts with timeout protection."""
    start_time = time.perf_counter()
    results = []
    
    for idx, layout in enumerate(prs.slide_layouts):
        # Check timeout at EACH iteration (critical for large templates)
        elapsed = time.perf_counter() - start_time
        if elapsed > timeout_seconds:
            warnings.append(
                f"Probe timeout at layout {idx} ({elapsed:.2f}s > {timeout_seconds}s) - "
                "returning partial results"
            )
            break  # Stop gracefully, return partial results
        
        results.append(process_layout(layout))
    
    return results
```
**Real implementation**: `ppt_capability_probe.py` lines 369-375 uses this pattern to stop deep analysis mid-probe when timeout expires.

**Layer 2: Transient Slides (Accurate Analysis)**
```python
def _add_transient_slide(prs, layout):
    """
    Safely add and remove a transient slide for deep analysis.
    Uses generator pattern to guarantee cleanup via finally block.
    
    Real implementation: ppt_capability_probe.py lines 294-313
    """
    slide = None
    added_index = -1
    try:
        # Add slide for analysis
        slide = prs.slides.add_slide(layout)
        added_index = len(prs.slides) - 1
        yield slide  # Caller analyzes the instantiated slide
        
    finally:
        # ALWAYS cleanup, even if analysis fails
        if added_index != -1 and added_index < len(prs.slides):
            try:
                rId = prs.slides._sldIdLst[added_index].rId
                prs.part.drop_rel(rId)
                del prs.slides._sldIdLst[added_index]
            except Exception:
                # Suppress cleanup errors to avoid masking analysis failures
                # File is not saved, so transient slide disappears anyway
                pass


# Usage pattern
for layout in prs.slide_layouts:
    with _add_transient_slide(prs, layout) as slide:
        # Get accurate placeholder positions from instantiated slide
        positions = extract_placeholder_positions(slide)
        # Slide automatically removed when exiting context
```
**Why transient slides?** Template positions are theoretical until instantiated. Transient slides let you get REAL positions without mutating the file (no `save()` call = no file change = atomic read).

**Layer 3: Partial Results + Warnings (Graceful Degradation)**
```python
def probe_presentation(filepath: Path, timeout_seconds: int = 30):
    """Probe with graceful degradation."""
    warnings = []
    info = []
    
    # For large files, limit scope
    all_layouts = list(prs.slide_layouts)
    max_layouts = 50  # Cap to prevent runaway analysis
    layouts_to_analyze = all_layouts[:max_layouts]
    
    if len(all_layouts) > max_layouts:
        info.append(f"Limited analysis to first {max_layouts} of {len(all_layouts)} layouts")
    
    # Perform analysis with timeout
    results = []
    start_time = time.perf_counter()
    for idx, layout in enumerate(layouts_to_analyze):
        if (time.perf_counter() - start_time) > timeout_seconds:
            warnings.append(f"Probe timeout - analyzed {idx} of {max_layouts} layouts")
            break  # Stop, return what we have
        results.append(analyze_layout(layout))
    
    # Always return partial results + metadata
    return {
        "status": "success",
        "layouts_analyzed": len(results),
        "layouts_total": len(all_layouts),
        "partial_results": True if len(results) < len(layouts_to_analyze) else False,
        "analysis_complete": len(results) == len(layouts_to_analyze),
        "layouts": results,
        "warnings": warnings,
        "info": info
    }
```
**Real implementation**: `ppt_capability_probe.py` lines 862-880 returns `analysis_complete` flag, list of partial results, and warnings.

**Complete Example: Safe Discovery Tool Template**
```python
def probe_with_resilience(filepath: Path, deep: bool = False, timeout_seconds: int = 30):
    """
    Complete resilience pattern for discovery tools.
    
    Combines:
    1. Pre-flight checks (file exists, not locked)
    2. Timeout protection
    3. Transient analysis (for deep mode)
    4. Graceful degradation
    5. Atomic verification
    6. Comprehensive error handling
    """
    
    # Pre-flight: Check file exists and is accessible
    if not filepath.exists():
        raise FileNotFoundError(f"File not found: {filepath}")
    
    try:
        with open(filepath, 'rb') as f:
            f.read(1)  # Try to read 1 byte - fails if locked
    except PermissionError:
        raise PermissionError(f"File is locked or inaccessible: {filepath}")
    
    # Atomic verification: Before state
    checksum_before = calculate_file_checksum(filepath)
    start_time = time.perf_counter()
    operation_id = str(uuid.uuid4())
    warnings = []
    
    prs = Presentation(str(filepath))
    results = []
    
    # Main analysis loop with timeout
    for idx, layout in enumerate(prs.slide_layouts):
        # Timeout check
        if (time.perf_counter() - start_time) > timeout_seconds:
            warnings.append(f"Probe timeout - analyzed {idx} layouts")
            break
        
        if deep:
            # Deep mode: Use transient slide for accurate analysis
            with _add_transient_slide(prs, layout) as slide:
                layout_data = analyze_layout_deep(slide)
        else:
            # Fast mode: Use template positions (less accurate)
            layout_data = analyze_layout_fast(layout)
        
        results.append(layout_data)
    
    duration_ms = (time.perf_counter() - start_time) * 1000
    
    # Atomic verification: After state
    checksum_after = calculate_file_checksum(filepath)
    atomic_safe = checksum_before == checksum_after
    
    return {
        "status": "success",
        "operation_id": operation_id,
        "duration_ms": int(duration_ms),
        "analysis_complete": len(results) == len(prs.slide_layouts),
        "layouts_analyzed": len(results),
        "layouts_total": len(prs.slide_layouts),
        "deep_analysis": deep,
        "atomic_verified": atomic_safe,
        "layouts": results,
        "warnings": warnings
    }
```
**Real implementation**: `ppt_capability_probe.py` lines 823-880 (full `probe_presentation()` function) demonstrates all 6 elements.

### Tool Classification Guidelines
When creating a new tool, classify it by phase:

```python
# Tool metadata should include phase classification
TOOL_METADATA = {
    "name": "ppt_add_shape",
    "version": "3.1.0",
    "primary_phase": "CREATE",
    "secondary_phases": ["VALIDATE"],  # For validation tools
    "requires_approval": False,
    "invalidates_indices": True,  # Affects shape indices
    "version_tracking": True  # Requires presentation version tracking
}
```

### Phase-Specific Requirements

**DISCOVER Phase Tools:**
- Must implement timeout handling (30 seconds default)
- Must have fallback probes (3 retries with exponential backoff)
- Must return comprehensive metadata including probe type
- Must handle probe failures gracefully

**CREATE Phase Tools:**
- Must track presentation versions (before/after)
- Must enforce approval tokens for destructive operations
- Must refresh shape indices after structural changes
- Must follow clone-before-edit principle

**VALIDATE Phase Tools:**
- Must return detailed violation reports with remediation suggestions
- Must categorize issues by severity (critical/warning/info)
- Must provide exact fix commands for accessibility issues
- Must enforce validation policies from manifest

## 9. Implementation Checklist
Before committing a new tool, verify:

### Governance & Safety
- [ ] **Clone-Before-Edit**: Does the tool work on `/work/` directory files only?
- [ ] **Approval Token**: Are destructive operations protected by token validation?
- [ ] **Version Tracking**: Does the tool capture `presentation_version_before` and `presentation_version_after`?
- [ ] **Index Freshness**: Does the tool refresh shape indices after structural operations?
- [ ] **Audit Trail**: Does the tool log all operations with timestamps and versions?

### Technical Requirements
- [ ] **JSON Argument Parsing**: If your tool takes a `Position` or `Size`, are you parsing the input string via `json.loads` in `argparse`?
- [ ] **Exit Codes**: Does it return correct exit codes (0-5) according to the matrix?
- [ ] **File Existence**: Do you check `if not filepath.exists()` before opening?
- [ ] **Self-Contained**: Does the tool run without needing to modify `core.py`?
- [ ] **Slide Bounds**: Do you check `if not 0 <= index < total_slides`?
- [ ] **Error Format**: Does the tool use the standard error response format with `error_type` and `suggestion`?

### v3.1.0+ Features
- [ ] **Opacity Handling**: Does the tool use `fill_opacity` instead of deprecated `transparency`?
- [ ] **Z-Order Management**: If using `set_z_order`, does the tool refresh indices afterward?
- [ ] **Speaker Notes**: For tools adding notes, do they support all modes (`append`, `prepend`, `overwrite`)?
- [ ] **Schema Validation**: Does the tool validate inputs against JSON schemas when applicable?

### Workflow Integration
- [ ] **Phase Classification**: Is the tool's primary/secondary phase documented?
- [ ] **Manifest Integration**: Does the tool update the change manifest when applicable?
- [ ] **Rollback Commands**: Are rollback commands provided for destructive operations?
- [ ] **Design Rationale**: For visual operations, is design rationale documented in responses?

## 10. Testing Requirements (NEW SECTION)
### Test Structure
```
tests/
├── test_core.py                  # Core library unit tests
├── test_shape_opacity.py         # Feature-specific tests  
├── test_tools/                   # CLI tool integration tests
│   ├── test_ppt_add_shape.py
│   └── ...
├── conftest.py                   # Shared fixtures
├── test_utils.py                 # Helper functions
└── assets/                       # Test files
    ├── sample.pptx
    └── template.pptx
```

### Required Test Coverage
| Category | What to Test |
|----------|--------------|
| **Happy Path** | Normal usage succeeds |
| **Edge Cases** | Boundary values (0, 1, max, empty) |
| **Error Cases** | Invalid inputs raise correct exceptions |
| **Validation** | Invalid ranges/formats rejected |
| **Backward Compat** | Deprecated features still work |
| **CLI Integration** | Tool produces valid JSON |
| **Governance** | Clone-before-edit enforced, tokens validated |
| **Version Tracking** | Presentation versions captured correctly |
| **Index Freshness** | Shape indices refreshed after structural changes |

### Test Pattern Example
```python
import pytest
from pathlib import Path

@pytest.fixture
def test_presentation(tmp_path):
    """Create a test presentation with blank slide."""
    pptx_path = tmp_path / "test.pptx"
    with PowerPointAgent() as agent:
        agent.create_new()
        agent.add_slide(layout_name="Blank")
        agent.save(pptx_path)
    return pptx_path

class TestAddShapeOpacity:
    """Tests for add_shape() opacity functionality."""
    
    def test_opacity_applied(self, test_presentation):
        """Test shape with valid opacity value."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            result = agent.add_shape(
                slide_index=0,
                shape_type="rectangle",
                position={"left": "10%", "top": "10%"},
                size={"width": "20%", "height": "20%"},
                fill_color="#0070C0",
                fill_opacity=0.5
            )
            agent.save()
        
        # Core method returns dict with styling details
        assert "shape_index" in result
        assert result["styling"]["fill_opacity"] == 0.5
        assert result["styling"]["fill_opacity_applied"] is True
    
    def test_approval_token_enforcement(self, test_presentation):
        """Test that destructive operations require approval tokens."""
        with PowerPointAgent(test_presentation) as agent:
            agent.open(test_presentation)
            
            with pytest.raises(PermissionError) as excinfo:
                agent.remove_shape(
                    slide_index=0,
                    shape_index=0,
                    approval_token=None  # Missing token
                )
            
            assert "Approval token required" in str(excinfo.value)
```

### Running Tests
```bash
# Run all tests
pytest tests/ -v

# Run specific file
pytest tests/test_shape_opacity.py -v

# Run with coverage
pytest tests/ --cov=core --cov-report=html

# Run parallel (faster)
pytest tests/ -v -n auto

# Stop on first failure
pytest tests/ -v -x
```

## 11. Contribution Workflow
Before Starting:
- Read this document — Understand the architecture
- Check existing tools — Don't duplicate functionality
- Review system prompt — Understand AI agent usage
- Set up environment:
  ```bash
  uv pip install -r requirements.txt
  uv pip install -r requirements-dev.txt
  ```

### PR Checklist
**Code Quality**
- [ ] Type hints on all function signatures
- [ ] Docstrings on all public functions
- [ ] Follows naming conventions
- [ ] `black` formatted
- [ ] `ruff` passes

**For New Tools**
- [ ] File named `ppt_<verb>_<noun>.py`
- [ ] Uses standard template structure with governance sections
- [ ] Outputs valid JSON to stdout only
- [ ] Exit code 0-5 according to matrix
- [ ] Validates paths with `pathlib.Path`
- [ ] All exceptions converted to JSON with standard format

## 12. E2E-Validated Troubleshooting Tips

These issues were discovered during end-to-end testing of the `powerpoint-skill`:

| Symptom | Root Cause | Fix |
|---------|-----------|-----|
| `ppt_add_slide.py --title` fails | No `--title` arg on this tool | Use `ppt_set_title.py` separately |
| `ppt_add_shape.py --shape-type` fails | Arg is `--shape` not `--shape-type` | Use `--shape rectangle` |
| `ppt_set_footer.py --show-page-number` fails | Arg is `--show-number` | Use `--show-number` |
| `ppt_remove_shape.py` silently succeeds without token | Tool didn't pass token to core | Fixed: now requires `--approval-token` |
| Color validation error on shapes | `RGBColor` is tuple, not object with `.red` | Fixed: use index access `[0]`, `[1]`, `[2]` |
| PDF/Image export fails | LibreOffice not installed | Install `libreoffice-impress` (optional dep) |
