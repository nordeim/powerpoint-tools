# 🔥 PowerPoint Agent Tools: Comprehensive Programming & Troubleshooting Guide (v3.1.0)

**Version**: 1.0  
**Applicable Core Version**: 3.1.0  
**Target Audience**: AI Coding Agents, Human Developers, DevOps Engineers  

## 1. Architectural Philosophy

The PowerPoint Agent Tools ecosystem is designed around a **Hub-and-Spoke architecture** optimized for stateless, atomic, and machine-parseable operations, combining the best of both source documents.

### 1.1 The Hub: `PowerPointAgent` (Core)
**Location**: `core/powerpoint_agent_core.py`  
**Role**: The "Operating System." It handles all direct interaction with the `.pptx` binary, manages file locking, enforces security (path traversal, approval tokens), and calculates state hashes (versioning).  
**Statefulness**: The Core instance is stateful while open (holds the `Presentation` object), but the Tools using it must treat it as ephemeral.

### 1.2 The Spokes: CLI Tools (`tools/*.py`)
**Location**: `tools/`  
**Role**: Thin wrappers around Core methods.  
**The Prime Directive**: JSON IN, JSON OUT.  
**Statelessness**: A tool must fully initialize, execute its task, save, and exit. It assumes no memory of previous commands.

### 1.3 The Four Pillars (Design Contract)
All tools must strictly adhere to these 4 principles to ensure compatibility with the AI Agent:

| Principle | Implementation | Why It Matters |
|-----------|----------------|----------------|
| **STATELESS** | Each call independent, no memory of previous calls | Reliable in distributed environments |
| **ATOMIC** | Open→Modify→Save→Close, one action per call | Predictable, recoverable workflows |
| **COMPOSABLE** | Tools can be chained, output feeds next input | Enables pipeline composition |
| **ACCESSIBLE** | WCAG 2.1 compliance, alt text, contrast, reading order | Creates inclusive presentations |
| **VISUAL-AWARE** | Typography scales, color theory, content density | Professional outputs |

## 2. The "Golden Rules" of Development

Violating these rules will break the CI/CD pipeline or cause the AI orchestrator to fail. These combine the strict hygiene rules of PROGRAMMING_GUIDE.md with the governance principles of PowerPoint_Tool_Development_Guide.md.

### 🔒 Rule 1: Output Hygiene is Non-Negotiable
**The Problem**: Tools are often chained via pipes (e.g., `tool.py | jq .`). If a library (like `pptx`) prints a warning, or if you use `print("Doing work...")`, the output becomes invalid JSON.

**The Requirement**:
- STDOUT is exclusively for the final JSON payload
- STDERR is for logs/debugs (but be careful, some pipelines capture `2>&1`)

**The Fix**: In v3.1.0, we apply Draconian Suppression at the top of every tool:
```python
# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that `jq` or other parsers only see your JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---
```

### 🔒 Rule 2: Clone-Before-Edit Principle
**MANDATORY**: Always work on cloned copies, never source files. This is the first, non-negotiable rule.

```python
# ✅ CORRECT: Clone first, then operate
from core.powerpoint_agent_core import PowerPointAgent

# Open the source, then clone
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

### 🔒 Rule 3: Fail Safely with JSON
If a tool crashes, it must still print a valid JSON object to stdout so the orchestrator knows why it failed.

**Bad**: Python Traceback dumped to shell.  
**Good**: `{"status": "error", "error": "IndexError...", "error_type": "IndexError"}`

### 🔒 Rule 4: Versioning is Mandatory
Every mutation (write) operation must capture the presentation state before and after the change.

**Why**: To detect race conditions and verify that the AI's intent was actually applied.  
**Implementation**: Core v3.1.0 methods return `presentation_version_before` and `presentation_version_after`. Tools must pass these through.

### 🔒 Rule 5: Approval Token System
**CRITICAL OPERATIONS REQUIRE APPROVAL**. The following operations require approval tokens:

- `ppt_delete_slide.py` 🔒 **Actively enforced (exit code 4)**
- `ppt_remove_shape.py` 🔒 **Actively enforced (exit code 4)**
- `ppt_merge_presentations.py` 🔒 **Actively enforced (exit code 4)**
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

## 3. Reference Tool Implementation (The "Perfect" Tool)

Use this template for creating new tools or refactoring existing ones. This combines the hygiene block from PROGRAMMING_GUIDE.md with the governance sections from PowerPoint_Tool_Development_Guide.md.

```python
#!/usr/bin/env python3
"""
Standard Tool Template v3.1.0
"""
import sys
import os

# --- HYGIENE BLOCK START ---
# CRITICAL: Redirect stderr to /dev/null immediately to prevent library noise.
# This guarantees that `jq` or other parsers only see your JSON on stdout.
sys.stderr = open(os.devnull, 'w')
# --- HYGIENE BLOCK END ---

import json
import argparse
import logging
from pathlib import Path
from typing import Dict, Any, Optional

# Configure logging to null (redundant but safe)
logging.basicConfig(level=logging.CRITICAL)

# Add parent directory to path to import core
sys.path.insert(0, str(Path(__file__).parent.parent))

from core.powerpoint_agent_core import (
    PowerPointAgent, 
    PowerPointAgentError,
    SlideNotFoundError,
    ShapeNotFoundError,
    LayoutNotFoundError,
    ApprovalTokenError
)
from core.strict_validator import ValidationError

def tool_logic(filepath: Path, param: str) -> Dict[str, Any]:
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
    
    # Clone-before-edit check
    if not str(filepath).startswith('/work/') and not str(filepath).startswith('work_'):
        raise PermissionError(
            "Safety violation: Direct editing of source files prohibited",
            details={"suggestion": "Always clone files first using ppt_clone_presentation.py before editing"}
        )
    
    # Context Manager handles Open/Save/Close/Locking
    with PowerPointAgent(filepath) as agent:
        agent.open(filepath)
        
        # 1. Capture Version
        info_before = agent.get_presentation_info()
        version_before = info_before["presentation_version"]
        
        # 2. Execute Core Method
        # V3.1.0 CHANGE: Core methods return DICTIONARIES, not just ints.
        result = agent.some_mutation_method(param)
        
        # 3. Save
        agent.save()
        
        # 4. Get Final State
        info_after = agent.get_presentation_info()
        version_after = info_after["presentation_version"]
    
    # 5. Construct Clean Response
    return {
        "status": "success",
        "file": str(filepath),
        "data": result,
        "presentation_version_before": version_before,
        "presentation_version_after": version_after
    }

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--file', required=True, type=Path)
    parser.add_argument('--param', required=True)
    parser.add_argument('--json', action='store_true', default=True)
    parser.add_argument('--approval-token', type=str, help='Approval token for destructive operations')
    args = parser.parse_args()

    try:
        # Validate approval token if required
        if args.requires_approval and not args.approval_token:
            raise ApprovalTokenError("Approval token required for destructive operation")
        
        result = tool_logic(args.file, args.param)
        # THE ONLY PRINT STATEMENT ALLOWED:
        sys.stdout.write(json.dumps(result, indent=2) + "\n")
        sys.exit(0)
    except Exception as e:
        # CATCH-ALL ERROR HANDLER
        error_res = {
            "status": "error",
            "error": str(e),
            "error_type": type(e).__name__
        }
        sys.stdout.write(json.dumps(error_res, indent=2) + "\n")
        sys.exit(1)

if __name__ == "__main__":
    main()
```

## 4. Core Library Internals & Gotchas

### 4.1 Geometry-Aware Versioning
**Concept**: A hash of the presentation state.  
**Gotcha**: In v2.0, only text content was hashed. If you moved a box, the version didn't change.  
**Fix**: In v3.1.0, `get_presentation_version` hashes `{left}:{top}:{width}:{height}` for every shape. Moving a shape by 1 pixel will change the version hash.

**Version Protocol**:
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

### 4.2 XML Manipulation (Opacity & Z-Order)
`python-pptx` does not support transparency or Z-order natively. We use `lxml` to hack the XML tree directly.

**Opacity**: We inject `<a:alpha val="50000"/>` into the color element. Note that Office uses a 0-100,000 scale (50000 = 50%). The Core converts 0.0-1.0 floats to this scale automatically.

**Z-Order**: We physically move the XML element within the `<p:spTree>`.  
**Warning**: Moving an element in the XML tree changes its index in the `shapes` collection.  
**Rule**: Always re-query `get_slide_info` after a Z-order change.

### 4.3 Shape Index Management Best Practices
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

### 4.4 Footer Strategy (The "Master Trap")
**The Trap**: `agent.prs.slide_masters[0].placeholders` might contain a footer. This makes you think the footer works.  
**The Reality**: Individual slides might have "Hide Background Graphics" on, or simply haven't instantiated that placeholder.  
**The Fix**: `ppt_set_footer.py` uses a Dual Strategy:
- Try setting the placeholder
- Check if any slides were actually updated (`slide_indices_updated`)
- If 0 slides updated, fall back to creating a Text Box Overlay

## 5. Data Structures Reference

When passing complex arguments to `PowerPointAgent` methods, use these dictionary schemas.

### Position Dictionary (`Dict[str, Any]`)
**Used in**: `add_text_box`, `insert_image`, `add_chart`, `add_shape`

**Percentage (Recommended)**: `{"left": "10%", "top": "20%"}`  
**Absolute (Inches)**: `{"left": 1.5, "top": 2.0}`  
**Anchor**: `{"anchor": "center", "offset_x": 0, "offset_y": -0.5}`  
**Anchors**: `top_left`, `top_center`, `top_right`, `center_left`, `center`, `center_right`, `bottom_left`, `bottom_center`, `bottom_right`  
**Grid**: `{"grid_row": 2, "grid_col": 3, "grid_size": 12}`

### Size Dictionary (`Dict[str, Any]`)
**Used in**: `add_text_box`, `insert_image`, `add_chart`, `add_shape`

**Percentage**: `{"width": "50%", "height": "50%"}`  
**Absolute**: `{"width": 5.0, "height": 3.0}`  
**Auto (Aspect Ratio)**: `{"width": "50%", "height": "auto"}`

### Colors
**Format**: Hex String `"#FF0000"` or `"#0070C0"`.

## 6. Core API Cheatsheet

You do not need to check `powerpoint_agent_core.py`. Use this reference for available methods on the `agent` instance.

### File & Info
| Method | Args | Returns |
|--------|------|---------|
| create_new() | template: Path=None | None |
| open() | filepath: Path | None |
| save() | filepath: Path=None | None |
| get_slide_count() | None | int |
| get_presentation_info() | None | Dict (metadata with presentation_version) |
| get_slide_info() | slide_index: int | Dict (shapes/text) |

### Slide Manipulation
| Method | Args | Returns |
|--------|------|---------|
| add_slide() | layout_name: str, index: int=None | Dict[str, Any] (slide_index, layout_name, total_slides, presentation_version_before/after) |
| delete_slide() | index: int, approval_token: str=None | Dict[str, Any] (deleted_index, previous_count, new_count, presentation_version_before/after) ⚠️ Requires approval token |
| duplicate_slide() | index: int | Dict[str, Any] (new_slide_index, total_slides, presentation_version_before/after) |
| reorder_slides() | from_index: int, to_index: int | Dict[str, Any] (from_index, to_index, total_slides, presentation_version_before/after) |
| set_slide_layout() | slide_index: int, layout_name: str | None |

### Content Creation
| Method | Args | Notes |
|--------|------|-------|
| add_text_box() | slide_index, text, position, size, font_name=None, font_size=18, bold=False, italic=False, color=None, alignment="left" | See Data Structures |
| add_bullet_list() | slide_index, items: List[str], position, size, bullet_style="bullet", font_size=18, font_name=None | Styles: bullet, numbered, none |
| set_title() | slide_index, title: str, subtitle: str=None | Uses layout placeholders |
| insert_image() | slide_index, image_path, position, size=None, alt_text=None, compress=False | Handles auto size. alt_text for accessibility |
| add_shape() | slide_index, shape_type, position, size, fill_color=None, fill_opacity=1.0, line_color=None, line_opacity=1.0, line_width=1.0, text=None | Types: rectangle, arrow, etc. Opacity range: 0.0-1.0 |
| replace_image() | slide_index, old_image_name: str, new_image_path, compress=False | Replace by name or partial match |
| add_chart() | slide_index, chart_type, data: Dict, position, size, title=None | Data: {"categories":[], "series":[]} |
| add_table() | slide_index, rows, cols, position, size, data: List[List]=None, header_row=True | Data is 2D array. header_row for styling hint |

### Formatting & Editing
| Method | Args | Notes |
|--------|------|-------|
| format_text() | slide_index, shape_index, font_name=None, font_size=None, bold=None, italic=None, color=None | Update text formatting |
| format_shape() | slide_index, shape_index, fill_color=None, fill_opacity=None, line_color=None, line_opacity=None, line_width=None | Opacity range: 0.0-1.0 ⚠️ transparency parameter DEPRECATED - use fill_opacity instead |
| replace_text() | find: str, replace: str, match_case: bool=False | Global text replacement |
| remove_shape() | slide_index, shape_index | Remove shape from slide ⚠️ Requires approval token |
| set_z_order() | slide_index, shape_index, action | Actions: bring_to_front, send_to_back, bring_forward, send_backward ⚠️ Refresh indices after |
| add_connector() | slide_index, connector_type, start_shape_index, end_shape_index | Types: straight, elbow, curve |
| reposition_shape() | slide_index, shape_index, position=None, size=None | Move/resize shapes by absolute inches |
| set_shape_text() | slide_index, shape_index, text | Update text in existing shapes |
| crop_image() | slide_index, shape_index, crop_box: Dict | crop_box: {"left": %, "top": %, "right": %, "bottom": %} |
| set_image_properties() | slide_index, shape_index, alt_text=None | Set accessibility |

### Validation
| Method | Returns |
|--------|---------|
| check_accessibility() | Dict (WCAG issues) |
| validate_presentation() | Dict (Empty slides, missing assets) |

### Chart & Presentation Operations
| Method | Args | Notes |
|--------|------|-------|
| update_chart_data() | slide_index, chart_index, data: Dict | Update existing chart data |
| format_chart() | slide_index, chart_index, title=None, legend_position=None | Modify chart appearance |
| add_notes() | slide_index, text, mode="append" | Modes: append, prepend, overwrite (v3.1.0+) |
| extract_notes() | None | Returns Dict[int, str] of all notes by slide |
| set_footer() | slide_index=None, text=None, show_number=False, show_date=False | Configure slide footer |
| set_background() | slide_index=None, color=None, image_path=None | Set slide or presentation background |

## 7. Workflow Context

### The 5-Phase Workflow
Tools are designed to work within a structured 5-phase workflow. Each tool should document which phase(s) it belongs to:

| Phase | Purpose | Tool Examples | Key Requirements |
|-------|---------|---------------|-------------------|
| DISCOVER | Deep inspection and capability probing | ppt_capability_probe.py, ppt_get_info.py, ppt_get_slide_info.py | Timeout handling, fallback probes, comprehensive metadata |
| PLAN | Manifest creation and design decisions | ppt_create_from_structure.py, ppt_validate_manifest.py | Schema validation, design rationale documentation |
| CREATE | Actual content creation and modification | ppt_add_shape.py, ppt_add_slide.py, ppt_replace_text.py | Version tracking, approval token enforcement, index freshness |
| VALIDATE | Quality assurance and compliance checking | ppt_validate_presentation.py, ppt_check_accessibility.py | WCAG 2.1 compliance, structural validation, contrast checking |
| DELIVER | Production handoff and documentation | ppt_export_pdf.py, ppt_extract_notes.py, ppt_generate_manifest.py | Complete audit trails, rollback commands, delivery packages |

### Probe Resilience Pattern
**CRITICAL**: Discovery tools must be resilient to large files and timeouts. Implement the Timeout + Transient Slides + Graceful Degradation pattern.

**Core Pattern: Timeout + Transient Slides + Warnings**

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

**Layer 2: Transient Slides (Accurate Analysis)**
```python
def _add_transient_slide(prs, layout):
    """
    Safely add and remove a transient slide for deep analysis.
    Uses generator pattern to guarantee cleanup via finally block.
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
```

**Layer 3: Partial Results + Warnings (Graceful Degradation)**
```python
def probe_presentation(filepath: Path, timeout_seconds: int = 30):
    """
    Probe with graceful degradation.
    """
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

## 8. Troubleshooting Playbook

### Error: `jq: parse error: Invalid numeric literal`
**Meaning**: Your tool printed something that isn't JSON to stdout.  
**Likely Culprits**:
- `print("Processing...")`
- A library warning (e.g., `DeprecationWarning`)
- `logging` configured to write to stdout (default behavior in some setups)  
**Fix**: Apply the Hygiene Block (redirect stderr to devnull) and ensure you use `sys.stdout.write(json.dumps(...))` only once.

### Error: `TypeError: '<=' not supported between 'int' and 'dict'`
**Meaning**: You treated a Core v3.1.0 return value as an integer.  
**Context**: `agent.add_slide()` used to return `slide_index` (int). Now it returns `{'slide_index': 1, 'version': ...}` (dict).  
**Fix**: Change `idx = agent.add_slide(...)` to `res = agent.add_slide(...); idx = res['slide_index']`.

### Error: `SlideNotFoundError`
**Meaning**: You requested an index that doesn't exist.  
**Context**: `python-pptx` is 0-indexed. The Ghost Slide: If your script crashes mid-loop, you might have partially created slides.  
**Fix**: Use `ppt_get_info.py` to verify the current slide count before assuming index N exists.

### Error: `ImportError: cannot import name 'PP_PLACEHOLDER'`
**Meaning**: You tried to import a constant from `core` that doesn't exist or isn't exported.  
**Fix**: Import standard constants directly from `pptx.enum.shapes` or define fallback constants within your tool if strict dependency isolation is needed.

### Tool-Specific Error Examples
**Permission Error (Exit Code 4)**:
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

**Shape Index Error (Exit Code 1)**:
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

**Version Mismatch Error (Exit Code 1)**:
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

## 9. Testing Requirements

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
| Happy Path | Normal usage succeeds |
| Edge Cases | Boundary values (0, 1, max, empty) |
| Error Cases | Invalid inputs raise correct exceptions |
| Validation | Invalid ranges/formats rejected |
| Backward Compat | Deprecated features still work |
| CLI Integration | Tool produces valid JSON |
| Governance | Clone-before-edit enforced, tokens validated |
| Version Tracking | Presentation versions captured correctly |
| Index Freshness | Shape indices refreshed after structural changes |

### Test Pattern Example
```python
import pytest
from pathlib import Path

@pytest.fixture
def test_presentation(tmp_path):
    """
    Create a test presentation with blank slide.
    """
    pptx_path = tmp_path / "test.pptx"
    with PowerPointAgent() as agent:
        agent.create_new()
        agent.add_slide(layout_name="Blank")
        agent.save(pptx_path)
    return pptx_path

class TestAddShapeOpacity:
    """
    Tests for add_shape() opacity functionality.
    """
    
    def test_opacity_applied(self, test_presentation):
        """
        Test shape with valid opacity value.
        """
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
        """
        Test that destructive operations require approval tokens.
        """
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

## 9.5 E2E-Validated Troubleshooting Tips

These issues were discovered during end-to-end testing of the `powerpoint-skill` creating a 7-slide presentation:

| Symptom | Root Cause | Fix |
|---------|-----------|-----|
| `ppt_add_slide.py --title` fails | No `--title` arg on this tool | Use `ppt_set_title.py` separately |
| `ppt_add_shape.py --shape-type` fails | Arg is `--shape` not `--shape-type` | Use `--shape rectangle` |
| `ppt_set_footer.py --show-page-number` fails | Arg is `--show-number` | Use `--show-number` |
| `ppt_remove_shape.py` silently succeeds without token | Tool didn't pass token to core | Fixed: now requires `--approval-token` |
| Color validation error on shapes | `RGBColor` is tuple, not object with `.red` | Fixed: use index access `[0]`, `[1]`, `[2]` |
| PDF/Image export fails | LibreOffice not installed | Install `libreoffice-impress` (optional dep) |
| Shape overflow on 4:3 slides | Percentage positioning uses 16:9 defaults | Use absolute inches for 4:3 presentations |
| Table overflows right margin | Table width exceeds slide width | Use 80% width max, or absolute inches |
| Overlay covers text content | Overlay added at top z-order | Use `ppt_set_z_order.py --action send_to_back` |

## 10. Contribution Workflow

### Before Starting:
- Read this document — Understand the architecture
- Check existing tools — Don't duplicate functionality
- Review system prompt — Understand AI agent usage
- Set up environment:
  ```
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

**Governance & Safety**
- [ ] Clone-Before-Edit: Does the tool work on `/work/` directory files only?
- [ ] Approval Token: Are destructive operations protected by token validation?
- [ ] Version Tracking: Does the tool capture `presentation_version_before` and `presentation_version_after`?
- [ ] Index Freshness: Does the tool refresh shape indices after structural operations?
- [ ] Audit Trail: Does the tool log all operations with timestamps and versions?

**v3.1.0+ Features**
- [ ] Opacity Handling: Does the tool use `fill_opacity` instead of deprecated `transparency`?
- [ ] Z-Order Management: If using `set_z_order`, does the tool refresh indices afterward?
- [ ] Speaker Notes: For tools adding notes, do they support all modes (`append`, `prepend`, `overwrite`)?
- [ ] Schema Validation: Does the tool validate inputs against JSON schemas when applicable?

