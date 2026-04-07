# PowerPoint Agent Core: Programming Handbook (v3.1.0)

**Version:** 3.1.0  
**Library:** `core/powerpoint_agent_core.py`  
**License:** MIT  
**E2E Validated:** April 7, 2026 — Full 7-slide presentation created successfully via `powerpoint-skill`  

---

## 1. Introduction

The `PowerPointAgent` core library is the foundational engine for the PowerPoint Agent Tools ecosystem. It provides a **stateless, atomic, and security-hardened** interface for manipulating `.pptx` files. Unlike the raw `python-pptx` library, this core handles file locking, complex positioning logic, accessibility compliance, and operation auditing (versioning).

### 1.1 Key Capabilities
*   **Context-Safe**: Handles file opening/closing/locking automatically.
*   **Observability**: Tracks presentation state via deterministic SHA-256 hashing (Geometry + Content).
*   **Governance**: Enforces "Approval Tokens" for destructive actions (`delete_slide`, `remove_shape`).
*   **Visual Fidelity**: Implements XML hacks for features missing in `python-pptx` (Opacity, Z-Order).
*   **Accessibility**: Built-in WCAG 2.1 AA checking and Color Contrast calculation.

---

## 2. Usage Pattern (The "Hub" Model)

Tools interacting with this core **must** use the Context Manager pattern to ensure file safety.

```python
from core.powerpoint_agent_core import PowerPointAgent, FileLockError

try:
    # Atomic Operation Pattern
    with PowerPointAgent(filepath) as agent:
        # 1. Acquire Lock & Load (with timeout protection)
        agent.open(filepath, acquire_lock=True)
        
        # 2. Mutate (Capture return dict)
        result = agent.add_shape(...)
        
        # 3. Save (Atomic Write)
        agent.save()
        
        # 4. Release Lock (Automatic on exit)
except FileLockError:
    # Handle contention gracefully
    pass
```

### 2.1 Logging Configuration
The core uses Python's standard `logging` module. Configure it in your tool wrapper to capture debug details without polluting standard output (JSON).

```python
import logging
import sys

# Configure specific logger for the core library
logger = logging.getLogger('core.powerpoint_agent_core')
logger.setLevel(logging.INFO)

# CRITICAL: Avoid polluting stdout (JSON output stream) - send logs to STDERR
handler = logging.StreamHandler(sys.stderr)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)
```

---

## 3. Critical Protocols (Non-Negotiable)

### 3.1 Version Tracking Protocol
Every mutation method **must** capture the presentation state before and after execution to maintain the audit trail.

**Standard Pattern**:
```python
# 1. Capture version BEFORE changes
version_before = agent.get_presentation_version()

# 2. Perform operations
result = agent.some_operation()

# 3. Capture version AFTER changes
version_after = agent.get_presentation_version()

# 4. Return both versions in response
return {
    "status": "success",
    "presentation_version_before": version_before,
    "presentation_version_after": version_after,
    "result": result
}
```

### 3.2 Shape Index Freshness Protocol
Structural operations invalidate shape indices. Tools **must** refresh their knowledge of the slide state after any of these operations before performing further actions on shapes.

**Invalidating Operations**:
| Operation | Effect | Required Action |
|-----------|--------|-----------------|
| `add_shape()` | Adds index at end | Refresh if targeting new shape |
| `remove_shape()` | Shifts subsequent indices down | **CRITICAL**: Refresh immediately |
| `set_z_order()` | Reorders indices | **CRITICAL**: Refresh immediately |
| `delete_slide()` | Invalidates all indices | Reload slide info |
| `add_slide()` | New slide context | Query new slide info |

**Correct Pattern**:
```python
# 1. Perform structural change
agent.remove_shape(slide_index=0, shape_index=5)

# 2. REFRESH INDICES (Do not assume index 6 is now 5)
slide_info = agent.get_slide_info(slide_index=0)

# 3. Find target shape by name/properties in refreshed list
target_shape = next(s for s in slide_info["shapes"] if s["name"] == "TargetBox")
agent.format_shape(slide_index=0, shape_index=target_shape["index"], ...)
```

---

## 4. Data Structures & Input Schemas

The Core uses flexible dictionary-based inputs to abstract layout complexity.

### 4.1 Positioning (`Position.from_dict`)
Used in `add_shape`, `add_text_box`, `insert_image`, etc.

| Type | Schema | Description |
|------|--------|-------------|
| **Percentage** | `{"left": "10%", "top": "20%"}` | Relative to slide dimensions. **Preferred.** |
| **Absolute** | `{"left": 1.5, "top": 2.0}` | Inches from top-left. |
| **Anchor** | `{"anchor": "center", "offset_x": 0, "offset_y": 0}` | Relative to standard points. |
| **Grid** | `{"grid_row": 2, "grid_col": 3, "grid_size": 12}` | 12-column grid system layout. |

**Anchor Points**: `top_left`, `top_center`, `top_right`, `center_left`, `center`, `center_right`, `bottom_left`, `bottom_center`, `bottom_right`.

### 4.2 Sizing (`Size.from_dict`)
| Type | Schema | Description |
|------|--------|-------------|
| **Percentage** | `{"width": "50%", "height": "50%"}` | % of slide width/height. |
| **Absolute** | `{"width": 5.0, "height": 3.0}` | Inches. |
| **Auto** | `{"width": "50%", "height": "auto"}` | Preserves aspect ratio. |

---

## 5. API Reference

### 5.1 File Operations

#### `open(filepath, acquire_lock=True)`
*   **Purpose**: Loads presentation and optionally acquires file lock.
*   **Safety**: Uses `os.open` with `O_CREAT|O_EXCL` (via `errno.EEXIST`) for cross-platform atomic locking. Implements 10-second timeout.
*   **Throws**: `PathValidationError`, `FileLockError` (on timeout), `PowerPointAgentError`.

#### `save(filepath=None)`
*   **Purpose**: Saves changes. If `filepath` is None, overwrites source.
*   **Safety**: Ensures parent directories exist.

#### `clone_presentation(output_path)`
*   **Purpose**: Creates a working copy.
*   **Returns**: A *new* `PowerPointAgent` instance pointed at the cloned file.

### 5.2 Slide Operations

#### `add_slide(layout_name, index=None)`
*   **Returns**: `Dict` (v3.1+) containing `slide_index`, `layout_name`, `total_slides`, `presentation_version_before/after`.
*   **Validation**: Raises `SlideNotFoundError` if `index` is out of bounds (removed silent clamping).

#### `delete_slide(index, approval_token=None)`
*   **Security**: **Requires** valid `approval_token` matching scope `delete:slide`.
*   **Throws**: `ApprovalTokenError` if token is invalid/missing.

#### `duplicate_slide(index)` / `reorder_slides(from_index, to_index)`
*   **Behavior**: Performs deep copy of shapes including text runs and styles.

### 5.3 Shape & Visual Operations

#### `add_shape(slide_index, shape_type, position, size, ...)`
*   **Arguments**:
    *   `fill_opacity` (float 0.0-1.0): **New in v3.1.0**. (1.0 = Opaque).
    *   `line_opacity` (float 0.0-1.0).
    *   `shape_type`: String key (e.g., `"rectangle"`, `"arrow_right"`).
*   **Returns**: Dictionary containing `shape_index` and applied styling.
*   **Internal**: Uses `_set_fill_opacity` to inject OOXML `<a:alpha>` tags.

#### `format_shape(slide_index, shape_index, ...)`
*   **Arguments**: `fill_color`, `fill_opacity`, `line_color`, etc.
*   **Deprecation**: `transparency` param is deprecated; explicit conversion to `1.0 - fill_opacity` occurs, logging a warning.

#### `remove_shape(slide_index, shape_index, approval_token=None)`
*   **Security**: **Requires** valid `approval_token` matching scope `remove:shape`.
*   **Warning**: Removing a shape shifts the indices of all subsequent shapes on that slide.

#### `set_z_order(slide_index, shape_index, action)`
*   **Actions**: `"bring_to_front"`, `"send_to_back"`, `"bring_forward"`, `"send_backward"`.
*   **Internal**: Physically moves the XML element in `<p:spTree>`.
*   **Critical Side Effect**: **Invalidates Shape Indices**. Tools must warn users to re-query `get_slide_info`.

#### `reposition_shape(slide_index, shape_index, position=None, size=None)`
*   **Purpose**: Move and/or resize an existing shape.
*   **Arguments**: `position` dict with `left`/`top` (inches), `size` dict with `width`/`height` (inches).
*   **Returns**: Dict with before/after dimensions.

#### `set_shape_text(slide_index, shape_index, text)`
*   **Purpose**: Update text content of an existing shape or text box.
*   **Arguments**: `text` string (supports `\n` for line breaks).
*   **Returns**: Dict with text preview and length metrics.

### 5.4 Text & Content

#### `add_text_box` / `add_bullet_list`
*   **Features**: Auto-fit text, specific font styling, alignment mapping.
*   **Returns**: `shape_index` of the created text container.

#### `replace_text(find, replace, match_case=False)`
*   **Scope**: Global (entire presentation) or scoped (slide/shape).
*   **Intelligence**: Tries to preserve formatting by replacing inside text runs first.

#### `add_notes(slide_index, text, mode="append")`
*   **Purpose**: Add speaker notes for accessibility/presenting.
*   **Modes**: `append` (default), `prepend`, `overwrite`.
*   **Returns**: Dictionary with text preview and length metrics.

#### `set_footer(text, show_number, show_date)`
*   **Mechanism**: Iterates through *all* slides to find placeholders (type 7, 6, 5).
*   **Returns**: `slides_processed` count. (Note: Does not create text boxes; relies on Tool layer for fallback).

### 5.5 Charts & Images

#### `add_chart(chart_type, data, ...)`
*   **Supported Types**: Column, Bar, Line, Pie, Area, Scatter, Doughnut.
*   **Data Format**: `{"categories": ["A", "B"], "series": [{"name": "S1", "values": [1, 2]}]}`.

#### `update_chart_data(slide_index, chart_index, data)`
*   **Strategy**:
    1.  Try `chart.replace_data()` (Best, preserves format).
    2.  Catch `AttributeError` (Older pptx versions).
    3.  Fallback: Recreate chart in-place (Preserves position/size/title, resets some style).

#### `insert_image` / `replace_image`
*   **Features**: Auto-aspect ratio calculation, optional compression (if Pillow installed), Alt Text setting.

---

## 6. Advanced Patterns

### 6.1 Transient Slide Pattern (Advanced Probe)
For accurate layout geometry analysis without corrupting the file, use this pattern:

```python
def analyze_layout_safe(prs, layout):
    slide = None
    added_index = -1
    try:
        # Create temporary slide
        slide = prs.slides.add_slide(layout)
        added_index = len(prs.slides) - 1
        
        # Analyze instantiated slide geometry
        return extract_metrics(slide)
    finally:
        # ALWAYS cleanup (even on failure)
        if added_index != -1:
            rId = prs.slides._sldIdLst[added_index].rId
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[added_index]
```
**Rules**: Never `save()` while a transient slide exists.

### 6.2 Production-Grade Probe Resilience
For production probes, implement this 3-layer pattern:

1.  **Timeout Protection** (30s default)
    *   Check elapsed time at each layout iteration
    *   Return partial results on timeout

2.  **Transient Slide Analysis** (as shown above)
    *   Create temporary slide for accurate geometry
    *   ALWAYS cleanup in finally block

3.  **Graceful Degradation**
    *   Limit layouts to 50 maximum
    *   Return `analysis_complete` flag with results
    *   Include warnings/info arrays in response

See `ppt_capability_probe.py` for complete implementation.

---

## 7. Security & Governance

### 7.1 Approval Token Generation
Tokens must be generated by a trusted service using HMAC-SHA256.

```python
import hmac, hashlib, base64, json, time

def generate_approval_token(scope: str, user: str, secret_key: bytes) -> str:
    """Generate HMAC-based approval token for development."""
    payload = {
        "scope": scope,
        "user": user,
        "issued": time.time(),
        "expiry": time.time() + 3600,  # 1 hour
        "single_use": True
    }
    # Serialize and Encode
    json_payload = json.dumps(payload)
    b64_payload = base64.urlsafe_b64encode(json_payload.encode()).decode()
    
    # Sign
    signature = hmac.new(secret_key, b64_payload.encode(), hashlib.sha256).hexdigest()
    
    # Combine
    return f"HMAC-SHA256:{b64_payload}.{signature}"
```

### 7.2 Path Validation
*   **Traversal Protection**: If `allowed_base_dirs` is set, checks `path.is_relative_to(base)`.
*   **Extension Check**: Enforces `.pptx`, `.pptm`, `.potx`.

---

## 8. Observability & Versioning

### 8.1 Presentation Versioning (`get_presentation_version`)
Returns a SHA-256 hash (prefix 16 chars) representing the state.

**Input for Hash**:
1.  Slide Count.
2.  Layout Names per slide.
3.  **Shape Geometry**: `{left}:{top}:{width}:{height}` (Detects moves/resizes).
4.  **Text Content**: SHA-256 of text runs.

### 8.2 Integration with AI Orchestration Layer
The version tracking protocol enables robust multi-agent workflows:
*   **State Verification**: Ensures no intervening edits occurred between "Plan" and "Act" phases.
*   **Conflict Detection**: If `expected_version != current_version`, the agent must abort and re-probe.
*   **Rollback**: Version hashes provide checkpoints for restoring previous states via backup files.

### 8.3 Error Handling Matrix

| Code | Category | Meaning | Response Format |
|---|---|---|---|
| 0 | Success | Completed | `{"status": "success", ...}` |
| 1 | Usage | Invalid args | `{"status": "error", "error_type": "ValueError", ...}` |
| 2 | Validation | Schema invalid | `{"status": "error", "error_type": "ValidationError", ...}` |
| 3 | Transient | Lock/Network | `{"status": "error", "error_type": "FileLockError", ...}` |
| 4 | Permission | Token missing | `{"status": "error", "error_type": "ApprovalTokenError", ...}` |
| 5 | Internal | Crash | `{"status": "error", "error_type": "PowerPointAgentError", ...}` |

### 8.4 Concrete Error Response Examples

**Permission Error (Exit Code 4):**
```json
{
  "status": "error",
  "error": "Approval token required for slide deletion",
  "error_type": "ApprovalTokenError",
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

---

## 9. Performance Characteristics

Understanding the cost of operations is vital for building efficient agents.

| Operation | Complexity | 10-Slide Deck | 50-Slide Deck | Notes |
|-----------|------------|---------------|---------------|-------|
| `get_presentation_version()` | O(N) Shapes | ~15ms | ~75ms | Scales linearly with total shape count. Called twice per mutation. |
| `capability_probe(deep=True)` | O(M) Layouts | ~120ms | ~600ms+ | Creates/destroys slides. Has 30s timeout. |
| `add_shape()` | O(1) | ~8ms | ~8ms | Constant time (XML injection). |
| `replace_text(global)` | O(N) TextRuns | ~25ms | ~125ms | Regex matching across all text runs. |
| `save()` | I/O Bound | ~50ms | ~200ms+ | Dominated by disk write speed and file size (images). |

**Optimization Guidelines**:
*   **Batching**: Not supported natively (stateless tools), but context managers in custom scripts can batch mutations before a single `save()`.
*   **Shallow Probes**: Use `deep=False` in `capability_probe` unless layout geometry is strictly required.
*   **Limits**: Avoid decks >100 slides or >50MB for interactive agent sessions to prevent timeouts.

---

## 10. Testing Strategies

### 10.1 Basic Test Pattern
Use `pytest` fixtures to create clean environments for testing tool logic.

```python
import pytest
from core.powerpoint_agent_core import PowerPointAgent

@pytest.fixture
def test_agent(tmp_path):
    """Create agent with blank slide."""
    pptx = tmp_path / "test.pptx"
    agent = PowerPointAgent()
    agent.create_new()
    agent.add_slide("Blank")
    agent.save(pptx)
    agent.close()
    return pptx

def test_version_changes_after_shape_removal(test_agent):
    with PowerPointAgent(test_agent) as agent:
        agent.open(test_agent)
        # Add shape to remove
        agent.add_shape(0, "rectangle", {"left":0, "top":0}, {"width":1, "height":1})
        agent.save()
        
        # Capture State 1
        version_before = agent.get_presentation_version()
        
        # Perform destructive operation
        # Note: Use valid token from Section 7.1 logic
        agent.remove_shape(
            slide_index=0, 
            shape_index=0,
            approval_token="HMAC-SHA256:..." 
        )
        
        # Capture State 2
        version_after = agent.get_presentation_version()
        
        assert version_before != version_after, "Version must change after modification"
```

---

## 11. Internal "Magic" (Troubleshooting)

### 11.1 Opacity Injection
`python-pptx` lacks transparency support. We use `lxml` to inject:
```xml
<a:solidFill>
  <a:srgbClr val="FF0000">
    <a:alpha val="50000"/> <!-- 50% Opacity -->
  </a:srgbClr>
</a:solidFill>
```
**Note**: Office uses 0-100,000 scale. Core converts 0.0-1.0 floats automatically.

### 11.2 Z-Order Manipulation
We physically move the `<p:sp>` element within the `<p:spTree>` XML list.
*   `bring_to_front`: Move to end of list.
*   `send_to_back`: Move to index 2 (after background/master refs).

### 11.3 Debugging OOXML
When visual features fail, inspect the underlying XML:
1.  Export the shape's element: `print(lxml.etree.tostring(shape.element, pretty_print=True))`
2.  Verify namespaces: Ensure `a:` corresponds to `http://schemas.openxmlformats.org/drawingml/2006/main`.
3.  Check for missing parents: Opacity requires `<a:solidFill>` to exist; if the shape has no fill, injection fails.

---

## 12. Workflow Integration Patterns

The core library supports the 5-phase workflow through:
*   **DISCOVER**: `get_presentation_info()`, `get_slide_info()`, `get_capabilities()` (via probe)
*   **PLAN**: Version tracking for manifest creation.
*   **CREATE**: All mutation methods with approval tokens.
*   **VALIDATE**: `validate_presentation()`, `check_accessibility()`.
*   **DELIVER**: `export_to_pdf()`, `extract_notes()`.

**Phase-specific requirements**:
*   **DISCOVER** tools must implement 30s timeout handling.
*   **CREATE** tools must track version changes.
*   **VALIDATE** tools must categorize issues by severity.

---

## 13. Backward Compatibility Policy

**v3.1.0 → v3.0.0 Compatibility**:
*   ✅ **Shape indices**: Methods now return dicts but preserve `shape_index` key.
*   ✅ **Versioning**: New `presentation_version_before/after` keys are additive.
*   ✅ **Transparency**: Deprecated but supported with conversion warnings.
*   ❌ **Silent index clamping**: Removed (intentional breaking change for safety).

**Migration Path**:
```python
# ❌ v3.0 pattern (NO LONGER WORKS in v3.1.0+)
idx = agent.add_slide()  # Used to return int — now returns Dict

# ✅ v3.1+ pattern (required)
result = agent.add_slide()  # Returns Dict[str, Any]
idx = result["slide_index"]
```

**Breaking Changes Handling**:
*   Use `try/except` for `SlideNotFoundError` instead of assuming index clamping.
*   Process `ApprovalTokenError` for destructive operations.
*   Always refresh shape indices after structural operations.

---

## 14. Troubleshooting Scenarios

### Scenario: "Shape index 10 out of range (0-8)"
**Diagnosis**: Shape indices shifted after a structural operation (`remove_shape` or `set_z_order`).
**Solution**:
1.  Identify the last structural operation.
2.  Call `get_slide_info()` immediately after it to get fresh indices.
3.  Target the shape by `name` or content instead of hardcoded index.

### Scenario: "File lock timeout after 10s"
**Diagnosis**: Another process (or a crashed previous run) holds the lock.
**Solution**:
1.  Check for `.pptx.lock` files and stale processes.
2.  Implement retry with exponential backoff in the tool wrapper.

### Scenario: "Chart formatting lost after update"
**Diagnosis**: The "Fallback Recreation" path was triggered because `replace_data` failed.
**Solution**:
1.  Verify `python-pptx` version.
2.  Accept that complex custom formatting may need manual re-application via `format_chart`.

---

## Appendix: Quick API Reference

| Category | Method | Critical Arguments | Destructive? |
|----------|--------|-------------------|--------------|
| **File** | `open()` | `filepath`, `acquire_lock` | No |
| **Slide** | `delete_slide()` | `index`, `approval_token` | **Yes** |
| **Slide** | `add_slide()` | `layout_name` | No |
| **Shape** | `add_shape()` | `shape_type`, `fill_opacity` | No |
| **Shape** | `set_z_order()` | `action` | No* |
| **Text** | `replace_text()` | `find`, `replace` | No |
| **Text** | `add_notes()` | `text`, `mode` | No |
| **Chart** | `update_chart_data()` | `data` | No |

*\*Requires immediate index refresh*
