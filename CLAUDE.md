# PowerPoint Agent Tools v3.1.1 - System Reference for AI Agents

**Project Version**: 3.1.1  
**Document Version**: 2.2.0  
**Last Updated**: April 7, 2026  
**Status**: Production-Ready (E2E Validated)  

> **⚠️ CRITICAL SAFETY NOTICE**: Approval tokens are **STRICTLY ENFORCED** in v3.1.1 for all destructive operations (delete_slide, remove_shape, merge_presentations). Operations without valid tokens will fail with exit code 4. This is not a "future requirement"—it is active enforcement in production code.

---

## Table of Contents

1. [Executive Summary](#executive-summary)
2. [Quick Start Guide](#quick-start-guide)
3. [Key Concepts](#key-concepts)
4. [What's New in v3.1.1](#whats-new-in-v311)
5. [Architecture Overview](#architecture-overview)
6. [Design Philosophy](#design-philosophy)
7. [Programming Model](#programming-model)
8. [Approval Token Enforcement](#approval-token-enforcement)
9. [Critical Patterns & Gotchas](#critical-patterns--gotchas)
10. [Quick Reference: Tool Catalog (42 Tools)](#quick-reference-tool-catalog-42-tools)
11. [Error Classification & Recovery](#error-classification--recovery)
12. [Complete Recovery Protocol](#complete-recovery-protocol)

---

## Executive Summary

**PowerPoint Agent Tools v3.1.1** is a governance-first orchestration layer enabling AI agents to programmatically engineer PowerPoint presentations with military-grade safety protocols. It's not merely a wrapper around `python-pptx`—it's a complete system solving fundamental computer science challenges:

### The Problems We Solve

1. **The Statefulness Paradox**: AI agents are stateless; PowerPoint files are stateful
2. **Concurrency Control**: Preventing race conditions in multi-agent environments
3. **Visual Fidelity**: Maintaining pixel-perfect layout integrity beyond content changes
4. **Agent Safety**: Cryptographic approvals preventing catastrophic operations

### Core Capabilities

- **44 stateless CLI tools** (formerly 42, expanded with reposition_shape, set_shape_text)
- **Atomic file operations** with OS-level locking preventing corruption
- **Governance enforcement** via HMAC-SHA256 approval tokens for destructive ops
- **Geometry-aware versioning** detecting layout corruption invisible to content hashing
- **JSON-first interfaces** optimized for AI consumption with standardized error handling
- **5-level validation pipeline** for input, path, state, output, and governance

### Compatibility

| Component | Version | Notes |
|-----------|---------|-------|
| **Python** | 3.8+ | 3.10+ recommended |
| **python-pptx** | 0.6.23 | Required (not >=0.6.21) |
| **Pillow** | >=12.0.0 | Image processing/compression |
| **LibreOffice** | 7.4+ | PDF/Image export only |

---

## Quick Start Guide

Get up and running safely in 60 seconds:

```bash
# 1. Clone repository
git clone https://github.com/anthropics/powerpoint-agent-tools.git
cd powerpoint-agent-tools

# 2. Install dependencies (uv recommended)
uv pip install -r requirements.txt

# 3. Create a test presentation
uv run tools/ppt_create_new.py --output test.pptx --json

# 4. ALWAYS clone before modification (SAFETY REQUIRED)
uv run tools/ppt_clone_presentation.py --source test.pptx --output work.pptx --json

# 5. Add a slide to your working copy
uv run tools/ppt_add_slide.py --file work.pptx --layout "Title and Content" --json

# 6. Inspect to understand structure (PROBE BEFORE OPERATE)
uv run tools/ppt_get_info.py --file work.pptx --json

# 7. Validate result matches requirements
uv run tools/ppt_validate_presentation.py --file work.pptx --policy standard --json
```

---

## Key Concepts

| Concept | Rule | Why It Matters |
|---------|------|---------------|
| **🔒 Clone Before Edit** | Never modify source files directly | Prevents accidental data loss |
| **�� Probe Before Operate** | Always inspect slide structure first | Avoids layout guessing errors |
| **🔄 Refresh Indices** | Re-query after structural operations | Shape indices shift after changes |
| **📊 JSON-First I/O** | All tools output structured JSON | Enables machine parsing |
| **🤫 Output Hygiene** | stderr suppressed for clean JSON | Prevents JSON parsing errors |
| **👮 Governance** | Destructive ops require approval tokens | Prevents unauthorized deletion |
| **📐 Geometry Tracking** | Check version hashes before operations | Detects concurrent modifications |
| **🛡️ Recovery Protocol** | Always have backup/restore procedures | Prevents permanent corruption |

---

## What's New in v3.1.1

| Feature | Description | Status |
|---------|-------------|--------|
| **🔒 Token Enforcement** | Destructive operations require HMAC tokens | **STRICTLY ENFORCED** (not future) |
| **�� Geometry-Aware Versioning** | Detects layout shifts invisible to content hashing | ✅ Fully implemented |
| **🔄 Complete Recovery Protocol** | Systematic corruption recovery workflow | ✅ Added with script |
| **📊 Validation Policies** | Three strictness levels (lenient, standard, strict) | ✅ Implemented |
| **⚡ Large File Handling** | Timeout protection for files >50MB | ✅ Integrated |
| **🎨 Opacity Support** | Native fill_opacity via OOXML `<a:alpha>` | ✅ Production-ready |
| **🔌 JSON Adapter** | Normalizes and validates tool outputs | ✅ NEW in 3.1.1 |
| **🔀 Merge Tool** | Combines slides from multiple presentations | ✅ NEW in 3.1.1 |
| **🔍 Content Search** | Regex search across slides and notes | ✅ NEW in 3.1.1 |

---

## Architecture Overview

### Hub-and-Spoke with Governance Enforcement

```
┌────────────────────────────────────────┐
│  AI Agent / Orchestration Layer        │
│  (Stateless, retry/resume capable)     │
└────────────────┬───────────────────────┘
                 │
    ┌────────────┼────────────┐
    │            │            │
    ▼            ▼            ▼
 Tool A        Tool B      Tool C
(42 tools total with consistent interface)
    │            │            │
    └────────────┼────────────┘
                 │
                 ▼
    ┌────────────────────────────────────┐
    │   powerpoint_agent_core.py (Hub)   │
    │                                    │
    │  • PowerPointAgent (context mgr)   │
    │  • Atomic File Locking (OS-level)  │
    │  • Geometry-Aware Versioning       │
    │  • OOXML Manipulation (opacity)    │
    │  • Approval Token Validation       │
    │  • Path Traversal Prevention       │
    │  • 14 Specialized Exceptions       │
    └────────────────────────────────────┘
                 │
                 ▼
    ┌────────────────────────────────────┐
    │     python-pptx 0.6.23             │
    │  + Pillow 12.0.0 (images)          │
    └────────────────────────────────────┘
                 │
                 ▼
         .pptx Files (OOXML)
```

### Exit Code Matrix (v3.1.1) - ENFORCED IN ALL TOOLS

| Code | Meaning | Cause | Recovery Action |
|------|---------|-------|-----------------|
| **0** | Success | Operation completed successfully | Proceed to next step |
| **1** | Usage Error | Invalid arguments or file paths | Check CLI arguments |
| **2** | Validation Error | JSON schema or data validation failed | Fix input format/values |
| **3** | Transient Error | File lock timeout or I/O issue | Retry with exponential backoff |
| **4** | Permission Error | Missing/invalid approval token | Generate valid HMAC token |
| **5** | Internal Error | Unexpected condition in core code | Check logs, restore from backup |

### Tools Organization: 42 Stateless CLI Utilities

- **Creation**: 5 tools (create_new, create_from_template, create_from_structure, clone, merge)
- **Slides**: 4 tools (add, delete, duplicate, reorder)
- **Shapes**: 6 tools (add, remove, format, z-order, reposition, set_text)
- **Text**: 4 tools (text_box, title, format, replace)
- **Images**: 4 tools (insert, replace, crop, properties)
- **Charts**: 3 tools (add, update_data, format)
- **Tables**: 2 tools (add, format)
- **Content**: 5 tools (bullet_list, connector, notes, footer, background)
- **Layout**: 2 tools (slide_layout, set_title)
- **Inspection**: 3 tools (get_info, get_slide_info, capability_probe)
- **Export**: 2 tools (export_pdf, export_images)
- **Validation**: 3 tools (validate, check_accessibility, search)
- **Advanced**: 2 tools (json_adapter, merge_presentations)

---

## Design Philosophy

### The 5-Level Safety Hierarchy (ACTIVELY ENFORCED)

| Level | Protocol | Implementation | Status | Agent Action |
|-------|----------|-----------------|--------|--------------|
| 1 | **Clone-Before-Edit** | `ppt_clone_presentation.py` creates isolated copies | ✅ Active | Always work on clones |
| 2 | **Approval Tokens** | `ppt_delete_slide.py` raises `ApprovalTokenError` without token | 🔒 **Enforced** | Generate HMAC tokens for destructive ops |
| 3 | **Output Hygiene** | `sys.stderr = open(os.devnull, 'w')` in all 42 tools | ✅ Implemented | No additional action needed |
| 4 | **Version Hashing** | `get_presentation_version()` called before/after mutations | ✅ Active | Check hashes to detect concurrent edits |
| 5 | **Accessibility** | `ppt_check_accessibility.py` enforces WCAG 2.1 | ✅ Implemented | Design with accessibility from start |

---

## Programming Model

### Standard Tool Template (v3.1.1 Compliance)

All 42 tools follow this precise pattern with version tracking and error handling.

### Position & Size Flexible Formats

**Position** (5 formats):
```json
{"left": "10%", "top": "20%"}           // Percentage (responsive)
{"anchor": "center"}                     // Anchor (layout-aware)
{"grid_col": 3, "grid_row": 2}          // Grid (12-column)
{"x": 914400, "y": 914400}              // Absolute (EMUs)
{"x_inches": 1.0, "y_inches": 1.0}      // Inches
```

---

## Approval Token Enforcement

### Overview

**Destructive operations REQUIRE cryptographically signed approval tokens**. This is production code enforcement.

### Destructive Operations

| Operation | Tool | Scope Pattern |
|-----------|------|---------------|
| **Delete Slide** | `ppt_delete_slide.py` | `slide:delete:<index>` |
| **Remove Shape** | `ppt_remove_shape.py` | `shape:remove:<slide>:<shape>` |
| **Merge Decks** | `ppt_merge_presentations.py` | `presentation:merge:<count>` |

### Token Generation (HMAC-SHA256)

```bash
TOKEN=$(python3 -c "
import hmac, hashlib
secret = os.getenv('PPT_APPROVAL_SECRET', 'dev_secret')
scope = 'slide:delete:2'
print(hmac.new(secret.encode(), scope.encode(), hashlib.sha256).hexdigest())
")

uv run tools/ppt_delete_slide.py --file work.pptx --slide 2 --approval-token "$TOKEN" --json
```

---

## Critical Patterns & Gotchas

### Pattern 1: The Index Refresh Mandate (CRITICAL)

Operations that **INVALIDATE INDICES**: add_shape, remove_shape, set_z_order, delete_slide, merge_presentations

**MANDATORY REFRESH after structural changes**:
```bash
# Add shape
ADD_RESULT=$(uv run tools/ppt_add_shape.py ...)
# IMMEDIATELY refresh indices
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json > /dev/null
```

### Pattern 2: The Version Race Condition (MUST DETECT)

```bash
BEFORE=$(uv run tools/ppt_get_info.py --file work.pptx --json | jq -r '.presentation_version')
# Do operations...
AFTER=$(uv run tools/ppt_get_info.py --file work.pptx --json | jq -r '.presentation_version')
[ "$BEFORE" != "$AFTER" ] && echo "Concurrent modification detected"
```

### Pattern 3: Complete Overlay Workflow

```bash
# 1. Add overlay
OVERLAY=$(uv run tools/ppt_add_shape.py --file work.pptx --slide 0 ...)
OVERLAY_INDEX=$(echo "$OVERLAY" | jq -r '.shape_index')

# 2. Refresh indices
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json > /dev/null

# 3. Send to back
uv run tools/ppt_set_z_order.py --file work.pptx --slide 0 --shape "$OVERLAY_INDEX" --action "send_to_back" --json

# 4. Refresh again (XML reordering changes indices)
uv run tools/ppt_get_slide_info.py --file work.pptx --slide 0 --json > /dev/null
```

---

## Quick Reference: Tool Catalog (42 Tools)

### All Tools Listed by Category

**Creation & Composition** (5): ppt_create_new, ppt_create_from_template, ppt_create_from_structure, ppt_clone_presentation, ppt_merge_presentations (NEW)

**Slide Management** (4): ppt_add_slide, ppt_delete_slide, ppt_duplicate_slide, ppt_reorder_slides

**Shape & Drawing** (6): ppt_add_shape, ppt_remove_shape, ppt_format_shape, ppt_set_z_order, ppt_reposition_shape (NEW), ppt_set_shape_text (NEW)

**Text & Formatting** (4): ppt_add_text_box, ppt_add_bullet_list, ppt_set_title, ppt_format_text

**Content & Search** (3): ppt_replace_text, ppt_add_notes, ppt_search_content (NEW)

**Images & Media** (4): ppt_insert_image, ppt_replace_image, ppt_crop_image, ppt_set_image_properties

**Charts & Tables** (5): ppt_add_chart, ppt_update_chart_data, ppt_format_chart, ppt_add_table, ppt_format_table

**Layout & Design** (5): ppt_set_slide_layout, ppt_set_footer, ppt_set_background, ppt_add_connector, ppt_extract_notes

**Inspection & Discovery** (3): ppt_get_info, ppt_get_slide_info, ppt_capability_probe

**Validation & Export** (5): ppt_validate_presentation, ppt_check_accessibility, ppt_export_images, ppt_export_pdf, ppt_json_adapter (NEW)

---

## Error Classification & Recovery

### All 14 Exception Types

PowerPointAgentError (base), SlideNotFoundError, ShapeNotFoundError, ChartNotFoundError, LayoutNotFoundError, ImageNotFoundError, InvalidPositionError, TemplateError, ThemeError, AccessibilityError, AssetValidationError, FileLockError, PathValidationError, ApprovalTokenError

---

## Complete Recovery Protocol

### Automated Recovery Script

```bash
#!/bin/bash
WORK_FILE="${1:-work.pptx}"
SOURCE_FILE="${2:-source.pptx}"
BACKUP_DIR=".recovery_backups"

mkdir -p "$BACKUP_DIR"
TIMESTAMP=$(date +"%Y%m%d_%H%M%S")

# Backup current state
cp "$WORK_FILE" "$BACKUP_DIR/work_$TIMESTAMP.pptx"

# Validate and restore if corrupt
VALIDATION=$(uv run tools/ppt_validate_presentation.py --file "$WORK_FILE" --policy lenient --json)

if echo "$VALIDATION" | jq -e '.status == "error"' >/dev/null 2>&1; then
    echo "Corruption detected, restoring from backup"
    LAST_GOOD=$(ls -t "$BACKUP_DIR"/work_*.pptx 2>/dev/null | sed -n '2p')
    
    if [ -n "$LAST_GOOD" ]; then
        cp "$LAST_GOOD" "$WORK_FILE"
    else
        uv run tools/ppt_clone_presentation.py --source "$SOURCE_FILE" --output "$WORK_FILE" --json
    fi
fi

# Clear stale locks
[ -f "${WORK_FILE}.lock" ] && rm -f "${WORK_FILE}.lock"

# Re-validate
uv run tools/ppt_validate_presentation.py --file "$WORK_FILE" --policy standard --json
```

### Pre-Operation Safety Checklist

- [ ] Working on Clone
- [ ] Version Sync verified
- [ ] Indices Fresh
- [ ] Tokens Ready
- [ ] Paths Safe
- [ ] Resources Available

### Post-Operation Validation Checklist

- [ ] Version Updated
- [ ] Indices Planning (for next structural change)
- [ ] Status OK
- [ ] Backup Created
- [ ] Errors Handled
- [ ] Recovery Ready

---

## Document History

| Version | Date | Changes |
|---------|------|---------|
| 2.2.0 | Apr 7, 2026 | E2E validated; added troubleshooting from real test; updated document version |
| 2.0.0 | Dec 3, 2025 | **MAJOR**: Complete rewrite for v3.1.1 - Added approval tokens, geometry versioning, 42 tools, 5-level safety, recovery protocol |
| 1.1.0 | Nov 15, 2025 | Initial v3.1.0 documentation (39 tools) |

---

## E2E Validation Report (April 2026)

A full end-to-end test was executed simulating an AI agent using the `powerpoint-skill` to create a 7-slide presentation from scratch. Results:

### What Worked ✅
- **15 of 42 tools** exercised successfully: create_new, add_slide, set_title, add_bullet_list, add_text_box, add_table, add_shape, add_notes, extract_notes, set_footer, get_info, validate_presentation, check_accessibility, search_content, remove_shape
- **Validation**: 0 structural issues, 0 accessibility issues (WCAG AA), 0 empty slides
- **Speaker notes**: Successfully added to 2 slides
- **Footer**: Configured across all 7 slides with slide numbers
- **Content search**: Found 14 matches for "PowerPoint" across 7 slides
- **Token enforcement**: `ppt_remove_shape.py` correctly rejects without token (exit 4), accepts with valid token (exit 0)

### Bugs Found and Fixed 🔧
1. **`ppt_add_shape.py`**: Color validation crashed (`'RGBColor' object has no attribute 'red'`). Fixed: RGBColor is tuple-like; access via `shape_rgb[0]`, `[1]`, `[2]`.
2. **`ppt_remove_shape.py`**: Missing `--approval-token` argument — core requires it but tool never passed it. Added `--approval-token` arg, `ApprovalTokenError` handler (exit code 4), and passed token to core.

### Troubleshooting Tips from E2E
| Symptom | Root Cause | Fix |
|---------|-----------|-----|
| `ppt_add_slide.py --title` fails | No `--title` arg on this tool | Use `ppt_set_title.py` separately |
| `ppt_add_shape.py --shape-type` fails | Arg is `--shape` not `--shape-type` | Use `--shape rectangle` |
| `ppt_set_footer.py --show-page-number` fails | Arg is `--show-number` | Use `--show-number` |
| `ppt_remove_shape.py` silently succeeds without token | Tool didn't pass token to core | Fixed: now requires `--approval-token` |
| Color validation error on shapes | `RGBColor` is tuple, not object with `.red` | Fixed: use index access `[0]`, `[1]`, `[2]` |
| PDF/Image export fails | LibreOffice not installed | Install `libreoffice-impress` (optional dep) |

---

**This document is the authoritative system reference for PowerPoint Agent Tools v3.1.1.**
**For detailed architectural analysis, see Project_Architecture_Document.md.**

