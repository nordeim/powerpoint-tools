# 🎨 PowerPoint Agent Tools

<div align="center">

![PowerPoint Agent Tools](https://img.shields.io/badge/PowerPoint-Agent_Tools-B7472A?style=for-the-badge&logo=microsoft-powerpoint&logoColor=white)
[![Python Version](https://img.shields.io/badge/python-3.8+-blue.svg?style=for-the-badge&logo=python&logoColor=white)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)
[![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen.svg?style=for-the-badge)](http://makeapullrequest.com)

**Production-grade PowerPoint manipulation for AI agents**

Build, edit, validate, and export PowerPoint presentations programmatically through simple CLI tools designed for AI consumption.

[Features](#-features) • [Quick Start](#-quick-start) • [Tool Catalog](#-tool-catalog) • [Documentation](#-tool-documentation) • [Examples](#-examples)

</div>

---

## 🎯 Why PowerPoint Agent Tools?

Traditional PowerPoint libraries require complex Python code and offer limited AI-agent-friendly APIs. **PowerPoint Agent Tools** provides a comprehensive suite of **44 stateless CLI tools** that bridge the gap between LLMs and `.pptx` files.

✅ **CLI-First Design** - AI agents call simple commands, no Python knowledge required  
✅ **JSON Everywhere** - All outputs machine-parsable for easy integration  
✅ **Flexible Positioning** - 5 positioning systems (%, anchor, grid, Excel-like, absolute)  
✅ **Structure-Driven** - Generate entire decks from a single JSON definition  
✅ **Validation & Safety** - Built-in accessibility checks (WCAG) and asset validation  
✅ **Visual Design** - Shapes, connectors, image manipulation, and formatting  
✅ **Export Capabilities** - Convert slides to PDF or High-Res Images  
✅ **Introspection** - Inspect slide content, shapes, and layouts before editing  

---

## 🚀 Quick Start

Create a professional presentation in 60 seconds:

```bash
# Install
pip install python-pptx Pillow

# 1. Create presentation
uv run tools/ppt_create_new.py --output pitch.pptx --json

# 2. Set title
uv run tools/ppt_set_title.py --file pitch.pptx --slide 0 --title "AI Revolution" --subtitle "Q4 Strategy" --json

# 3. Add new slide
uv run tools/ppt_add_slide.py --file pitch.pptx --layout "Title and Content" --json

# 4. Add chart
uv run tools/ppt_add_chart.py \
  --file pitch.pptx \
  --slide 1 \
  --chart-type column \
  --data-string '{"categories":["Q1","Q2"],"series":[{"name":"Growth","values":[10,50]}]}' \
  --position '{"left":"10%","top":"25%"}' \
  --size '{"width":"80%","height":"60%"}' \
  --json

# 5. Validate accessibility
uv run tools/ppt_check_accessibility.py --file pitch.pptx --json
```

**Result:** A valid, accessible presentation created entirely via CLI.

---

## 📦 Installation

### Requirements
- **Python:** 3.8 or higher
- **Dependencies:** `python-pptx`, `Pillow`
- **Optional:** `LibreOffice` (required for PDF/Image export)

### Install via pip (Recommended)

```bash
# Clone repository
git clone https://github.com/nordeim/powerpoint-agent-tools.git
cd powerpoint-agent-tools

# Install dependencies
pip install -r requirements.txt
```

---

## 🛠️ Tool Catalog

The suite consists of **42 tools** organized by capability.

### 1. Creation & Architecture
| Tool | Purpose |
|------|---------|
| `ppt_create_new.py` | Create blank presentation |
| `ppt_create_from_template.py` | Create from corporate .pptx template |
| `ppt_create_from_structure.py` | Generate full deck from JSON structure |
| `ppt_clone_presentation.py` | Create exact copy (backup/safe-mode) |

### 2. Slide Management
| Tool | Purpose |
|------|---------|
| `ppt_add_slide.py` | Add slide with specific layout |
| `ppt_delete_slide.py` | Remove slide by index ⚠️ requires token |
| `ppt_duplicate_slide.py` | Clone existing slide |
| `ppt_reorder_slides.py` | Move slide to new position |
| `ppt_set_slide_layout.py` | Change layout of existing slide |
| `ppt_set_title.py` | Set title/subtitle text |

### 3. Text & Content
| Tool | Purpose |
|------|---------|
| `ppt_add_text_box.py` | Add text with formatting |
| `ppt_add_bullet_list.py` | Add bullet/numbered lists |
| `ppt_format_text.py` | Style existing text (font, color, bold) |
| `ppt_replace_text.py` | Global find/replace |
| `ppt_add_notes.py` | Add speaker notes |

### 4. Images & Media
| Tool | Purpose |
|------|---------|
| `ppt_insert_image.py` | Insert image with auto-aspect ratio |
| `ppt_replace_image.py` | Swap image (e.g., update logo) |
| `ppt_crop_image.py` | Crop image |
| `ppt_set_image_properties.py` | Set Alt Text (accessibility) |

### 5. Visual Design
| Tool | Purpose |
|------|---------|
| `ppt_add_shape.py` | Add rectangles, circles, arrows (with opacity) |
| `ppt_format_shape.py` | Update fill/border colors and opacity |
| `ppt_remove_shape.py` | Remove shape ⚠️ requires token |
| `ppt_add_connector.py` | Draw lines between shapes |
| `ppt_reposition_shape.py` | Move and/or resize shapes |
| `ppt_set_shape_text.py` | Update text in existing shapes |
| `ppt_set_z_order.py` | Bring to front / send to back |
| `ppt_set_background.py` | Set slide background color or image |

### 6. Data Visualization
| Tool | Purpose |
|------|---------|
| `ppt_add_chart.py` | Add Column, Line, Pie charts |
| `ppt_update_chart_data.py` | Update chart data |
| `ppt_format_chart.py` | Update chart title/legend |
| `ppt_add_table.py` | Add data table |
| `ppt_format_table.py` | Style table headers and cells |

### 7. Inspection & Analysis
| Tool | Purpose |
|------|---------|
| `ppt_get_info.py` | Get file metadata & layout list |
| `ppt_get_slide_info.py` | Inspect slide content (shapes/text) |
| `ppt_capability_probe.py` | Deep probe: layouts, theme, fonts, capabilities |
| `ppt_extract_notes.py` | Extract speaker notes |
| `ppt_search_content.py` | Search text across slides and notes |

### 8. Validation & Export
| Tool | Purpose |
|------|---------|
| `ppt_validate_presentation.py` | Check assets & structure (lenient/standard/strict) |
| `ppt_check_accessibility.py` | Audit WCAG 2.1 compliance |
| `ppt_export_pdf.py` | Convert deck to PDF (requires LibreOffice) |
| `ppt_export_images.py` | Convert slides to PNG/JPG (requires LibreOffice) |
| `ppt_json_adapter.py` | Validate and normalize tool JSON output |

### 9. Advanced
| Tool | Purpose |
|------|---------|
| `ppt_merge_presentations.py` | Combine slides from multiple decks ⚠️ requires token |
| `ppt_set_footer.py` | Configure footer text and slide numbers |

---

## 📖 Tool Documentation

All tools accept `--json` for structured output. Paths can be absolute or relative.

### 🏗️ Creation & Architecture

#### `ppt_create_from_structure.py`
Generate a complete presentation in one pass using a JSON definition file.
```bash
uv run tools/ppt_create_from_structure.py --structure deck_spec.json --output output.pptx --json
```

#### `ppt_create_from_template.py`
Create a new deck based on a corporate template.
```bash
uv run tools/ppt_create_from_template.py --template corp_master.pptx --output draft.pptx --slides 5 --json
```

#### `ppt_create_new.py`
Create a blank presentation.
```bash
uv run tools/ppt_create_new.py --output new.pptx --layout "Title Slide" --json
```

#### `ppt_clone_presentation.py`
Clone a presentation (useful for "Save As" workflows).
```bash
uv run tools/ppt_clone_presentation.py --source base.pptx --output v2.pptx --json
```

### 🎞️ Slide Management

#### `ppt_add_slide.py`
Add a slide with a specific layout.
```bash
uv run tools/ppt_add_slide.py --file deck.pptx --layout "Title and Content" --title "Agenda" --json
```

#### `ppt_delete_slide.py`
Delete a slide by index (0-based).
```bash
uv run tools/ppt_delete_slide.py --file deck.pptx --index 2 --json
```

#### `ppt_duplicate_slide.py`
Clone a slide to the end of the deck.
```bash
uv run tools/ppt_duplicate_slide.py --file deck.pptx --index 0 --json
```

#### `ppt_reorder_slides.py`
Move a slide from one position to another.
```bash
uv run tools/ppt_reorder_slides.py --file deck.pptx --from-index 4 --to-index 1 --json
```

#### `ppt_set_slide_layout.py`
Change the layout of an existing slide.
```bash
uv run tools/ppt_set_slide_layout.py --file deck.pptx --slide 0 --layout "Title Only" --json
```

### 📝 Text & Content

#### `ppt_add_text_box.py`
Add text with flexible positioning.
```bash
uv run tools/ppt_add_text_box.py --file deck.pptx --slide 0 --text "Draft" --position '{"top":"10%","left":"80%"}' --size '{"width":"10%","height":"5%"}' --json
```

#### `ppt_add_bullet_list.py`
Add formatted lists.
```bash
uv run tools/ppt_add_bullet_list.py --file deck.pptx --slide 1 --items "Point A,Point B" --position '{"grid":"C4"}' --size '{"width":"50%","height":"50%"}' --json
```

#### `ppt_format_text.py`
Format text in a specific shape. Use `ppt_get_slide_info.py` to find the `shape` index.
```bash
uv run tools/ppt_format_text.py --file deck.pptx --slide 0 --shape 1 --color "#FF0000" --bold --json
```

#### `ppt_replace_text.py`
Global find and replace.
```bash
uv run tools/ppt_replace_text.py --file deck.pptx --find "2023" --replace "2024" --json
```

### 🖼️ Images & Media

#### `ppt_insert_image.py`
Insert an image.
```bash
uv run tools/ppt_insert_image.py --file deck.pptx --slide 0 --image logo.png --position '{"anchor":"top_right"}' --size '{"width":"15%","height":"auto"}' --alt-text "Logo" --json
```

#### `ppt_replace_image.py`
Replace an image by its current name (useful for logo updates).
```bash
uv run tools/ppt_replace_image.py --file deck.pptx --slide 0 --old-image "Picture 1" --new-image new_logo.png --json
```

#### `ppt_set_image_properties.py`
Set Alt Text (accessibility) or transparency.
```bash
uv run tools/ppt_set_image_properties.py --file deck.pptx --slide 0 --shape 2 --alt-text "Detailed Description" --json
```

### 🎨 Visual Design

#### `ppt_add_shape.py`
Add shapes like rectangles, arrows, or stars.
```bash
uv run tools/ppt_add_shape.py --file deck.pptx --slide 0 --shape arrow_right --position '{"left":1.0,"top":1.0}' --size '{"width":2.0,"height":1.0}' --fill-color "#0000FF" --json
```

#### `ppt_format_shape.py`
Update shape colors and borders.
```bash
uv run tools/ppt_format_shape.py --file deck.pptx --slide 0 --shape 1 --fill-color "#00FF00" --line-width 3 --json
```

#### `ppt_add_connector.py`
Draw a line connecting two shapes.
```bash
uv run tools/ppt_add_connector.py --file deck.pptx --slide 0 --from-shape 0 --to-shape 1 --json
```

### 📊 Data Visualization

#### `ppt_add_chart.py`
Add a chart from JSON data.
```bash
uv run tools/ppt_add_chart.py --file deck.pptx --slide 1 --chart-type pie --data data.json --position '{"left":"10%","top":"10%"}' --size '{"width":"80%","height":"80%"}' --json
```

#### `ppt_format_chart.py`
Update chart title or legend position.
```bash
uv run tools/ppt_format_chart.py --file deck.pptx --slide 1 --chart 0 --title "New Data" --legend bottom --json
```

#### `ppt_add_table.py`
Add a data table.
```bash
uv run tools/ppt_add_table.py --file deck.pptx --slide 2 --rows 3 --cols 3 --data table.json --position '{"grid":"C3"}' --size '{"width":"50%","height":"50%"}' --json
```

### 🔍 Inspection & Analysis

#### `ppt_get_info.py`
Get presentation metadata (size, slide count, layout names).
```bash
uv run tools/ppt_get_info.py --file deck.pptx --json
```

#### `ppt_get_slide_info.py`
Inspect a slide to find shape indices and content. **Critical for editing existing slides.**
```bash
uv run tools/ppt_get_slide_info.py --file deck.pptx --slide 0 --json
```

#### `ppt_extract_notes.py`
Extract speaker notes to JSON.
```bash
uv run tools/ppt_extract_notes.py --file deck.pptx --json
```

### 🛡️ Validation & Export

#### `ppt_check_accessibility.py`
Run WCAG 2.1 checks (Contrast, Alt Text).
```bash
uv run tools/ppt_check_accessibility.py --file deck.pptx --json
```

#### `ppt_validate_presentation.py`
General health check (missing assets, empty slides).
```bash
uv run tools/ppt_validate_presentation.py --file deck.pptx --json
```

#### `ppt_export_pdf.py`
Convert to PDF (Requires LibreOffice).
```bash
uv run tools/ppt_export_pdf.py --file deck.pptx --output deck.pdf --json
```

#### `ppt_export_images.py`
Export slides as images.
```bash
uv run tools/ppt_export_images.py --file deck.pptx --output-dir slides/ --format png --json
```

---

## 🎯 Positioning Systems

PowerPoint Agent Tools offers **5 flexible positioning systems**.

| System | Format | Best For |
|--------|--------|----------|
| **Percentage** | `{"left":"10%", "top":"20%"}` | **Responsive Layouts (Recommended)** |
| **Anchor** | `{"anchor":"bottom_right"}` | Headers/Footers/Logos |
| **Grid** | `{"grid_row":2, "grid_col":2}` | Structured Grids (12x12) |
| **Excel-Ref** | `{"grid":"C4"}` | Excel Users |
| **Absolute** | `{"left":1.5, "top":2.0}` | Exact Design Specs (Inches) |

---

## 🏗️ Architecture

```
powerpoint-agent-tools/
├── core/
│   ├── __init__.py
│   ├── powerpoint_agent_core.py      # Core library (4,437 lines)
│   └── strict_validator.py           # JSON schema validation
├── tools/                              # 42 CLI tools
│   ├── ppt_create_*.py                 # Creation
│   ├── ppt_add_*.py                    # Content Addition
│   ├── ppt_format_*.py                 # Styling
│   ├── ppt_export_*.py                 # Export
│   └── ...
├── skills/
│   └── powerpoint-skill/               # AI agent skill (SKILL.md)
├── schemas/                            # JSON schemas for validation
├── CLAUDE.md                           # AI Agent System Reference
├── requirements.txt
└── README.md
```

---

## 🔒 Approval Tokens (Destructive Operations)

Operations that delete or merge require an approval token:

```bash
# Generate token
TOKEN=$(python3 -c "
import hmac, hashlib, os
secret = os.getenv('PPT_APPROVAL_SECRET', 'dev_secret')
print(hmac.new(secret.encode(), b'slide:delete:2', hashlib.sha256).hexdigest())
")

# Use with destructive operations
uv run tools/ppt_delete_slide.py --file deck.pptx --slide 2 --approval-token "$TOKEN" --json
uv run tools/ppt_remove_shape.py --file deck.pptx --slide 0 --shape 3 --approval-token "$TOKEN" --json
uv run tools/ppt_merge_presentations.py --sources '[...]' --output merged.pptx --approval-token "$TOKEN" --json
```

**Scope patterns**: `slide:delete:<index>`, `shape:remove:<slide>:<shape>`, `merge:presentations:<count>`

---

## 💡 Troubleshooting

| Issue | Solution |
|-------|----------|
| `jq: parse error: Invalid numeric literal` | Tool printed non-JSON to stdout. Ensure `--json` flag is used. |
| `Shape index X out of range` | Shape indices shift after add/remove/z-order. Run `ppt_get_slide_info.py` to refresh. |
| `Approval token required` (exit code 4) | Generate token with `scripts/generate_token.py` or inline Python. |
| `LibreOffice not found` | Install: `sudo apt install libreoffice-impress` (required for PDF/image export only). |
| `ppt_add_slide.py --title` doesn't work | `ppt_add_slide.py` has no `--title` arg. Use `ppt_set_title.py` separately. |
| `ppt_add_shape.py --shape-type` doesn't work | Use `--shape` not `--shape-type` (e.g., `--shape rectangle`). |
| `ppt_set_footer.py --show-page-number` fails | Use `--show-number` not `--show-page-number`. |

---

## 📄 License

MIT License. See [LICENSE](LICENSE) for details.
