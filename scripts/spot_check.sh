#!/bin/bash

echo "SPOT CHECK: Validating Critical Tool Calls"
echo "==========================================="

FILE="/home/project/powerpoint-agent-tools/AGENT_SYSTEM_PROMPT.md"

# Check 1: ppt_add_slide calls
echo -e "\n✓ ppt_add_slide calls (should have --file, --layout, optional --index, --title)"
grep -A 4 "uv run tools/ppt_add_slide.py" "$FILE" | head -25 | grep -E "^\s*--" | head -15

# Check 2: ppt_set_title calls
echo -e "\n✓ ppt_set_title calls (should have --file, --slide, --title, optional --subtitle)"
grep -A 5 "uv run tools/ppt_set_title.py" "$FILE" | head -30 | grep -E "^\s*--" | head -20

# Check 3: ppt_add_notes calls (especially for --mode parameter)
echo -e "\n✓ ppt_add_notes calls (check --mode parameter validity)"
grep -B 2 "uv run tools/ppt_add_notes.py" "$FILE" | grep -A 6 "uv run tools/ppt_add_notes.py" | grep -E "(--mode|--text)" | head -10

# Check 4: ppt_add_text_box with position and size JSON
echo -e "\n✓ ppt_add_text_box position/size syntax (JSON should be valid)"
grep -A 5 "uv run tools/ppt_add_text_box.py" "$FILE" | grep -E "(--position|--size)" | head -10

# Check 5: Look for any remaining LAST references
echo -e "\n✓ Verify NO remaining 'LAST' references"
LAST_COUNT=$(grep -c "slide LAST\|index LAST" "$FILE" 2>/dev/null || echo "0")
echo "Found 'LAST' references: $LAST_COUNT (should be 0)"

# Check 6: Validate critical argument patterns
echo -e "\n✓ Verify --json is used consistently"
TOOL_CALLS=$(grep -c "uv run tools/ppt_" "$FILE")
JSON_CALLS=$(grep "uv run tools/ppt_" "$FILE" | grep -c "\-\-json")
echo "Total tool calls: $TOOL_CALLS"
echo "Tool calls with --json: $JSON_CALLS"

# Check 7: Look for unclosed quotes
echo -e "\n✓ Check for quote balance in critical sections"
SECTION_START=$(grep -n "## SECTION VIII" "$FILE" | cut -d: -f1)
SECTION_END=$(grep -n "## SECTION IX\|## APPENDIX" "$FILE" | head -1 | cut -d: -f1)
echo "Checking Section VIII (Pattern Library) lines $SECTION_START-$SECTION_END..."
# Simple heuristic: check for matching braces
OPEN_BRACES=$(sed -n "${SECTION_START},${SECTION_END}p" "$FILE" | grep -o '{' | wc -l)
CLOSE_BRACES=$(sed -n "${SECTION_START},${SECTION_END}p" "$FILE" | grep -o '}' | wc -l)
echo "  Open braces: $OPEN_BRACES, Close braces: $CLOSE_BRACES"
[ "$OPEN_BRACES" -eq "$CLOSE_BRACES" ] && echo "  ✅ Braces balanced" || echo "  ⚠️ Brace mismatch detected"

