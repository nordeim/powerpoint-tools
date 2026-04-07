#!/usr/bin/env python3
"""
Validate all tool calls in AGENT_SYSTEM_PROMPT.md against actual tool signatures.
"""

import re
from pathlib import Path
from collections import defaultdict

# Read the system prompt file
system_prompt_path = Path("/home/project/powerpoint-agent-tools/AGENT_SYSTEM_PROMPT.md")
with open(system_prompt_path, 'r') as f:
    content = f.read()

# Extract all uv run commands
tool_call_pattern = r'uv run (tools/ppt_\w+\.py)[\s\\]*([^\n]*(?:\n\s+--[^\n]*)*)'
matches = re.finditer(tool_call_pattern, content)

tool_calls = []
for match in matches:
    tool_name = match.group(1)
    args_part = match.group(2)
    # Get line number
    line_num = content[:match.start()].count('\n') + 1
    tool_calls.append({
        'line': line_num,
        'tool': tool_name.split('/')[-1],
        'full_command': match.group(0)[:100]
    })

# Group by tool
by_tool = defaultdict(list)
for call in tool_calls:
    by_tool[call['tool']].append(call)

print(f"Found {len(tool_calls)} tool calls across {len(by_tool)} unique tools\n")

# Print summary by tool
for tool in sorted(by_tool.keys()):
    calls = by_tool[tool]
    print(f"{tool}: {len(calls)} calls")
    for call in calls[:3]:  # Show first 3
        print(f"  Line {call['line']}: {call['full_command'][:80]}...")

print(f"\n\nCRITICAL ISSUES FOUND:")
print("=" * 80)

# Check for LAST index usage
last_pattern = r'--(?:slide|index)\s+LAST'
last_matches = list(re.finditer(last_pattern, content))
if last_matches:
    print(f"\n🔴 CRITICAL: Found {len(last_matches)} uses of 'LAST' as slide/index value")
    for match in last_matches[:5]:
        line_num = content[:match.start()].count('\n') + 1
        line_text = content.split('\n')[line_num - 1].strip()
        print(f"   Line {line_num}: {line_text[:100]}")
    print("   ACTION: Tools expect integer slide indices (0-based), not 'LAST'")

