#!/usr/bin/env bash
# probe_wrapper.sh
# Usage: probe_wrapper.sh /absolute/path/to/presentation.pptx
# Behavior:
# - Validates absolute path and readability
# - Runs ppt_capability_probe.py with timeout and retries
# - On failure, runs fallback probes: ppt_get_info.py and ppt_get_slide_info.py
# - Emits JSON to stdout; on error emits structured JSON and non-zero exit

set -euo pipefail

FILE="${1:-}"
TIMEOUT_SECONDS=15
MAX_RETRIES=3
SLEEP_BASE=2
TMPDIR="$(mktemp -d)"
PROBE_OUT="$TMPDIR/probe.json"

function emit_error {
  jq -n --arg code "$1" --arg msg "$2" --argjson retryable "$3" \
    '{error:{error_code:$code, message:$msg, retryable:$retryable}}'
}

if [[ -z "$FILE" ]]; then
  emit_error "USAGE_ERROR" "Missing file argument" false
  exit 1
fi

if [[ "${FILE:0:1}" != "/" ]]; then
  emit_error "RELATIVE_PATH_NOT_ALLOWED" "Absolute path required" false
  exit 1
fi

if [[ ! -r "$FILE" ]]; then
  emit_error "PERMISSION_DENIED" "File not readable" false
  exit 1
fi

# Disk space check on containing filesystem
MIN_SPACE_MB=100
avail_mb=$(df --output=avail -m "$(dirname "$FILE")" | tail -1 | tr -d ' ')
if [[ -z "$avail_mb" || "$avail_mb" -lt "$MIN_SPACE_MB" ]]; then
  emit_error "LOW_DISK_SPACE" "Available space less than ${MIN_SPACE_MB}MB" false
  exit 1
fi

# Check tool availability
if ! command -v ppt_capability_probe.py >/dev/null 2>&1; then
  emit_error "TOOL_MISSING" "ppt_capability_probe.py not found in PATH" false
  exit 1
fi

# Attempt probe with retries and exponential backoff
attempt=0
while [[ $attempt -lt $MAX_RETRIES ]]; do
  attempt=$((attempt+1))
  if timeout "${TIMEOUT_SECONDS}s" ppt_capability_probe.py --file "$FILE" --deep --json > "$PROBE_OUT" 2>&1; then
    cat "$PROBE_OUT"
    rm -rf "$TMPDIR"
    exit 0
  else
    sleep_time=$((SLEEP_BASE ** attempt))
    sleep "$sleep_time"
  fi
done

# Fallback probes
if command -v ppt_get_info.py >/dev/null 2>&1 && command -v ppt_get_slide_info.py >/dev/null 2>&1; then
  info_json="$TMPDIR/info.json"
  slide0_json="$TMPDIR/slide0.json"
  if ppt_get_info.py --file "$FILE" --json > "$info_json" 2>&1 && ppt_get_slide_info.py --file "$FILE" --slide 0 --json > "$slide0_json" 2>&1; then
    # Merge minimal metadata into a single JSON object
    jq -s '.[0] + {probe_fallback:true, slide0:.[1]}' "$info_json" "$slide0_json"
    rm -rf "$TMPDIR"
    exit 0
  else
    emit_error "PROBE_FALLBACK_FAILED" "Both deep probe and fallback probes failed" true
    rm -rf "$TMPDIR"
    exit 3
  fi
else
  emit_error "PROBE_AND_FALLBACK_TOOLS_MISSING" "Fallback tools not available" false
  rm -rf "$TMPDIR"
  exit 1
fi
