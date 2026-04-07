#!/usr/bin/env bash
# preflight_check.sh
# Usage: preflight_check.sh /absolute/path/to/presentation.pptx
# Performs:
# - absolute path enforcement
# - read/write permission checks
# - disk space check
# - probe via probe_wrapper.sh
# - outputs JSON summary

set -euo pipefail

FILE="${1:-}"
MIN_SPACE_MB=100

if [[ -z "$FILE" ]]; then
  echo '{"error":"Missing file argument"}' && exit 1
fi

if [[ "${FILE:0:1}" != "/" ]]; then
  echo '{"error":"Absolute path required"}' && exit 1
fi

if [[ ! -r "$FILE" ]]; then
  echo '{"error":"File not readable"}' && exit 1
fi

if [[ ! -w "$(dirname "$FILE")" ]]; then
  echo '{"error":"No write permission to destination directory"}' && exit 1
fi

avail_mb=$(df --output=avail -m "$(dirname "$FILE")" | tail -1 | tr -d ' ')
if [[ -z "$avail_mb" || "$avail_mb" -lt "$MIN_SPACE_MB" ]]; then
  echo "{\"error\":\"Low disk space: ${avail_mb}MB available\"}" && exit 1
fi

# Run probe wrapper and capture JSON
if probe_wrapper.sh "$FILE" > /tmp/preflight_probe.json 2>&1; then
  jq -n --arg file "$FILE" --argjson probe "$(cat /tmp/preflight_probe.json | jq '.')" \
    '{file:$file, preflight: {status: "ok"}, probe: $probe}'
  exit 0
else
  echo '{"error":"Probe failed"}' && exit 2
fi
