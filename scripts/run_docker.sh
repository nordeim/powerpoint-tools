#!/bin/bash
set -e

# Get the root directory of the project
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )/.." && pwd )"

echo "🚀 Starting PowerPoint Agent Container..."
echo "📂 Mounting: $DIR -> /app"

# Run interactive container
# --rm: Remove container after exit
# -it: Interactive TTY
# -v: Mount project root to /app
docker run --rm -it \
    -v "$DIR:/app" \
    ppt-agent-tools:latest

