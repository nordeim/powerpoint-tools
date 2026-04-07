#!/bin/bash
set -e

# docker pull nordeim/ppt-agent-tools
# docker compose up -d
# Rebuild and recreate (useful after Dockerfile changes):
#docker compose up -d --build

# Get the root directory of the project
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )/.." && pwd )"

echo "🐳 Building PowerPoint Agent Tools Docker Image..."
#docker build -t ppt-agent-tools:latest -f "$DIR/docker/Dockerfile" "$DIR"
#docker build -t jeremiahwong/ppt-agent-tools:latest -f docker/Dockerfile .
#DIR=`pwd` &&  docker build -t ppt-agent-tools:latest -f "$DIR/docker/Dockerfile" "$DIR"
DIR=`pwd`&& docker build --ignorefile "$DIR/docker/.dockerignore" -t ppt-agent-tools:latest -f "$DIR/docker/Dockerfile" "$DIR"


echo "✅ Build Complete. Image: ppt-agent-tools:latest"

docker run -d --name ppt-agent-tools jeremiahwong/ppt-agent-tools:latest /bin/bash -c "while true; do sleep 86400; done"
