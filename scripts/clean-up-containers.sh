#!/bin/bash
docker image prune -f            # removes dangling images
docker builder prune -af         # reclaims builder cache (careful: removes all build cache)
docker images --format "table {{.Repository}}\t{{.Tag}}\t{{.ID}}\t{{.Size}}"
