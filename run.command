#!/bin/bash
# Double-clickable launcher for macOS (Finder). Runs run.sh in this folder.
cd "$(dirname "$0")"
exec /bin/bash "./run.sh"
