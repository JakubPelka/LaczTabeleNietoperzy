#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"

if [ -x "$PROJECT_DIR/.venv/bin/python3" ]; then
    exec "$PROJECT_DIR/.venv/bin/python3" "$PROJECT_DIR/start.py"
else
    exec python3 "$PROJECT_DIR/start.py"
fi
