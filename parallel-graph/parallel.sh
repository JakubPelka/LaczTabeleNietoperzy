#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$PROJECT_DIR/.venv"

if ! command -v python3 >/dev/null 2>&1; then
    echo "Nie znaleziono python3. Zainstaluj Python 3.10 lub nowszy." >&2
    exit 1
fi

if ! python3 - <<'PY' >/dev/null 2>&1
import tkinter
PY
then
    echo "Brakuje Tkinter. Na Ubuntu/Debian zainstaluj: sudo apt install python3-tk" >&2
    exit 1
fi

if [[ ! -x "$VENV_DIR/bin/python" ]]; then
    python3 -m venv "$VENV_DIR" || {
        echo "Nie udało się utworzyć środowiska. Na Ubuntu/Debian zainstaluj python3-venv." >&2
        exit 1
    }
fi

if ! "$VENV_DIR/bin/python" -c "import openpyxl, plotly" >/dev/null 2>&1; then
    "$VENV_DIR/bin/python" -m pip install --upgrade pip
    "$VENV_DIR/bin/python" -m pip install -e "$PROJECT_DIR"
fi

exec "$VENV_DIR/bin/python" -m parallel_graph

