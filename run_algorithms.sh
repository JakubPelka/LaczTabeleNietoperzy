#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
VENV_DIR="$PROJECT_DIR/.venv"
PYTHON_BIN="$VENV_DIR/bin/python"

if [[ ! -x "$PYTHON_BIN" ]]; then
    python3 -m venv "$VENV_DIR"
    "$PYTHON_BIN" -m pip install --upgrade pip
    "$PYTHON_BIN" -m pip install -r "$PROJECT_DIR/requirements.txt"
fi

printf '%s\n' \
    'Wybierz algorytm:' \
    '  1) Tabele i wykresy zbiorcze' \
    '  2) Tabele oraz wykresy zbiorcze i/lub noc-po-nocy'
read -r -p 'Wybór [1-2]: ' selection

case "$selection" in
    1) script="$PROJECT_DIR/algorithms/merge_summary_charts.py" ;;
    2) script="$PROJECT_DIR/algorithms/merge_summary_and_nightly_charts.py" ;;
    *) printf 'Nieprawidłowy wybór.\n' >&2; exit 2 ;;
esac

exec "$PYTHON_BIN" "$script"
