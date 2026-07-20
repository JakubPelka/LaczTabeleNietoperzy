#!/usr/bin/env bash
set -euo pipefail

PROJECT_DIR="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
VENDOR_DIR="$PROJECT_DIR/vendor"

if ! command -v python3 >/dev/null 2>&1; then
    echo "Nie znaleziono python3. Zainstaluj Python 3.10 lub nowszy." >&2
    exit 1
fi

if ! python3 -m pip --version >/dev/null 2>&1; then
    echo "Brakuje pip dla python3. Zainstaluj pip i uruchom skrypt ponownie." >&2
    exit 1
fi

mkdir -p "$VENDOR_DIR"
echo "Pobieranie bibliotek do: $VENDOR_DIR"
python3 -m pip install \
    --disable-pip-version-check \
    --upgrade \
    --target "$VENDOR_DIR" \
    --requirement "$PROJECT_DIR/requirements.txt"

exec python3 "$PROJECT_DIR/parallel.py"
