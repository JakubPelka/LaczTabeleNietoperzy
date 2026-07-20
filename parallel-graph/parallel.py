"""Offline launcher for Parallel Bat Graph using project-local packages."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
VENDOR_DIR = PROJECT_DIR / "vendor"
REQUIRED_MODULES = {
    "openpyxl": "openpyxl",
    "plotly": "plotly",
}


def show_error(message: str) -> None:
    try:
        from tkinter import messagebox

        messagebox.showerror("Parallel Bat Graph", message)
    except Exception:
        print(message, file=sys.stderr)


def main() -> int:
    # ``parallel.sh`` installs third-party packages here. Keeping this path
    # first makes the downloaded copy available to direct offline launches.
    sys.path.insert(0, str(VENDOR_DIR))

    missing = [
        label for module, label in REQUIRED_MODULES.items() if not (VENDOR_DIR / module).exists()
    ]
    if missing:
        show_error(
            "Brakuje lokalnych bibliotek programu:\n\n- "
            + "\n- ".join(missing)
            + "\n\nPołącz komputer z internetem i uruchom raz parallel.sh. "
            "Później parallel.py będzie działać offline."
        )
        return 1

    if importlib.util.find_spec("tkinter") is None:
        show_error(
            "W używanym Pythonie brakuje Tkinter (Tcl/Tk).\n\n"
            "Tkinter jest częścią instalacji Pythona i nie może zostać "
            "pobrany przez pip. Zainstaluj Python z obsługą Tcl/Tk."
        )
        return 1

    sys.path.insert(0, str(PROJECT_DIR / "src"))
    from parallel_graph.app import main as run_application

    run_application()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
