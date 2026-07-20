"""Offline launcher for Parallel Bat Graph using the current Python."""

from __future__ import annotations

import importlib.util
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
REQUIRED_MODULES = {
    "openpyxl": "openpyxl",
    "plotly": "plotly",
    "tkinter": "Tkinter (Tcl/Tk)",
}


def show_error(message: str) -> None:
    try:
        from tkinter import messagebox

        messagebox.showerror("Parallel Bat Graph", message)
    except Exception:
        print(message, file=sys.stderr)


def main() -> int:
    missing = [
        label for module, label in REQUIRED_MODULES.items() if importlib.util.find_spec(module) is None
    ]
    if missing:
        show_error(
            "W używanym Pythonie brakuje bibliotek:\n\n- "
            + "\n- ".join(missing)
            + "\n\nStarter działa offline i nie pobiera pakietów. "
            "Uruchom go Pythonem z USB, który ma już te biblioteki."
        )
        return 1

    sys.path.insert(0, str(PROJECT_DIR / "src"))
    from parallel_graph.app import main as run_application

    run_application()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
