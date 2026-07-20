"""Offline launcher for the two LaczTabeleNietoperzy algorithms.

This file deliberately never creates a virtual environment or invokes pip.  It
uses the Python interpreter that opened it, which makes it suitable for a
portable Python installation stored on a USB drive.
"""

from __future__ import annotations

import importlib.util
import runpy
import sys
from pathlib import Path


PROJECT_DIR = Path(__file__).resolve().parent
REQUIRED_MODULES = {
    "pandas": "pandas",
    "openpyxl": "openpyxl",
    "matplotlib": "matplotlib",
    "tkinter": "Tkinter (Tcl/Tk)",
}


def missing_dependencies() -> list[str]:
    return [label for module, label in REQUIRED_MODULES.items() if importlib.util.find_spec(module) is None]


def show_error(message: str) -> None:
    try:
        from tkinter import messagebox

        messagebox.showerror("LaczTabeleNietoperzy", message)
    except Exception:
        print(message, file=sys.stderr)


def choose_algorithm() -> Path | None:
    import tkinter as tk

    root = tk.Tk()
    root.title("LaczTabeleNietoperzy")
    root.resizable(False, False)
    selected: list[Path] = []

    tk.Label(root, text="Wybierz algorytm:", font=("TkDefaultFont", 11, "bold")).pack(
        padx=24, pady=(20, 12), anchor="w"
    )

    choices = (
        ("Tabele i wykresy zbiorcze", "merge_summary_charts.py"),
        ("Tabele oraz wykresy zbiorcze i/lub noc-po-nocy", "merge_summary_and_nightly_charts.py"),
    )

    def select(filename: str) -> None:
        selected.append(PROJECT_DIR / "algorithms" / filename)
        root.destroy()

    for label, filename in choices:
        tk.Button(root, text=label, width=52, command=lambda name=filename: select(name)).pack(
            padx=24, pady=5
        )
    tk.Button(root, text="Anuluj", width=18, command=root.destroy).pack(pady=(12, 20))
    root.mainloop()
    return selected[0] if selected else None


def main() -> int:
    missing = missing_dependencies()
    if missing:
        show_error(
            "W używanym Pythonie brakuje bibliotek:\n\n- "
            + "\n- ".join(missing)
            + "\n\nStarter działa offline i nie pobiera pakietów. "
            "Uruchom go Pythonem z USB, który ma już te biblioteki."
        )
        return 1

    script = choose_algorithm()
    if script is None:
        return 0
    runpy.run_path(str(script), run_name="__main__")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
