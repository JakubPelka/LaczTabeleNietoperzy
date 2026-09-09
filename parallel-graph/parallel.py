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
    sys.path.insert(0, str(VENDOR_DIR))
    sys.path.insert(0, str(PROJECT_DIR / "src"))

    if len(sys.argv) > 1:
        import argparse
        parser = argparse.ArgumentParser(description="Parallel Bat Graph Headless CLI")
        parser.add_argument("--input-files", nargs="+", required=True, help="Input XLSX or CSV files")
        parser.add_argument("--output-dir", required=True, help="Output directory")
        parser.add_argument("--time-start", help="Start time HH:MM")
        parser.add_argument("--time-end", help="End time HH:MM")
        parser.add_argument("--language", default="pl", help="Language (pl, sv, en)")
        args = parser.parse_args()

        time_range = (args.time_start, args.time_end) if args.time_start and args.time_end else None
        from parallel_graph.export import generate_outputs
        from parallel_graph.models import SourceSpec
        sources = [SourceSpec(path=Path(f), name=Path(f).stem) for f in args.input_files]
        generate_outputs(sources, Path(args.output_dir), language=args.language, time_range=time_range)
        return 0

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

    from parallel_graph.app import main as run_application

    run_application()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

