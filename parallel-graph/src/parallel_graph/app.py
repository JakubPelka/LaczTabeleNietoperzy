"""Tkinter desktop application."""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from .export import generate_outputs
from .models import SourceSpec
from .translations import LANGUAGES, tr


class ParallelGraphApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Parallel Bat Graph")
        self.root.minsize(820, 520)
        self.language_var = tk.StringVar(value=LANGUAGES["pl"])
        self.output_var = tk.StringVar()
        self.status_var = tk.StringVar()
        self.source_rows: list[tuple[Path, tk.StringVar, ttk.Frame, ttk.Button]] = []
        self._build_ui()
        self._set_language()

    def _language(self) -> str:
        selected = self.language_var.get()
        return next((code for code, label in LANGUAGES.items() if label == selected), "pl")

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        header = ttk.Frame(outer)
        header.pack(fill="x")
        self.title_label = ttk.Label(header, font=("TkDefaultFont", 15, "bold"))
        self.title_label.pack(side="left", anchor="w")
        self.language_combo = ttk.Combobox(
            header, textvariable=self.language_var, values=list(LANGUAGES.values()),
            state="readonly", width=12,
        )
        self.language_combo.pack(side="right")
        self.language_label = ttk.Label(header)
        self.language_label.pack(side="right", padx=(0, 7))
        self.language_combo.bind("<<ComboboxSelected>>", lambda _event: self._set_language())
        self.intro_label = ttk.Label(
            outer,
            wraplength=780,
        )
        self.intro_label.pack(anchor="w", pady=(3, 12))

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(0, 8))
        self.add_button = ttk.Button(actions, command=self._choose_files)
        self.add_button.pack(side="left")
        self.clear_button = ttk.Button(actions, command=self._clear_sources)
        self.clear_button.pack(side="left", padx=8)

        self.source_box = ttk.LabelFrame(outer)
        self.source_box.pack(fill="both", expand=True)
        canvas = tk.Canvas(self.source_box, highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.source_box, orient="vertical", command=canvas.yview)
        self.source_frame = ttk.Frame(canvas, padding=7)
        self.source_window = canvas.create_window((0, 0), window=self.source_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        self.source_frame.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda event: canvas.itemconfigure(self.source_window, width=event.width))
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        self.output_box = ttk.LabelFrame(outer, padding=8)
        self.output_box.pack(fill="x", pady=(12, 8))
        ttk.Entry(self.output_box, textvariable=self.output_var).pack(side="left", fill="x", expand=True)
        self.output_button = ttk.Button(self.output_box, command=self._choose_output)
        self.output_button.pack(side="left", padx=(8, 0))

        footer = ttk.Frame(outer)
        footer.pack(fill="x")
        ttk.Label(footer, textvariable=self.status_var).pack(side="left", fill="x", expand=True)
        self.generate_button = ttk.Button(footer, command=self._start_generation)
        self.generate_button.pack(side="right")

    def _set_language(self) -> None:
        language = self._language()
        self.root.title(tr(language, "window_title"))
        self.title_label.configure(text=tr(language, "app_title"))
        self.intro_label.configure(text=tr(language, "intro"))
        self.language_label.configure(text=tr(language, "language"))
        self.add_button.configure(text=tr(language, "add_files"))
        self.clear_button.configure(text=tr(language, "clear_list"))
        self.source_box.configure(text=tr(language, "data_sources"))
        self.output_box.configure(text=tr(language, "output_folder"))
        self.output_button.configure(text=tr(language, "choose"))
        self.generate_button.configure(text=tr(language, "generate"))
        for _path, _name, _frame, remove_button in self.source_rows:
            remove_button.configure(text=tr(language, "remove"))
        self.status_var.set(tr(language, "initial_status"))

    def _choose_files(self) -> None:
        selected = filedialog.askopenfilenames(
            title=tr(self._language(), "file_dialog_title"),
            filetypes=[
                (tr(self._language(), "excel_files"), "*.xlsx"),
                (tr(self._language(), "all_files"), "*.*"),
            ],
        )
        existing = {path.resolve() for path, _name, _frame, _button in self.source_rows}
        for selected_path in selected:
            path = Path(selected_path).resolve()
            if path not in existing:
                self._add_source(path)
                existing.add(path)
        if selected and not self.output_var.get():
            self.output_var.set(str(Path(selected[0]).resolve().parent))

    def _add_source(self, path: Path) -> None:
        row = ttk.Frame(self.source_frame)
        row.pack(fill="x", pady=3)
        name_var = tk.StringVar(value=path.stem)
        ttk.Entry(row, textvariable=name_var, width=30).pack(side="left", padx=(0, 8))
        ttk.Label(row, text=str(path), anchor="w").pack(side="left", fill="x", expand=True)
        remove_button = ttk.Button(
            row, text=tr(self._language(), "remove"), command=lambda: self._remove_source(row)
        )
        remove_button.pack(side="right", padx=(8, 0))
        self.source_rows.append((path, name_var, row, remove_button))

    def _remove_source(self, frame: ttk.Frame) -> None:
        self.source_rows = [item for item in self.source_rows if item[2] is not frame]
        frame.destroy()

    def _clear_sources(self) -> None:
        for _path, _name, frame, _button in self.source_rows:
            frame.destroy()
        self.source_rows.clear()

    def _choose_output(self) -> None:
        selected = filedialog.askdirectory(title=tr(self._language(), "output_dialog_title"))
        if selected:
            self.output_var.set(selected)

    def _start_generation(self) -> None:
        language = self._language()
        sources = [
            SourceSpec(path, name.get().strip())
            for path, name, _frame, _button in self.source_rows
        ]
        output = self.output_var.get().strip()
        if not sources:
            messagebox.showerror(tr(language, "missing_files_title"), tr(language, "missing_files"))
            return
        if not output:
            messagebox.showerror(tr(language, "missing_output_title"), tr(language, "missing_output"))
            return
        self.generate_button.configure(state="disabled")
        self.status_var.set(tr(language, "processing"))
        threading.Thread(
            target=self._generate, args=(sources, Path(output), language), daemon=True
        ).start()

    def _generate(self, sources: list[SourceSpec], output: Path, language: str) -> None:
        try:
            paths, reports = generate_outputs(sources, output, language)
        except Exception as exc:  # Display input/dependency errors in the GUI.
            self.root.after(0, self._generation_failed, str(exc), language)
            return
        self.root.after(0, self._generation_done, paths, reports, language)

    def _generation_failed(self, message: str, language: str) -> None:
        self.generate_button.configure(state="normal")
        self.status_var.set(tr(language, "failed"))
        messagebox.showerror(tr(language, "error"), message)

    def _generation_done(self, paths: list[Path], reports: list[object], language: str) -> None:
        self.generate_button.configure(state="normal")
        loaded = sum(getattr(report, "registrations_loaded", 0) for report in reports)
        self.status_var.set(tr(language, "ready_status", count=loaded, sources=len(reports)))
        html_path = paths[0]
        should_open = messagebox.askyesno(
            tr(language, "chart_ready"),
            tr(language, "open_chart", path=html_path),
        )
        if should_open:
            self._open_path(html_path)

    @staticmethod
    def _open_path(path: Path) -> None:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])


def main() -> None:
    root = tk.Tk()
    ParallelGraphApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
