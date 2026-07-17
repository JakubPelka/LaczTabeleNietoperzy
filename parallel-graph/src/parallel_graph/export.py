"""Generate all output artifacts."""

from __future__ import annotations

import csv
from datetime import datetime
from pathlib import Path
from typing import Iterable

from .aggregation import aggregate_by_minute
from .chart import write_chart
from .excel_reader import read_sources
from .models import ImportReport, MinuteCount, SourceSpec
from .translations import tr


def _write_csv(path: Path, counts: Iterable[MinuteCount]) -> None:
    with path.open("w", encoding="utf-8-sig", newline="") as file_obj:
        writer = csv.writer(file_obj)
        writer.writerow(["minute", "source", "manual_id", "count"])
        for row in counts:
            writer.writerow([row.minute.isoformat(sep=" ", timespec="minutes"), row.source, row.species, row.count])


def _write_report(path: Path, reports: list[ImportReport], counts: list[MinuteCount], language: str) -> None:
    total_loaded = sum(report.registrations_loaded for report in reports)
    lines = [
        tr(language, "report_title"),
        f"{tr(language, 'created')}: {datetime.now().isoformat(sep=' ', timespec='seconds')}",
        f"{tr(language, 'report_sources')}: {len(reports)}",
        f"{tr(language, 'loaded_total')}: {total_loaded}",
        f"{tr(language, 'aggregated_points')}: {len(counts)}",
        "",
    ]
    for report in reports:
        lines.extend([
            f"[{report.source}]",
            f"{tr(language, 'file')}: {report.path}",
            f"{tr(language, 'sheet')}: {report.sheet}",
            f"{tr(language, 'loaded')}: {report.registrations_loaded}",
            f"{tr(language, 'blank_skipped')}: {report.blank_species_skipped}",
            f"{tr(language, 'noise_skipped')}: {report.noise_skipped}",
            f"{tr(language, 'invalid_time_skipped')}: {report.invalid_time_skipped}",
            "",
        ])
    path.write_text("\n".join(lines), encoding="utf-8")


def generate_outputs(
    sources: list[SourceSpec], output_directory: Path, language: str = "pl"
) -> tuple[list[Path], list[ImportReport]]:
    if not sources:
        raise ValueError("Wybierz co najmniej jeden plik Excel.")
    names = [source.name.strip() for source in sources]
    if any(not name for name in names):
        raise ValueError("Każde źródło musi mieć nazwę.")
    if len(set(name.casefold() for name in names)) != len(names):
        raise ValueError("Nazwy źródeł muszą być unikalne.")
    output_directory = Path(output_directory).expanduser().resolve()
    output_directory.mkdir(parents=True, exist_ok=True)
    registrations, reports = read_sources(sources)
    counts = aggregate_by_minute(registrations)
    if not counts:
        raise ValueError("Nie znaleziono rejestracji z prawidłowym czasem i MANUAL ID.")
    html_path = output_directory / "parallel_bat_activity.html"
    csv_path = output_directory / "parallel_bat_activity_data.csv"
    report_path = output_directory / "parallel_bat_activity_report.txt"
    write_chart(html_path, counts, tr(language, "chart_title"), language)
    _write_csv(csv_path, counts)
    _write_report(report_path, reports, counts, language)
    return [html_path, csv_path, report_path], reports
