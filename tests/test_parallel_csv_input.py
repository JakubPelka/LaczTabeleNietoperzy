from datetime import datetime
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
import sys


PROJECT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(PROJECT_ROOT / "parallel-graph" / "src"))

from parallel_graph.excel_reader import WorkbookFormatError, read_source
from parallel_graph.export import generate_outputs
from parallel_graph.models import SourceSpec


class ParallelCsvInputTests(TestCase):
    def test_reads_semicolon_csv_with_bom_and_header_after_preamble(self) -> None:
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "source.csv"
            path.write_text(
                "Export from detector\n"
                "DATE;TIME;MANUAL ID;AUTO ID\n"
                "2026-07-14;01:02:03;Nyctalus noctula FOD;NYCNOC\n"
                '2026-07-14;01:02:04;"Myotis daubentonii, Pipistrellus pygmaeus";MYODAU\n'
                "2026-07-14;01:02:05;Noise;EPTNIL\n"
                "2026-07-14;01:02:06;;EPTNIL\n"
                "2026-07-14;bad;Eptesicus nilssonii;EPTNIL\n",
                encoding="utf-8-sig",
            )

            registrations, report = read_source(SourceSpec(path, "CSV source"))

            self.assertEqual(
                [item.species for item in registrations],
                [
                    "Nyctalus noctula FOD",
                    "Myotis daubentonii",
                    "Pipistrellus pygmaeus",
                ],
            )
            self.assertEqual(registrations[0].timestamp, datetime(2026, 7, 14, 1, 2, 3))
            self.assertEqual(report.noise_skipped, 1)
            self.assertEqual(report.blank_species_skipped, 1)
            self.assertEqual(report.invalid_time_skipped, 1)

    def test_generates_all_outputs_from_windows_encoded_csv(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            source_path = directory / "source.CSV"
            source_path.write_bytes(
                (
                    "DATE,TIME,MANUAL ID\n"
                    '2026-07-14,01:02:03,"Plecotus auritus FÖD"\n'
                ).encode("cp1252")
            )

            paths, reports = generate_outputs(
                [SourceSpec(source_path, "CSV source")], directory / "output"
            )

            self.assertTrue(all(path.is_file() for path in paths))
            self.assertEqual(reports[0].registrations_loaded, 1)

    def test_rejects_unsupported_input_format(self) -> None:
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "source.txt"
            path.write_text("DATE,TIME,MANUAL ID\n", encoding="utf-8")
            with self.assertRaises(WorkbookFormatError):
                read_source(SourceSpec(path, "Bad"))
