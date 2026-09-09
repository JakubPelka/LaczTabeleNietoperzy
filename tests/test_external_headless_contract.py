import os
import sys
import shutil
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
if str(PROJECT_ROOT / "parallel-graph" / "src") not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT / "parallel-graph" / "src"))

from algorithms.merge_summary_and_nightly_charts import (
    validate_hhmm,
    get_unique_input_stems,
    run_analysis,
)
from parallel_graph.excel_reader import is_time_in_range
from datetime import datetime


class ExternalHeadlessContractTests(TestCase):
    def test_1_manual_recording_window_validation(self) -> None:
        self.assertEqual(validate_hhmm("21:00"), (21, 0))
        self.assertEqual(validate_hhmm("04:30"), (4, 30))
        self.assertEqual(validate_hhmm("00:00"), (0, 0))
        self.assertEqual(validate_hhmm("23:59"), (23, 59))

        invalid_inputs = ["25:00", "21:70", "foo", "21", "21:00:00"]
        for invalid in invalid_inputs:
            with self.assertRaises(ValueError):
                validate_hhmm(invalid)

    def test_2_overnight_range_2100_to_0430(self) -> None:
        start, stop = "21:00", "04:30"
        self.assertTrue(is_time_in_range(datetime(2026, 7, 14, 22, 30), start, stop))
        self.assertTrue(is_time_in_range(datetime(2026, 7, 15, 2, 15), start, stop))
        self.assertTrue(is_time_in_range(datetime(2026, 7, 14, 21, 0), start, stop))
        self.assertTrue(is_time_in_range(datetime(2026, 7, 15, 4, 30), start, stop))

        self.assertFalse(is_time_in_range(datetime(2026, 7, 14, 14, 0), start, stop))
        self.assertFalse(is_time_in_range(datetime(2026, 7, 14, 20, 59), start, stop))
        self.assertFalse(is_time_in_range(datetime(2026, 7, 15, 4, 31), start, stop))

    def test_3_overnight_range_2000_to_0600(self) -> None:
        start, stop = "20:00", "06:00"
        self.assertTrue(is_time_in_range(datetime(2026, 7, 14, 21, 0), start, stop))
        self.assertTrue(is_time_in_range(datetime(2026, 7, 15, 5, 30), start, stop))
        self.assertFalse(is_time_in_range(datetime(2026, 7, 14, 12, 0), start, stop))

    def test_4_to_12_headless_run_outputs(self) -> None:
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            dir_a = temp_path / "dir_a"
            dir_b = temp_path / "dir_b"
            dir_a.mkdir()
            dir_b.mkdir()

            file1 = dir_a / "bat_data.csv"
            file2 = dir_b / "bat_data.csv"  # duplicate stem

            csv_content1 = (
                "DATE;TIME;MANUAL ID\n"
                "2026-07-14;22:15:00;Nyctalus noctula FOD\n"
                "2026-07-15;02:30:00;Myotis daubentonii SOC\n"
            )
            csv_content2 = (
                "DATE;TIME;MANUAL ID\n"
                "2026-07-14;23:00:00;Pipistrellus pygmaeus FORBI\n"
            )
            file1.write_text(csv_content1, encoding="utf-8")
            file2.write_text(csv_content2, encoding="utf-8")

            # Stem mapping check
            input_files = [str(file1), str(file2)]
            stems = get_unique_input_stems(input_files)
            self.assertEqual(stems[str(file1)], "bat_data")
            self.assertEqual(stems[str(file2)], "bat_data_2")

            # Run 1: HTML disabled
            run_analysis({
                "input_files": input_files,
                "base_dir": str(temp_path),
                "base_name": "test_run",
                "custom_time_range": ("21:00", "04:30"),
                "do_plots_summary": True,
                "do_plots_pernight": True,
                "generate_html": False,
                "open_files": False,
            })

            combined_dir = temp_path / "results" / "combined"
            inputs_dir = temp_path / "results" / "inputs"

            # 9. Combined NVI/ART files exist
            self.assertTrue((combined_dir / "test_run_NVI.xlsx").is_file())
            self.assertTrue((combined_dir / "test_run_ART.xlsx").is_file())

            # 6. HTML disabled => no parallel files
            self.assertFalse((combined_dir / "parallel_bat_activity.html").exists())
            self.assertFalse((combined_dir / "parallel_bat_activity_data.csv").exists())
            self.assertFalse((combined_dir / "parallel_bat_activity_report.txt").exists())

            # 11 & 12. Two input sources produce two deterministic parent folders
            parent_folders = sorted([p.name for p in inputs_dir.iterdir() if p.is_dir()])
            self.assertEqual(parent_folders, ["bat_data", "bat_data_2"])

            # 10. Per-input folder structure
            for stem in parent_folders:
                stem_path = inputs_dir / stem
                self.assertTrue((stem_path / "summary" / "line").is_dir())
                self.assertTrue((stem_path / "summary" / "stacked_ART").is_dir())
                self.assertTrue((stem_path / "summary" / "stacked_NVI").is_dir())
                self.assertTrue((stem_path / "nights").is_dir())

            # Run 2: HTML enabled
            run_analysis({
                "input_files": input_files,
                "base_dir": str(temp_path),
                "base_name": "test_run",
                "custom_time_range": ("21:00", "04:30"),
                "do_plots_summary": True,
                "do_plots_pernight": True,
                "generate_html": True,
                "open_files": False,
            })

            # 7. HTML enabled => HTML + CSV + report
            html_file = combined_dir / "parallel_bat_activity.html"
            csv_file = combined_dir / "parallel_bat_activity_data.csv"
            report_file = combined_dir / "parallel_bat_activity_report.txt"

            self.assertTrue(html_file.is_file())
            self.assertTrue(csv_file.is_file())
            self.assertTrue(report_file.is_file())

            # 8. HTML is self-contained
            html_text = html_file.read_text(encoding="utf-8")
            self.assertIn("Plotly.newPlot", html_text)
            self.assertNotIn('<script src="https://cdn.plot.ly', html_text)
