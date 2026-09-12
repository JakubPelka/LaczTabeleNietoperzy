import os
import sys
import shutil
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
from datetime import datetime

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))
if str(PROJECT_ROOT / "parallel-graph" / "src") not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT / "parallel-graph" / "src"))

from algorithms.merge_summary_and_nightly_charts import (
    validate_hhmm,
    get_unique_input_stems,
    build_manual_interval_sequence,
    run_analysis,
)
from parallel_graph.excel_reader import is_time_in_range


class ExternalHeadlessContractTests(TestCase):
    def test_1_manual_recording_window_validation(self) -> None:
        self.assertEqual(validate_hhmm("21:00"), (21, 0))
        self.assertEqual(validate_hhmm("04:30"), (4, 30))
        self.assertEqual(validate_hhmm("00:00"), (0, 0))
        self.assertEqual(validate_hhmm("23:59"), (23, 59))
        self.assertEqual(validate_hhmm("09:00"), (9, 0))

        invalid_inputs = ["9:00", "5:30", "25:00", "21:70", "foo", "21", "21:00:00"]
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

    def test_4_global_stem_collision_safety(self) -> None:
        files = ["/dir1/sample.csv", "/dir2/sample.csv", "/dir3/sample_2.csv"]
        stems = get_unique_input_stems(files)
        assigned = list(stems.values())
        self.assertEqual(len(set(assigned)), 3)
        self.assertEqual(assigned, ["sample", "sample_2", "sample_2_2"])

    def test_5_to_12_headless_run_outputs(self) -> None:
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

            input_files = [str(file1), str(file2)]

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
            inputs_dir = temp_path / "results"

            # Combined NVI/ART files exist
            self.assertTrue((combined_dir / "test_run_NVI.xlsx").is_file())
            self.assertTrue((combined_dir / "test_run_ART.xlsx").is_file())

            # HTML disabled => no parallel files
            self.assertFalse((combined_dir / "parallel_bat_activity.html").exists())
            self.assertFalse((combined_dir / "parallel_bat_activity_data.csv").exists())
            self.assertFalse((combined_dir / "parallel_bat_activity_report.txt").exists())

            # Two input sources produce two deterministic parent folders directly under results/
            parent_folders = sorted([p.name for p in inputs_dir.iterdir() if p.is_dir() and p.name != "combined"])
            self.assertEqual(parent_folders, ["bat_data", "bat_data_2"])

            # Per-input folder structure
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

            html_file = combined_dir / "parallel_bat_activity.html"
            csv_file = combined_dir / "parallel_bat_activity_data.csv"
            report_file = combined_dir / "parallel_bat_activity_report.txt"

            self.assertTrue(html_file.is_file())
            self.assertTrue(csv_file.is_file())
            self.assertTrue(report_file.is_file())

            # HTML is self-contained
            html_text = html_file.read_text(encoding="utf-8")
            self.assertIn("Plotly.newPlot", html_text)
            self.assertNotIn('<script src="https://cdn.plot.ly', html_text)

            # Report does not leak absolute server paths
            report_text = report_file.read_text(encoding="utf-8")
            self.assertIn("Plik: bat_data.csv", report_text)
            self.assertNotIn(str(temp_path), report_text)

    def test_13_html_generator_fails_closed(self) -> None:
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            empty_file = temp_path / "empty.csv"
            empty_file.write_text("DATE;TIME;MANUAL ID\n", encoding="utf-8")

            with self.assertRaises(ValueError):
                run_analysis({
                    "input_files": [str(empty_file)],
                    "base_dir": str(temp_path),
                    "base_name": "test_empty",
                    "custom_time_range": ("21:00", "04:30"),
                    "do_plots_summary": False,
                    "do_plots_pernight": False,
                    "generate_html": True,
                    "open_files": False,
                })

    def test_14_static_chart_filtering_excludes_out_of_window_data(self) -> None:
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            bat_file = temp_path / "bat_data.csv"
            csv_content = (
                "DATE;TIME;MANUAL ID\n"
                "2026-07-14;22:15:00;Nyctalus noctula FOD\n"
                "2026-07-14;14:00:00;Eptesicus nilssonii SOC\n"
            )
            bat_file.write_text(csv_content, encoding="utf-8")

            # Manual mode 21:00 -> 04:30
            run_analysis({
                "input_files": [str(bat_file)],
                "base_dir": str(temp_path),
                "base_name": "test_filtered",
                "custom_time_range": ("21:00", "04:30"),
                "do_plots_summary": True,
                "do_plots_pernight": False,
                "generate_html": False,
                "open_files": False,
            })

            line_dir = temp_path / "results" / "bat_data" / "summary" / "line"
            self.assertTrue((line_dir / "Nyctalus noctula.png").exists())
            # Eptesicus nilssonii was at 14:00 (out of window), must NOT be generated
            self.assertFalse((line_dir / "Cnephaeus nilssonii.png").exists())
            self.assertFalse((line_dir / "Eptesicus nilssonii.png").exists())

            # Auto mode (None) includes all rows
            run_analysis({
                "input_files": [str(bat_file)],
                "base_dir": str(temp_path),
                "base_name": "test_auto",
                "custom_time_range": None,
                "do_plots_summary": True,
                "do_plots_pernight": False,
                "generate_html": False,
                "open_files": False,
            })
            self.assertTrue((line_dir / "Cnephaeus nilssonii.png").exists())

    def test_15_diagram_base_custom_path(self) -> None:
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            custom_diagram_dir = temp_path / "custom_inputs"
            bat_file = temp_path / "bat_data.csv"
            csv_content = "DATE;TIME;MANUAL ID\n2026-07-14;22:15:00;Nyctalus noctula FOD\n"
            bat_file.write_text(csv_content, encoding="utf-8")

            run_analysis({
                "input_files": [str(bat_file)],
                "base_dir": str(temp_path),
                "diagram_base": str(custom_diagram_dir),
                "base_name": "test_diagram_base",
                "custom_time_range": ("21:00", "04:30"),
                "do_plots_summary": True,
                "do_plots_pernight": False,
                "generate_html": False,
                "open_files": False,
            })

            line_dir = custom_diagram_dir / "bat_data" / "summary" / "line"
            self.assertTrue((line_dir / "Nyctalus noctula.png").exists())

    def test_16_build_manual_interval_sequence_overnight(self) -> None:
        seq = build_manual_interval_sequence(("21:00", "04:30"))
        self.assertEqual(seq[0], "21:00")
        self.assertEqual(seq[-1], "04:30")
        self.assertIn("23:45", seq)
        self.assertIn("00:00", seq)
        self.assertIn("00:15", seq)

        # Order check: 23:45 comes before 00:00, 00:00 comes before 04:30
        idx_2345 = seq.index("23:45")
        idx_0000 = seq.index("00:00")
        idx_0430 = seq.index("04:30")
        self.assertTrue(idx_2345 < idx_0000 < idx_0430)

        # Same-day range
        seq_sameday = build_manual_interval_sequence(("18:00", "23:00"))
        self.assertEqual(seq_sameday[0], "18:00")
        self.assertEqual(seq_sameday[-1], "23:00")
        self.assertNotIn("00:00", seq_sameday)

    def test_17_overnight_manual_aggregation_non_empty(self) -> None:
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            bat_file = temp_path / "bat_data.csv"
            # Late-evening record at 22:15 and post-midnight record at 02:30
            csv_content = (
                "DATE;TIME;MANUAL ID\n"
                "2026-07-14;22:15:00;Nyctalus noctula FOD\n"
                "2026-07-15;02:30:00;Nyctalus noctula SOC\n"
            )
            bat_file.write_text(csv_content, encoding="utf-8")

            run_analysis({
                "input_files": [str(bat_file)],
                "base_dir": str(temp_path),
                "base_name": "test_overnight_manual",
                "custom_time_range": ("21:00", "04:30"),
                "do_plots_summary": True,
                "do_plots_pernight": True,
                "generate_html": False,
                "open_files": False,
            })

            line_dir = temp_path / "results" / "bat_data" / "summary" / "line"
            self.assertTrue((line_dir / "Nyctalus noctula.png").is_file())
            self.assertTrue((line_dir / "alla_arter.png").is_file())

    def test_18_evolution_7_8_9_contract(self) -> None:
        with TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            bat_file = temp_path / "bat_data.csv"
            csv_content = (
                "DATE;TIME;MANUAL ID\n"
                "2026-07-14;22:15:00;Nyctalus noctula FOD\n"
                "2026-07-14;22:30:00;Nyctalus noctula SOF\n"
                "2026-07-15;02:30:00;Myotis daubentonii SOC\n"
            )
            bat_file.write_text(csv_content, encoding="utf-8")

            # Run with ymax_mode = zoomed and dataset containing SOF
            run_analysis({
                "input_files": [str(bat_file)],
                "base_dir": str(temp_path),
                "base_name": "test_ev789",
                "custom_time_range": ("21:00", "04:30"),
                "ymax_mode": "zoomed",
                "do_plots_summary": True,
                "do_plots_pernight": True,
                "generate_html": False,
                "open_files": False,
            })

            stem_dir = temp_path / "results" / "bat_data"
            summary_art = stem_dir / "summary" / "stacked_ART"
            summary_nvi = stem_dir / "summary" / "stacked_NVI"

            # Check new all-species stacked grouped charts (#9)
            self.assertTrue((summary_art / "alla_arter.png").is_file())
            self.assertTrue((summary_nvi / "alla_arter.png").is_file())

            # Check per-species stacked charts remain present
            self.assertTrue((summary_art / "Nyctalus noctula.png").is_file())
            self.assertTrue((summary_art / "Myotis daubentonii.png").is_file())

            # Check per-night all-species stacked chart
            nights_dir = stem_dir / "nights"
            for nd in nights_dir.iterdir():
                if nd.is_dir():
                    self.assertTrue((nd / "stacked_ART" / "alla_arter.png").is_file())
                    self.assertTrue((nd / "stacked_NVI" / "alla_arter.png").is_file())



