from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase
import pandas as pd

from algorithms.merge_summary_and_nightly_charts import run_analysis


class HeadlessMergeTests(TestCase):
    def test_run_analysis_headless(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            df1 = pd.DataFrame(
                {
                    "DATE": ["2026-07-14", "2026-07-15"],
                    "TIME": ["22:15:00", "01:30:00"],
                    "MANUAL ID": ["Pipistrellus pygmaeus FOD", "Myotis daubentonii SOC"],
                }
            )
            df2 = pd.DataFrame(
                {
                    "DATE": ["2026-07-14", "2026-07-15"],
                    "TIME": ["23:00:00", "02:15:00"],
                    "MANUAL ID": ["Pipistrellus pygmaeus", "Myotis daubentonii FOD"],
                }
            )
            path1 = directory / "site_a.csv"
            path2 = directory / "site_b.csv"
            df1.to_csv(path1, index=False, sep=";", encoding="utf-8")
            df2.to_csv(path2, index=False, sep=";", encoding="utf-8")

            out_dir = directory / "out"
            settings = {
                "input_files": [str(path1), str(path2)],
                "base_dir": str(out_dir),
                "base_name": "test_output",
                "custom_time_range": None,
                "diagram_base": None,
                "do_plots_summary": True,
                "do_plots_pernight": True,
                "colors": {},
                "open_files": False,
            }

            run_analysis(settings)

            results_dir = out_dir / "results" / "combined"
            self.assertTrue(results_dir.exists())
            self.assertTrue((results_dir / "test_output_NVI.xlsx").exists())
            self.assertTrue((results_dir / "test_output_ART.xlsx").exists())
            self.assertTrue((out_dir / "results" / "inputs").exists())
