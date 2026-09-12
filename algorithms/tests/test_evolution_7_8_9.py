"""Focused synthetic regression tests for LaczTabeleNietoperzy #7, #8, #9."""

import os
import tempfile
import pandas as pd
import pytest
from pathlib import Path

from algorithms.merge_summary_and_nightly_charts import (
    extract_species_and_type,
    resolve_dataset_class_config,
    compute_ymax_with_headroom,
    run_analysis,
)


def create_synthetic_xlsx(file_path: str, rows: list[dict]):
    df = pd.DataFrame(rows)
    df.to_excel(file_path, index=False)


def test_issue7_ymax_headroom_calculation():
    assert compute_ymax_with_headroom(0) == 1
    assert compute_ymax_with_headroom(-5) == 1
    assert compute_ymax_with_headroom(1) == 2  # 1 * 1.10 = 1.1 -> ceil(1.1) = 2
    assert compute_ymax_with_headroom(5) == 6  # 5 * 1.10 = 5.5 -> ceil(5.5) = 6
    assert compute_ymax_with_headroom(10) == 11  # 10 * 1.10 = 11.0 -> 11
    assert compute_ymax_with_headroom(20) == 22


def test_issue8_sof_parsing_and_extraction():
    assert extract_species_and_type("NYCNOC SOC") == [("NYCNOC", "SOC")]
    assert extract_species_and_type("NYCNOC SOF") == [("NYCNOC", "SOF")]
    assert extract_species_and_type("NYCNOC FOD") == [("NYCNOC", "FOD")]
    assert extract_species_and_type("NYCNOC FORBI") == [("NYCNOC", "FORBI")]
    assert extract_species_and_type("NYCNOC socialt - flyg") == [("NYCNOC", "SOF")]
    assert extract_species_and_type("NYCNOC socialt - läte") == [("NYCNOC", "SOC")]


def test_issue8_dataset_without_sof_color_fallback():
    rows = [
        {"DATE": "2026-06-01", "TIME": "22:00", "MANUAL ID": "NYCNOC SOC"},
        {"DATE": "2026-06-01", "TIME": "22:15", "MANUAL ID": "PLEUAR FOD"},
    ]
    df = pd.DataFrame(rows)
    cfg = resolve_dataset_class_config([df])

    assert cfg["has_sof"] is False
    assert cfg["type_order"] == ["Socialt", "Födosökande", "Förbiflygande"]
    assert cfg["colors_nvi"][0] == "#FF0000"  # original SOC NVI color
    assert cfg["colors_art"][0] == "#EB09D8"  # original SOC ART color


def test_issue8_dataset_with_sof_color_rules():
    rows = [
        {"DATE": "2026-06-01", "TIME": "22:00", "MANUAL ID": "NYCNOC SOC"},
        {"DATE": "2026-06-01", "TIME": "22:15", "MANUAL ID": "PLEUAR SOF"},
    ]
    df = pd.DataFrame(rows)
    cfg = resolve_dataset_class_config([df])

    assert cfg["has_sof"] is True
    assert cfg["type_order"] == ["Socialt - läte", "Socialt - flyg", "Födosökande", "Förbiflygande"]
    assert cfg["code_map"]["SOC"] == "Socialt - läte"
    assert cfg["code_map"]["SOF"] == "Socialt - flyg"
    # SOC takes new darker defaults
    assert cfg["art_color_dict"]["Socialt - läte"] == "#8D0B82"
    assert cfg["nvi_color_dict"]["Socialt - läte"] == "#9A0B08"
    # SOF takes original SOC colors
    assert cfg["art_color_dict"]["Socialt - flyg"] == "#EB09D8"
    assert cfg["nvi_color_dict"]["Socialt - flyg"] == "#FF0000"


def test_issues_7_8_9_headless_full_execution_flow():
    with tempfile.TemporaryDirectory() as tmp_dir:
        input_path = os.path.join(tmp_dir, "test_bats.xlsx")
        out_base = os.path.join(tmp_dir, "output")

        rows = [
            {"DATE": "2026-06-01", "TIME": "22:00", "MANUAL ID": "NYCNOC SOC"},
            {"DATE": "2026-06-01", "TIME": "22:15", "MANUAL ID": "NYCNOC SOF"},
            {"DATE": "2026-06-01", "TIME": "22:30", "MANUAL ID": "PLEUAR FOD"},
            {"DATE": "2026-06-02", "TIME": "23:00", "MANUAL ID": "VESMUR FORBI"},
        ]
        create_synthetic_xlsx(input_path, rows)

        settings = {
            "input_files": [input_path],
            "base_dir": out_base,
            "base_name": "test_res",
            "custom_time_range": ("21:00", "04:30"),
            "ymax_mode": "zoomed",
            "do_plots_summary": True,
            "do_plots_pernight": True,
            "generate_html": False,
            "colors": {},
            "open_files": False,
        }

        run_analysis(settings)

        res_dir = Path(out_base) / "results"
        assert (res_dir / "combined" / "test_res_NVI.xlsx").exists()
        assert (res_dir / "combined" / "test_res_ART.xlsx").exists()

        stem_dir = res_dir / "test_bats"
        # Summary charts check (#9 alla_arter.png in summary)
        assert (stem_dir / "summary" / "line" / "alla_arter.png").exists()
        assert (stem_dir / "summary" / "stacked_ART" / "alla_arter.png").exists()
        assert (stem_dir / "summary" / "stacked_NVI" / "alla_arter.png").exists()

        # Per-species summary charts check
        assert (stem_dir / "summary" / "stacked_ART" / "NYCNOC.png").exists()
        assert (stem_dir / "summary" / "stacked_ART" / "PLEUAR.png").exists()
        assert (stem_dir / "summary" / "stacked_ART" / "VESMUR.png").exists()

        # Per-night charts check (#9 alla_arter.png per night)
        nights_dir = stem_dir / "nights"
        night_subdirs = [d for d in nights_dir.iterdir() if d.is_dir()]
        assert len(night_subdirs) == 2  # 2026-06-01 and 2026-06-02

        for nd in night_subdirs:
            assert (nd / "line" / "alla_arter.png").exists()
            assert (nd / "stacked_ART" / "alla_arter.png").exists()
            assert (nd / "stacked_NVI" / "alla_arter.png").exists()
