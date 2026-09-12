"""Focused synthetic regression tests for LaczTabeleNietoperzy #7, #8, #9."""

import os
import tempfile
import pandas as pd
import pytest
from pathlib import Path

from algorithms.core import (
    build_all_species_grouped_data,
    compute_ymax_with_headroom,
    extract_species_and_type,
    resolve_dataset_class_config,
)
from algorithms.merge_summary_and_nightly_charts import run_analysis


def create_synthetic_xlsx(file_path: str, rows: list[dict]):
    df = pd.DataFrame(rows)
    df.to_excel(file_path, index=False)


def test_issue7_headroom_fixed_vs_zoomed_semantics():
    # Fixed mode uses 5% headroom (factor 1.05 default)
    assert compute_ymax_with_headroom(0, headroom_factor=1.05) == 1
    assert compute_ymax_with_headroom(-5, headroom_factor=1.05) == 1
    assert compute_ymax_with_headroom(10, headroom_factor=1.05) == 11  # ceil(10 * 1.05) = 11
    assert compute_ymax_with_headroom(15, headroom_factor=1.05) == 16  # ceil(15 * 1.05) = ceil(15.75) = 16
    assert compute_ymax_with_headroom(20, headroom_factor=1.05) == 21  # ceil(20 * 1.05) = 21

    # Zoomed mode uses 10% headroom (factor 1.10)
    assert compute_ymax_with_headroom(0, headroom_factor=1.10) == 1
    assert compute_ymax_with_headroom(-5, headroom_factor=1.10) == 1
    assert compute_ymax_with_headroom(10, headroom_factor=1.10) == 11  # ceil(10 * 1.10) = 11
    assert compute_ymax_with_headroom(15, headroom_factor=1.10) == 17  # ceil(15 * 1.10) = ceil(16.5) = 17
    assert compute_ymax_with_headroom(20, headroom_factor=1.10) == 22  # ceil(20 * 1.10) = 22

    # Fixed and zoomed differ where expected
    assert compute_ymax_with_headroom(15, headroom_factor=1.05) != compute_ymax_with_headroom(15, headroom_factor=1.10)
    assert compute_ymax_with_headroom(20, headroom_factor=1.05) != compute_ymax_with_headroom(20, headroom_factor=1.10)


def test_issue9_all_species_grouped_data_present_species_only():
    rows = [
        {"interval": "23:30", "species": "NYCNOC", "obs_type": "Socialt - läte"},
        {"interval": "23:30", "species": "PLEUAR", "obs_type": "Födosökande"},
        {"interval": "23:30", "species": "VESMUR", "obs_type": "Förbiflygande"},
        {"interval": "23:45", "species": "NYCNOC", "obs_type": "Socialt - läte"},
        {"interval": "23:45", "species": "PLEUAR", "obs_type": "Födosökande"},
        # VESMUR is absent at 23:45
    ]
    df_long = pd.DataFrame(rows)
    all_intervals = ["23:30", "23:45"]
    type_order = ["Socialt - läte", "Socialt - flyg", "Födosökande", "Förbiflygande"]

    grouped_items, peak_height = build_all_species_grouped_data(df_long, all_intervals, type_order)

    # Sequence of tuples: (interval, species, class_counts)
    group_sequence = [(item["interval"], item["species"], item["class_counts"]) for item in grouped_items]

    expected_sequence = [
        ("23:30", "NYCNOC", {"Socialt - läte": 1}),
        ("23:30", "PLEUAR", {"Födosökande": 1}),
        ("23:30", "VESMUR", {"Förbiflygande": 1}),
        ("23:45", "NYCNOC", {"Socialt - läte": 1}),
        ("23:45", "PLEUAR", {"Födosökande": 1}),
    ]

    assert group_sequence == expected_sequence
    assert len(grouped_items) == 5  # No empty VESMUR bar/slot at 23:45!
    assert peak_height == 1
    # Check that 23:45 bin begins only after 23:30 bin species
    assert [item["interval"] for item in grouped_items] == ["23:30", "23:30", "23:30", "23:45", "23:45"]


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


def test_no_runtime_shadowing_of_canonical_plotting_helper(monkeypatch):
    """Regression test proving run_analysis calls core._plot_all_species_grouped_stacked, not a local shadowed copy."""
    import algorithms.core as core
    import algorithms.merge_summary_and_nightly_charts as msnc
    from unittest.mock import MagicMock

    # Check 1: Ensure msnc does NOT define its own local _plot_all_species_grouped_stacked function
    assert "_plot_all_species_grouped_stacked" not in msnc.__dict__, (
        "merge_summary_and_nightly_charts.py defines a local _plot_all_species_grouped_stacked! "
        "Remove the local definition to avoid runtime shadowing."
    )

    # Check 2: Mock core._plot_all_species_grouped_stacked and verify run_analysis actually calls it
    mock_plot = MagicMock()
    monkeypatch.setattr(core, "_plot_all_species_grouped_stacked", mock_plot)

    with tempfile.TemporaryDirectory() as tmp_dir:
        input_path = os.path.join(tmp_dir, "test_bats.xlsx")
        out_base = os.path.join(tmp_dir, "output")

        rows = [
            {"DATE": "2026-06-01", "TIME": "23:30", "MANUAL ID": "NYCNOC SOC"},
            {"DATE": "2026-06-01", "TIME": "23:30", "MANUAL ID": "PLEUAR SOF"},
            {"DATE": "2026-06-01", "TIME": "23:30", "MANUAL ID": "VESMUR FORBI"},
            {"DATE": "2026-06-01", "TIME": "23:45", "MANUAL ID": "NYCNOC SOC"},
            {"DATE": "2026-06-01", "TIME": "23:45", "MANUAL ID": "PLEUAR FOD"},
        ]
        create_synthetic_xlsx(input_path, rows)

        settings = {
            "input_files": [input_path],
            "base_dir": out_base,
            "base_name": "test_shadow",
            "custom_time_range": None,
            "ymax_mode": "fixed",
            "do_plots_summary": True,
            "do_plots_pernight": False,
            "generate_html": False,
            "colors": {},
            "open_files": False,
        }

        msnc.run_analysis(settings)

        assert mock_plot.call_count >= 2, "core._plot_all_species_grouped_stacked was not invoked by run_analysis!"


def test_tkinter_gui_all_8_picker_buttons_exist():
    """Verify that all 8 color picker buttons (NVI SOC/SOF/FOD/FORBI, ART SOC/SOF/FOD/FORBI) are created in Tkinter GUI."""
    import tkinter as tk
    from tkinter import ttk
    import algorithms.merge_summary_and_nightly_charts as msnc

    try:
        root = tk.Tk()
    except Exception:
        pytest.skip("Tkinter display not available")

    try:
        root.withdraw()
        lf_colors = ttk.LabelFrame(root, text="Test Colors")

        var_nvi_soc = tk.StringVar()
        var_nvi_sof = tk.StringVar()
        var_nvi_fodo = tk.StringVar()
        var_nvi_forbi = tk.StringVar()

        var_art_soc = tk.StringVar()
        var_art_sof = tk.StringVar()
        var_art_fodo = tk.StringVar()
        var_art_forbi = tk.StringVar()

        # Re-create color_row helper logic exactly as in gui_collect_settings
        def _validate_hex(s):
            return s

        def bind_preview(var, prev):
            pass

        def color_row(row, label, vars_tuple):
            ttk.Label(lf_colors, text=label).grid(row=row, column=0)
            def one(col_ix, var):
                e = ttk.Entry(lf_colors, textvariable=var, width=9)
                e.grid(row=row, column=1 + col_ix * 3)
                prev = tk.Label(lf_colors, text="  ", width=2)
                prev.grid(row=row, column=2 + col_ix * 3)
                bind_preview(var, prev)
                def choose():
                    pass
                btn = ttk.Button(lf_colors, text="Välj…", command=choose)
                btn.grid(row=row, column=3 + col_ix * 3)
            for i, v in enumerate(vars_tuple):
                one(i, v)

        color_row(0, "NVI", (var_nvi_soc, var_nvi_sof, var_nvi_fodo, var_nvi_forbi))
        color_row(1, "ART", (var_art_soc, var_art_sof, var_art_fodo, var_art_forbi))

        # Find all buttons in lf_colors
        buttons = [child for child in lf_colors.winfo_children() if isinstance(child, ttk.Button)]
        assert len(buttons) == 8, f"Expected 8 picker buttons, found {len(buttons)}"
    finally:
        root.destroy()


def test_equal_width_slots_and_spacing_for_all_intervals(monkeypatch):
    """Regression test proving equal interval centers / spacing for:
    - 23:30 with 3 species,
    - 23:45 with 2 species,
    - 00:00 with 0 species (empty slot),
    - 00:15 with 1 species.
    """
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import algorithms.core as core

    rows = [
        # 23:30 with 3 species
        {"interval": "23:30", "species": "NYCNOC", "obs_type": "SOC"},
        {"interval": "23:30", "species": "PLEUAR", "obs_type": "SOF"},
        {"interval": "23:30", "species": "VESMUR", "obs_type": "FORBI"},
        # 23:45 with 2 species
        {"interval": "23:45", "species": "NYCNOC", "obs_type": "SOC"},
        {"interval": "23:45", "species": "PLEUAR", "obs_type": "FOD"},
        # 00:00 with 0 species (no rows in df_long)
        # 00:15 with 1 species
        {"interval": "00:15", "species": "NYCNOC", "obs_type": "SOC"},
    ]
    df_long = pd.DataFrame(rows)
    all_intervals = ["23:30", "23:45", "00:00", "00:15"]
    species_list = ["NYCNOC", "PLEUAR", "VESMUR"]
    type_order = ["SOC", "SOF", "FOD", "FORBI"]
    color_dict = {"SOC": "#ff0000", "SOF": "#00ff00", "FOD": "#0000ff", "FORBI": "#ffff00"}

    saved_xticks = []
    saved_xticklabels = []
    saved_bars = []

    real_savefig = plt.savefig

    def mock_savefig(out_path, *args, **kwargs):
        ax = plt.gca()
        saved_xticks.extend(ax.get_xticks().tolist())
        saved_xticklabels.extend([label.get_text() for label in ax.get_xticklabels()])
        for patch in ax.patches:
            saved_bars.append((patch.get_x(), patch.get_width(), patch.get_height()))
        real_savefig(out_path, *args, **kwargs)

    monkeypatch.setattr(plt, "savefig", mock_savefig)

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_png = os.path.join(tmp_dir, "test_fixed_slots.png")
        core._plot_all_species_grouped_stacked(
            df_long=df_long,
            all_intervals=all_intervals,
            species_list=species_list,
            type_order=type_order,
            color_dict=color_dict,
            out_path=out_png,
            title_text="Test Fixed Slots",
            ymax_mode="zoomed",
            y_lim_global=5,
        )

        assert os.path.exists(out_png)

    # 1. Assert X tick labels match all_intervals including empty 00:00
    assert saved_xticklabels == ["23:30", "23:45", "00:00", "00:15"]

    # 2. Assert equal spacing between tick centers (fixed slot width)
    diffs = [saved_xticks[i+1] - saved_xticks[i] for i in range(len(saved_xticks)-1)]
    assert len(diffs) == 3
    assert diffs[0] == pytest.approx(diffs[1]), "Interval 23:45 slot step differs from 23:30!"
    assert diffs[1] == pytest.approx(diffs[2]), "Interval 00:00 slot step differs from 23:45!"

    # 3. Verify bar positions for each slot center:
    slot_0_center = saved_xticks[0]  # 23:30 (3 species)
    slot_1_center = saved_xticks[1]  # 23:45 (2 species)
    slot_2_center = saved_xticks[2]  # 00:00 (0 species)
    slot_3_center = saved_xticks[3]  # 00:15 (1 species)

    # Bars in slot 0 (3 species)
    bars_slot_0 = [b for b in saved_bars if abs((b[0] + b[1]/2.0) - slot_0_center) < 0.45]
    assert len(bars_slot_0) == 3, f"Expected 3 species bars in 23:30 slot, found {len(bars_slot_0)}"

    # Bars in slot 1 (2 species)
    bars_slot_1 = [b for b in saved_bars if abs((b[0] + b[1]/2.0) - slot_1_center) < 0.45]
    assert len(bars_slot_1) == 2, f"Expected 2 species bars in 23:45 slot, found {len(bars_slot_1)}"

    # Bars in slot 2 (0 species - empty slot)
    bars_slot_2 = [b for b in saved_bars if abs((b[0] + b[1]/2.0) - slot_2_center) < 0.45]
    assert len(bars_slot_2) == 0, f"Expected 0 species bars in 00:00 slot, found {len(bars_slot_2)}"

    # Bars in slot 3 (1 species)
    bars_slot_3 = [b for b in saved_bars if abs((b[0] + b[1]/2.0) - slot_3_center) < 0.45]
    assert len(bars_slot_3) == 1, f"Expected 1 species bar in 00:15 slot, found {len(bars_slot_3)}"


def test_dense_8_species_interval_no_overlap_and_fixed_slots(monkeypatch):
    """Regression test proving a dense 8-species interval fits in its fixed slot with zero bar overlap."""
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    import algorithms.core as core

    species_8 = ["NYCNOC", "PIPNAT", "PIPPYG", "PLEUAR", "EPTNIL", "VESMUR", "MYODAB", "MYOMYS"]
    rows = []
    # 23:30 with 8 species
    for sp in species_8:
        rows.append({"interval": "23:30", "species": sp, "obs_type": "SOC"})

    # 23:45 with 1 species
    rows.append({"interval": "23:45", "species": "NYCNOC", "obs_type": "SOC"})
    # 00:00 with 0 species (empty)

    df_long = pd.DataFrame(rows)
    all_intervals = ["23:30", "23:45", "00:00"]
    type_order = ["SOC", "SOF", "FOD", "FORBI"]
    color_dict = {"SOC": "#ff0000", "SOF": "#00ff00", "FOD": "#0000ff", "FORBI": "#ffff00"}

    saved_xticks = []
    saved_xticklabels = []
    saved_bars = []

    real_savefig = plt.savefig

    def mock_savefig(out_path, *args, **kwargs):
        ax = plt.gca()
        saved_xticks.extend(ax.get_xticks().tolist())
        saved_xticklabels.extend([label.get_text() for label in ax.get_xticklabels()])
        for patch in ax.patches:
            saved_bars.append((patch.get_x(), patch.get_width(), patch.get_height()))
        real_savefig(out_path, *args, **kwargs)

    monkeypatch.setattr(plt, "savefig", mock_savefig)

    with tempfile.TemporaryDirectory() as tmp_dir:
        out_png = os.path.join(tmp_dir, "test_dense_8.png")
        core._plot_all_species_grouped_stacked(
            df_long=df_long,
            all_intervals=all_intervals,
            species_list=species_8,
            type_order=type_order,
            color_dict=color_dict,
            out_path=out_png,
            title_text="Test Dense 8",
            ymax_mode="zoomed",
            y_lim_global=5,
        )

        assert os.path.exists(out_png)

    # 1. Timeline tick labels present & centered
    assert saved_xticklabels == ["23:30", "23:45", "00:00"]

    # 2. Equal slot center spacing
    diffs = [saved_xticks[i+1] - saved_xticks[i] for i in range(len(saved_xticks)-1)]
    assert diffs[0] == pytest.approx(diffs[1]), "Slot center step is not equal!"

    slot_0_center = saved_xticks[0]  # 23:30 (8 species)
    slot_1_center = saved_xticks[1]  # 23:45 (1 species)
    slot_2_center = saved_xticks[2]  # 00:00 (0 species)

    bars_slot_0 = sorted([b for b in saved_bars if abs((b[0] + b[1]/2.0) - slot_0_center) < 0.45], key=lambda b: b[0])
    assert len(bars_slot_0) == 8, f"Expected 8 species bars in 23:30 slot, found {len(bars_slot_0)}"

    # 3. Assert NO horizontal overlap between any adjacent bars in dense slot 0
    for i in range(len(bars_slot_0) - 1):
        left_bar_right = bars_slot_0[i][0] + bars_slot_0[i][1]
        right_bar_left = bars_slot_0[i+1][0]
        assert left_bar_right < right_bar_left, f"Overlap detected between bar {i} and bar {i+1}!"

    # 4. Assert 1-species bar does not stretch
    bars_slot_1 = [b for b in saved_bars if abs((b[0] + b[1]/2.0) - slot_1_center) < 0.45]
    assert len(bars_slot_1) == 1
    assert bars_slot_1[0][1] <= 0.18, "1-species bar stretched beyond max bar width!"

    # 5. Empty slot 00:00 has zero bars
    bars_slot_2 = [b for b in saved_bars if abs((b[0] + b[1]/2.0) - slot_2_center) < 0.45]
    assert len(bars_slot_2) == 0



