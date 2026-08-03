"""Read bat registration tables from supported input formats."""

from __future__ import annotations

from pathlib import Path

import pandas as pd


def read_input_table(path: str | Path) -> pd.DataFrame:
    """Read the first XLSX sheet or a delimiter-separated CSV table."""
    input_path = Path(path)
    suffix = input_path.suffix.lower()
    if suffix == ".xlsx":
        return pd.read_excel(input_path, sheet_name=0, engine="openpyxl")
    if suffix == ".csv":
        try:
            return pd.read_csv(
                input_path,
                sep=None,
                engine="python",
                encoding="utf-8-sig",
            )
        except UnicodeDecodeError:
            # CSV files exported by older Windows/Excel installations commonly
            # use the local Windows code page instead of UTF-8.
            return pd.read_csv(
                input_path,
                sep=None,
                engine="python",
                encoding="cp1252",
            )
    raise ValueError(
        f"Formatet stöds inte: {input_path.name}. Välj en XLSX- eller CSV-fil."
    )
