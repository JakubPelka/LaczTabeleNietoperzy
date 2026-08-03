from pathlib import Path
from tempfile import TemporaryDirectory
from unittest import TestCase

import pandas as pd

from algorithms.input_reader import read_input_table


class InputReaderTests(TestCase):
    def test_reads_xlsx_and_semicolon_csv_with_the_same_columns(self) -> None:
        with TemporaryDirectory() as temp_dir:
            directory = Path(temp_dir)
            expected = pd.DataFrame(
                {
                    "DATE": ["2026-07-14"],
                    "TIME": ["01:02:03"],
                    "MANUAL ID": ["Nyctalus noctula FOD"],
                }
            )
            xlsx_path = directory / "source.xlsx"
            csv_path = directory / "source.csv"
            expected.to_excel(xlsx_path, index=False)
            expected.to_csv(csv_path, index=False, sep=";", encoding="utf-8-sig")

            xlsx_result = read_input_table(xlsx_path)
            csv_result = read_input_table(csv_path)

            self.assertEqual(list(xlsx_result.columns), list(expected.columns))
            self.assertEqual(list(csv_result.columns), list(expected.columns))
            self.assertEqual(csv_result.loc[0, "MANUAL ID"], "Nyctalus noctula FOD")

    def test_reads_windows_encoded_csv(self) -> None:
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "source.csv"
            path.write_bytes(
                "DATE;TIME;MANUAL ID\n2026-07-14;01:02:03;Plecotus auritus FÖD\n".encode(
                    "cp1252"
                )
            )

            result = read_input_table(path)

            self.assertEqual(result.loc[0, "MANUAL ID"], "Plecotus auritus FÖD")

    def test_rejects_unsupported_format(self) -> None:
        with TemporaryDirectory() as temp_dir:
            path = Path(temp_dir) / "source.txt"
            path.write_text("DATE,TIME,MANUAL ID\n", encoding="utf-8")
            with self.assertRaises(ValueError):
                read_input_table(path)
