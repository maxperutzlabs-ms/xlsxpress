from __future__ import annotations

import datetime as dt
import os
import time
import warnings
from collections.abc import Callable
from pathlib import Path

import polars as pl
import pytest

from xlsxpress.core import paths, writers
from xlsxpress.core.cleanup import CleanupPolicy
from xlsxpress.core.errors import ExcelLimitError, ReadError, UnsupportedFormatError
from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import run_conversion

pytestmark = pytest.mark.integration

_TSV = (
    "protein\tgene\tid\tintensity\tcount\tscore\tnote\n"
    "P12345\tSEPT2\t007\t1.10\t3\tNaN\t=1+1\n"
    "Q99999\tMARCH1\t2310009E13\t1e+05\t-12\t0.5\t\n"
    "O00001\t1-Mar\t000\t-inf\t0\tNA\tTRUE\n"
)

_EXPECTED_CELLS = {
    "A1": ("text", "protein"),
    "G1": ("text", "note"),
    "A2": ("text", "P12345"),
    "B2": ("text", "SEPT2"),
    "B3": ("text", "MARCH1"),
    "B4": ("text", "1-Mar"),
    "C2": ("text", "007"),
    "C3": ("text", "2310009E13"),
    "C4": ("text", "000"),
    "D2": ("number", "1.1"),
    "D3": ("number", "100000"),
    "D4": ("text", "-inf"),
    "E2": ("number", "3"),
    "E3": ("number", "-12"),
    "E4": ("number", "0"),
    "F3": ("number", "0.5"),
    "G2": ("text", "=1+1"),
    "G4": ("text", "TRUE"),
}


@pytest.fixture
def tsv_file(tmp_path: Path) -> Path:
    path = tmp_path / "input" / "data.tsv"
    path.parent.mkdir()
    path.write_text(_TSV, encoding="utf-8")
    return path


def _xlsx_files(folder: Path) -> list[Path]:
    return sorted(folder.glob("*.xlsx"))


class TestRunConversionValues:
    def test_run_conversion_writes_text_file_values_unchanged_cell_by_cell(
        self, tsv_file: Path, xlsx_reader: Callable
    ) -> None:
        result = run_conversion(ConversionOptions(tsv_file))

        sheet = xlsx_reader(result.output_path)
        assert {ref: (sheet.cells[ref].kind, sheet.cells[ref].value) for ref in _EXPECTED_CELLS} == _EXPECTED_CELLS
        assert {"F2", "F4", "G3"}.isdisjoint(sheet.cells)
        assert result.warnings == ()

    def test_run_conversion_writes_parquet_types_cell_by_cell(self, tmp_path: Path, xlsx_reader: Callable) -> None:
        input_path = tmp_path / "data.parquet"
        pl.DataFrame(
            {
                "id": ["007"],
                "value": [0.1],
                "flag": [True],
                "day": [dt.date(2024, 1, 2)],
                "zoned": [dt.datetime(2024, 1, 2, 10, tzinfo=dt.UTC)],
            }
        ).write_parquet(input_path)

        sheet = xlsx_reader(run_conversion(ConversionOptions(input_path)).output_path)

        assert [(sheet.cells[ref].kind, sheet.cells[ref].value) for ref in ("A2", "B2", "C2", "E2")] == [
            ("text", "007"),
            ("number", "0.1"),
            ("boolean", "1"),
            ("text", "2024-01-02T10:00:00+00:00"),
        ]
        assert (sheet.cells["D2"].kind, sheet.cells["D2"].value, sheet.cells["D2"].number_format) == (
            "number",
            "45293",
            "yyyy-mm-dd",
        )

    def test_run_conversion_limits_rows_in_preview_mode(self, tsv_file: Path, xlsx_reader: Callable) -> None:
        sheet = xlsx_reader(run_conversion(ConversionOptions(tsv_file, nrows=1)).output_path)

        assert "A2" in sheet.cells
        assert "A3" not in sheet.cells

    def test_run_conversion_skips_header_formatting_in_plain_mode(self, tsv_file: Path, xlsx_reader: Callable) -> None:
        sheet = xlsx_reader(run_conversion(ConversionOptions(tsv_file, format_header=False)).output_path)

        assert not sheet.cells["A1"].bold
        assert sheet.autofilter is None


class TestRunConversionOutputLocation:
    def test_run_conversion_writes_next_to_the_input_without_overwriting(self, tsv_file: Path) -> None:
        existing = tsv_file.with_suffix(".xlsx")
        existing.write_bytes(b"existing workbook")

        result = run_conversion(ConversionOptions(tsv_file))

        assert result.output_path == tsv_file.parent / "data-02.xlsx"
        assert existing.read_bytes() == b"existing workbook"

    def test_run_conversion_writes_unique_files_to_the_shared_temp_folder(self, tsv_file: Path) -> None:
        first = run_conversion(ConversionOptions(tsv_file, to_temp=True)).output_path
        second = run_conversion(ConversionOptions(tsv_file, to_temp=True)).output_path

        assert first.parent == second.parent == paths.get_temp_dir()
        assert first.name.startswith("data_20")
        assert first != second
        assert _xlsx_files(tsv_file.parent) == []

    def test_run_conversion_cleans_old_temp_files_before_writing(self, tsv_file: Path) -> None:
        temp_dir = paths.get_temp_dir()
        old = run_conversion(ConversionOptions(tsv_file, to_temp=True)).output_path
        recent = run_conversion(ConversionOptions(tsv_file, to_temp=True)).output_path
        long_ago_ns = time.time_ns() - 30 * 86_400 * 10**9
        os.utime(old, ns=(long_ago_ns, long_ago_ns))

        new = run_conversion(ConversionOptions(tsv_file, to_temp=True, cleanup_policy=CleanupPolicy(max_age_days=14)))

        assert _xlsx_files(temp_dir) == sorted([recent, new.output_path])


class TestRunConversionWarnings:
    @pytest.mark.parametrize(
        ("content", "expected"),
        [("a\ta\n1\t2\n", "appear more than once"), ("a\n" + "x" * 40_000 + "\n", "were truncated")],
        ids=["repeated-names", "long-text"],
    )
    def test_run_conversion_returns_conversion_warnings_instead_of_emitting_them(
        self, tmp_path: Path, content: str, expected: str
    ) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text(content, encoding="utf-8")

        with warnings.catch_warnings():
            warnings.simplefilter("error")
            result = run_conversion(ConversionOptions(input_path))

        assert len(result.warnings) == 1
        assert expected in result.warnings[0]


class TestRunConversionErrors:
    @pytest.mark.parametrize(
        ("content", "name", "error"),
        [
            ("", "data.tsv", ReadError),
            ("a\tb\n1\t2\t3\n", "data.tsv", ReadError),
            ("a\n", "data.xlsx", UnsupportedFormatError),
            ("a\tb\tc\n1\t2\t3\n", "data.tsv", ExcelLimitError),
        ],
        ids=["empty", "extra-fields", "unsupported", "too-many-columns"],
    )
    @pytest.mark.parametrize("to_temp", [False, True])
    def test_run_conversion_leaves_no_xlsx_file_behind_on_errors(
        self,
        tmp_path: Path,
        monkeypatch: pytest.MonkeyPatch,
        content: str,
        name: str,
        error: type[Exception],
        to_temp: bool,
    ) -> None:
        monkeypatch.setattr(writers, "EXCEL_MAX_COLUMNS", 2)
        input_path = tmp_path / "input" / name
        input_path.parent.mkdir()
        input_path.write_text(content, encoding="utf-8")

        with pytest.raises(error):
            run_conversion(ConversionOptions(input_path, to_temp=to_temp))

        assert _xlsx_files(input_path.parent) == [input_path] * (name == "data.xlsx")
        assert _xlsx_files(paths.get_temp_dir()) == []

    def test_run_conversion_raises_read_error_for_a_missing_input_file(self, tmp_path: Path) -> None:
        with pytest.raises(ReadError, match="does not exist"):
            run_conversion(ConversionOptions(tmp_path / "missing.tsv"))
