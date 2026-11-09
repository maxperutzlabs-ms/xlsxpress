from __future__ import annotations

import datetime as dt
import os
import shutil
import time
import warnings
from collections.abc import Callable
from pathlib import Path
from typing import Any

import polars as pl
import pytest

from xlsxpress.core import excel, paths, writers
from xlsxpress.core.cleanup import CleanupPolicy
from xlsxpress.core.errors import (
    ConversionError,
    ExcelLimitError,
    OpenOutputError,
    ReadError,
    UnsupportedFormatError,
    WriteError,
)
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


class TestRunConversionWorkFolder:
    @pytest.mark.parametrize("to_temp", [False, True])
    def test_run_conversion_creates_the_output_name_only_after_the_workbook_is_complete(
        self, tsv_file: Path, monkeypatch: pytest.MonkeyPatch, to_temp: bool
    ) -> None:
        output_dir = paths.get_temp_dir() if to_temp else tsv_file.parent
        original_write_excel = writers.write_excel
        files_while_writing: list[list[Path]] = []

        def write_excel(*args: Any, **kwargs: Any) -> Path:
            files_while_writing.append(_xlsx_files(output_dir))
            return original_write_excel(*args, **kwargs)

        monkeypatch.setattr(writers, "write_excel", write_excel)

        result = run_conversion(ConversionOptions(tsv_file, to_temp=to_temp))

        assert files_while_writing == [[]]
        assert _xlsx_files(output_dir) == [result.output_path]
        assert list(paths.get_work_dir().iterdir()) == []

    def test_run_conversion_deletes_work_files_older_than_one_day(self, tsv_file: Path) -> None:
        work_dir = paths.get_work_dir()
        stale, fresh = work_dir / "tmpstale.xlsx", work_dir / "tmpfresh.xlsx"
        for path in (stale, fresh):
            path.write_bytes(b"unfinished")
        two_days_ago_ns = time.time_ns() - 2 * 86_400 * 10**9
        os.utime(stale, ns=(two_days_ago_ns, two_days_ago_ns))

        run_conversion(ConversionOptions(tsv_file))

        assert list(work_dir.iterdir()) == [fresh]

    @pytest.mark.parametrize(("to_temp", "patched"), [(True, (os, "replace")), (False, (shutil, "copyfile"))])
    def test_run_conversion_raises_write_error_and_cleans_up_when_moving_or_copying_fails(
        self, tsv_file: Path, monkeypatch: pytest.MonkeyPatch, to_temp: bool, patched: tuple[Any, str]
    ) -> None:
        def fail(source: Path, target: Path) -> None:
            Path(target).write_bytes(b"partial")
            raise OSError(28, "No space left on device")

        monkeypatch.setattr(*patched, fail)
        output_dir = paths.get_temp_dir() if to_temp else tsv_file.parent

        with pytest.raises(WriteError, match=r"Could not write .*data.*\.xlsx.*No space left on device"):
            run_conversion(ConversionOptions(tsv_file, to_temp=to_temp))

        assert _xlsx_files(output_dir) == []
        assert list(paths.get_work_dir().iterdir()) == []


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
        assert list(paths.get_work_dir().iterdir()) == []

    def test_run_conversion_raises_read_error_for_a_missing_input_file(self, tmp_path: Path) -> None:
        with pytest.raises(ReadError, match="does not exist"):
            run_conversion(ConversionOptions(tmp_path / "missing.tsv"))


class TestRunConversionOpen:
    def test_run_conversion_opens_the_written_file_with_open_after(
        self, tsv_file: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        opened: list[Path] = []
        monkeypatch.setattr(excel, "open_in_default_app", opened.append)

        result = run_conversion(ConversionOptions(tsv_file, open_after=True))

        assert opened == [result.output_path]

    def test_run_conversion_raises_open_output_error_but_keeps_the_file_when_opening_fails(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text("a\ta\n1\t2\n", encoding="utf-8")

        def fail_to_open(path: Path) -> None:
            raise OSError(1155, "No application is associated with the specified file for this operation")

        monkeypatch.setattr(excel, "open_in_default_app", fail_to_open)

        with pytest.raises(OpenOutputError) as excinfo:
            run_conversion(ConversionOptions(input_path, open_after=True))

        error = excinfo.value
        assert not isinstance(error, ConversionError)
        assert error.output_path == tmp_path / "data.xlsx"
        assert error.output_path.stat().st_size > 0
        assert len(error.warnings) == 1
        assert "appear more than once" in error.warnings[0]
        assert "could not be opened" in str(error)
        assert "No application is associated" in str(error)
        assert "no program is set up to open .xlsx files" in str(error)
