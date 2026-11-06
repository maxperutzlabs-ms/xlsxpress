from __future__ import annotations

import datetime as dt
import tempfile
import warnings
from collections.abc import Callable
from decimal import Decimal
from pathlib import Path
from typing import Any

import polars as pl
import pytest
from xlsxwriter.worksheet import Worksheet

from xlsxpress.core import writers
from xlsxpress.core.errors import ConversionWarning, ExcelLimitError, WriteError
from xlsxpress.core.writers import EXCEL_MAX_COLUMNS, EXCEL_MAX_ROWS, EXCEL_MAX_STRING_LENGTH, write_excel

pytestmark = pytest.mark.unit


@pytest.fixture
def write_sheet(tmp_path: Path, xlsx_reader: Callable) -> Callable:
    def write(df: pl.DataFrame, **kwargs: bool):
        output_path = tmp_path / "output.xlsx"
        write_excel(df, output_path, **kwargs)
        return xlsx_reader(output_path)

    return write


class TestWriteExcelText:
    @pytest.mark.parametrize(
        "value",
        [
            "SEPT2",
            "MARCH1",
            "1-Mar",
            "007",
            "1.10",
            "2310009E13",
            "-5",
            "+41 123",
            "TRUE",
            "=1+1",
            "@foo",
            "https://example.org",
            " leading space",
            "trailing space ",
            "line\nbreak",
            "tab\tseparated",
            "<&>\"'",
            "_x0041_",
            "control\x01character",
            "Müller αβ 中文 😀",
        ],
    )
    def test_write_excel_writes_strings_as_unchanged_text(self, write_sheet: Callable, value: str) -> None:
        sheet = write_sheet(pl.DataFrame({"s": [value]}))

        assert sheet.cells["A2"].kind == "text"
        assert sheet.cells["A2"].value == value

    @pytest.mark.parametrize("dtype", [pl.Categorical, pl.Enum(["a", "b"])])
    def test_write_excel_writes_categorical_values_as_text(self, write_sheet: Callable, dtype: pl.DataType) -> None:
        sheet = write_sheet(pl.DataFrame({"s": pl.Series(["b", "a"], dtype=dtype)}))

        assert [(sheet.cells[ref].kind, sheet.cells[ref].value) for ref in ("A2", "A3")] == [
            ("text", "b"),
            ("text", "a"),
        ]

    def test_write_excel_leaves_null_strings_blank(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"s": ["a", None, "c"]}))

        assert "A3" not in sheet.cells
        assert sheet.cells["A4"].value == "c"

    def test_write_excel_truncates_text_over_the_excel_limit_with_warning(self, write_sheet: Callable) -> None:
        with pytest.warns(ConversionWarning, match=r"1 value\(s\) in column 's'"):
            sheet = write_sheet(pl.DataFrame({"s": ["A" * (EXCEL_MAX_STRING_LENGTH + 1)]}))

        assert sheet.cells["A2"].value == "A" * EXCEL_MAX_STRING_LENGTH

    def test_write_excel_keeps_text_at_the_excel_limit_without_warning(self, write_sheet: Callable) -> None:
        with warnings.catch_warnings():
            warnings.simplefilter("error")
            sheet = write_sheet(pl.DataFrame({"s": ["A" * EXCEL_MAX_STRING_LENGTH]}))

        assert len(sheet.cells["A2"].value) == EXCEL_MAX_STRING_LENGTH


class TestWriteExcelNumbers:
    @pytest.mark.parametrize(
        ("dtype", "value"),
        [
            (pl.Int8, -5),
            (pl.UInt8, 255),
            (pl.Int64, 0),
            (pl.Int64, 123_456_789_012_345),
            (pl.Int64, -999_999_999_999_999),
        ],
    )
    def test_write_excel_writes_integers_up_to_15_digits_as_exact_numbers(
        self, write_sheet: Callable, dtype: pl.DataType, value: int
    ) -> None:
        sheet = write_sheet(pl.DataFrame({"n": pl.Series([value], dtype=dtype)}))

        assert sheet.cells["A2"].kind == "number"
        assert int(sheet.cells["A2"].value) == value
        assert sheet.cells["A2"].number_format == "General"

    @pytest.mark.parametrize(
        ("dtype", "value"),
        [
            (pl.Int64, 1_234_567_890_123_456_789),
            (pl.UInt64, 18_446_744_073_709_551_615),
            (pl.Int128, 12_345_678_901_234_567_890),
        ],
    )
    def test_write_excel_writes_integers_over_15_digits_as_numbers_with_excel_precision(
        self, write_sheet: Callable, dtype: pl.DataType, value: int
    ) -> None:
        sheet = write_sheet(pl.DataFrame({"n": pl.Series([value], dtype=dtype)}))

        assert sheet.cells["A2"].kind == "number"
        assert float(sheet.cells["A2"].value) == pytest.approx(value, rel=1e-15)

    @pytest.mark.parametrize("value", [1.23456789, -0.000123456789, 1.5e300, 1 / 3])
    def test_write_excel_writes_floats_with_full_precision_in_general_format(
        self, write_sheet: Callable, value: float
    ) -> None:
        sheet = write_sheet(pl.DataFrame({"x": [value]}))

        assert sheet.cells["A2"].kind == "number"
        assert float(sheet.cells["A2"].value) == value
        assert sheet.cells["A2"].number_format == "General"

    def test_write_excel_writes_floats_with_at_most_16_significant_digits(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"x": [0.1 + 0.2]}))

        assert sheet.cells["A2"].value == "0.3"

    def test_write_excel_writes_float32_values_as_their_decimal_value(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"x": pl.Series([0.1, 1.1], dtype=pl.Float32)}))

        assert [sheet.cells["A2"].value, sheet.cells["A3"].value] == ["0.1", "1.1"]

    @pytest.mark.parametrize("dtype", [pl.Float32, pl.Float64])
    def test_write_excel_leaves_nan_and_null_blank(self, write_sheet: Callable, dtype: pl.DataType) -> None:
        sheet = write_sheet(pl.DataFrame({"x": pl.Series([float("nan"), None, 1.0], dtype=dtype)}))

        assert "A2" not in sheet.cells
        assert "A3" not in sheet.cells
        assert sheet.cells["A4"].value == "1"

    @pytest.mark.parametrize("dtype", [pl.Float32, pl.Float64])
    def test_write_excel_writes_infinity_as_text(self, write_sheet: Callable, dtype: pl.DataType) -> None:
        sheet = write_sheet(pl.DataFrame({"x": pl.Series([float("inf"), float("-inf")], dtype=dtype)}))

        assert [(sheet.cells[ref].kind, sheet.cells[ref].value) for ref in ("A2", "A3")] == [
            ("text", "inf"),
            ("text", "-inf"),
        ]

    def test_write_excel_writes_decimals_as_numbers(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"d": pl.Series([Decimal("1.25")], dtype=pl.Decimal(10, 2))}))

        assert (sheet.cells["A2"].kind, sheet.cells["A2"].value) == ("number", "1.25")


class TestWriteExcelBooleans:
    def test_write_excel_writes_booleans_as_excel_booleans(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"b": [True, False, None]}))

        assert [(sheet.cells[ref].kind, sheet.cells[ref].value) for ref in ("A2", "A3")] == [
            ("boolean", "1"),
            ("boolean", "0"),
        ]
        assert "A4" not in sheet.cells


class TestWriteExcelDates:
    def test_write_excel_writes_dates_as_excel_dates(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"d": [dt.date(2024, 1, 2)]}))

        assert (sheet.cells["A2"].kind, sheet.cells["A2"].value) == ("number", "45293")
        assert sheet.cells["A2"].number_format == "yyyy-mm-dd"

    def test_write_excel_writes_naive_datetimes_as_excel_datetimes(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"t": [dt.datetime(2024, 1, 2, 12, 0)]}))

        assert float(sheet.cells["A2"].value) == 45293.5
        assert sheet.cells["A2"].number_format == "yyyy-mm-dd hh:mm:ss"

    def test_write_excel_shows_milliseconds_for_datetime_columns_with_fractional_seconds(
        self, write_sheet: Callable
    ) -> None:
        sheet = write_sheet(
            pl.DataFrame({"t": [dt.datetime(2024, 1, 2, 12, 0), dt.datetime(2024, 1, 2, 12, 0, 0, 250_000)]})
        )

        assert sheet.cells["A3"].number_format == "yyyy-mm-dd hh:mm:ss.000"
        assert float(sheet.cells["A3"].value) == pytest.approx(45293.5 + 0.25 / 86_400, abs=1e-9)

    @pytest.mark.parametrize(
        ("value", "expected"),
        [
            (dt.date(1899, 12, 31), "1899-12-31"),
            (dt.date(1900, 1, 1), "1900-01-01"),
            (dt.date(1900, 2, 28), "1900-02-28"),
            (dt.datetime(1900, 1, 1, 10, 30), "1900-01-01T10:30:00"),
        ],
    )
    def test_write_excel_writes_dates_before_march_1900_as_iso_text(
        self, write_sheet: Callable, value: dt.date, expected: str
    ) -> None:
        sheet = write_sheet(pl.DataFrame({"d": [value]}))

        assert (sheet.cells["A2"].kind, sheet.cells["A2"].value) == ("text", expected)

    def test_write_excel_writes_march_1900_as_first_excel_date(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"d": [dt.date(1900, 3, 1)]}))

        assert (sheet.cells["A2"].kind, sheet.cells["A2"].value) == ("number", "61")

    def test_write_excel_writes_timezone_aware_datetimes_as_iso_text_with_offset(self, write_sheet: Callable) -> None:
        utc = pl.Series(
            [dt.datetime(2024, 1, 2, 9, 0), dt.datetime(2024, 7, 2, 9, 0, 0, 500)], dtype=pl.Datetime("us", "UTC")
        )
        sheet = write_sheet(pl.DataFrame({"t": utc.dt.convert_time_zone("Europe/Vienna")}))

        assert [(sheet.cells[ref].kind, sheet.cells[ref].value) for ref in ("A2", "A3")] == [
            ("text", "2024-01-02T10:00:00+01:00"),
            ("text", "2024-07-02T11:00:00.000500+02:00"),
        ]


class TestWriteExcelOtherTypes:
    @pytest.mark.parametrize(
        ("series", "expected"),
        [
            (pl.Series([dt.time(10, 30, 5)]), "10:30:05"),
            (pl.Series([dt.timedelta(days=1, seconds=5)]), "1 day, 0:00:05"),
            (pl.Series([[1, 2]]), "[1, 2]"),
            (pl.Series([{"a": 1, "b": "x"}]), "{'a': 1, 'b': 'x'}"),
            (pl.Series([b"\x00\xff"]), "00ff"),
        ],
    )
    def test_write_excel_writes_other_types_as_text(
        self, write_sheet: Callable, series: pl.Series, expected: str
    ) -> None:
        sheet = write_sheet(pl.DataFrame({"v": series}))

        assert (sheet.cells["A2"].kind, sheet.cells["A2"].value) == ("text", expected)

    def test_write_excel_leaves_all_null_columns_blank(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"empty": [None, None]}))

        assert set(sheet.cells) == {"A1"}


class TestWriteExcelHeader:
    def test_write_excel_writes_column_names_as_unchanged_text(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"=x": [1], "1": [2], "007": [3]}))

        assert [(sheet.cells[ref].kind, sheet.cells[ref].value) for ref in ("A1", "B1", "C1")] == [
            ("text", "=x"),
            ("text", "1"),
            ("text", "007"),
        ]

    def test_write_excel_keeps_empty_column_names_empty(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"a": [1], "": [2]}))

        assert "B1" not in sheet.cells or sheet.cells["B1"].value == ""

    def test_write_excel_formats_freezes_and_filters_the_header_by_default(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"a": [1, 2], "b": [3, 4]}))

        assert sheet.cells["A1"].bold
        assert not sheet.cells["A2"].bold
        assert sheet.frozen_rows == 1
        assert sheet.autofilter == "A1:B3"

    def test_write_excel_skips_all_header_formatting_in_plain_mode(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"a": [1, 2], "b": [3, 4]}), format_header=False)

        assert not sheet.cells["A1"].bold
        assert sheet.frozen_rows == 0
        assert sheet.autofilter is None

    def test_write_excel_writes_header_of_table_without_rows(self, write_sheet: Callable) -> None:
        sheet = write_sheet(pl.DataFrame({"a": [], "b": []}))

        assert set(sheet.cells) == {"A1", "B1"}
        assert sheet.autofilter == "A1:B1"


class TestWriteExcelColumnWidths:
    def test_write_excel_keeps_column_widths_between_minimum_and_maximum(self, write_sheet: Callable) -> None:
        columns = {"a": ["x"], "b": [None], "c": ["x" * 500], "d": ["x" * 5_000]}
        sheet = write_sheet(pl.DataFrame(columns, schema=dict.fromkeys(columns, pl.String)))

        short, empty, long, longer = sheet.column_widths
        assert short == empty
        assert long == longer
        assert short < long

    @pytest.mark.parametrize("value", [dt.datetime(2024, 1, 2, 12, 0), dt.datetime(2024, 1, 2, 12, 0, 0, 250_000)])
    def test_write_excel_widens_datetime_columns_beyond_the_maximum_to_show_full_dates(
        self, write_sheet: Callable, value: dt.datetime
    ) -> None:
        sheet = write_sheet(pl.DataFrame({"t": [value], "text": ["x" * 500]}))

        datetime_width, capped_text_width = sheet.column_widths
        assert datetime_width > capped_text_width


class TestWriteExcelLimits:
    def test_write_excel_raises_for_too_many_rows_without_writing_a_file(self, tmp_path: Path) -> None:
        output_path = tmp_path / "output.xlsx"

        with pytest.raises(ExcelLimitError, match="1,048,576 rows"):
            write_excel(pl.DataFrame({"a": pl.zeros(EXCEL_MAX_ROWS, eager=True)}), output_path)
        assert not output_path.exists()

    def test_write_excel_raises_for_too_many_columns_without_writing_a_file(self, tmp_path: Path) -> None:
        output_path = tmp_path / "output.xlsx"
        df = pl.DataFrame({f"c{i}": [1] for i in range(EXCEL_MAX_COLUMNS + 1)})

        with pytest.raises(ExcelLimitError, match="16,385 columns"):
            write_excel(df, output_path)
        assert not output_path.exists()

    def test_write_excel_limits_match_the_xlsxwriter_worksheet_limits(self) -> None:
        worksheet = Worksheet()

        assert (EXCEL_MAX_ROWS, EXCEL_MAX_COLUMNS, EXCEL_MAX_STRING_LENGTH) == (
            worksheet.xls_rowmax,
            worksheet.xls_colmax,
            worksheet.xls_strmax,
        )

    def test_write_excel_writes_every_row_up_to_the_row_limit(
        self, write_sheet: Callable, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(writers, "EXCEL_MAX_ROWS", 4)

        sheet = write_sheet(pl.DataFrame({"a": [1, 2, 3]}))

        assert sheet.cells["A4"].value == "3"
        with pytest.raises(ExcelLimitError):
            write_sheet(pl.DataFrame({"a": [1, 2, 3, 4]}))


class TestWriteExcelFile:
    def test_write_excel_raises_write_error_when_the_folder_does_not_exist(self, tmp_path: Path) -> None:
        output_path = tmp_path / "missing" / "output.xlsx"

        with pytest.raises(WriteError, match="Could not write .*output.xlsx.*folder exists and is writable"):
            write_excel(pl.DataFrame({"a": [1]}), output_path)
        assert not output_path.exists()

    def test_write_excel_puts_xlsxwriter_temp_files_into_tmpdir(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch, xlsx_reader: Callable
    ) -> None:
        tmpdir = tmp_path / "work"
        tmpdir.mkdir()
        used_dirs: list[str | None] = []
        original_mkstemp = tempfile.mkstemp

        def mkstemp(*args: Any, **kwargs: Any) -> tuple[int, str]:
            used_dirs.append(kwargs.get("dir"))
            return original_mkstemp(*args, **kwargs)

        monkeypatch.setattr(tempfile, "mkstemp", mkstemp)

        write_excel(pl.DataFrame({"a": ["x", "y"]}), tmp_path / "output.xlsx", tmpdir=tmpdir)

        assert used_dirs
        assert set(used_dirs) == {str(tmpdir)}
        assert xlsx_reader(tmp_path / "output.xlsx").cells["A3"].value == "y"
        assert list(tmpdir.iterdir()) == []
