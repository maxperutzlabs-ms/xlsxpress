"""Writing DataFrames to formatted xlsx files with xlsxwriter.

Each polars data type is written in one defined way, following the conversion
decisions in `docs/development-plan.md`. Rows are streamed to disk (xlsxwriter's
`constant_memory` mode), so memory use does not grow with the size of the table.
"""

from __future__ import annotations

import contextlib
import datetime as dt
import math
import warnings
from collections.abc import Callable
from pathlib import Path
from typing import Any

import polars as pl
import xlsxwriter
from xlsxwriter.exceptions import XlsxFileError
from xlsxwriter.utility import cell_autofit_width
from xlsxwriter.workbook import Workbook
from xlsxwriter.worksheet import Worksheet

from xlsxpress.core.errors import ConversionWarning, ExcelLimitError, WriteError

#: Maximum number of rows of an Excel worksheet, including the header row.
EXCEL_MAX_ROWS = 1_048_576
#: Maximum number of columns of an Excel worksheet.
EXCEL_MAX_COLUMNS = 16_384
#: Maximum number of characters in an Excel cell.
EXCEL_MAX_STRING_LENGTH = 32_767

_WORKBOOK_OPTIONS = {
    "constant_memory": True,
    "strings_to_formulas": False,
    "strings_to_numbers": False,
    "strings_to_urls": False,
}
_HEADER_FORMAT = {
    "bold": True,
    "align": "center",
    "valign": "vcenter",
    "text_wrap": True,
    "bottom": 0,
    "top": 0,
}
_DATE_FORMAT = "yyyy-mm-dd"
_DATETIME_FORMAT = "yyyy-mm-dd hh:mm:ss"
_DATETIME_MILLISECONDS_FORMAT = "yyyy-mm-dd hh:mm:ss.000"
_ISO_DATETIME_WITH_OFFSET = "%Y-%m-%dT%H:%M:%S%.f%:z"
#: Earlier Excel serial dates are unreliable because of Excel's 1900 leap year bug.
_FIRST_EXCEL_DATE = dt.date(1900, 3, 1)
_MAX_COL_PIXELS = 125
_MIN_COL_PIXELS = 40
_WIDTH_SAMPLE_ROWS = 1_000

_CellWriter = Callable[[int, int, Any], None]


def write_excel(
    df: pl.DataFrame,
    output_path: Path,
    format_header: bool = True,
) -> Path:
    """Write a DataFrame to an xlsx file and return the output path.

    Values are written according to their data type:

    - Strings as text, never as formulas, numbers or links.
    - Integers, decimals and floats as numbers in Excel's `General` format. Excel
      keeps at most 15 significant digits of a number.
    - Null and NaN as blank cells, `inf` and `-inf` as text.
    - Booleans as Excel booleans.
    - Dates and datetimes without a time zone as Excel dates. Excel's resolution is
      one millisecond. Dates before 1900-03-01 are written as ISO 8601 text.
    - Datetimes with a time zone as ISO 8601 text including the UTC offset.
    - Times, durations, lists, structs, binary and all other types as text.

    Column widths are estimated from the header and the first rows, within fixed
    minimum and maximum widths.

    Args:
        df: The table to write.
        output_path: Target xlsx path. The parent directory must exist.
        format_header: If True (default), formats the header row in bold with
            wrapped text, freezes it and adds an autofilter.

    Returns:
        The output path.

    Raises:
        ExcelLimitError: If the table has more rows or columns than an Excel
            worksheet holds. No file is written in that case.
        WriteError: If the file cannot be written, for example because the folder
            is not writable or the disk is full. No file is left behind.

    Warns:
        ConversionWarning: If text longer than Excel's limit of 32,767 characters
            was truncated.
    """
    _check_excel_limits(df)
    df = _prepare_columns(df)
    workbook = xlsxwriter.Workbook(str(output_path), _WORKBOOK_OPTIONS)
    try:
        sheet_writer = _SheetWriter(workbook, workbook.add_worksheet())
        sheet_writer.write(df, format_header)
        workbook.close()
    except BaseException as error:
        with contextlib.suppress(Exception):
            workbook.close()
        with contextlib.suppress(OSError):
            output_path.unlink(missing_ok=True)
        if isinstance(error, OSError | XlsxFileError):
            raise WriteError(_describe_write_error(output_path, error)) from error
        raise
    sheet_writer.warn_about_truncated_text(df.columns)
    return output_path


class _SheetWriter:
    """Writes one DataFrame to a worksheet, choosing a cell writer per column."""

    def __init__(self, workbook: Workbook, worksheet: Worksheet) -> None:
        self._workbook = workbook
        self._worksheet = worksheet
        self._date_format = workbook.add_format({"num_format": _DATE_FORMAT})
        self._datetime_format = workbook.add_format({"num_format": _DATETIME_FORMAT})
        self._datetime_ms_format = workbook.add_format(
            {"num_format": _DATETIME_MILLISECONDS_FORMAT}
        )
        self._truncated: dict[int, int] = {}

    def write(self, df: pl.DataFrame, format_header: bool) -> None:
        """Write header, column widths, header formatting and all rows."""
        self._set_column_widths(df)
        self._write_header(df, format_header)
        cell_writers = [self._cell_writer(df[name]) for name in df.columns]
        for row_index, row in enumerate(df.iter_rows(), start=1):
            for col_index, value in enumerate(row):
                if value is not None:
                    cell_writers[col_index](row_index, col_index, value)

    def warn_about_truncated_text(self, columns: list[str]) -> None:
        """Emit one `ConversionWarning` per column with truncated text."""
        for col_index, count in self._truncated.items():
            warnings.warn(
                f"{count:,} value(s) in column {columns[col_index]!r} are longer "
                f"than Excel's limit of {EXCEL_MAX_STRING_LENGTH:,} characters and "
                "were truncated.",
                ConversionWarning,
                stacklevel=3,
            )

    def _set_column_widths(self, df: pl.DataFrame) -> None:
        """Fit columns to header and sample values, within min and max widths.

        Date columns are always wide enough for their full date, because Excel
        shows `###` instead of a date that does not fit, while text overflows into
        the next cell.
        """
        sample = df.head(_WIDTH_SAMPLE_ROWS)
        for col_index, name in enumerate(df.columns):
            texts = [name] + [str(v) for v in sample[name].to_list() if v is not None]
            width = max(cell_autofit_width(text) for text in texts)
            width = max(_MIN_COL_PIXELS, min(width, _MAX_COL_PIXELS))
            date_text = _date_display_text(df[name])
            if date_text is not None:
                width = max(width, cell_autofit_width(date_text))
            self._worksheet.set_column_pixels(col_index, col_index, width)

    def _write_header(self, df: pl.DataFrame, format_header: bool) -> None:
        header_format = None
        if format_header:
            header_format = self._workbook.add_format(_HEADER_FORMAT)
            self._worksheet.freeze_panes(1, 0)
            if df.width:
                self._worksheet.autofilter(0, 0, df.height, df.width - 1)
        for col_index, name in enumerate(df.columns):
            self._worksheet.write_string(0, col_index, name, header_format)

    def _cell_writer(self, series: pl.Series) -> _CellWriter:
        """Return the function that writes one non-null value of `series`."""
        dtype = series.dtype
        if dtype == pl.String or isinstance(dtype, pl.Categorical | pl.Enum):
            return self._write_text
        if dtype.is_integer():
            return self._write_number
        if dtype.is_float():
            return self._write_float
        if dtype.is_decimal():
            return self._write_decimal
        if dtype == pl.Boolean:
            return self._write_boolean
        if dtype == pl.Date:
            return self._write_date
        if isinstance(dtype, pl.Datetime):
            return self._datetime_writer(_has_fractional_seconds(series))
        if dtype == pl.Time:
            return self._write_iso_text
        if dtype == pl.Binary:
            return self._write_hex_text
        return self._write_str_text

    def _write_text(self, row: int, col: int, value: str) -> None:
        if self._worksheet.write_string(row, col, value) == -2:
            self._truncated[col] = self._truncated.get(col, 0) + 1

    def _write_number(self, row: int, col: int, value: float) -> None:
        self._worksheet.write_number(row, col, value)

    def _write_float(self, row: int, col: int, value: float) -> None:
        if math.isnan(value):
            return
        if math.isinf(value):
            self._worksheet.write_string(row, col, "inf" if value > 0 else "-inf")
            return
        self._worksheet.write_number(row, col, value)

    def _write_decimal(self, row: int, col: int, value: Any) -> None:
        self._worksheet.write_number(row, col, float(value))

    def _write_boolean(self, row: int, col: int, value: bool) -> None:
        self._worksheet.write_boolean(row, col, value)

    def _write_date(self, row: int, col: int, value: dt.date) -> None:
        if value < _FIRST_EXCEL_DATE:
            self._write_text(row, col, value.isoformat())
        else:
            self._worksheet.write_datetime(row, col, value, self._date_format)

    def _datetime_writer(self, has_fraction: bool) -> _CellWriter:
        number_format = (
            self._datetime_ms_format if has_fraction else self._datetime_format
        )

        def write_datetime(row: int, col: int, value: dt.datetime) -> None:
            if value.date() < _FIRST_EXCEL_DATE:
                self._write_text(row, col, value.isoformat())
            else:
                self._worksheet.write_datetime(row, col, value, number_format)

        return write_datetime

    def _write_iso_text(self, row: int, col: int, value: dt.time) -> None:
        self._write_text(row, col, value.isoformat())

    def _write_hex_text(self, row: int, col: int, value: bytes) -> None:
        self._write_text(row, col, value.hex())

    def _write_str_text(self, row: int, col: int, value: object) -> None:
        self._write_text(row, col, str(value))


def _date_display_text(series: pl.Series) -> str | None:
    """Return a placeholder as long as the Excel date display, or None for non-dates."""
    if series.dtype == pl.Date:
        return "0000-00-00"
    if isinstance(series.dtype, pl.Datetime):
        return (
            "0000-00-00 00:00:00.000"
            if _has_fractional_seconds(series)
            else "0000-00-00 00:00:00"
        )
    return None


def _has_fractional_seconds(series: pl.Series) -> bool:
    return (series.dt.nanosecond() != 0).any()


def _check_excel_limits(df: pl.DataFrame) -> None:
    max_data_rows = EXCEL_MAX_ROWS - 1
    if df.height > max_data_rows:
        raise ExcelLimitError(
            f"The table has {df.height:,} rows, but an Excel worksheet holds at most "
            f"{max_data_rows:,} rows below the header. Use the preview option to "
            "convert only the first rows."
        )
    if df.width > EXCEL_MAX_COLUMNS:
        raise ExcelLimitError(
            f"The table has {df.width:,} columns, but an Excel worksheet holds at "
            f"most {EXCEL_MAX_COLUMNS:,} columns."
        )


def _prepare_columns(df: pl.DataFrame) -> pl.DataFrame:
    """Convert columns that polars must not hand to Python as they are.

    - Float32 columns become Float64 through their shortest decimal text. A direct
      cast turns the Float32 value `0.1` into `0.10000000149011612`.
    - Datetime columns with a time zone become ISO 8601 text. Converting them to
      Python objects needs the `tzdata` package on Windows, and polars panics
      without it.
    """
    conversions: list[pl.Expr] = []
    for name, dtype in df.schema.items():
        if dtype == pl.Float32:
            conversions.append(pl.col(name).cast(pl.String).cast(pl.Float64))
        elif isinstance(dtype, pl.Datetime) and dtype.time_zone is not None:
            conversions.append(pl.col(name).dt.to_string(_ISO_DATETIME_WITH_OFFSET))
    return df.with_columns(conversions) if conversions else df


def _describe_write_error(output_path: Path, error: Exception) -> str:
    # xlsxwriter wraps the OSError of a failed file creation in `FileCreateError`.
    cause = error.args[0] if isinstance(error, XlsxFileError) and error.args else error
    reason = cause.strerror if isinstance(cause, OSError) and cause.strerror else cause
    return (
        f"Could not write {str(output_path)!r}: {reason}. Check that the folder "
        "exists and is writable, and that there is enough free disk space."
    )
