"""Writing DataFrames to formatted xlsx files."""

from __future__ import annotations

from pathlib import Path

import polars as pl
import xlsxwriter
from xlsxwriter.utility import cell_autofit_width

_HEADER_FORMAT = {
    "bold": True,
    "align": "center",
    "valign": "vcenter",
    "text_wrap": True,
    "bottom": 0,
    "top": 0,
}
_MAX_COL_PIXELS: int = 125
_MIN_COL_PIXELS: int = 40


def write_excel(
    df: pl.DataFrame, output_path: Path, format_header: bool = True
) -> Path:
    """Write a DataFrame to an xlsx file and return the output path.

    Replaces NaN with blank cells to avoid Excel showing formula errors `=#NUM!`.
    Floating-point columns use Excel's `General` format to avoid fixed decimal rounding.
    Automatically adjusts column widths to fit content, capped to a max width of 200
    pixels to avoid excessively wide columns.

    Args:
        df: The table to write.
        output_path: Target xlsx path (parent directory must exist).
        format_header: Apply bold/shaded formatting to and freeze the header row.
    """
    df = df.with_columns(pl.col(pl.Float32, pl.Float64).fill_nan(None))
    workbook = xlsxwriter.Workbook(output_path)
    try:
        df.write_excel(
            workbook=workbook,
            header_format=_HEADER_FORMAT if format_header else None,
            freeze_panes=(1, 0) if format_header else None,
            autofit=True,
            dtype_formats={pl.Float32: "General", pl.Float64: "General"},
        )
        _apply_capped_autofit(workbook, df)
    finally:
        workbook.close()

    return output_path


def _apply_capped_autofit(workbook: xlsxwriter.Workbook, df: pl.DataFrame) -> None:
    """Set each column to autofit width, capped to [_MIN_COL_PIXELS, _MAX_COL_PIXELS]."""
    worksheet = workbook.worksheets()[0]

    for col_idx, col in enumerate(df.columns):
        # Measure header and all cell values as strings.
        values = [col] + ["" if v is None else str(v) for v in df[col].to_list()]
        raw_width = max(
            (cell_autofit_width(v) for v in values), default=_MIN_COL_PIXELS
        )
        capped = max(_MIN_COL_PIXELS, min(raw_width, _MAX_COL_PIXELS))
        worksheet.set_column_pixels(col_idx, col_idx, capped)
