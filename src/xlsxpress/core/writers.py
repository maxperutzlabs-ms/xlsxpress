"""Writing DataFrames to formatted xlsx files."""

from __future__ import annotations

from pathlib import Path

import polars as pl

_HEADER_FORMAT = {
    "bold": True,
    "align": "center",
    "valign": "vcenter",
    "text_wrap": True,
    "bottom": 0,
    "top": 0,
}


def write_excel(
    df: pl.DataFrame, output_path: Path, format_header: bool = True
) -> Path:
    """Write a DataFrame to an xlsx file and return the output path.

    Args:
        df: The table to write.
        output_path: Target xlsx path (parent directory must exist).
        format_header: Apply bold/shaded formatting to and freeze the
            header row.
    """
    if format_header:
        df.write_excel(
            workbook=output_path,
            header_format=_HEADER_FORMAT,
            freeze_panes=(1, 0),
            autofit=True,
        )
    else:
        df.write_excel(workbook=output_path, autofit=True)
    return output_path
