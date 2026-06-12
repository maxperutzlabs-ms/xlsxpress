"""Conversion options shared by all frontends (CLI, context menu, GUI)."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


@dataclass(slots=True)
class ConversionOptions:
    """Options controlling a single file to '.xlsx' conversion.

    Attributes:
        input_path: The tabular source file to convert.
        nrows: If set, only read the first `nrows` data rows (preview mode).
        to_temp: Write the '.xlsx' file into the temp directory instead of next to the
            input file.
        open_after: Open the resulting '.xlsx' file in the default application (Excel on
            Windows) after writing.
        format_header: Apply header formatting to the first row.
    """

    input_path: Path
    nrows: int | None = None
    to_temp: bool = False
    temp_dir: Path | None = None
    open_after: bool = False
    format_header: bool = True
