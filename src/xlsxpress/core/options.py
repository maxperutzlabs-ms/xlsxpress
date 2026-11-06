"""Conversion options shared by all frontends (CLI, context menu, GUI)."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from xlsxpress.core.cleanup import CleanupPolicy


@dataclass(slots=True)
class ConversionOptions:
    """Options controlling a single file to xlsx conversion.

    Attributes:
        input_path: The tabular source file to convert.
        nrows: If set, only read the first `nrows` data rows (preview mode).
        to_temp: Write the xlsx file into the temp folder instead of next to the
            input file (see `paths.get_temp_dir`). Temp files get a unique name,
            and old ones are deleted according to `cleanup_policy`, edited or
            not.
        open_after: Open the resulting xlsx file in the default application (Excel
            on Windows) after writing.
        format_header: Apply header formatting to the first row.
        cleanup_policy: When to delete old files from the temp folder. Only used
            with `to_temp`.
    """

    input_path: Path
    nrows: int | None = None
    to_temp: bool = False
    open_after: bool = False
    format_header: bool = True
    cleanup_policy: CleanupPolicy = field(default_factory=CleanupPolicy)
