"""The main conversion pipeline: read, write xlsx, optionally open."""

from __future__ import annotations

import contextlib
import warnings
from dataclasses import dataclass
from pathlib import Path

from xlsxpress.core import cleanup, excel, paths, readers, writers
from xlsxpress.core.errors import ConversionWarning, WriteError
from xlsxpress.core.options import ConversionOptions


@dataclass(frozen=True, slots=True)
class ConversionResult:
    """Outcome of a successful conversion.

    Attributes:
        output_path: Path of the written xlsx file.
        warnings: Messages of all `ConversionWarning`s, for example about truncated
            text or a file that is not valid UTF-8. Frontends show them to the user.
    """

    output_path: Path
    warnings: tuple[str, ...] = ()


def run_conversion(options: ConversionOptions) -> ConversionResult:
    """Convert a tabular file to an xlsx file according to `options`.

    The xlsx file is written next to the input file, or into the temp folder with
    `options.to_temp`. An existing file is never overwritten; a free name is chosen
    instead. When writing to the temp folder, old temp files are cleaned up first.

    `ConversionWarning`s are collected and returned instead of being emitted, so
    every frontend can show them.

    Args:
        options: The conversion options.

    Returns:
        The output path and the collected warnings.

    Raises:
        errors.ConversionError: If the input cannot be read, does not fit into an
            Excel worksheet, or the xlsx file cannot be written (see its subclasses
            in `xlsxpress.core.errors`). No partial xlsx file is left behind.
    """
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always", ConversionWarning)
        output_path = _convert(options)
    messages: list[str] = []
    for warning in caught:
        if issubclass(warning.category, ConversionWarning):
            messages.append(str(warning.message))
        else:
            warnings.warn_explicit(
                warning.message, warning.category, warning.filename, warning.lineno
            )
    if options.open_after:
        excel.open_in_default_app(output_path)
    return ConversionResult(output_path, tuple(messages))


def _convert(options: ConversionOptions) -> Path:
    input_path = options.input_path.resolve()
    df = readers.read_table(input_path, nrows=options.nrows)

    if options.to_temp:
        output_dir = paths.get_temp_dir()
        cleanup.clean_temp_dir(output_dir, options.cleanup_policy)
    else:
        output_dir = input_path.parent
    try:
        output_path = paths.reserve_output_path(
            input_path, output_dir, unique=options.to_temp
        )
    except OSError as error:
        raise WriteError(
            f"Could not create an xlsx file in {str(output_dir)!r}: "
            f"{error.strerror or error}. Check that the folder is writable, or "
            "convert to the temp folder instead."
        ) from error

    try:
        writers.write_excel(df, output_path, format_header=options.format_header)
    except BaseException:
        with contextlib.suppress(OSError):
            output_path.unlink(missing_ok=True)
        raise
    return output_path
