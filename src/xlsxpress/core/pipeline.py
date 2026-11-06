"""The main conversion pipeline: read, write xlsx, optionally open."""

from __future__ import annotations

import contextlib
import os
import shutil
import tempfile
import warnings
from dataclasses import dataclass
from pathlib import Path

from xlsxpress.core import cleanup, excel, paths, readers, writers
from xlsxpress.core.errors import ConversionWarning, OpenOutputError, WriteError
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

    The workbook is first written to the work folder (`paths.get_work_dir`) and
    moved (temp folder) or copied (next to the input) to its output path only when
    complete, so a file under the output name is never unfinished, even if the
    process is killed during the conversion.

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
        errors.OpenOutputError: If `options.open_after` is set and the written file
            could not be opened. The error holds the output path and the warnings.
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
    result = ConversionResult(output_path, tuple(messages))
    if options.open_after:
        try:
            excel.open_in_default_app(output_path)
        except OSError as error:
            raise OpenOutputError(
                f"The xlsx file was written to '{output_path}', but it could not be "
                f"opened: {error.strerror or error}. Maybe no program is set up to "
                "open .xlsx files. Open the file by hand, or choose a default app "
                "for .xlsx files.",
                output_path,
                result.warnings,
            ) from error
    return result


def _convert(options: ConversionOptions) -> Path:
    input_path = options.input_path.resolve()
    df = readers.read_table(input_path, nrows=options.nrows)

    if options.to_temp:
        output_dir = paths.get_temp_dir()
        cleanup.clean_temp_dir(output_dir, options.cleanup_policy)
    else:
        output_dir = input_path.parent

    # The workbook is written to the work folder first, so the output name never
    # holds an unfinished file, even if the process is killed while writing.
    work_dir = paths.get_work_dir()
    cleanup.clean_work_dir(work_dir)
    try:
        handle, name = tempfile.mkstemp(suffix=".xlsx", dir=work_dir)
    except OSError as error:
        raise WriteError(
            f"Could not create a file in the XlsxPress work folder '{work_dir}': "
            f"{error.strerror or error}. Check that there is enough free disk space."
        ) from error
    os.close(handle)
    work_file = Path(name)
    try:
        writers.write_excel(
            df, work_file, format_header=options.format_header, tmpdir=work_dir
        )
        return _publish(work_file, input_path, output_dir, options.to_temp)
    finally:
        with contextlib.suppress(OSError):
            work_file.unlink(missing_ok=True)


def _publish(
    work_file: Path, input_path: Path, output_dir: Path, to_temp: bool
) -> Path:
    try:
        output_path = paths.reserve_output_path(input_path, output_dir, unique=to_temp)
    except OSError as error:
        raise WriteError(
            f"Could not create an xlsx file in '{output_dir}': "
            f"{error.strerror or error}. Check that the folder is writable, or "
            "convert to the temp folder instead."
        ) from error
    # The temp folder sits next to the work folder, on the same drive and with the
    # same permissions, so the file is moved: a copy would be a second new file,
    # which the virus scanner checks again when Excel opens it (about 0.6 s for a
    # 2.5 MB workbook). Elsewhere the file is copied, so it gets the permissions of
    # the output folder and also works on another drive or a network share.
    try:
        if to_temp:
            os.replace(work_file, output_path)
        else:
            shutil.copyfile(work_file, output_path)
    except BaseException as error:
        with contextlib.suppress(OSError):
            output_path.unlink(missing_ok=True)
        if isinstance(error, OSError):
            raise WriteError(
                f"Could not write '{output_path}': {error.strerror or error}. Check "
                "that the folder is writable and that there is enough free disk space."
            ) from error
        raise
    return output_path
