"""Decisions of the conversion dialog, kept free of GUI code so they can be tested.

The dialog (`xlsxpress.app.gui.dialog`) reads its widgets into a `FormState`, turns
it into conversion options or saved defaults with the functions here, and shows the
`Outcome` of a conversion.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Literal

from xlsxpress.app.context_menu import ENTRIES
from xlsxpress.config import DialogDefaults
from xlsxpress.core.cleanup import CleanupPolicy
from xlsxpress.core.errors import ConversionError, OpenOutputError
from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import ConversionResult

#: Label of the context menu entry that converts with the saved defaults.
OPEN_ENTRY_LABEL = next(entry.label for entry in ENTRIES if entry.app_command == "open")


@dataclass(frozen=True, slots=True)
class FormState:
    """The values of the dialog's widgets.

    Attributes:
        open_after: "Open in Excel after conversion" is on.
        to_temp: "Save to temp directory" is on.
        format_header: "Format header row" is on.
        preview: "Preview only (limit rows)" is on.
        nrows_text: The text of the Rows field.
    """

    open_after: bool
    to_temp: bool
    format_header: bool
    preview: bool
    nrows_text: str


@dataclass(frozen=True, slots=True)
class Outcome:
    """What the dialog shows after a conversion.

    Attributes:
        kind: `success` if the file was written (and opened, if requested),
            `open_failed` if it was written but could not be opened, `error` if
            the conversion failed.
        message: The text to show.
        output_path: The written xlsx file, or None after an error.
        warnings: Messages of the conversion's warnings.
        close_automatically: True if the dialog has nothing to report and closes.
    """

    kind: Literal["success", "open_failed", "error"]
    message: str
    output_path: Path | None = None
    warnings: tuple[str, ...] = ()
    close_automatically: bool = False


def form_from_defaults(defaults: DialogDefaults) -> FormState:
    """Return the form state that shows `defaults`."""
    return FormState(
        open_after=defaults.open_after,
        to_temp=defaults.to_temp,
        format_header=defaults.format_header,
        preview=defaults.preview,
        nrows_text=str(defaults.nrows),
    )


def parse_nrows(text: str) -> int:
    """Parse the Rows field as a positive whole number.

    Args:
        text: The text of the Rows field.

    Returns:
        The number of rows.

    Raises:
        ValueError: If `text` is not a positive whole number.
    """
    try:
        nrows = int(text.strip())
    except ValueError:
        nrows = 0
    if nrows < 1:
        raise ValueError(f"Rows must be a positive whole number, not {text!r}.")
    return nrows


def build_options(
    input_path: Path, form: FormState, cleanup_policy: CleanupPolicy
) -> ConversionOptions:
    """Return the conversion options for the current form.

    The Rows field is only used, and only checked, with "Preview only".

    Args:
        input_path: The file to convert.
        form: The values of the dialog's widgets.
        cleanup_policy: When to delete old files from the temp folder.

    Returns:
        The conversion options.

    Raises:
        ValueError: If "Preview only" is on and the Rows field is invalid.
    """
    return ConversionOptions(
        input_path=input_path,
        nrows=parse_nrows(form.nrows_text) if form.preview else None,
        to_temp=form.to_temp,
        open_after=form.open_after,
        format_header=form.format_header,
        cleanup_policy=cleanup_policy,
    )


def defaults_from_form(form: FormState, saved: DialogDefaults) -> DialogDefaults:
    """Return the defaults to save for the current form.

    Without "Preview only" the Rows field is disabled and can't be corrected, so an
    invalid value there keeps the saved number of rows instead of failing.

    Args:
        form: The values of the dialog's widgets.
        saved: The currently saved defaults.

    Returns:
        The new defaults.

    Raises:
        ValueError: If "Preview only" is on and the Rows field is invalid.
    """
    if form.preview:
        nrows = parse_nrows(form.nrows_text)
    else:
        try:
            nrows = parse_nrows(form.nrows_text)
        except ValueError:
            nrows = saved.nrows
    return DialogDefaults(
        open_after=form.open_after,
        preview=form.preview,
        nrows=nrows,
        to_temp=form.to_temp,
        format_header=form.format_header,
    )


def outcome_of_result(result: ConversionResult, opened: bool) -> Outcome:
    """Return the outcome of a successful conversion.

    The dialog closes automatically only if the file was opened in Excel and there
    are no warnings; otherwise it stays open to show where the file is.

    Args:
        result: The result of the conversion.
        opened: True if the file was opened in Excel.

    Returns:
        The outcome to show.
    """
    message = "Converted and opened in Excel." if opened else "Converted."
    return Outcome(
        kind="success",
        message=message,
        output_path=result.output_path,
        warnings=result.warnings,
        close_automatically=opened and not result.warnings,
    )


def outcome_of_error(error: Exception, log_path: Path) -> Outcome:
    """Return the outcome of a conversion that raised `error`.

    Args:
        error: The raised error.
        log_path: The log file, named for unexpected errors.

    Returns:
        The outcome to show. A file that was written but could not be opened is
        not a failed conversion: its outcome has the output path.
    """
    if isinstance(error, OpenOutputError):
        return Outcome(
            kind="open_failed",
            message=str(error),
            output_path=error.output_path,
            warnings=error.warnings,
        )
    if isinstance(error, ConversionError):
        return Outcome(kind="error", message=f"Conversion failed: {error}")
    return Outcome(
        kind="error",
        message=(
            f"Conversion failed with an unexpected error: {error!r}. "
            f"The details are in the log file: {log_path}"
        ),
    )


def about_text(
    version: str, config_path: Path, temp_dir: Path, cleanup_policy: CleanupPolicy
) -> str:
    """Return the text of the dialog's About box.

    Args:
        version: The installed XlsxPress version.
        config_path: Where the defaults are saved.
        temp_dir: The temp folder.
        cleanup_policy: When files in the temp folder are deleted.

    Returns:
        The text.
    """
    return (
        f"XlsxPress {version}\n\n"
        "The options in this dialog apply to this conversion. "
        '"Update default options" saves them as the defaults of this dialog and of '
        f'the context menu entry "{OPEN_ENTRY_LABEL}". That entry always opens the '
        'file in Excel, whatever "Open in Excel after conversion" is set to.\n\n'
        '"Reset" sets the options to the built-in defaults. They are saved only when '
        'you click "Update default options".\n\n'
        f"Files in the temp folder are deleted after {cleanup_policy.max_age_days} "
        f"days, or when there are more than {cleanup_policy.max_files}, even if you "
        "edited them. Save a workbook somewhere else to keep it.\n\n"
        f"Temp folder: {temp_dir}\n"
        f"Defaults file: {config_path}"
    )
