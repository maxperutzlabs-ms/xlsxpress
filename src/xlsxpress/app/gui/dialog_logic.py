"""Decisions of the conversion dialog, kept free of GUI code so they can be tested.

The dialog (`xlsxpress.app.gui.dialog`) reads its widgets into a `FormState`, turns
it into conversion options or saved defaults with the functions here, and shows the
`Outcome` of a conversion. This module also has the texts that depend on the
options, and the truncation of long paths for labels.
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
from xlsxpress.core.progress import Stage

#: Label of the context menu entry that converts with the saved defaults.
OPEN_ENTRY_LABEL = next(entry.label for entry in ENTRIES if entry.app_command == "open")

#: Marks the part of a path that was left out.
ELLIPSIS = "\u2026"

_STAGE_TEXTS: dict[Stage, str] = {
    "reading": f"Reading file{ELLIPSIS}",
    "writing": f"Writing workbook{ELLIPSIS}",
    "saving": f"Saving workbook{ELLIPSIS}",
    "opening": f"Opening in Excel{ELLIPSIS}",
}

#: Style of the dialog's result banner.
BannerKind = Literal["success", "warning", "error"]


@dataclass(frozen=True, slots=True)
class FormState:
    """The values of the dialog's widgets.

    Attributes:
        open_after: "Open in Excel after conversion" is checked.
        to_temp: "Temp folder" is selected under "Save workbook".
        format_header: "Format header row" is checked.
        preview: "Convert only the first N rows" is checked.
        nrows_text: The text of the row limit field.
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
        title: The title of the result banner.
        message: The explanation of an error, or an empty string.
        output_path: The written xlsx file, or None after an error.
        warnings: Messages of the conversion's warnings.
        close_automatically: True if the dialog has nothing to report and closes.
    """

    kind: Literal["success", "open_failed", "error"]
    title: str
    message: str = ""
    output_path: Path | None = None
    warnings: tuple[str, ...] = ()
    close_automatically: bool = False

    @property
    def banner(self) -> BannerKind:
        """The banner style: a written file with a problem is a warning."""
        if self.kind == "error":
            return "error"
        if self.kind == "open_failed" or self.warnings:
            return "warning"
        return "success"


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
    """Parse the row limit field as a positive whole number.

    Args:
        text: The text of the row limit field.

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
        raise ValueError(
            f"Invalid number of rows {text!r}. "
            "Enter a whole number of rows greater than 0."
        )
    return nrows


def build_options(
    input_path: Path, form: FormState, cleanup_policy: CleanupPolicy
) -> ConversionOptions:
    """Return the conversion options for the current form.

    The row limit field is only used, and only checked, when its checkbox is
    checked.

    Args:
        input_path: The file to convert.
        form: The values of the dialog's widgets.
        cleanup_policy: When to delete old files from the temp folder.

    Returns:
        The conversion options.

    Raises:
        ValueError: If the row limit is checked and its field is invalid.
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

    Without the row limit checked, its field is disabled and can't be corrected, so
    an invalid value there keeps the saved number of rows instead of failing.

    Args:
        form: The values of the dialog's widgets.
        saved: The currently saved defaults.

    Returns:
        The new defaults.

    Raises:
        ValueError: If the row limit is checked and its field is invalid.
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
    if result.warnings:
        title = "Converted with warnings"
    elif opened:
        title = "Converted and opened in Excel"
    else:
        title = "Converted"
    return Outcome(
        kind="success",
        title=title,
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
            title="Converted, but not opened",
            message=str(error),
            output_path=error.output_path,
            warnings=error.warnings,
        )
    if isinstance(error, ConversionError):
        return Outcome(kind="error", title="Conversion failed", message=str(error))
    return Outcome(
        kind="error",
        title="Conversion failed",
        message=(
            f"Unexpected error: {error!r}. The details are in the log file: {log_path}"
        ),
    )


def save_location_help(
    to_temp: bool,
    source_name: str,
    cleanup_policy: CleanupPolicy,
    output_name: str | None = None,
) -> str:
    """Return the help text below the "Save workbook" choice.

    Args:
        to_temp: True if "Temp folder" is selected.
        source_name: The name of the file to convert.
        cleanup_policy: When files in the temp folder are deleted.
        output_name: The name of the workbook last written next to the original,
            if there is one. Ignored for the temp folder.

    Returns:
        The help text.
    """
    if to_temp:
        return (
            f"Deleted after {cleanup_policy.max_age_days} days or when there are "
            f"more than {cleanup_policy.max_files} workbooks. Save a copy elsewhere "
            "to keep it."
        )
    saved = f"Saved as {output_name} next to" if output_name else "Saved next to"
    return f"{saved} {source_name}. Existing files are never overwritten."


def about_sections(cleanup_policy: CleanupPolicy) -> tuple[tuple[str, str], ...]:
    """Return the titles and texts of the About dialog's explanations.

    Args:
        cleanup_policy: When files in the temp folder are deleted.

    Returns:
        Pairs of title and text.
    """
    return (
        (
            "Save as default",
            "Stores these options for this dialog and for "
            f"\u201c{OPEN_ENTRY_LABEL}\u201d. That entry always opens Excel.",
        ),
        (
            "Reset",
            "Restores the built-in options. Nothing is saved until you click "
            "Save as default.",
        ),
        (
            "Temp folder",
            f"Workbooks are deleted after {cleanup_policy.max_age_days} days or "
            f"beyond {cleanup_policy.max_files} files, even if edited.",
        ),
    )


def progress_status(stage: Stage, fraction: float | None) -> str:
    """Return the status line of the progress window.

    Args:
        stage: The current step of the conversion.
        fraction: The completed part of the stage, from 0 to 1, or None.

    Returns:
        The status text, with a percentage if `fraction` is known.
    """
    text = _STAGE_TEXTS[stage]
    if fraction is None:
        return text
    return f"{text} {int(min(max(fraction, 0.0), 1.0) * 100)}%"


def format_elapsed(seconds: float) -> str:
    """Return a duration as minutes and seconds, for example `1:07`."""
    minutes, rest = divmod(int(seconds), 60)
    return f"{minutes}:{rest:02d}"


def truncate_middle(text: str, max_chars: int) -> str:
    r"""Shorten a path to at most `max_chars` characters by leaving out its middle.

    Keeps the start (the root and the first folder) and the end (the last two
    components, or the last one), joined with an ellipsis, for example
    `C:\Users\…\demo\demo.xlsx`. If that is still too long, keeps characters from
    both ends instead.

    Args:
        text: The path, with `\` or `/` as separator.
        max_chars: The maximum length of the result.

    Returns:
        `text` if it is short enough, otherwise the shortened path.

    Raises:
        ValueError: If `max_chars` is less than 1.
    """
    _check_max_chars(max_chars)
    if len(text) <= max_chars:
        return text
    sep = _separator(text)
    parts = text.split(sep)
    # The head is "C:\Users", "/home" or "\\server\share": leading empty parts
    # belong to the root, and a share counts as part of the root.
    n_root = len(text) - len(text.lstrip(sep))
    n_head = n_root + 2 if n_root > 1 else 2
    for n_tail in (2, 1):
        if len(parts) - n_head - n_tail < 1:
            continue
        candidate = sep.join([*parts[:n_head], ELLIPSIS, *parts[-n_tail:]])
        if len(candidate) <= max_chars:
            return candidate
    n_end = (max_chars - 1) // 2
    n_start = max_chars - 1 - n_end
    return text[:n_start] + ELLIPSIS + text[len(text) - n_end :]


def truncate_start(text: str, max_chars: int) -> str:
    r"""Shorten a path to at most `max_chars` characters by leaving out its start.

    Keeps as many of the last components as fit, after an ellipsis, for example
    `…\AppData\Local\xlsxpress`. If not even the last component fits, keeps its
    last characters.

    Args:
        text: The path, with `\` or `/` as separator.
        max_chars: The maximum length of the result.

    Returns:
        `text` if it is short enough, otherwise the shortened path.

    Raises:
        ValueError: If `max_chars` is less than 1.
    """
    _check_max_chars(max_chars)
    if len(text) <= max_chars:
        return text
    sep = _separator(text)
    parts = text.split(sep)
    shortened = ""
    for n_tail in range(1, len(parts)):
        candidate = sep.join([ELLIPSIS, *parts[-n_tail:]])
        if len(candidate) > max_chars:
            break
        shortened = candidate
    if shortened:
        return shortened
    return ELLIPSIS + text[len(text) - (max_chars - 1) :]


def _check_max_chars(max_chars: int) -> None:
    if max_chars < 1:
        raise ValueError(f"Invalid max_chars {max_chars!r}. Must be at least 1.")


def _separator(path: str) -> str:
    return "\\" if "\\" in path else "/"
