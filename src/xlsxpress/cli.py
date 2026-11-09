"""Command line interface: a thin routing layer without conversion logic."""

from __future__ import annotations

import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Annotated

from cyclopts import App, Parameter, validators

from xlsxpress import __VERSION__
from xlsxpress.app import context_menu
from xlsxpress.config import get_config_path, load_config
from xlsxpress.core.errors import ConversionError, OpenOutputError
from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import run_conversion

app = App(
    name="xlsxpress",
    help="Open tabular files (csv, tsv, txt, parquet) as safely formatted Excel files.",
    version=__VERSION__,
)

context_menu_app = App(
    name="context-menu",
    help="Add XlsxPress to the File Explorer context menu (Windows only).",
)
app.command(context_menu_app)

_Flag = Parameter(negative_bool=(), show_default=False)
_Nrows = Annotated[
    int | None, Parameter(name=["--nrows", "-n"], validator=validators.Number(gte=1))
]


def main(argv: Sequence[str] | None = None) -> int:
    """Run the command line application.

    Args:
        argv: Command line arguments. If None (default), uses the process arguments.

    Returns:
        The exit code: 0 on success, 1 if the conversion failed or the written file
        could not be opened, and cyclopts' exit code for invalid arguments.
    """
    result = app(argv, result_action="return_value")
    return result if isinstance(result, int) else 0


@app.command(name="open")
def open_file(
    file: Path,
    /,
    *,
    nrows: _Nrows = None,
    plain: Annotated[bool, _Flag] = False,
) -> int:
    """Convert to a temporary xlsx file and open it in Excel.

    The xlsx file is written to the XlsxPress temp folder, which is cleaned up
    automatically. Save the workbook somewhere else to keep your changes.

    Args:
        file: Tabular input file (csv, tsv, txt or parquet).
        nrows: Only include the first N data rows (preview).
        plain: Skip header formatting (bold header, frozen first row, filter).
    """
    options = ConversionOptions(
        file, nrows=nrows, to_temp=True, open_after=True, format_header=not plain
    )
    return _run(options)


@app.command
def convert(
    file: Path,
    /,
    *,
    nrows: _Nrows = None,
    plain: Annotated[bool, _Flag] = False,
    temp: Annotated[bool, _Flag] = False,
    open_after: Annotated[
        bool, Parameter(name="--open", negative_bool=(), show_default=False)
    ] = False,
) -> int:
    """Convert to an xlsx file next to the input file.

    An existing xlsx file is never overwritten; a numbered name is used instead.

    Args:
        file: Tabular input file (csv, tsv, txt or parquet).
        nrows: Only include the first N data rows (preview).
        plain: Skip header formatting (bold header, frozen first row, filter).
        temp: Write to the XlsxPress temp folder instead, which is cleaned up
            automatically.
        open_after: Open the xlsx file in Excel after converting.
    """
    options = ConversionOptions(
        file,
        nrows=nrows,
        to_temp=temp,
        open_after=open_after,
        format_header=not plain,
    )
    return _run(options)


def _run(options: ConversionOptions) -> int:
    """Run the conversion and report the result, without tracebacks."""
    options.cleanup_policy = load_config().cleanup
    try:
        result = run_conversion(options)
    except ConversionError as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    except OpenOutputError as error:
        _report_written(error.output_path, error.warnings)
        print(f"Error: {error}", file=sys.stderr)
        return 1
    _report_written(result.output_path, result.warnings)
    return 0


def _report_written(output_path: Path, warnings: tuple[str, ...]) -> None:
    for message in warnings:
        print(f"Warning: {message}", file=sys.stderr)
    print(f"Wrote {output_path}")


@context_menu_app.command
def register() -> int:
    """Add the XlsxPress entries to the context menu of csv, tsv, txt and parquet files.

    The entries run XlsxPress from this installation. Running the command again
    replaces existing entries.
    """
    if not _check_platform():
        return 1
    try:
        pythonw = context_menu.find_pythonw()
    except FileNotFoundError as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    if not context_menu.is_uv_tool_environment():
        print(
            f"Warning: XlsxPress is not running from a uv tool installation "
            f"('{sys.prefix}'). If this environment is deleted or moved, the "
            "context menu entries stop working. Install XlsxPress with "
            "'uv tool install xlsxpress' and register from there.",
            file=sys.stderr,
        )
    try:
        context_menu.register(pythonw)
    except OSError as error:
        print(
            f"Error: Could not write the context menu entries: {error}", file=sys.stderr
        )
        return 1
    extensions = ", ".join(context_menu.EXTENSIONS)
    print(f"Registered the context menu entries for {extensions} files:")
    for entry in context_menu.ENTRIES:
        print(f"  {entry.label}")
    print('On Windows 11, they are under "Show more options".')
    return 0


@context_menu_app.command
def unregister() -> int:
    """Remove the XlsxPress entries from the context menu."""
    if not _check_platform():
        return 1
    removed = context_menu.unregister()
    if removed:
        print(f"Removed {removed} context menu entries.")
    else:
        print("No context menu entries were registered.")
    return 0


@context_menu_app.command
def status() -> int:
    """Show the registered context menu entries and the config file location."""
    if not _check_platform():
        return 1
    entries = context_menu.read_entries()
    total = len(context_menu.EXTENSIONS) * len(context_menu.ENTRIES)
    if not entries:
        print("Context menu: not registered")
    else:
        print(f"Context menu: {len(entries)} of {total} entries registered")
        print("Commands:")
        for command in dict.fromkeys(entry.command for entry in entries):
            print(f"  {command}")
    _print_status_problems(entries, total)
    config_path = get_config_path()
    missing = "" if config_path.is_file() else " (not created yet)"
    print(f"Config file: {config_path}{missing}")
    return 0


def _check_platform() -> bool:
    if context_menu.is_supported_platform():
        return True
    print("Error: The context menu is only available on Windows.", file=sys.stderr)
    return False


def _print_status_problems(
    entries: list[context_menu.RegisteredEntry], total: int
) -> None:
    """Warn about missing entries and entries that run a missing program."""
    register_hint = "Run 'xlsxpress context-menu register' to fix it."
    if entries and len(entries) < total:
        print(f"Warning: Some entries are missing. {register_hint}", file=sys.stderr)
    programs = dict.fromkeys(entry.program for entry in entries)
    for program in programs:
        if program is None or not program.is_file():
            print(
                f"Warning: The entries run '{program}', which does not exist, so they "
                f"do not work. {register_hint}",
                file=sys.stderr,
            )
