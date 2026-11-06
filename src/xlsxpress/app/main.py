"""Entry point of the context menu entries, run as `pythonw -m xlsxpress.app`.

Commands:
    open FILE    Convert with the saved default options and open the result.
    dialog FILE  Show the conversion dialog.

The app has no console, so errors and warnings are shown in message boxes, and
unexpected errors are logged to the log file.
"""

from __future__ import annotations

import argparse
import logging
import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Literal

from xlsxpress import config

logger = logging.getLogger("xlsxpress.app")


def main(argv: Sequence[str] | None = None) -> int:
    """Run the app.

    Args:
        argv: Command line arguments. If None (default), uses the process arguments.

    Returns:
        The exit code: 0 on success, 1 on an error.
    """
    parser = argparse.ArgumentParser(prog="python -m xlsxpress.app")
    subparsers = parser.add_subparsers(dest="command", required=True)
    for name in ("open", "dialog"):
        subparsers.add_parser(name).add_argument("filename", type=Path)
    args = parser.parse_args(argv)

    _setup_logging()
    try:
        if args.command == "open":
            return _run_open(args.filename)
        return _run_dialog(args.filename)
    except Exception:  # noqa: BLE001 - last resort: log + GUI-visible error
        logger.exception("Unhandled error")
        _show_error_box(f"XlsxPress failed. See the log file:\n{config.get_log_path()}")
        return 1


def _run_open(filename: Path) -> int:
    """Context menu 'Open in Excel': convert with the saved default options."""
    from xlsxpress.core.errors import ConversionError, OpenOutputError
    from xlsxpress.core.options import ConversionOptions
    from xlsxpress.core.pipeline import run_conversion

    app_config = config.load_config()
    options = ConversionOptions(
        input_path=filename,
        nrows=app_config.dialog.nrows if app_config.dialog.preview else None,
        to_temp=app_config.dialog.to_temp,
        open_after=app_config.dialog.open_after,
        format_header=app_config.dialog.format_header,
        cleanup_policy=app_config.cleanup,
    )
    try:
        result = run_conversion(options)
    except ConversionError as error:
        logger.exception("Conversion of %s failed", filename)
        _show_error_box(str(error))
        return 1
    except OpenOutputError as error:
        logger.exception("Opening %s failed", error.output_path)
        if error.warnings:
            _show_warning_box("\n\n".join(error.warnings))
        _show_error_box(str(error))
        return 1
    if result.warnings:
        _show_warning_box("\n\n".join(result.warnings))
    return 0


def _run_dialog(filename: Path) -> int:
    from xlsxpress.app.gui.dialog import run_dialog

    run_dialog(filename, config.load_config())
    return 0


def _setup_logging() -> None:
    log_path = config.get_log_path()
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        filename=log_path,
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


def _show_error_box(message: str) -> None:
    _show_message_box("error", message)


def _show_warning_box(message: str) -> None:
    _show_message_box("warning", message)


def _show_message_box(kind: Literal["error", "warning"], message: str) -> None:
    try:
        import tkinter
        from tkinter import messagebox

        root = tkinter.Tk()
        root.withdraw()
        if kind == "error":
            messagebox.showerror("XlsxPress", message)
        else:
            messagebox.showwarning("XlsxPress", message)
        root.destroy()
    except Exception:  # noqa: BLE001 - headless fallback
        print(message, file=sys.stderr)
