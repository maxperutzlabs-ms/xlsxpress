"""Entry point for the pyapp executable (NOT the standalone CLI).

Routing:
    (no args)        -> launcher GUI (after install check/repair)
    open <file>      -> convert to appdata temp + open in Excel
    dialog <file>    -> options dialog GUI
    --version        -> print version (used by the launcher after updates)
"""

from __future__ import annotations

import argparse
import logging
import sys
from pathlib import Path
from typing import Literal

from xlsxpress.app import appdata, config, installation

logger = logging.getLogger("xlsxpress.app")


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="xlsxpress-app", add_help=True)
    parser.add_argument(
        "--version", action="version", version=installation.package_version()
    )
    subparsers = parser.add_subparsers(dest="command")
    for name in ("open", "dialog"):
        sub = subparsers.add_parser(name)
        sub.add_argument("filename", type=Path)
    args = parser.parse_args(argv)

    _setup_logging()
    try:
        if args.command == "open":
            return _run_open(args.filename)
        if args.command == "dialog":
            return _run_dialog(args.filename)
        return _run_launcher()
    except Exception:  # noqa: BLE001 - last resort: log + GUI-visible error
        logger.exception("Unhandled error")
        _show_error_box(f"XlsxPress failed. See log file:\n{appdata.get_log_path()}")
        return 1


def _run_open(filename: Path) -> int:
    """Context menu 'Open in Excel': convert with the saved default options.

    Conversion errors and warnings are shown in a message box, because the context
    menu has no console.
    """
    from xlsxpress.core.errors import ConversionError
    from xlsxpress.core.options import ConversionOptions
    from xlsxpress.core.pipeline import run_conversion

    appdata.ensure_directory_structure()
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
    if result.warnings:
        _show_warning_box("\n\n".join(result.warnings))
    return 0


def _run_dialog(filename: Path) -> int:
    from xlsxpress.app.gui.dialog import run_dialog

    appdata.ensure_directory_structure()
    run_dialog(filename, config.load_config())
    return 0


def _run_launcher() -> int:
    from xlsxpress.app.gui.launcher import run_launcher

    state = installation.check_and_repair()  # step 1: verify/fix before GUI
    run_launcher(state)
    return 0


def _setup_logging() -> None:
    appdata.ensure_directory_structure()
    logging.basicConfig(
        filename=appdata.get_log_path(),
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
    )


def _show_error_box(message: str) -> None:
    """Show an error in a message box; the GUI app has no visible console."""
    _show_message_box("error", message)


def _show_warning_box(message: str) -> None:
    """Show conversion warnings in a message box."""
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


if __name__ == "__main__":
    sys.exit(main())
