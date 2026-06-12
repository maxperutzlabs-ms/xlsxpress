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

from xlsxpress.app import appdata, cleanup, config, installation

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
    """Context menu 'Open in Excel': fixed options, appdata temp dir."""
    from xlsxpress.core.options import ConversionOptions
    from xlsxpress.core.pipeline import run_conversion

    appdata.ensure_directory_structure()
    app_config = config.load_config()
    cleanup.clean_temp_dir(appdata.get_temp_dir(), app_config.cleanup)
    options = ConversionOptions(
        input_path=filename,
        nrows=None,
        to_temp=True,
        open_after=True,
        format_header=True,
        temp_dir=appdata.get_temp_dir(),
    )
    run_conversion(options)
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
    """Stdout is discarded under PYAPP_IS_GUI; errors must be visible."""
    try:
        import tkinter
        from tkinter import messagebox

        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("XlsxPress", message)
        root.destroy()
    except Exception:  # noqa: BLE001 - headless fallback
        print(message, file=sys.stderr)


if __name__ == "__main__":
    sys.exit(main())
