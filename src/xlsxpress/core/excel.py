"""Opening files in the platform's default application (Excel on Windows).

Also shows files in the platform's file manager.
"""

from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path


def open_in_default_app(path: Path) -> None:
    """Open `path` with the OS default application, non-blocking."""
    if sys.platform == "win32":
        os.startfile(path)  # noqa: S606
    elif sys.platform == "darwin":
        subprocess.Popen(["open", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path)])


def show_in_folder(path: Path) -> None:
    """Show `path` in the file manager, selected where possible, non-blocking.

    Args:
        path: The file to show.

    Raises:
        OSError: If the file manager cannot be started.
    """
    if sys.platform == "win32":
        # Explorer expects the path as a separate argument after "/select,".
        subprocess.Popen(["explorer", "/select,", str(path)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", "-R", str(path)])
    else:
        subprocess.Popen(["xdg-open", str(path.parent)])
