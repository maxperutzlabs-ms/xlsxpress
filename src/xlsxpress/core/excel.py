"""Opening files in the platform's default application (Excel on Windows)."""

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
