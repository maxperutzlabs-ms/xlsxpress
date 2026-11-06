"""Application data directory layout. Only the `app` package may import this."""

from __future__ import annotations

from pathlib import Path

import platformdirs

from xlsxpress.core import paths

APPNAME = "xlsxpress"
APPAUTHOR = "perutzms"


def get_appdata_dir() -> Path:
    """Return the app data directory, e.g. %LOCALAPPDATA%/perutzms/xlsxpress."""
    return Path(platformdirs.user_data_dir(APPNAME, APPAUTHOR))


def get_temp_dir() -> Path:
    """Return the temp folder shared with the CLI (see `paths.get_temp_dir`)."""
    return paths.get_temp_dir()


def get_config_path() -> Path:
    return get_appdata_dir() / "config.toml"


def get_install_manifest_path() -> Path:
    return get_appdata_dir() / "install.toml"


def get_installed_exe_path() -> Path:
    """Path where the pyapp executable is expected after installation."""
    return get_appdata_dir() / "xlsxpress.exe"


def get_log_path() -> Path:
    return get_appdata_dir() / "xlsxpress.log"


def ensure_directory_structure() -> None:
    """Create the appdata directory tree (idempotent)."""
    get_appdata_dir().mkdir(parents=True, exist_ok=True)
    get_temp_dir()
