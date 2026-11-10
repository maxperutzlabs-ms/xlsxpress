"""User settings shared by the CLI and the app: the config file and the log file.

Both live in the user data folder, e.g. `%LOCALAPPDATA%/perutzms/xlsxpress`. Temporary
xlsx files live elsewhere, see `xlsxpress.core.paths.get_temp_dir`.
"""

from __future__ import annotations

import tomllib
from dataclasses import asdict, dataclass
from pathlib import Path

import platformdirs

from xlsxpress.core.cleanup import CleanupPolicy

_APPNAME = "xlsxpress"
_APPAUTHOR = "perutzms"


@dataclass(slots=True)
class DialogDefaults:
    """Default options of the conversion dialog and the context menu's "Open in Excel".

    "Open in Excel" uses all options except `open_after`: it always opens the file.

    Attributes:
        open_after: Open the xlsx file in Excel after converting in the dialog.
        preview: Only convert the first `nrows` data rows.
        nrows: The number of data rows to convert with `preview`.
        to_temp: Write to the temp folder instead of next to the input file.
        format_header: Format the header row (bold, wrapped text, frozen, with an
            autofilter).
    """

    open_after: bool = True
    preview: bool = False
    nrows: int = 100
    to_temp: bool = True
    format_header: bool = True


@dataclass(slots=True)
class AppConfig:
    """The contents of the config file.

    Attributes:
        dialog: Default conversion options of the dialog and the context menu.
        cleanup: When to delete old files from the temp folder.
    """

    dialog: DialogDefaults
    cleanup: CleanupPolicy


def default_config() -> AppConfig:
    """Return the config used when the config file has no other values."""
    return AppConfig(dialog=DialogDefaults(), cleanup=CleanupPolicy())


def load_config() -> AppConfig:
    """Load the config file, using defaults for anything missing or invalid.

    Never raises, because a broken config file must not break conversions. A
    section that isn't a table, a value of the wrong type, a negative number or
    an `nrows` of 0 is ignored.

    Returns:
        The loaded config.
    """
    config = default_config()
    try:
        with get_config_path().open("rb") as file:
            data = tomllib.load(file)
    except (OSError, tomllib.TOMLDecodeError):
        return config
    for section, target in (("dialog", config.dialog), ("cleanup", config.cleanup)):
        values = data.get(section, {})
        if not isinstance(values, dict):
            continue
        for key, value in values.items():
            if _is_valid_value(target, key, value):
                setattr(target, key, value)
    return config


def write_config(config: AppConfig, path: Path | None = None) -> None:
    """Write `config` as TOML, creating the data folder if needed.

    Args:
        config: The config to write.
        path: The file to write. If None (default), writes the config file, see
            `get_config_path`.

    Raises:
        OSError: If the folder or the file cannot be written.
    """
    path = path or get_config_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    lines: list[str] = ["# XlsxPress configuration", ""]
    for section_name, section in (
        ("dialog", config.dialog),
        ("cleanup", config.cleanup),
    ):
        lines.append(f"[{section_name}]")
        for key, value in asdict(section).items():
            # The standard library can read TOML but not write it.
            lines.append(
                f"{key} = {str(value).lower() if isinstance(value, bool) else value}"
            )
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def get_data_dir() -> Path:
    """Return the user data folder, e.g. `%LOCALAPPDATA%/perutzms/xlsxpress`."""
    return Path(platformdirs.user_data_dir(_APPNAME, _APPAUTHOR))


def get_config_path() -> Path:
    """Return the path of the config file."""
    return get_data_dir() / "config.toml"


def get_log_path() -> Path:
    """Return the path of the app's log file."""
    return get_data_dir() / "xlsxpress.log"


def _is_valid_value(target: object, key: str, value: object) -> bool:
    """Whether `value` has the type of the default and is in range.

    The type must match exactly, because `true` is an `int` in Python. Integers
    must be at least 0, and `nrows` at least 1.
    """
    if not hasattr(target, key) or type(value) is not type(getattr(target, key)):
        return False
    if isinstance(value, int) and not isinstance(value, bool):
        return value >= (1 if key == "nrows" else 0)
    return True
