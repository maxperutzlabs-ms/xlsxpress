"""User config: hardcoded defaults, creation, loading with fallback."""

from __future__ import annotations

import tomllib
from dataclasses import asdict, dataclass
from pathlib import Path

from xlsxpress.app import appdata


@dataclass(slots=True)
class DialogDefaults:
    """Default state of the dialog GUI controls (also the package defaults)."""

    open_after: bool = True
    preview: bool = False
    nrows: int = 100
    to_temp: bool = True
    format_header: bool = True


@dataclass(slots=True)
class CleanupPolicy:
    """When to delete files from the appdata temp directory."""

    max_age_days: int = 14
    max_files: int = 100


@dataclass(slots=True)
class AppConfig:
    dialog: DialogDefaults
    cleanup: CleanupPolicy


def default_config() -> AppConfig:
    return AppConfig(dialog=DialogDefaults(), cleanup=CleanupPolicy())


def load_config() -> AppConfig:
    """Load config from appdata, falling back to defaults for anything missing or
    unreadable. Never raises: a broken config must not break context-menu conversions.
    """
    config = default_config()
    path = appdata.get_config_path()
    try:
        with path.open("rb") as file:
            data = tomllib.load(file)
    except OSError, tomllib.TOMLDecodeError:
        return config
    for section, target in (("dialog", config.dialog), ("cleanup", config.cleanup)):
        for key, value in data.get(section, {}).items():
            if hasattr(target, key) and isinstance(value, type(getattr(target, key))):
                setattr(target, key, value)
    return config


def write_config(config: AppConfig, path: Path | None = None) -> None:
    """Write the config as TOML (stdlib cannot write TOML; format manually)."""
    path = path or appdata.get_config_path()
    lines: list[str] = ["# XlsxPress configuration", ""]
    for section_name, section in (
        ("dialog", config.dialog),
        ("cleanup", config.cleanup),
    ):
        lines.append(f"[{section_name}]")
        for key, value in asdict(section).items():
            lines.append(
                f"{key} = {str(value).lower() if isinstance(value, bool) else value}"
            )
        lines.append("")
    path.write_text("\n".join(lines), encoding="utf-8")


def ensure_config_exists() -> None:
    if not appdata.get_config_path().exists():
        write_config(default_config())
