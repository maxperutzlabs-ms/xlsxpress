"""Locations of output files: the shared temp folder and collision-free names."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import platformdirs

#: Upper bound for the counter appended to a name that already exists.
_MAX_ATTEMPTS = 1_000


def get_cache_dir() -> Path:
    """Return the user cache directory, e.g. `%LOCALAPPDATA%/xlsxpress/Cache`."""
    return platformdirs.user_cache_path("xlsxpress", appauthor=False)


def get_temp_dir() -> Path:
    """Return the folder for temporary xlsx files and create it if needed.

    The CLI, the context menu and the dialog all share this folder, so one cleanup
    policy covers every temporary file.
    """
    temp_dir = get_cache_dir() / "temp"
    temp_dir.mkdir(parents=True, exist_ok=True)
    return temp_dir


def get_work_dir() -> Path:
    """Return the folder for unfinished xlsx files and create it if needed.

    Workbooks are written here first and moved or copied to their output path only
    when complete, so an interrupted conversion never leaves a broken file under the
    output name. Files left behind by a killed process are deleted by
    `cleanup.clean_work_dir`.
    """
    work_dir = get_cache_dir() / "work"
    work_dir.mkdir(parents=True, exist_ok=True)
    return work_dir


def reserve_output_path(
    input_path: Path, output_dir: Path, unique: bool = False
) -> Path:
    """Create an empty xlsx file for `input_path` in `output_dir` and return its path.

    The name is taken from the input file (`data.tsv` becomes `data.xlsx`). With
    `unique=True` a timestamp is added (`data_2026-11-06_142530.xlsx`), because Excel
    cannot open two workbooks with the same name at once, even from different
    folders. If the name exists already, a counter is appended (`data-02.xlsx`), so
    an existing file is never overwritten.

    The file is created exclusively, so conversions running at the same time never
    get the same path. The caller moves or copies the finished workbook over the
    empty file, or deletes it if the conversion fails.

    Args:
        input_path: The file being converted.
        output_dir: The folder for the xlsx file. It must exist.
        unique: If True, adds a timestamp to the name. Defaults to False.

    Returns:
        The path of the created, empty file.

    Raises:
        FileExistsError: If no free name was found after many attempts.
        OSError: If the file cannot be created in `output_dir`.
    """
    stem = input_path.stem
    if unique:
        stem = f"{stem}_{dt.datetime.now():%Y-%m-%d_%H%M%S}"
    for attempt in range(1, _MAX_ATTEMPTS + 1):
        name = f"{stem}.xlsx" if attempt == 1 else f"{stem}-{attempt:02d}.xlsx"
        candidate = output_dir / name
        try:
            candidate.open("xb").close()
        except FileExistsError:
            continue
        return candidate
    raise FileExistsError(
        f"Could not find a free file name for {stem!r} in '{output_dir}'. "
        "Delete or move old xlsx files from this folder."
    )
