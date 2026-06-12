"""Filesystem path helpers."""

from __future__ import annotations

import tempfile
from pathlib import Path


def get_temp_dir() -> Path:
    """Return the directory for temporary xlsx output files.

    Currently a dedicated subdirectory of the platform temp directory. Later this will
    point into the app data directory, enabling cleanup policies (e.g. delete files
    older than N days).
    """
    temp_dir = Path(tempfile.gettempdir()) / "xlsxpress"
    temp_dir.mkdir(parents=True, exist_ok=True)
    return temp_dir


def resolve_output_path(input_path: Path, output_dir: Path) -> Path:
    """Return a non-colliding xlsx output path for `input_path` in `output_dir`.

    Uses the input file's stem ("data.tsv" -> "data.xlsx"). If that file already exists,
    appends a counter ("data (2).xlsx") so we never attempt to overwrite an existing
    file (it might be open and locked by Excel).
    """
    candidate = output_dir / f"{input_path.stem}.xlsx"
    counter = 2
    while candidate.exists():
        candidate = output_dir / f"{input_path.stem}-{counter:02d}.xlsx"
        counter += 1
    return candidate
