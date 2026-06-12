"""Temp directory housekeeping."""

from __future__ import annotations

import time
from pathlib import Path

from xlsxpress.app.config import CleanupPolicy


def clean_temp_dir(temp_dir: Path, policy: CleanupPolicy) -> int:
    """Delete temp files per policy; return the number of deleted files.

    Policy: delete files older than `max_age_days`; additionally, if more than
    `max_files` remain, delete the oldest until the cap is met. Files locked by Excel
    simply fail to delete and are skipped.
    """
    if not temp_dir.is_dir():
        return 0
    files = sorted(
        (f for f in temp_dir.iterdir() if f.is_file()),
        key=lambda f: f.stat().st_mtime,
    )
    deleted = 0
    cutoff = time.time() - policy.max_age_days * 86400
    keep: list[Path] = []
    for file in files:
        if file.stat().st_mtime < cutoff:
            deleted += _try_delete(file)
        else:
            keep.append(file)
    for file in keep[: max(0, len(keep) - policy.max_files)]:
        deleted += _try_delete(file)
    return deleted


def _try_delete(file: Path) -> int:
    try:
        file.unlink()
        return 1
    except OSError:
        return 0
