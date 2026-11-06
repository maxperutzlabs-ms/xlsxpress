"""Cleanup of old xlsx files in the temp folder and of leftovers in the work folder.

Temp files are disposable: a quick way to look at a file in Excel. Cleanup deletes
them by age and count, whether or not they were edited. The work folder holds
unfinished workbooks, which only stay behind when a conversion was killed. Cleanup
tolerates files that disappear or are locked while it runs, so several conversions
can run at the same time.
"""

from __future__ import annotations

import time
from dataclasses import dataclass
from pathlib import Path

#: Age after which a file in the work folder is certainly not in use any more.
WORK_FILE_MAX_AGE_DAYS = 1

_SECONDS_PER_DAY = 86_400
#: Prefix of the lock files Excel creates next to an open workbook.
_EXCEL_LOCK_PREFIX = "~$"


@dataclass(slots=True)
class CleanupPolicy:
    """When to delete files from the temp folder.

    Attributes:
        max_age_days: Delete files older than this many days.
        max_files: If more files remain, delete the oldest until this many are left.
    """

    max_age_days: int = 14
    max_files: int = 100


def clean_temp_dir(temp_dir: Path, policy: CleanupPolicy) -> int:
    """Delete old xlsx files from the temp folder according to `policy`.

    Xlsx files last modified more than `policy.max_age_days` ago are deleted. If
    more than `policy.max_files` remain, the oldest are deleted until the limit is
    met. Other files, subfolders and Excel's lock files (`~$name.xlsx`) are not
    touched. Files that are locked, for example because they are open in Excel, or
    that disappear while cleanup runs, are skipped.

    Args:
        temp_dir: The temp folder.
        policy: The cleanup policy.

    Returns:
        The number of deleted files.
    """
    if not temp_dir.is_dir():
        return 0
    candidates: list[tuple[int, Path]] = []
    for file in temp_dir.glob("*.xlsx"):
        if file.name.startswith(_EXCEL_LOCK_PREFIX):
            continue
        try:
            if file.is_file():
                candidates.append((file.stat().st_mtime_ns, file))
        except OSError:
            continue
    candidates.sort()

    cutoff_ns = int((time.time() - policy.max_age_days * _SECONDS_PER_DAY) * 1e9)
    expired = [file for mtime_ns, file in candidates if mtime_ns < cutoff_ns]
    remaining = [file for mtime_ns, file in candidates if mtime_ns >= cutoff_ns]
    excess = remaining[: max(0, len(remaining) - policy.max_files)]
    return sum(_try_delete(file) for file in expired + excess)


def clean_work_dir(work_dir: Path) -> int:
    """Delete files that interrupted conversions left in the work folder.

    A conversion deletes its own work files when it ends, even when it fails. Only a
    process that was killed, for example by closing the dialog during a conversion,
    leaves files behind. Every file directly in the folder that was last modified
    more than `WORK_FILE_MAX_AGE_DAYS` ago is deleted. Younger files may belong to a
    conversion that is still running and are kept. Subfolders and locked files are
    not touched.

    Args:
        work_dir: The work folder, see `paths.get_work_dir`.

    Returns:
        The number of deleted files.
    """
    if not work_dir.is_dir():
        return 0
    cutoff_ns = int((time.time() - WORK_FILE_MAX_AGE_DAYS * _SECONDS_PER_DAY) * 1e9)
    expired: list[Path] = []
    for file in work_dir.iterdir():
        try:
            if file.is_file() and file.stat().st_mtime_ns < cutoff_ns:
                expired.append(file)
        except OSError:
            continue
    return sum(_try_delete(file) for file in expired)


def _try_delete(path: Path) -> bool:
    try:
        path.unlink()
    except OSError:
        return False
    return True
