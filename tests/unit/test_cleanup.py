from __future__ import annotations

import os
import time
from pathlib import Path

import pytest

from xlsxpress.core.cleanup import CleanupPolicy, clean_temp_dir, clean_work_dir

pytestmark = pytest.mark.unit

_DAY_NS = 86_400 * 10**9


def _temp_file(temp_dir: Path, name: str, age_days: float = 0) -> Path:
    path = temp_dir / name
    path.write_bytes(b"xlsx content")
    mtime_ns = time.time_ns() - int(age_days * _DAY_NS)
    os.utime(path, ns=(mtime_ns, mtime_ns))
    return path


class TestCleanTempDir:
    def test_clean_temp_dir_deletes_files_older_than_the_maximum_age(self, tmp_path: Path) -> None:
        old = _temp_file(tmp_path, "old.xlsx", age_days=15)
        new = _temp_file(tmp_path, "new.xlsx", age_days=13)

        deleted = clean_temp_dir(tmp_path, CleanupPolicy(max_age_days=14))

        assert deleted == 1
        assert not old.exists()
        assert new.exists()

    def test_clean_temp_dir_deletes_the_oldest_files_above_the_maximum_count(self, tmp_path: Path) -> None:
        files = [_temp_file(tmp_path, f"{age}.xlsx", age_days=age) for age in (3, 1, 2)]

        deleted = clean_temp_dir(tmp_path, CleanupPolicy(max_files=1))

        assert deleted == 2
        assert [file.exists() for file in files] == [False, True, False]

    def test_clean_temp_dir_ignores_other_files_folders_and_excel_lock_files(self, tmp_path: Path) -> None:
        kept = [
            _temp_file(tmp_path, "notes.txt", age_days=30),
            _temp_file(tmp_path, "~$data.xlsx", age_days=30),
        ]
        folder = tmp_path / "folder.xlsx"
        folder.mkdir()

        deleted = clean_temp_dir(tmp_path, CleanupPolicy(max_age_days=1, max_files=0))

        assert deleted == 0
        assert all(path.exists() for path in [*kept, folder])

    def test_clean_temp_dir_skips_files_that_cannot_be_deleted(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        locked = _temp_file(tmp_path, "locked.xlsx", age_days=30)
        old = _temp_file(tmp_path, "old.xlsx", age_days=30)
        original_unlink = Path.unlink

        def unlink(self: Path, missing_ok: bool = False) -> None:
            if self == locked:
                raise PermissionError(13, "The file is open in Excel")
            original_unlink(self, missing_ok=missing_ok)

        monkeypatch.setattr(Path, "unlink", unlink)

        assert clean_temp_dir(tmp_path, CleanupPolicy(max_age_days=1)) == 1
        assert locked.exists()
        assert not old.exists()

    def test_clean_temp_dir_returns_zero_for_a_missing_folder(self, tmp_path: Path) -> None:
        assert clean_temp_dir(tmp_path / "missing", CleanupPolicy()) == 0


class TestCleanWorkDir:
    def test_clean_work_dir_deletes_files_of_any_name_older_than_one_day(self, tmp_path: Path) -> None:
        old = [_temp_file(tmp_path, "tmpab12.xlsx", age_days=1.1), _temp_file(tmp_path, "tmpcd34", age_days=3)]
        recent = _temp_file(tmp_path, "tmpef56.xlsx", age_days=0.9)

        deleted = clean_work_dir(tmp_path)

        assert deleted == 2
        assert not any(path.exists() for path in old)
        assert recent.exists()

    def test_clean_work_dir_skips_subfolders_and_locked_files(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        locked = _temp_file(tmp_path, "locked.xlsx", age_days=2)
        folder = tmp_path / "folder"
        folder.mkdir()
        original_unlink = Path.unlink

        def unlink(self: Path, missing_ok: bool = False) -> None:
            if self == locked:
                raise PermissionError(13, "The file is in use")
            original_unlink(self, missing_ok=missing_ok)

        monkeypatch.setattr(Path, "unlink", unlink)

        assert clean_work_dir(tmp_path) == 0
        assert locked.exists()
        assert folder.is_dir()

    def test_clean_work_dir_returns_zero_for_a_missing_folder(self, tmp_path: Path) -> None:
        assert clean_work_dir(tmp_path / "missing") == 0
