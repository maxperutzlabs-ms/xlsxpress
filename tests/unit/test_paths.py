from __future__ import annotations

import re
from pathlib import Path

import pytest

from xlsxpress.core import paths
from xlsxpress.core.paths import get_temp_dir, reserve_output_path

pytestmark = pytest.mark.unit


class TestGetTempDir:
    def test_get_temp_dir_creates_the_temp_folder_inside_the_cache_dir(self, isolated_cache_dir: Path) -> None:
        temp_dir = get_temp_dir()

        assert temp_dir == isolated_cache_dir / "temp"
        assert temp_dir.is_dir()


class TestReserveOutputPath:
    def test_reserve_output_path_uses_the_input_name_and_creates_an_empty_file(self, tmp_path: Path) -> None:
        path = reserve_output_path(Path("data.tsv"), tmp_path)

        assert path == tmp_path / "data.xlsx"
        assert path.read_bytes() == b""

    def test_reserve_output_path_appends_a_counter_instead_of_overwriting(self, tmp_path: Path) -> None:
        (tmp_path / "data.xlsx").write_bytes(b"existing")

        first = reserve_output_path(Path("data.tsv"), tmp_path)
        second = reserve_output_path(Path("data.tsv"), tmp_path)

        assert (first.name, second.name) == ("data-02.xlsx", "data-03.xlsx")
        assert (tmp_path / "data.xlsx").read_bytes() == b"existing"

    def test_reserve_output_path_adds_a_timestamp_for_unique_names(self, tmp_path: Path) -> None:
        first = reserve_output_path(Path("data.tsv"), tmp_path, unique=True)
        second = reserve_output_path(Path("data.tsv"), tmp_path, unique=True)

        assert re.fullmatch(r"data_\d{4}-\d{2}-\d{2}_\d{6}\.xlsx", first.name)
        assert first != second

    def test_reserve_output_path_raises_when_no_free_name_is_left(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        monkeypatch.setattr(paths, "_MAX_ATTEMPTS", 2)
        (tmp_path / "data.xlsx").touch()
        (tmp_path / "data-02.xlsx").touch()

        with pytest.raises(FileExistsError, match="free file name for 'data'"):
            reserve_output_path(Path("data.tsv"), tmp_path)
