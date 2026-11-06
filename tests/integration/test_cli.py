from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import pytest

from xlsxpress import __VERSION__
from xlsxpress.cli import main
from xlsxpress.config import AppConfig, CleanupPolicy, DialogDefaults, write_config
from xlsxpress.core import excel, paths

pytestmark = pytest.mark.integration


@pytest.fixture
def input_file(tmp_path: Path) -> Path:
    path = tmp_path / "input" / "data.tsv"
    path.parent.mkdir()
    path.write_text("id\tvalue\n007\t1.5\n008\t2\n009\t3\n", encoding="utf-8")
    return path


@pytest.fixture
def opened(monkeypatch: pytest.MonkeyPatch) -> list[Path]:
    paths_opened: list[Path] = []
    monkeypatch.setattr(excel, "open_in_default_app", paths_opened.append)
    return paths_opened


def _temp_files() -> list[Path]:
    return sorted(paths.get_temp_dir().glob("*.xlsx"))


class TestConvert:
    def test_convert_reports_the_written_file_and_an_error_when_opening_fails(
        self, input_file: Path, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
    ) -> None:
        def fail_to_open(path: Path) -> None:
            raise OSError(1155, "No application is associated with the specified file for this operation")

        monkeypatch.setattr(excel, "open_in_default_app", fail_to_open)

        exit_code = main(["convert", str(input_file), "--open"])

        output_path = input_file.with_suffix(".xlsx")
        captured = capsys.readouterr()
        assert exit_code == 1
        assert captured.out == f"Wrote {output_path}\n"
        assert captured.err.startswith("Error: The xlsx file was written")
        assert "could not be opened" in captured.err
        assert output_path.exists()

    def test_convert_writes_next_to_the_input_and_prints_the_path(
        self, input_file: Path, opened: list[Path], capsys: pytest.CaptureFixture[str], xlsx_reader: Callable
    ) -> None:
        exit_code = main(["convert", str(input_file)])

        output_path = input_file.with_suffix(".xlsx")
        captured = capsys.readouterr()
        assert exit_code == 0
        assert captured.out == f"Wrote {output_path}\n"
        assert captured.err == ""
        assert xlsx_reader(output_path).cells["A2"].value == "007"
        assert opened == []

    def test_convert_with_temp_and_open_writes_to_the_temp_folder_and_opens_it(
        self, input_file: Path, opened: list[Path]
    ) -> None:
        exit_code = main(["convert", str(input_file), "--temp", "--open"])

        assert exit_code == 0
        assert opened == _temp_files()
        assert len(opened) == 1
        assert not input_file.with_suffix(".xlsx").exists()

    @pytest.mark.parametrize("flag", ["--nrows", "-n"])
    def test_convert_limits_rows(self, input_file: Path, xlsx_reader: Callable, flag: str) -> None:
        main(["convert", str(input_file), flag, "2"])

        cells = xlsx_reader(input_file.with_suffix(".xlsx")).cells
        assert "A3" in cells
        assert "A4" not in cells

    def test_convert_with_plain_skips_header_formatting(self, input_file: Path, xlsx_reader: Callable) -> None:
        main(["convert", str(input_file), "--plain"])

        sheet = xlsx_reader(input_file.with_suffix(".xlsx"))
        assert not sheet.cells["A1"].bold
        assert sheet.frozen_rows == 0

    def test_convert_prints_warnings_to_stderr(self, tmp_path: Path, capsys: pytest.CaptureFixture[str]) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text("a\ta\n1\t2\n", encoding="utf-8")

        exit_code = main(["convert", str(input_path)])

        captured = capsys.readouterr()
        assert exit_code == 0
        assert captured.err.startswith("Warning: Column names appear more than once")
        assert captured.out.startswith("Wrote ")

    @pytest.mark.parametrize(("content", "name"), [(None, "missing.tsv"), ("a\tb\n1\t2\t3\n", "data.tsv")])
    def test_convert_prints_conversion_errors_and_exits_with_1(
        self, tmp_path: Path, capsys: pytest.CaptureFixture[str], content: str | None, name: str
    ) -> None:
        input_path = tmp_path / name
        if content is not None:
            input_path.write_text(content, encoding="utf-8")

        exit_code = main(["convert", str(input_path)])

        captured = capsys.readouterr()
        assert exit_code == 1
        assert captured.err.startswith("Error: ")
        assert "Traceback" not in captured.err
        assert captured.out == ""

    def test_convert_rejects_a_row_count_below_one(self, input_file: Path) -> None:
        with pytest.raises(SystemExit) as exit_info:
            main(["convert", str(input_file), "--nrows", "0"])

        assert exit_info.value.code not in (0, None)
        assert not input_file.with_suffix(".xlsx").exists()


class TestOpen:
    def test_open_writes_to_the_temp_folder_and_opens_it(self, input_file: Path, opened: list[Path]) -> None:
        exit_code = main(["open", str(input_file)])

        assert exit_code == 0
        assert opened == _temp_files()
        assert len(opened) == 1
        assert not input_file.with_suffix(".xlsx").exists()

    def test_open_uses_the_cleanup_policy_from_the_app_config(self, input_file: Path, opened: list[Path]) -> None:
        write_config(AppConfig(dialog=DialogDefaults(), cleanup=CleanupPolicy(max_files=1)))

        for _ in range(3):
            main(["open", str(input_file)])

        assert len(_temp_files()) == 2


class TestVersion:
    def test_version_prints_the_package_version(self, capsys: pytest.CaptureFixture[str]) -> None:
        exit_code = main(["--version"])

        assert exit_code == 0
        assert capsys.readouterr().out.strip() == __VERSION__
