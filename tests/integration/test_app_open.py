from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from xlsxpress.app import main as app_main
from xlsxpress.config import AppConfig, CleanupPolicy, DialogDefaults, write_config
from xlsxpress.core import excel, paths

pytestmark = pytest.mark.integration


def _count_data_rows(xlsx_path: Path) -> int:
    with ZipFile(xlsx_path) as workbook:
        sheet = workbook.read("xl/worksheets/sheet1.xml").decode()
    return sheet.count("<row ") - 1


def _fail_on_error_box(message: str) -> None:
    pytest.fail(f"Error box shown: {message}")


@pytest.fixture(autouse=True)
def no_logging_or_error_boxes(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(app_main, "_setup_logging", lambda: None)
    monkeypatch.setattr(app_main, "_show_error_box", _fail_on_error_box)


@pytest.fixture(autouse=True)
def opened(monkeypatch: pytest.MonkeyPatch) -> list[Path]:
    paths_opened: list[Path] = []
    monkeypatch.setattr(excel, "open_in_default_app", paths_opened.append)
    return paths_opened


class TestMainOpen:
    @pytest.mark.parametrize("open_after", [False, True])
    def test_open_always_opens_the_file_whatever_the_saved_open_after_option(
        self, tmp_path: Path, opened: list[Path], open_after: bool
    ) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text("a\n1\n", encoding="utf-8")
        write_config(AppConfig(dialog=DialogDefaults(open_after=open_after), cleanup=CleanupPolicy()))

        exit_code = app_main.main(["open", str(input_path)])

        assert exit_code == 0
        assert opened == list(paths.get_temp_dir().glob("*.xlsx"))
        assert len(opened) == 1

    @pytest.mark.parametrize(("preview", "expected_rows"), [(False, 150), (True, 100)])
    def test_open_limits_rows_only_when_preview_is_enabled(
        self, tmp_path: Path, preview: bool, expected_rows: int
    ) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text("value\n" + "".join(f"{i}\n" for i in range(150)), encoding="utf-8")
        defaults = DialogDefaults(open_after=False, preview=preview, nrows=100)
        write_config(AppConfig(dialog=defaults, cleanup=CleanupPolicy()))

        exit_code = app_main.main(["open", str(input_path)])

        outputs = list(paths.get_temp_dir().glob("*.xlsx"))
        assert exit_code == 0
        assert len(outputs) == 1
        assert _count_data_rows(outputs[0]) == expected_rows

    def test_open_shows_the_conversion_error_message_in_an_error_box(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text("a\tb\n1\t2\t3\n", encoding="utf-8")
        messages: list[str] = []
        monkeypatch.setattr(app_main, "_show_error_box", messages.append)

        exit_code = app_main.main(["open", str(input_path)])

        assert exit_code == 1
        assert len(messages) == 1
        assert "more tab-separated fields than the header row" in messages[0]

    def test_open_shows_conversion_warnings_in_a_warning_box(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text("a\ta\n1\t2\n", encoding="utf-8")
        write_config(AppConfig(dialog=DialogDefaults(open_after=False), cleanup=CleanupPolicy()))
        messages: list[str] = []
        monkeypatch.setattr(app_main, "_show_warning_box", messages.append)

        exit_code = app_main.main(["open", str(input_path)])

        assert exit_code == 0
        assert len(messages) == 1
        assert "appear more than once" in messages[0]

    def test_open_shows_that_the_file_was_written_when_it_cannot_be_opened(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text("a\n1\n", encoding="utf-8")

        def fail_to_open(path: Path) -> None:
            raise OSError(1155, "No application is associated with the specified file for this operation")

        monkeypatch.setattr(excel, "open_in_default_app", fail_to_open)
        messages: list[str] = []
        monkeypatch.setattr(app_main, "_show_error_box", messages.append)

        exit_code = app_main.main(["open", str(input_path)])

        outputs = list(paths.get_temp_dir().glob("*.xlsx"))
        assert exit_code == 1
        assert len(outputs) == 1
        assert len(messages) == 1
        assert str(outputs[0]) in messages[0]
        assert "could not be opened" in messages[0]
        assert "failed" not in messages[0]
