from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile

import pytest

from xlsxpress.app import appdata
from xlsxpress.app import main as app_main
from xlsxpress.app.config import AppConfig, CleanupPolicy, DialogDefaults, write_config

pytestmark = pytest.mark.integration


def _count_data_rows(xlsx_path: Path) -> int:
    with ZipFile(xlsx_path) as workbook:
        sheet = workbook.read("xl/worksheets/sheet1.xml").decode()
    return sheet.count("<row ") - 1


def _fail_on_error_box(message: str) -> None:
    pytest.fail(f"Error box shown: {message}")


@pytest.fixture
def appdata_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    directory = tmp_path / "appdata"
    monkeypatch.setattr(appdata, "get_appdata_dir", lambda: directory)
    monkeypatch.setattr(app_main, "_setup_logging", lambda: None)
    monkeypatch.setattr(app_main, "_show_error_box", _fail_on_error_box)
    appdata.ensure_directory_structure()
    return directory


class TestMainOpen:
    @pytest.mark.parametrize(("preview", "expected_rows"), [(False, 150), (True, 100)])
    def test_open_limits_rows_only_when_preview_is_enabled(
        self, tmp_path: Path, appdata_dir: Path, preview: bool, expected_rows: int
    ) -> None:
        input_path = tmp_path / "data.tsv"
        input_path.write_text("value\n" + "".join(f"{i}\n" for i in range(150)), encoding="utf-8")
        defaults = DialogDefaults(open_after=False, preview=preview, nrows=100)
        write_config(AppConfig(dialog=defaults, cleanup=CleanupPolicy()))

        exit_code = app_main.main(["open", str(input_path)])

        outputs = list(appdata.get_temp_dir().glob("*.xlsx"))
        assert exit_code == 0
        assert len(outputs) == 1
        assert _count_data_rows(outputs[0]) == expected_rows
