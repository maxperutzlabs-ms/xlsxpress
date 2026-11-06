from __future__ import annotations

from pathlib import Path

import pytest

from xlsxpress.app import context_menu
from xlsxpress.app.context_menu import ENTRIES, RegisteredEntry

pytestmark = pytest.mark.unit


class TestBuildCommand:
    def test_build_command_quotes_the_python_and_the_file_placeholder(self) -> None:
        pythonw = Path(r"C:\Users\Jane Doe\uv\tools\xlsxpress\Scripts\pythonw.exe")

        command = context_menu.build_command(pythonw, "open")

        assert command == rf'"{pythonw}" -m xlsxpress.app open "%1"'

    def test_build_command_round_trips_through_command_program(self) -> None:
        pythonw = Path(r"C:\Program Files\Python\pythonw.exe")

        assert context_menu.command_program(context_menu.build_command(pythonw, "dialog")) == pythonw


class TestCommandProgram:
    @pytest.mark.parametrize(
        ("command", "expected"),
        [
            (r'"C:\a b\pythonw.exe" -m xlsxpress.app open "%1"', r"C:\a b\pythonw.exe"),
            (r'C:\perutzms\xlsxpress\xlsxpress.exe open "%1"', r"C:\perutzms\xlsxpress\xlsxpress.exe"),
            (r'"C:\a\xlsxpress.exe" open "%1"', r"C:\a\xlsxpress.exe"),
            (r'"C:\unclosed quote\pythonw.exe -m', r"C:\unclosed"),
        ],
    )
    def test_command_program_returns_the_first_token(self, command: str, expected: str) -> None:
        assert context_menu.command_program(command) == Path(expected)

    @pytest.mark.parametrize("command", ["", "   "])
    def test_command_program_returns_none_for_an_empty_command(self, command: str) -> None:
        assert context_menu.command_program(command) is None


class TestRegisteredEntry:
    @pytest.mark.parametrize(
        ("command", "is_pyapp"),
        [
            (r'"C:\Users\x\AppData\Local\perutzms\xlsxpress\xlsxpress.exe" open "%1"', True),
            (r'"C:\Users\x\AppData\Local\perutzms\xlsxpress\XlsxPress.EXE" open "%1"', True),
            (r'"C:\tools\xlsxpress\Scripts\pythonw.exe" -m xlsxpress.app open "%1"', False),
            ("", False),
        ],
    )
    def test_is_pyapp_detects_the_old_executable(self, command: str, is_pyapp: bool) -> None:
        assert RegisteredEntry(".csv", ENTRIES[0], command).is_pyapp is is_pyapp


class TestFindPythonw:
    def test_find_pythonw_returns_pythonw_next_to_the_executable(self, tmp_path: Path) -> None:
        (tmp_path / "python.exe").touch()
        (tmp_path / "pythonw.exe").touch()

        assert context_menu.find_pythonw(tmp_path / "python.exe") == tmp_path / "pythonw.exe"

    def test_find_pythonw_raises_if_pythonw_is_missing(self, tmp_path: Path) -> None:
        (tmp_path / "python.exe").touch()

        with pytest.raises(FileNotFoundError, match="uv tool install xlsxpress"):
            context_menu.find_pythonw(tmp_path / "python.exe")


class TestIsUvToolEnvironment:
    def test_is_uv_tool_environment_is_true_with_a_uv_receipt(self, tmp_path: Path) -> None:
        (tmp_path / "uv-receipt.toml").write_text("[tool]\n", encoding="utf-8")

        assert context_menu.is_uv_tool_environment(tmp_path)

    def test_is_uv_tool_environment_is_false_without_a_uv_receipt(self, tmp_path: Path) -> None:
        (tmp_path / "pyvenv.cfg").touch()

        assert not context_menu.is_uv_tool_environment(tmp_path)
