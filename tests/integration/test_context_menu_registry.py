from __future__ import annotations

import sys
from collections.abc import Iterator
from pathlib import Path

import pytest

from xlsxpress.app import context_menu
from xlsxpress.cli import main

pytestmark = pytest.mark.integration

windows_only = pytest.mark.skipif(sys.platform != "win32", reason="The context menu uses the Windows registry")

_TEST_ROOT = r"Software\xlsxpress-tests"


def _read_value(key_path: str, name: str) -> str | None:
    import winreg

    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            return winreg.QueryValueEx(key, name)[0]
    except OSError:
        return None


@pytest.fixture
def test_registry_root(monkeypatch: pytest.MonkeyPatch) -> Iterator[str]:
    context_menu._delete_key_tree(_TEST_ROOT)
    monkeypatch.setattr(context_menu, "_REGISTRY_ROOT", _TEST_ROOT)
    monkeypatch.setattr(context_menu, "_notify_shell_changed", lambda: None)
    yield _TEST_ROOT
    context_menu._delete_key_tree(_TEST_ROOT)


@pytest.fixture
def pythonw(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    path = tmp_path / "env" / "Scripts" / "pythonw.exe"
    path.parent.mkdir(parents=True)
    path.touch()
    monkeypatch.setattr(context_menu, "find_pythonw", lambda: path)
    return path


@pytest.fixture
def tool_environment(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr(context_menu, "is_uv_tool_environment", lambda: True)


@windows_only
class TestRegistryRoundTrip:
    def test_register_writes_every_entry_and_unregister_removes_them(
        self, test_registry_root: str, pythonw: Path
    ) -> None:
        context_menu.register(pythonw)

        entries = context_menu.read_entries()
        assert len(entries) == len(context_menu.EXTENSIONS) * len(context_menu.ENTRIES)
        assert {(entry.extension, entry.entry, entry.command) for entry in entries} == {
            (extension, entry, context_menu.build_command(pythonw, entry.app_command))
            for extension in context_menu.EXTENSIONS
            for entry in context_menu.ENTRIES
        }
        assert context_menu.unregister() == len(entries)
        assert context_menu.read_entries() == []
        assert context_menu.unregister() == 0

    def test_register_sets_the_menu_label(self, test_registry_root: str, pythonw: Path) -> None:
        context_menu.register(pythonw)

        entry = context_menu.ENTRIES[1]
        key_path = rf"{test_registry_root}\.parquet\shell\{entry.registry_key}"
        assert _read_value(key_path, "") == entry.label


@windows_only
class TestContextMenuCommands:
    def test_register_status_and_unregister_round_trip(
        self, test_registry_root: str, pythonw: Path, tool_environment: None, capsys: pytest.CaptureFixture[str]
    ) -> None:
        assert main(["context-menu", "register"]) == 0
        registered = capsys.readouterr()
        assert main(["context-menu", "status"]) == 0
        status = capsys.readouterr()
        assert main(["context-menu", "unregister"]) == 0
        unregistered = capsys.readouterr()
        assert main(["context-menu", "status"]) == 0
        status_after = capsys.readouterr()

        assert "Open in Excel (XlsxPress)" in registered.out
        assert registered.err == ""
        assert "Context menu: 8 of 8 entries registered" in status.out
        assert context_menu.build_command(pythonw, "open") in status.out
        assert status.err == ""
        assert unregistered.out == "Removed 8 context menu entries.\n"
        assert "Context menu: not registered" in status_after.out

    def test_register_warns_outside_a_uv_tool_environment(
        self, test_registry_root: str, pythonw: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        exit_code = main(["context-menu", "register"])

        assert exit_code == 0
        assert capsys.readouterr().err.startswith("Warning: XlsxPress is not running from a uv tool installation")

    def test_register_fails_without_pythonw(
        self, test_registry_root: str, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
    ) -> None:
        def missing_pythonw() -> Path:
            raise FileNotFoundError("Found no 'pythonw.exe'")

        monkeypatch.setattr(context_menu, "find_pythonw", missing_pythonw)

        exit_code = main(["context-menu", "register"])

        assert exit_code == 1
        assert capsys.readouterr().err == "Error: Found no 'pythonw.exe'\n"
        assert context_menu.read_entries() == []

    def test_status_warns_about_missing_entries(
        self, test_registry_root: str, pythonw: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        context_menu.register(pythonw)
        entry = context_menu.ENTRIES[0]
        context_menu._delete_key_tree(rf"{test_registry_root}\.csv\shell\{entry.registry_key}")

        main(["context-menu", "status"])

        captured = capsys.readouterr()
        assert "Context menu: 7 of 8 entries registered" in captured.out
        assert "Warning: Some entries are missing." in captured.err

    def test_status_warns_if_the_registered_python_no_longer_exists(
        self, test_registry_root: str, pythonw: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        context_menu.register(pythonw)
        pythonw.unlink()

        main(["context-menu", "status"])

        assert f"Warning: The entries run '{pythonw}', which does not exist" in capsys.readouterr().err


class TestUnsupportedPlatform:
    @pytest.mark.parametrize("command", ["register", "unregister", "status"])
    def test_context_menu_commands_fail_outside_windows(
        self, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str], command: str
    ) -> None:
        monkeypatch.setattr(context_menu, "is_supported_platform", lambda: False)

        exit_code = main(["context-menu", command])

        assert exit_code == 1
        assert capsys.readouterr().err == "Error: The context menu is only available on Windows.\n"
