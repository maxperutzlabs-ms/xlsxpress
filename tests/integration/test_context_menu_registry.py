from __future__ import annotations

import os
import sys
from collections.abc import Iterator
from pathlib import Path
from types import ModuleType, TracebackType
from typing import Any

import pytest

from xlsxpress.app import context_menu
from xlsxpress.cli import main

pytestmark = pytest.mark.integration

windows_only = pytest.mark.skipif(sys.platform != "win32", reason="The context menu uses the Windows registry")
real_registry_only = pytest.mark.skipif(
    os.environ.get("XLSXPRESS_REGISTRY_TESTS") != "1",
    reason="Writes to the real registry; set XLSXPRESS_REGISTRY_TESTS=1 to run (done in CI only)",
)

_REAL_REGISTRY_TEST_ROOT = r"Software\xlsxpress-tests"


class _FakeKey:
    """A registry key of `FakeWinreg`. Names are case-insensitive, like in the registry."""

    def __init__(self, name: str) -> None:
        self.name = name
        self.values: dict[str, tuple[Any, int]] = {}
        self.subkeys: dict[str, _FakeKey] = {}

    def __enter__(self) -> _FakeKey:
        return self

    def __exit__(
        self, exc_type: type[BaseException] | None, exc: BaseException | None, traceback: TracebackType | None
    ) -> None:
        return None


class FakeWinreg(ModuleType):
    """In-memory stand-in for the parts of `winreg` that `context_menu` uses.

    It mirrors the behaviour the code relies on: `CreateKey` creates missing parent keys, `OpenKey` and `QueryValueEx`
    raise `FileNotFoundError` for missing keys and values, `EnumKey` raises `OSError` after the last subkey,
    `DeleteKey` refuses keys that still have subkeys, and a value name of `None` or `""` is the key's default value.
    """

    REG_SZ = 1
    KEY_READ = 0x20019

    def __init__(self) -> None:
        super().__init__("winreg")
        self.HKEY_CURRENT_USER = _FakeKey("HKEY_CURRENT_USER")

    def CreateKey(self, key: _FakeKey, sub_key: str) -> _FakeKey:
        for name in sub_key.split("\\"):
            key = key.subkeys.setdefault(name.lower(), _FakeKey(name))
        return key

    def OpenKey(self, key: _FakeKey, sub_key: str, reserved: int = 0, access: int = KEY_READ) -> _FakeKey:
        for name in sub_key.split("\\"):
            if name.lower() not in key.subkeys:
                raise FileNotFoundError(2, "The system cannot find the file specified")
            key = key.subkeys[name.lower()]
        return key

    def SetValueEx(self, key: _FakeKey, value_name: str | None, reserved: int, value_type: int, value: Any) -> None:
        key.values[(value_name or "").lower()] = (value, value_type)

    def QueryValueEx(self, key: _FakeKey, value_name: str | None) -> tuple[Any, int]:
        try:
            return key.values[(value_name or "").lower()]
        except KeyError:
            raise FileNotFoundError(2, "The system cannot find the file specified") from None

    def EnumKey(self, key: _FakeKey, index: int) -> str:
        subkeys = list(key.subkeys.values())
        if index >= len(subkeys):
            raise OSError(259, "No more data is available")
        return subkeys[index].name

    def DeleteKey(self, key: _FakeKey, sub_key: str) -> None:
        parent_path, _, name = sub_key.rpartition("\\")
        parent = self.OpenKey(key, parent_path) if parent_path else key
        target = self.OpenKey(parent, name)
        if target.subkeys:
            raise PermissionError(5, "Access is denied")
        del parent.subkeys[name.lower()]


@pytest.fixture
def fake_winreg(monkeypatch: pytest.MonkeyPatch) -> FakeWinreg:
    fake = FakeWinreg()
    monkeypatch.setitem(sys.modules, "winreg", fake)
    monkeypatch.setattr(context_menu, "_notify_shell_changed", lambda: None)
    return fake


@pytest.fixture
def real_registry_root(monkeypatch: pytest.MonkeyPatch) -> Iterator[str]:
    context_menu._delete_key_tree(_REAL_REGISTRY_TEST_ROOT)
    monkeypatch.setattr(context_menu, "_REGISTRY_ROOT", _REAL_REGISTRY_TEST_ROOT)
    monkeypatch.setattr(context_menu, "_notify_shell_changed", lambda: None)
    yield _REAL_REGISTRY_TEST_ROOT
    context_menu._delete_key_tree(_REAL_REGISTRY_TEST_ROOT)


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


def _shell_key_path(extension: str, entry: context_menu.MenuEntry) -> str:
    return rf"{context_menu._REGISTRY_ROOT}\{extension}\shell\{entry.registry_key}"


def _expected_entries(pythonw: Path) -> set[tuple[str, context_menu.MenuEntry, str]]:
    return {
        (extension, entry, context_menu.build_command(pythonw, entry.app_command))
        for extension in context_menu.EXTENSIONS
        for entry in context_menu.ENTRIES
    }


@windows_only
class TestRegistryRoundTrip:
    def test_register_writes_every_entry_and_unregister_removes_them(
        self, fake_winreg: FakeWinreg, pythonw: Path
    ) -> None:
        context_menu.register(pythonw)

        entries = context_menu.read_entries()
        assert {(entry.extension, entry.entry, entry.command) for entry in entries} == _expected_entries(pythonw)
        assert context_menu.unregister() == len(entries)
        assert context_menu.read_entries() == []
        assert context_menu.unregister() == 0

    def test_register_sets_the_menu_label_as_the_default_value(self, fake_winreg: FakeWinreg, pythonw: Path) -> None:
        context_menu.register(pythonw)

        entry = context_menu.ENTRIES[1]
        with fake_winreg.OpenKey(fake_winreg.HKEY_CURRENT_USER, _shell_key_path(".parquet", entry)) as key:
            assert fake_winreg.QueryValueEx(key, "") == (entry.label, fake_winreg.REG_SZ)

    def test_register_overwrites_existing_entries(self, fake_winreg: FakeWinreg, pythonw: Path, tmp_path: Path) -> None:
        context_menu.register(tmp_path / "old" / "pythonw.exe")

        context_menu.register(pythonw)

        entries = context_menu.read_entries()
        assert {(entry.extension, entry.entry, entry.command) for entry in entries} == _expected_entries(pythonw)

    def test_unregister_keeps_other_entries_of_the_extension(self, fake_winreg: FakeWinreg, pythonw: Path) -> None:
        other_key = rf"{context_menu._REGISTRY_ROOT}\.csv\shell\OtherApp"
        fake_winreg.CreateKey(fake_winreg.HKEY_CURRENT_USER, other_key)
        context_menu.register(pythonw)

        context_menu.unregister()

        assert fake_winreg.OpenKey(fake_winreg.HKEY_CURRENT_USER, other_key).name == "OtherApp"


@windows_only
@real_registry_only
class TestRealRegistryRoundTrip:
    def test_register_and_unregister_work_on_the_real_registry(self, real_registry_root: str, pythonw: Path) -> None:
        import winreg

        context_menu.register(pythonw)

        entries = context_menu.read_entries()
        entry = context_menu.ENTRIES[0]
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, _shell_key_path(".csv", entry)) as key:
            label = winreg.QueryValueEx(key, "")[0]
        assert {(entry.extension, entry.entry, entry.command) for entry in entries} == _expected_entries(pythonw)
        assert label == entry.label
        assert context_menu.unregister() == len(entries)
        assert context_menu.read_entries() == []


@windows_only
class TestContextMenuCommands:
    def test_register_status_and_unregister_round_trip(
        self, fake_winreg: FakeWinreg, pythonw: Path, tool_environment: None, capsys: pytest.CaptureFixture[str]
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
        self, fake_winreg: FakeWinreg, pythonw: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        exit_code = main(["context-menu", "register"])

        assert exit_code == 0
        assert capsys.readouterr().err.startswith("Warning: XlsxPress is not running from a uv tool installation")

    def test_register_fails_without_pythonw(
        self, fake_winreg: FakeWinreg, monkeypatch: pytest.MonkeyPatch, capsys: pytest.CaptureFixture[str]
    ) -> None:
        def missing_pythonw() -> Path:
            raise FileNotFoundError("Found no 'pythonw.exe'")

        monkeypatch.setattr(context_menu, "find_pythonw", missing_pythonw)

        exit_code = main(["context-menu", "register"])

        assert exit_code == 1
        assert capsys.readouterr().err == "Error: Found no 'pythonw.exe'\n"
        assert context_menu.read_entries() == []

    def test_status_warns_about_missing_entries(
        self, fake_winreg: FakeWinreg, pythonw: Path, capsys: pytest.CaptureFixture[str]
    ) -> None:
        context_menu.register(pythonw)
        context_menu._delete_key_tree(_shell_key_path(".csv", context_menu.ENTRIES[0]))

        main(["context-menu", "status"])

        captured = capsys.readouterr()
        assert "Context menu: 7 of 8 entries registered" in captured.out
        assert "Warning: Some entries are missing." in captured.err

    def test_status_warns_if_the_registered_python_no_longer_exists(
        self, fake_winreg: FakeWinreg, pythonw: Path, capsys: pytest.CaptureFixture[str]
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
