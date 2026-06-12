"""Windows file context menu registration (HKCU, no admin required).

On non-Windows platforms all functions are no-ops returning False.
"""

from __future__ import annotations

import sys
from pathlib import Path
from typing import NamedTuple

EXTENSIONS = (".txt", ".csv", ".tsv", ".parquet")


class ContextMenuConfig(NamedTuple):
    registry_key: str
    menu_label: str
    app_subcommand: str


_ENTRIES: tuple[ContextMenuConfig, ...] = (
    ContextMenuConfig("XlsxPress.Open", "Open in Excel (XlsxPress)", "open"),
    ContextMenuConfig("XlsxPress.Dialog", "Open in Excel with options...", "dialog"),
)
_CLASSES_ROOT = r"Software\Classes\SystemFileAssociations"


def is_supported_platform() -> bool:
    return sys.platform == "win32"


def register(exe_path: Path) -> bool:
    """Create context menu entries for all supported extensions.

    Idempotent: re-registering overwrites existing keys (also serves as the upgrade
    path if the exe location changes).
    """
    if not is_supported_platform():
        return False
    import winreg

    for extension in EXTENSIONS:
        for entry in _ENTRIES:
            shell_key = rf"{_CLASSES_ROOT}\{extension}\shell\{entry.registry_key}"
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, shell_key) as key:
                winreg.SetValueEx(key, None, 0, winreg.REG_SZ, entry.menu_label)
                winreg.SetValueEx(key, "Icon", 0, winreg.REG_SZ, str(exe_path))
            with winreg.CreateKey(
                winreg.HKEY_CURRENT_USER, rf"{shell_key}\command"
            ) as key:
                command = f'"{exe_path}" {entry.app_subcommand} "%1"'
                winreg.SetValueEx(key, None, 0, winreg.REG_SZ, command)
    _notify_shell_changed()
    return True


def unregister() -> bool:
    """Remove all context menu entries (idempotent, ignores missing keys)."""
    if not is_supported_platform():
        return False
    import winreg

    for extension in EXTENSIONS:
        for entry in _ENTRIES:
            shell_key = rf"{_CLASSES_ROOT}\{extension}\shell\{entry.registry_key}"
            _delete_key_tree(winreg.HKEY_CURRENT_USER, shell_key)
    _notify_shell_changed()
    return True


def is_registered() -> bool:
    """Return True if at least one context menu entry exists in the registry."""
    if not is_supported_platform():
        return False
    import winreg

    for extension in EXTENSIONS:
        for entry in _ENTRIES:
            try:
                shell_key = (
                    rf"{_CLASSES_ROOT}\{extension}\shell\{entry.registry_key}\command"
                )
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, shell_key):
                    return True
            except OSError:
                continue
    return False


def registered_exe_path() -> Path | None:
    """Return the exe path used by the first found registry entry, if any."""
    if not is_supported_platform():
        return None
    import winreg

    for extension in EXTENSIONS:
        for entry in _ENTRIES:
            try:
                shell_key = (
                    rf"{_CLASSES_ROOT}\{extension}\shell\{entry.registry_key}\command"
                )
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, shell_key) as key:
                    command, _ = winreg.QueryValueEx(key, None)
                return Path(command.split('"')[1])
            except OSError, IndexError:
                continue
    return None


def _delete_key_tree(root: int, key_path: str) -> None:
    """Recursively delete a registry key and its subkeys; ignore missing keys."""
    import winreg

    try:
        with winreg.OpenKey(
            root, key_path, 0, winreg.KEY_READ | winreg.KEY_WRITE
        ) as key:
            while True:
                try:
                    subkey_name = winreg.EnumKey(key, 0)
                except OSError:
                    break
                _delete_key_tree(root, rf"{key_path}\{subkey_name}")
        winreg.DeleteKey(root, key_path)
    except OSError:
        pass


def _notify_shell_changed() -> None:
    """Tell Explorer that file associations changed (refreshes the menu)."""
    import ctypes

    SHCNE_ASSOCCHANGED, SHCNF_IDLIST = 0x08000000, 0x0
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
