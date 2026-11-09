"""Windows File Explorer context menu entries for XlsxPress.

The entries are written to the current user's part of the registry (HKCU), so no
admin rights are needed. Each entry runs the app with the Python of the environment
that registered it, e.g. `"...\\Scripts\\pythonw.exe" -m xlsxpress.app open "%1"`.
`pythonw.exe` runs without a console window.

The registry functions are only available on Windows; check `is_supported_platform`
first.
"""

from __future__ import annotations

import shlex
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import NamedTuple

#: File types that get the context menu entries.
EXTENSIONS = (".csv", ".tsv", ".txt", ".parquet")


class MenuEntry(NamedTuple):
    """One context menu entry.

    Attributes:
        registry_key: Name of the entry's registry key below `shell`.
        label: Text shown in the context menu.
        app_command: Command of `python -m xlsxpress.app` that the entry runs.
    """

    registry_key: str
    label: str
    app_command: str


#: The entries registered for every extension.
ENTRIES = (
    MenuEntry("XlsxPress.Open", "Open in Excel (XlsxPress)", "open"),
    MenuEntry("XlsxPress.Dialog", "Open conversion dialog (XlsxPress)", "dialog"),
)

#: Registry key below HKCU that holds the entries per extension. Tests replace it.
_REGISTRY_ROOT = r"Software\Classes\SystemFileAssociations"


@dataclass(frozen=True, slots=True)
class RegisteredEntry:
    """A context menu entry found in the registry.

    Attributes:
        extension: The file extension, e.g. `.csv`.
        entry: The menu entry.
        command: The registered command line.
    """

    extension: str
    entry: MenuEntry
    command: str

    @property
    def program(self) -> Path | None:
        """The program the command runs, or None if the command is empty."""
        return command_program(self.command)


def is_supported_platform() -> bool:
    """Return True if context menu entries can be registered (Windows only)."""
    return sys.platform == "win32"


def find_pythonw(executable: str | Path | None = None) -> Path:
    """Return the windowless Python next to the given Python executable.

    Args:
        executable: A `python.exe`. If None (default), uses the running Python
            (`sys.executable`).

    Returns:
        The path of `pythonw.exe` in the same folder.

    Raises:
        FileNotFoundError: If there is no `pythonw.exe` next to `executable`.
    """
    executable = Path(executable or sys.executable)
    pythonw = executable.with_name("pythonw.exe")
    if not pythonw.is_file():
        raise FileNotFoundError(
            f"Found no 'pythonw.exe' next to the running Python '{executable}'. "
            "Install XlsxPress with 'uv tool install xlsxpress' and register the "
            "context menu from that installation."
        )
    return pythonw


def build_command(pythonw: Path, app_command: str) -> str:
    """Return the registry command line that runs `app_command` for the clicked file.

    Args:
        pythonw: The windowless Python of the XlsxPress environment.
        app_command: The app command, `open` or `dialog`.

    Returns:
        The command line, e.g. `"C:\\...\\pythonw.exe" -m xlsxpress.app open "%1"`.
    """
    return f'"{pythonw}" -m xlsxpress.app {app_command} "%1"'


def command_program(command: str) -> Path | None:
    """Return the program a registry command line runs.

    Args:
        command: A command line whose program is quoted or contains no spaces.

    Returns:
        The program's path, or None if `command` is empty.
    """
    try:
        parts = shlex.split(command, posix=False)
    except ValueError:
        parts = command.split()
    if not parts:
        return None
    return Path(parts[0].strip('"'))


def is_uv_tool_environment(prefix: str | Path | None = None) -> bool:
    """Return True if a Python environment was installed with `uv tool install`.

    uv writes a `uv-receipt.toml` into every tool environment. A tool environment
    stays where it is until `uv tool uninstall`, unlike a project or conda
    environment, which may be deleted or moved and leave the entries broken.

    Args:
        prefix: The environment's folder. If None (default), checks the running
            environment (`sys.prefix`).

    Returns:
        True if `prefix` is a uv tool environment.
    """
    return (Path(prefix or sys.prefix) / "uv-receipt.toml").is_file()


def register(pythonw: Path) -> None:
    """Create or overwrite the context menu entries for all extensions.

    Args:
        pythonw: The windowless Python that runs the entries, see `find_pythonw`.

    Raises:
        OSError: If a registry key cannot be written.
    """
    import winreg

    for extension in EXTENSIONS:
        for entry in ENTRIES:
            shell_key = _shell_key(extension, entry)
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER, shell_key) as key:
                winreg.SetValueEx(key, None, 0, winreg.REG_SZ, entry.label)
            with winreg.CreateKey(
                winreg.HKEY_CURRENT_USER, rf"{shell_key}\command"
            ) as key:
                command = build_command(pythonw, entry.app_command)
                winreg.SetValueEx(key, None, 0, winreg.REG_SZ, command)
    _notify_shell_changed()


def unregister() -> int:
    """Remove all context menu entries. Missing entries are ignored.

    Returns:
        The number of entries removed.
    """
    removed = 0
    for extension in EXTENSIONS:
        for entry in ENTRIES:
            removed += _delete_key_tree(_shell_key(extension, entry))
    _notify_shell_changed()
    return removed


def read_entries() -> list[RegisteredEntry]:
    """Return the context menu entries that exist in the registry.

    Returns:
        The registered entries, in the order of `EXTENSIONS` and `ENTRIES`.
    """
    import winreg

    found: list[RegisteredEntry] = []
    for extension in EXTENSIONS:
        for entry in ENTRIES:
            command_key = rf"{_shell_key(extension, entry)}\command"
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER, command_key) as key:
                    command, _ = winreg.QueryValueEx(key, "")
            except OSError:
                continue
            found.append(RegisteredEntry(extension, entry, str(command)))
    return found


def _shell_key(extension: str, entry: MenuEntry) -> str:
    return rf"{_REGISTRY_ROOT}\{extension}\shell\{entry.registry_key}"


def _delete_key_tree(key_path: str) -> bool:
    """Delete a registry key below HKCU with all subkeys. Return False if missing."""
    import winreg

    root = winreg.HKEY_CURRENT_USER
    try:
        with winreg.OpenKey(root, key_path, 0, winreg.KEY_READ) as key:
            subkeys: list[str] = []
            while True:
                try:
                    subkeys.append(winreg.EnumKey(key, len(subkeys)))
                except OSError:
                    break
    except OSError:
        return False
    for subkey in subkeys:
        _delete_key_tree(rf"{key_path}\{subkey}")
    winreg.DeleteKey(root, key_path)
    return True


def _notify_shell_changed() -> None:
    """Tell File Explorer that file associations changed, so it updates the menu."""
    import ctypes

    SHCNE_ASSOCCHANGED, SHCNF_IDLIST = 0x08000000, 0x0
    ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
