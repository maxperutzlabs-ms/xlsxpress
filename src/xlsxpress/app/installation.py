"""Installation state management: install.toml, install/repair/uninstall/update."""

from __future__ import annotations

import os
import shutil
import subprocess
import tomllib
from dataclasses import dataclass
from importlib import metadata
from pathlib import Path

from xlsxpress.app import appdata, config, context_menu


def package_version() -> str:
    try:
        return metadata.version("xlsxpress")
    except metadata.PackageNotFoundError:
        return "0.0.0-dev"


def running_exe_path() -> Path | None:
    """Path of the pyapp executable currently running us, if any.

    Requires the binary to be built with PYAPP_PASS_LOCATION=1, which make PyApp set the
    PYAPP env var to the executable path instead of "1".
    """
    value = os.environ.get("PYAPP", "")
    if value and value != "1":
        path = Path(value)
        if path.is_file():
            return path
    return None


@dataclass(slots=True)
class InstallState:
    """Contents of install.toml."""

    version: str = ""
    installed: bool = False
    context_registered: bool = False

    def to_toml(self) -> str:
        return (
            f'version = "{self.version}"\n'
            f"installed = {str(self.installed).lower()}\n"
            f"context_registered = {str(self.context_registered).lower()}\n"
        )


def read_state() -> InstallState:
    try:
        with appdata.get_install_manifest_path().open("rb") as file:
            data = tomllib.load(file)
        return InstallState(
            version=str(data.get("version", "")),
            installed=bool(data.get("installed", False)),
            context_registered=bool(data.get("context_registered", False)),
        )
    except OSError, tomllib.TOMLDecodeError:
        return InstallState()


def write_state(state: InstallState) -> None:
    appdata.ensure_directory_structure()
    appdata.get_install_manifest_path().write_text(state.to_toml(), encoding="utf-8")


def check_and_repair() -> InstallState:
    """Reconcile install.toml with reality (filesystem + registry) and return the
    corrected state. This is the single source of truth the launcher relies on.

    Rules:
      - installed = exe actually present in appdata
      - context_registered = entries actually in registry
      - if context entries exist but the exe is gone -> remove the entries
        (they would be dead links)
    """
    state = read_state()
    exe_exists = appdata.get_installed_exe_path().is_file()
    registered = context_menu.is_registered()

    if registered and not exe_exists:
        context_menu.unregister()
        registered = False

    state.installed = exe_exists
    state.context_registered = registered
    if exe_exists and not state.version:
        state.version = package_version()
    write_state(state)
    return state


def install(register_context: bool) -> InstallState:
    """Perform the installation: directories, exe copy, config, manifest."""
    appdata.ensure_directory_structure()
    config.ensure_config_exists()

    target_exe = appdata.get_installed_exe_path()
    source_exe = running_exe_path()
    if source_exe and source_exe.resolve() != target_exe.resolve():
        shutil.copy2(source_exe, target_exe)

    if register_context and target_exe.is_file():
        context_menu.register(target_exe)

    return check_and_repair()


def uninstall() -> None:
    """Remove context menu entries and the install manifest.

    Note: does not delete the appdata directory itself while running, since the running
    exe may live there; the launcher offers this and explains the leftover exe must be
    deleted manually.
    """
    context_menu.unregister()
    manifest = appdata.get_install_manifest_path()
    manifest.unlink(missing_ok=True)
    shutil.rmtree(appdata.get_temp_dir(), ignore_errors=True)


def self_update() -> tuple[bool, str]:
    """Run `pyapp.exe self update`, blocking. Returns (success, message).

    Must run the installed exe; afterwards queries the new version by invoking the exe
    with --version (which runs the updated code).
    """
    exe = appdata.get_installed_exe_path()
    if not exe.is_file():
        return False, "No installed executable found."
    try:
        result = subprocess.run(
            [str(exe), "self", "update"], capture_output=True, text=True, timeout=600
        )
    except (OSError, subprocess.TimeoutExpired) as error:
        return False, f"Update failed: {error}"
    if result.returncode != 0:
        return False, f"Update failed:\n{result.stderr.strip()[-500:]}"

    version_result = subprocess.run(
        [str(exe), "--version"], capture_output=True, text=True, timeout=120
    )
    new_version = version_result.stdout.strip() or package_version()
    state = read_state()
    state.version = new_version
    write_state(state)
    return True, f"Updated to version {new_version}."
