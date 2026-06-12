"""Launcher / maintenance window shown when the exe runs without arguments."""

from __future__ import annotations

import threading
from tkinter import messagebox

import customtkinter as ctk

from xlsxpress.app import appdata, cleanup, config, context_menu, installation
from xlsxpress.app.installation import InstallState


def run_launcher(state: InstallState) -> None:
    window = _LauncherWindow(state)
    window.mainloop()


class _LauncherWindow(ctk.CTk):
    def __init__(self, state: InstallState) -> None:
        super().__init__()
        self.title("XlsxPress Launcher")
        self.resizable(False, False)
        self._center(380, 420)

        ctk.CTkLabel(
            self, text="XlsxPress", font=ctk.CTkFont(size=18, weight="bold")
        ).pack(pady=(16, 4))
        self._status_label = ctk.CTkLabel(self, text="", justify="left")
        self._status_label.pack(pady=(0, 12))

        self._install_button = self._button("Install", self._on_install)
        self._context_button = self._button("", self._on_toggle_context)
        self._update_button = self._button("Check for updates", self._on_update)
        self._button("Clean temp files", self._on_clean_temp)
        self._button("Open config file", self._on_open_config)
        self._button("Uninstall", self._on_uninstall, fg_color="#8b3a3a")

        self._busy_label = ctk.CTkLabel(self, text="")
        self._busy_label.pack(pady=8)

        self._refresh(state)

    # -- UI helpers -------------------------------------------------------
    def _button(self, text: str, command, **kwargs) -> ctk.CTkButton:
        button = ctk.CTkButton(self, text=text, command=command, width=220, **kwargs)
        button.pack(pady=4)
        return button

    def _center(self, width: int, height: int) -> None:
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _refresh(self, state: InstallState) -> None:
        registered = "yes" if state.context_registered else "no"
        installed = "yes" if state.installed else "no"
        version = state.version or installation.package_version()
        self._status_label.configure(
            text=f"Version: {version}\nInstalled: {installed}\nContext menu: {registered}"
        )
        self._install_button.configure(
            state="disabled" if state.installed else "normal",
            text="Installed" if state.installed else "Install",
        )
        self._context_button.configure(
            text="Remove context menu"
            if state.context_registered
            else "Register context menu",
            state="normal"
            if state.installed and context_menu.is_supported_platform()
            else "disabled",
        )
        self._update_button.configure(state="normal" if state.installed else "disabled")

    def _run_busy(self, label: str, work, on_done) -> None:
        """Run `work` in a thread with all buttons disabled and a busy label."""
        self._set_buttons_enabled(False)
        self._busy_label.configure(text=label)

        def task() -> None:
            result = work()
            self.after(
                0, lambda: (self._busy_label.configure(text=""), on_done(result))
            )
            self.after(0, lambda: self._set_buttons_enabled(True))

        threading.Thread(target=task, daemon=True).start()

    def _set_buttons_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for widget in self.winfo_children():
            if isinstance(widget, ctk.CTkButton):
                widget.configure(state=state)
        if enabled:
            self._refresh(installation.read_state())

    # -- actions ----------------------------------------------------------
    def _on_install(self) -> None:
        register = messagebox.askyesno(
            "XlsxPress", "Add XlsxPress to the Windows file context menu?"
        )
        self._refresh(installation.install(register_context=register))

    def _on_toggle_context(self) -> None:
        state = installation.read_state()
        if state.context_registered:
            context_menu.unregister()
        else:
            context_menu.register(appdata.get_installed_exe_path())
        self._refresh(installation.check_and_repair())

    def _on_update(self) -> None:
        self._run_busy(
            "Updating... this may take a few minutes.",
            installation.self_update,
            lambda result: messagebox.showinfo("XlsxPress update", result[1]),
        )

    def _on_clean_temp(self) -> None:
        deleted = cleanup.clean_temp_dir(
            appdata.get_temp_dir(), config.load_config().cleanup
        )
        messagebox.showinfo("XlsxPress", f"Deleted {deleted} temp file(s).")

    def _on_open_config(self) -> None:
        from xlsxpress.core.excel import open_in_default_app

        config.ensure_config_exists()
        open_in_default_app(appdata.get_config_path())

    def _on_uninstall(self) -> None:
        if not messagebox.askyesno(
            "XlsxPress", "Remove context menu entries and installation data?"
        ):
            return
        installation.uninstall()
        messagebox.showinfo(
            "XlsxPress",
            "Uninstalled. You may now delete the XlsxPress folder in your "
            f"local app data directory:\n{appdata.get_appdata_dir()}",
        )
        self.destroy()
