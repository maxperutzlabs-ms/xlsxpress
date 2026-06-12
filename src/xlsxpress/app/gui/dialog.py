"""Conversion options dialog (customtkinter)."""

from __future__ import annotations

import threading
from pathlib import Path

import customtkinter as ctk

from xlsxpress.app import appdata, cleanup
from xlsxpress.app.config import AppConfig
from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import run_conversion


def run_dialog(filename: Path, app_config: AppConfig) -> None:
    dialog = _OptionsDialog(filename, app_config)
    dialog.mainloop()


class _OptionsDialog(ctk.CTk):
    def __init__(self, filename: Path, app_config: AppConfig) -> None:
        super().__init__()
        self._filename = filename
        self._config = app_config
        defaults = app_config.dialog

        self.title("XlsxPress")
        self.resizable(False, False)
        self._center(360, 360)

        ctk.CTkLabel(self, text=filename.name, font=ctk.CTkFont(weight="bold")).pack(
            padx=16, pady=(16, 8)
        )

        self._open_var = ctk.BooleanVar(value=defaults.open_after)
        self._preview_var = ctk.BooleanVar(value=defaults.preview)
        self._temp_var = ctk.BooleanVar(value=defaults.to_temp)
        self._header_var = ctk.BooleanVar(value=defaults.format_header)

        self._switch("Open in Excel after conversion", self._open_var)
        self._switch("Save to temp directory (protects original)", self._temp_var)
        self._switch("Format header row", self._header_var)
        self._switch("Preview only (limit rows)", self._preview_var)

        nrows_frame = ctk.CTkFrame(self, fg_color="transparent")
        nrows_frame.pack(fill="x", padx=16, pady=4)
        ctk.CTkLabel(nrows_frame, text="Rows:").pack(side="left")
        self._nrows_entry = ctk.CTkEntry(nrows_frame, width=80)
        self._nrows_entry.insert(0, str(defaults.nrows))
        self._nrows_entry.pack(side="left", padx=8)

        self._status_label = ctk.CTkLabel(self, text="")
        self._status_label.pack(pady=(8, 0))

        button_frame = ctk.CTkFrame(self, fg_color="transparent")
        button_frame.pack(fill="x", padx=16, pady=12)
        ctk.CTkButton(
            button_frame,
            text="Open config",
            width=100,
            fg_color="gray",
            command=self._open_config,
        ).pack(side="left")
        self._submit_button = ctk.CTkButton(
            button_frame, text="Convert", command=self._submit
        )
        self._submit_button.pack(side="right")

    def _switch(self, text: str, variable: ctk.BooleanVar) -> None:
        ctk.CTkSwitch(self, text=text, variable=variable).pack(
            anchor="w", padx=16, pady=4
        )

    def _center(self, width: int, height: int) -> None:
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")

    def _open_config(self) -> None:
        from xlsxpress.core.excel import open_in_default_app  # works for any file

        open_in_default_app(appdata.get_config_path())

    def _submit(self) -> None:
        try:
            nrows = int(self._nrows_entry.get()) if self._preview_var.get() else None
            if nrows is not None and nrows < 1:
                raise ValueError
        except ValueError:
            self._status_label.configure(text="Rows must be a positive integer.")
            return

        options = ConversionOptions(
            input_path=self._filename,
            nrows=nrows,
            to_temp=self._temp_var.get(),
            open_after=self._open_var.get(),
            format_header=self._header_var.get(),
            temp_dir=appdata.get_temp_dir(),
        )
        self._submit_button.configure(state="disabled")
        self._status_label.configure(text="Converting... this may take a moment.")
        threading.Thread(target=self._convert, args=(options,), daemon=True).start()

    def _convert(self, options: ConversionOptions) -> None:
        try:
            cleanup.clean_temp_dir(appdata.get_temp_dir(), self._config.cleanup)
            run_conversion(options)
        except Exception as error:  # noqa: BLE001
            self.after(0, self._on_error, str(error))
        else:
            self.after(0, self.destroy)

    def _on_error(self, message: str) -> None:
        self._submit_button.configure(state="normal")
        self._status_label.configure(text=f"Error: {message[:80]}")
