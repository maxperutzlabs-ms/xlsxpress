"""Conversion options dialog (customtkinter).

The conversion runs in a worker thread, so the window stays responsive. The thread
never touches widgets: it puts its `Outcome` into a queue, which the Tk main thread
polls with `after()`. The decisions (options, saved defaults, which outcome to show)
live in `dialog_logic`, where they are tested.

Closing the window during a conversion ends the process and the conversion with it.
No unfinished file is left under the output name, because the pipeline writes to
its work folder first.
"""

from __future__ import annotations

import logging
import queue
import threading
import tkinter as tk
from collections.abc import Callable
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

from xlsxpress import __VERSION__, config
from xlsxpress.app.gui.dialog_logic import (
    OPEN_ENTRY_LABEL,
    FormState,
    Outcome,
    about_text,
    build_options,
    defaults_from_form,
    form_from_defaults,
    outcome_of_error,
    outcome_of_result,
)
from xlsxpress.config import AppConfig, DialogDefaults, write_config
from xlsxpress.core import excel, paths
from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import run_conversion

logger = logging.getLogger(__name__)

_WIDTH = 400
_PADDING = 16
_WRAP = _WIDTH - 2 * _PADDING
_POLL_MS = 100
#: Text colour of error messages in light and dark mode.
_ERROR_COLOR = ("#b00020", "#ff6b6b")
#: Text colour of the disabled Rows field in light and dark mode.
_DISABLED_TEXT_COLOR = ("gray60", "gray45")


def run_dialog(filename: Path, app_config: AppConfig) -> None:
    """Show the conversion dialog for `filename` until the user closes it.

    Args:
        filename: The file to convert.
        app_config: The loaded config, for the dialog's defaults and the cleanup
            policy.
    """
    dialog = _OptionsDialog(filename, app_config)
    dialog.mainloop()


class _OptionsDialog(ctk.CTk):
    def __init__(self, filename: Path, app_config: AppConfig) -> None:
        super().__init__()
        self._filename = filename
        self._config = app_config
        self._outcomes: queue.Queue[Outcome] = queue.Queue()

        self.title("XlsxPress")
        self.resizable(False, False)
        self.minsize(_WIDTH, 0)

        top_frame = ctk.CTkFrame(self, fg_color="transparent")
        top_frame.pack(fill="x", padx=_PADDING, pady=(_PADDING, 8))
        ctk.CTkLabel(
            top_frame,
            text=filename.name,
            font=ctk.CTkFont(weight="bold"),
            wraplength=_WRAP - 80,
            justify="left",
        ).pack(side="left")
        ctk.CTkButton(
            top_frame, text="About", width=60, fg_color="gray", command=self._show_about
        ).pack(side="right")

        self._open_var = ctk.BooleanVar()
        self._temp_var = ctk.BooleanVar()
        self._header_var = ctk.BooleanVar()
        self._preview_var = ctk.BooleanVar()
        self._switches = [
            self._switch("Open in Excel after conversion", self._open_var),
            self._switch("Save to temp directory", self._temp_var),
            self._switch("Format header row", self._header_var),
            self._switch(
                "Preview only (limit rows)",
                self._preview_var,
                command=self._update_nrows_state,
            ),
        ]

        nrows_frame = ctk.CTkFrame(self, fg_color="transparent")
        nrows_frame.pack(fill="x", padx=_PADDING, pady=4)
        ctk.CTkLabel(nrows_frame, text="Rows:").pack(side="left")
        self._nrows_entry = ctk.CTkEntry(nrows_frame, width=80)
        self._nrows_entry.pack(side="left", padx=8)
        self._nrows_color = self._nrows_entry.cget("text_color")

        self._status_label = ctk.CTkLabel(
            self, text="", wraplength=_WRAP, justify="left", anchor="w"
        )
        self._status_label.pack(fill="x", padx=_PADDING, pady=(8, 0))
        self._status_color = self._status_label.cget("text_color")

        # Shown after a conversion that wrote a file.
        self._result_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._path_label = ctk.CTkLabel(
            self._result_frame, text="", wraplength=_WRAP, justify="left", anchor="w"
        )
        self._path_label.pack(fill="x")
        self._folder_button = ctk.CTkButton(
            self._result_frame, text="Show in folder", width=120
        )
        self._folder_button.pack(anchor="w", pady=(4, 0))
        self._warnings_box = ctk.CTkTextbox(
            self._result_frame, height=110, wrap="word", width=_WRAP
        )

        self._button_frame = ctk.CTkFrame(self, fg_color="transparent")
        self._button_frame.pack(fill="x", side="bottom", padx=_PADDING, pady=12)
        self._reset_button = ctk.CTkButton(
            self._button_frame,
            text="Reset",
            width=60,
            fg_color="gray",
            command=self._reset_options,
        )
        self._reset_button.pack(side="left")
        self._update_button = ctk.CTkButton(
            self._button_frame,
            text="Update default options",
            width=150,
            fg_color="gray",
            command=self._update_defaults,
        )
        self._update_button.pack(side="left", padx=8)
        self._bind_tooltip(
            self._update_button,
            "Saves these options as the defaults of this dialog and of\n"
            f'the context menu entry "{OPEN_ENTRY_LABEL}"',
        )
        self._submit_button = ctk.CTkButton(
            self._button_frame, text="Convert", width=100, command=self._submit
        )
        self._submit_button.pack(side="right")

        self._set_form(form_from_defaults(app_config.dialog))
        self._center()

    # Form

    def _switch(
        self,
        text: str,
        variable: ctk.BooleanVar,
        command: Callable[[], None] | None = None,
    ) -> ctk.CTkSwitch:
        switch = ctk.CTkSwitch(self, text=text, variable=variable, command=command)
        switch.pack(anchor="w", padx=_PADDING, pady=4)
        return switch

    def _read_form(self) -> FormState:
        return FormState(
            open_after=self._open_var.get(),
            to_temp=self._temp_var.get(),
            format_header=self._header_var.get(),
            preview=self._preview_var.get(),
            nrows_text=self._nrows_entry.get(),
        )

    def _set_form(self, form: FormState) -> None:
        self._open_var.set(form.open_after)
        self._temp_var.set(form.to_temp)
        self._header_var.set(form.format_header)
        self._preview_var.set(form.preview)
        self._nrows_entry.configure(state="normal")
        self._nrows_entry.delete(0, "end")
        self._nrows_entry.insert(0, form.nrows_text)
        self._update_nrows_state()

    def _update_nrows_state(self) -> None:
        self._set_nrows_enabled(self._preview_var.get())

    def _set_nrows_enabled(self, enabled: bool) -> None:
        # A disabled CTkEntry looks unchanged, so its text is greyed out by hand.
        self._nrows_entry.configure(
            state="normal" if enabled else "disabled",
            text_color=self._nrows_color if enabled else _DISABLED_TEXT_COLOR,
        )

    def _reset_options(self) -> None:
        self._set_form(form_from_defaults(DialogDefaults()))
        self._show_status(
            "Options reset to the built-in defaults. Click "
            '"Update default options" to save them.'
        )

    def _update_defaults(self) -> None:
        try:
            defaults = defaults_from_form(self._read_form(), self._config.dialog)
        except ValueError as error:
            self._show_status(str(error), error=True)
            return
        updated = AppConfig(dialog=defaults, cleanup=self._config.cleanup)
        try:
            write_config(updated)
        except OSError as error:
            logger.exception("Saving the defaults failed")
            self._show_status(
                f"Could not save the defaults to '{config.get_config_path()}': "
                f"{error.strerror or error}.",
                error=True,
            )
            return
        self._config = updated
        self._show_status("Default options saved.")

    # Conversion

    def _submit(self) -> None:
        try:
            options = build_options(
                self._filename, self._read_form(), self._config.cleanup
            )
        except ValueError as error:
            self._show_status(str(error), error=True)
            return
        self._submit_button.configure(state="disabled")
        self._show_status("Converting... this may take a moment.")
        threading.Thread(
            target=self._convert_in_background, args=(options,), daemon=True
        ).start()
        self.after(_POLL_MS, self._poll_outcome)

    def _convert_in_background(self, options: ConversionOptions) -> None:
        # Runs in the worker thread: no widget calls here.
        try:
            result = run_conversion(options)
        except Exception as error:  # noqa: BLE001 - every error is shown to the user
            logger.exception("Conversion of %s failed", options.input_path)
            self._outcomes.put(outcome_of_error(error, config.get_log_path()))
        else:
            self._outcomes.put(outcome_of_result(result, opened=options.open_after))

    def _poll_outcome(self) -> None:
        try:
            outcome = self._outcomes.get_nowait()
        except queue.Empty:
            self.after(_POLL_MS, self._poll_outcome)
            return
        self._show_outcome(outcome)

    def _show_outcome(self, outcome: Outcome) -> None:
        if outcome.close_automatically:
            self.destroy()
            return
        self._show_status(outcome.message, error=outcome.kind != "success")
        if outcome.output_path is None:
            self._submit_button.configure(state="normal")
            return

        output_path = outcome.output_path
        self._path_label.configure(text=f"Saved as: {output_path}")
        self._folder_button.configure(command=lambda: self._show_in_folder(output_path))
        if outcome.warnings:
            self._warnings_box.configure(state="normal")
            self._warnings_box.delete("1.0", "end")
            self._warnings_box.insert(
                "1.0", "Warnings:\n\n" + "\n\n".join(outcome.warnings)
            )
            self._warnings_box.configure(state="disabled")
            self._warnings_box.pack(fill="x", pady=(8, 0))
        self._result_frame.pack(fill="x", padx=_PADDING, pady=(8, 0))

        # The conversion is done: the options no longer apply, Convert becomes Close.
        for widget in (*self._switches, self._reset_button):
            widget.configure(state="disabled")
        self._set_nrows_enabled(False)
        self._submit_button.configure(
            text="Close", state="normal", command=self.destroy
        )

    def _show_in_folder(self, path: Path) -> None:
        try:
            excel.show_in_folder(path)
        except OSError as error:
            logger.exception("Showing %s in its folder failed", path)
            self._show_status(
                f"Could not open the folder of '{path}': {error.strerror or error}.",
                error=True,
            )

    # Helpers

    def _show_status(self, text: str, error: bool = False) -> None:
        color = _ERROR_COLOR if error else self._status_color
        self._status_label.configure(text=text, text_color=color)

    def _show_about(self) -> None:
        text = about_text(
            __VERSION__,
            config.get_config_path(),
            paths.get_temp_dir(),
            self._config.cleanup,
        )
        messagebox.showinfo("About XlsxPress", text, parent=self)

    def _center(self) -> None:
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) // 2
        y = (self.winfo_screenheight() - self.winfo_reqheight()) // 3
        # Set only the position, unscaled, so the window keeps its natural size and
        # grows when results are shown.
        self.wm_geometry(f"+{x}+{y}")

    def _bind_tooltip(self, widget: ctk.CTkBaseClass, text: str) -> None:
        tip: list[tk.Toplevel | None] = [None]

        def show(event: tk.Event) -> None:
            if tip[0]:
                return
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + widget.winfo_height() + 4
            win = tk.Toplevel(widget)
            win.wm_overrideredirect(True)
            win.wm_geometry(f"+{x}+{y}")
            tk.Label(
                win,
                text=text,
                background="#e0e0e0",
                relief="solid",
                borderwidth=1,
                font=("TkDefaultFont", 9),
            ).pack()
            tip[0] = win

        def hide(event: tk.Event) -> None:
            if tip[0]:
                tip[0].destroy()
                tip[0] = None

        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)
