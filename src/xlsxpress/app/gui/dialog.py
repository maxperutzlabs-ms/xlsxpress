"""Conversion options dialog (customtkinter).

The conversion runs in a worker thread, so the window stays responsive. The thread
never touches widgets: it puts its `Outcome` into a queue, which the Tk main thread
polls with `after()`. The decisions (options, saved defaults, which outcome to show)
live in `dialog_logic`, where they are tested. Colours, fonts and sizes are in
`theme`, the repeated widgets in `widgets`.

A conversion that takes long shows a `ProgressPopup`. Closing the popup cancels the
conversion and returns to the dialog. Closing the dialog during a conversion
cancels it too: the dialog disappears at once, and the process ends when the
conversion has stopped and cleaned up.
"""

from __future__ import annotations

import logging
import queue
import threading
import time
from pathlib import Path
from typing import Literal

import customtkinter as ctk

from xlsxpress import __VERSION__, config
from xlsxpress.app.gui import theme
from xlsxpress.app.gui.about import AboutDialog
from xlsxpress.app.gui.dialog_logic import (
    BannerKind,
    FormState,
    Outcome,
    build_options,
    defaults_from_form,
    form_from_defaults,
    outcome_of_error,
    outcome_of_result,
    parse_nrows,
    save_location_help,
)
from xlsxpress.app.gui.progress_window import SHOW_DELAY_MS, ProgressPopup
from xlsxpress.app.gui.theme import Fonts
from xlsxpress.app.gui.widgets import (
    RoundBadge,
    SegmentedChoice,
    card,
    card_title,
    checkbox,
    fit_text,
    make_keyboard_button,
    primary_button,
    secondary_button,
    text_label,
    tooltip,
)
from xlsxpress.config import AppConfig, DialogDefaults, write_config
from xlsxpress.core import excel, paths
from xlsxpress.core.errors import ConversionCancelled
from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import run_conversion
from xlsxpress.core.progress import Progress

logger = logging.getLogger(__name__)

_POLL_MS = 100
_SAVED_FEEDBACK_MS = 1500
_TEMP_FOLDER = "Temp folder"
_NEXT_TO_ORIGINAL = "Next to original"

#: Background, border, title colour, icon colour and icon text of each banner kind.
_BANNER_STYLES: dict[BannerKind, tuple[str, str, str, str, str]] = {
    "success": (theme.OK_BG, theme.OK_BORDER, theme.TEAL_DARK, theme.TEAL, "✓"),
    "warning": (
        theme.WARN_BG,
        theme.WARN_BORDER,
        theme.WARN_TEXT,
        theme.WARN_ICON,
        "!",
    ),
    "error": (theme.ERR_BG, theme.ERR_BORDER, theme.ERR_TEXT, theme.RED, "×"),
}
_BANNER_ICON_SPACE = 14 + 18 + 10 + 14
_BANNER_BUTTON_SPACE = 110 + 10
#: Width of the banner's text column, without and with the "Show in folder" button.
_BANNER_TEXT_WIDTH = theme.WINDOW_WIDTH - 2 * theme.PAD - 2 - _BANNER_ICON_SPACE
_BANNER_TEXT_WIDTH_BUTTON = _BANNER_TEXT_WIDTH - _BANNER_BUTTON_SPACE
#: Up to this many lines, banner details are a label, beyond it a scrolling box.
_BANNER_MAX_LABEL_LINES = 3


def run_dialog(filename: Path, app_config: AppConfig) -> None:
    """Show the conversion dialog for `filename` until the user closes it.

    Args:
        filename: The file to convert.
        app_config: The loaded config, for the dialog's defaults and the cleanup
            policy.
    """
    ctk.set_appearance_mode("light")
    dialog = _OptionsDialog(filename, app_config)
    dialog.mainloop()


class _OptionsDialog(ctk.CTk):
    def __init__(self, filename: Path, app_config: AppConfig) -> None:
        super().__init__(fg_color=theme.BG)
        self._filename = filename
        self._config = app_config
        #: Outcomes of the worker thread; None for a cancelled conversion.
        self._outcomes: queue.Queue[Outcome | None] = queue.Queue()
        self._fonts = Fonts()
        self._busy = False
        self._progress = Progress()
        self._popup: ProgressPopup | None = None
        #: True once the dialog was closed during a conversion.
        self._closing = False
        self._converting_to_temp = True
        #: Name of the last workbook written next to the original.
        self._output_name: str | None = None
        #: True while the row limit field holds an invalid value that was rejected.
        self._nrows_invalid = False
        #: What the banner shows; a row limit error disappears once the value is valid.
        self._banner_content: Literal["none", "conversion", "nrows", "message"] = "none"
        self._banner_path = ""

        self.title("XlsxPress")
        self.resizable(False, False)
        self.geometry(f"{theme.WINDOW_WIDTH}x{theme.WINDOW_HEIGHT}")

        self._open_var = ctk.BooleanVar()
        self._header_var = ctk.BooleanVar()
        self._preview_var = ctk.BooleanVar()
        self._nrows_var = ctk.StringVar()
        self._nrows_var.trace_add("write", lambda *_args: self._on_nrows_changed())

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=theme.PAD, pady=theme.PAD)
        # Created in visual order, which is also the Tab order.
        self._build_header(content)
        self._build_file_card(content)
        self._build_options_card(content)
        self._save_card = self._build_save_card(content)
        self._build_banner(content)
        ctk.CTkFrame(content, fg_color="transparent", width=1, height=1).pack(
            fill="both", expand=True
        )
        self._build_button_bar(content)

        # The main button converts, or closes the dialog after a conversion.
        self.bind("<Return>", lambda _event: self._submit_button.invoke())
        self.bind("<KP_Enter>", lambda _event: self._submit_button.invoke())
        self.bind("<Escape>", lambda _event: self._close())
        self.protocol("WM_DELETE_WINDOW", self._close)

        self._set_form(form_from_defaults(app_config.dialog))
        self._center()

    # Layout

    def _build_header(self, parent: ctk.CTkFrame) -> None:
        fonts = self._fonts
        header = ctk.CTkFrame(
            parent, fg_color=theme.NAVY, corner_radius=theme.CARD_RADIUS
        )
        header.pack(fill="x")
        text_label(header, "XlsxPress", fonts.h_app, theme.WHITE).pack(
            side="left", padx=(theme.CARD_PADX, 0)
        )
        text_label(header, f"v{__VERSION__}", fonts.small, theme.ON_NAVY_MUTED).pack(
            side="left", anchor="s", padx=(8, 0), pady=(0, 16)
        )
        about = ctk.CTkButton(
            header,
            text="About",
            width=72,
            height=32,
            corner_radius=6,
            font=fonts.body,
            fg_color=theme.NAVY_BTN,
            hover_color=theme.NAVY_BTN_HOVER,
            border_width=1,
            border_color=theme.NAVY_BTN_BORDER,
            text_color=theme.WHITE,
            command=self._show_about,
        )
        make_keyboard_button(about, ring_color=theme.WHITE)
        about.pack(side="right", padx=theme.CARD_PADX, pady=theme.CARD_PADY)

    def _build_file_card(self, parent: ctk.CTkFrame) -> None:
        fonts = self._fonts
        file_card = card(parent)
        file_card.pack(fill="x", pady=(theme.GAP, 0))
        padx = theme.CARD_PADX
        text_label(file_card, "FILE", fonts.caption, theme.TEXT_MUTED).pack(
            anchor="w", padx=padx, pady=(theme.CARD_PADY, 3)
        )
        name = self._filename.name
        name_label = text_label(
            file_card,
            fit_text(name, fonts.h_file, theme.CARD_INNER_WIDTH),
            fonts.h_file,
            theme.TEXT,
        )
        name_label.pack(anchor="w", padx=padx)
        folder = str(self._filename.absolute().parent)
        folder_label = text_label(
            file_card,
            fit_text(folder, fonts.small, theme.CARD_INNER_WIDTH),
            fonts.small,
            theme.TEXT_MUTED,
        )
        folder_label.pack(anchor="w", padx=padx, pady=(3, theme.CARD_PADY))
        tooltip(name_label, lambda: name)
        tooltip(folder_label, lambda: folder)

    def _build_options_card(self, parent: ctk.CTkFrame) -> None:
        fonts = self._fonts
        options_card = card(parent)
        options_card.pack(fill="x", pady=(theme.GAP, 0))
        padx = theme.CARD_PADX
        card_title(options_card, 1, "Options", fonts).pack(
            anchor="w", padx=padx, pady=(theme.CARD_PADY, 0)
        )

        header_box = checkbox(
            options_card, "Format header row", self._header_var, fonts
        )
        header_box.pack(anchor="w", padx=padx, pady=(12, 0))
        text_label(
            options_card,
            "Bold, wrapped, frozen, with filter buttons",
            fonts.small,
            theme.TEXT_MUTED,
        ).pack(anchor="w", padx=(padx + theme.CHECKBOX_TEXT_INDENT, padx))
        open_box = checkbox(
            options_card, "Open in Excel after conversion", self._open_var, fonts
        )
        open_box.pack(anchor="w", padx=padx, pady=(12, 0))

        row = ctk.CTkFrame(options_card, fg_color="transparent")
        row.pack(anchor="w", padx=padx, pady=(10, theme.CARD_PADY))
        preview_box = checkbox(
            row,
            "Convert only the first",
            self._preview_var,
            fonts,
            command=self._on_preview_toggled,
        )
        preview_box.pack(side="left")
        self._checkboxes = (header_box, open_box, preview_box)
        self._nrows_entry = ctk.CTkEntry(
            row,
            textvariable=self._nrows_var,
            width=64,
            height=28,
            corner_radius=6,
            border_width=1,
            font=fonts.body,
        )
        self._nrows_entry.pack(side="left", padx=(10, 0))
        text_label(row, "rows", fonts.body, theme.TEXT_MUTED).pack(
            side="left", padx=(10, 0)
        )

    def _build_save_card(self, parent: ctk.CTkFrame) -> ctk.CTkFrame:
        fonts = self._fonts
        save_card = card(parent)
        save_card.pack(fill="x", pady=(theme.GAP, 0))
        padx = theme.CARD_PADX
        card_title(save_card, 2, "Save workbook", fonts).pack(
            anchor="w", padx=padx, pady=(theme.CARD_PADY, 0)
        )
        self._location = SegmentedChoice(
            save_card,
            [_TEMP_FOLDER, _NEXT_TO_ORIGINAL],
            fonts,
            command=lambda _value: self._update_location_help(),
        )
        self._location.pack(fill="x", padx=padx, pady=(10, 0))
        self._location_help = text_label(
            save_card,
            "",
            fonts.small,
            theme.TEXT_MUTED,
            wraplength=theme.CARD_INNER_WIDTH,
        )
        self._location_help.pack(fill="x", padx=padx, pady=(8, theme.CARD_PADY))
        return save_card

    def _build_banner(self, parent: ctk.CTkFrame) -> None:
        # Packed only when there is something to show.
        fonts = self._fonts
        self._banner = ctk.CTkFrame(
            parent, corner_radius=theme.CARD_RADIUS, border_width=1
        )
        self._banner.grid_columnconfigure(1, weight=1)
        self._banner_icon = RoundBadge(self._banner, 18, fonts.icon)
        text_column = ctk.CTkFrame(self._banner, fg_color="transparent")
        text_column.grid(row=0, column=1, sticky="ew", pady=theme.CARD_PADY - 2)
        self._banner_title = text_label(text_column, "", fonts.body_bold, theme.TEXT)
        self._banner_title.pack(fill="x")
        self._banner_path_label = text_label(text_column, "", fonts.small, theme.TEXT_2)
        tooltip(self._banner_path_label, lambda: self._banner_path)
        self._banner_detail_label = text_label(
            text_column, "", fonts.small, theme.TEXT_2
        )
        self._banner_detail_box = ctk.CTkTextbox(
            text_column,
            height=80,
            wrap="word",
            border_width=0,
            border_spacing=0,
            corner_radius=0,
            font=fonts.small,
            text_color=theme.TEXT_2,
            scrollbar_button_color=theme.CARD_BORDER,
            scrollbar_button_hover_color=theme.SECONDARY_BORDER,
        )
        self._folder_button = ctk.CTkButton(
            self._banner,
            text="Show in folder",
            width=110,
            height=30,
            corner_radius=6,
            font=fonts.small_bold,
            fg_color=theme.CARD,
            hover_color=theme.OK_BG,
            border_width=1,
            border_color=theme.TEAL,
            text_color=theme.TEAL,
        )
        make_keyboard_button(self._folder_button, ring_color=theme.TEAL_DARK)

    def _build_button_bar(self, parent: ctk.CTkFrame) -> None:
        fonts = self._fonts
        bar = ctk.CTkFrame(parent, fg_color="transparent")
        bar.pack(fill="x")
        self._reset_button = secondary_button(
            bar, "Reset", self._reset_options, fonts, width=60
        )
        self._reset_button.pack(side="left")
        self._save_button = secondary_button(
            bar, "Save as default", self._save_defaults, fonts, width=120
        )
        self._save_button.pack(side="left", padx=(8, 0))
        self._submit_button = primary_button(
            bar, "Convert", self._submit, fonts, width=107
        )
        self._submit_button.pack(side="right")

    # Form

    def _read_form(self) -> FormState:
        return FormState(
            open_after=self._open_var.get(),
            to_temp=self._location.get() == _TEMP_FOLDER,
            format_header=self._header_var.get(),
            preview=self._preview_var.get(),
            nrows_text=self._nrows_var.get(),
        )

    def _set_form(self, form: FormState) -> None:
        self._open_var.set(form.open_after)
        self._header_var.set(form.format_header)
        self._preview_var.set(form.preview)
        self._nrows_var.set(form.nrows_text)
        self._location.set(_TEMP_FOLDER if form.to_temp else _NEXT_TO_ORIGINAL)
        self._update_nrows_state()
        self._update_location_help()

    def _on_preview_toggled(self) -> None:
        self._update_nrows_state()
        if self._preview_var.get():
            self._nrows_entry.focus_set()
            self._nrows_entry.select_range(0, "end")

    def _on_nrows_changed(self) -> None:
        if not self._nrows_invalid:
            return
        try:
            parse_nrows(self._nrows_var.get())
        except ValueError:
            return
        self._nrows_invalid = False
        self._update_nrows_state()
        if self._banner_content == "nrows":
            self._hide_banner()

    def _mark_nrows_invalid(self) -> None:
        self._nrows_invalid = True
        self._update_nrows_state()
        self._nrows_entry.focus_set()

    def _update_nrows_state(self) -> None:
        if self._preview_var.get():
            border = theme.ERR_BORDER if self._nrows_invalid else theme.TEAL
            self._nrows_entry.configure(
                state="normal",
                fg_color=theme.CARD,
                border_color=border,
                text_color=theme.TEXT,
            )
        else:
            # A disabled CTkEntry looks unchanged, so it is greyed out by hand.
            self._nrows_entry.configure(
                state="disabled",
                fg_color=theme.BG,
                border_color=theme.CARD_BORDER,
                text_color=theme.DISABLED_TEXT,
            )

    def _update_location_help(self) -> None:
        to_temp = self._location.get() == _TEMP_FOLDER
        text = save_location_help(
            to_temp, self._filename.name, self._config.cleanup, self._output_name
        )
        self._location_help.configure(text=text)
        self._fit_height()

    def _reset_options(self) -> None:
        self._nrows_invalid = False
        self._set_form(form_from_defaults(DialogDefaults()))
        if self._banner_content == "nrows":
            self._hide_banner()

    def _save_defaults(self) -> None:
        try:
            defaults = defaults_from_form(self._read_form(), self._config.dialog)
        except ValueError as error:
            self._mark_nrows_invalid()
            self._show_banner("error", "Defaults not saved", detail=str(error))
            self._banner_content = "nrows"
            return
        updated = AppConfig(dialog=defaults, cleanup=self._config.cleanup)
        try:
            write_config(updated)
        except OSError as error:
            logger.exception("Saving the defaults failed")
            self._show_banner(
                "error",
                "Defaults not saved",
                detail=(
                    f"Could not save the defaults to '{config.get_config_path()}': "
                    f"{error.strerror or error}."
                ),
            )
            return
        self._config = updated
        self._save_button.configure(text="Saved ✓")
        self.after(
            _SAVED_FEEDBACK_MS,
            lambda: self._save_button.configure(text="Save as default"),
        )

    # Conversion

    def _submit(self) -> None:
        if self._busy:
            return
        try:
            options = build_options(
                self._filename, self._read_form(), self._config.cleanup
            )
        except ValueError as error:
            self._mark_nrows_invalid()
            self._show_banner("error", "Conversion not started", detail=str(error))
            self._banner_content = "nrows"
            return
        self._busy = True
        self._converting_to_temp = options.to_temp
        self._submit_button.configure(
            state="disabled", text="Converting…", fg_color=theme.RED_HOVER
        )
        progress = Progress()
        started = time.monotonic()
        self._progress = progress
        threading.Thread(
            target=self._convert_in_background, args=(options, progress), daemon=True
        ).start()
        self.after(SHOW_DELAY_MS, lambda: self._show_progress(progress, started))
        self.after(_POLL_MS, self._poll_outcome)

    def _convert_in_background(
        self, options: ConversionOptions, progress: Progress
    ) -> None:
        # Runs in the worker thread: no widget calls here.
        try:
            result = run_conversion(options, progress)
        except ConversionCancelled:
            self._outcomes.put(None)
        except Exception as error:  # noqa: BLE001 - every error is shown to the user
            logger.exception("Conversion of %s failed", options.input_path)
            self._outcomes.put(outcome_of_error(error, config.get_log_path()))
        else:
            self._outcomes.put(outcome_of_result(result, opened=options.open_after))

    def _poll_outcome(self) -> None:
        try:
            outcome = self._outcomes.get_nowait()
        except queue.Empty:
            if self._popup is not None:
                self._popup.refresh()
            self.after(_POLL_MS, self._poll_outcome)
            return
        self._show_outcome(outcome)

    def _show_progress(self, progress: Progress, started: float) -> None:
        # Only if this conversion is still running and the dialog is still shown.
        if not self._busy or progress is not self._progress or progress.cancelled:
            return
        self._popup = ProgressPopup(
            self,
            self._fonts,
            self._filename.name,
            progress,
            started,
            on_cancel=self._cancel_conversion,
        )

    def _cancel_conversion(self) -> None:
        # The conversion stops at its next check and then reports its outcome.
        self._progress.cancel()
        self._close_popup()
        self._submit_button.configure(text="Cancelling…")

    def _close_popup(self) -> None:
        if self._popup is not None:
            self._popup.destroy()
            self._popup = None

    def _show_outcome(self, outcome: Outcome | None) -> None:
        self._close_popup()
        if self._closing or (outcome is not None and outcome.close_automatically):
            self.destroy()
            return
        self._busy = False
        self._submit_button.configure(
            state="normal", text="Convert", fg_color=theme.RED
        )
        if outcome is None:
            return
        if outcome.output_path is not None:
            if not self._converting_to_temp:
                self._output_name = outcome.output_path.name
                self._update_location_help()
            self._lock_options()
        detail = "\n\n".join(filter(None, (outcome.message, *outcome.warnings)))
        self._show_banner(
            outcome.banner, outcome.title, path=outcome.output_path, detail=detail
        )
        self._banner_content = "conversion"

    def _lock_options(self) -> None:
        # The file is written: the options no longer apply, Convert becomes Close.
        # After an error nothing was written, and Convert stays available.
        for widget in (*self._checkboxes, self._reset_button):
            widget.configure(state="disabled")
        self._location.set_enabled(False)
        self._nrows_entry.configure(
            state="disabled",
            fg_color=theme.BG,
            border_color=theme.CARD_BORDER,
            text_color=theme.DISABLED_TEXT,
        )
        self._submit_button.configure(text="Close", command=self.destroy)

    def _show_in_folder(self, path: Path) -> None:
        try:
            excel.show_in_folder(path)
        except OSError as error:
            logger.exception("Showing %s in its folder failed", path)
            self._show_banner(
                "error",
                "Could not open the folder",
                detail=(
                    f"Could not open the folder of '{path}': {error.strerror or error}."
                ),
            )

    # Banner

    def _show_banner(
        self, kind: BannerKind, title: str, path: Path | None = None, detail: str = ""
    ) -> None:
        background, border, title_color, icon_color, icon = _BANNER_STYLES[kind]
        self._banner.configure(fg_color=background, border_color=border)
        self._banner_icon.set(icon, icon_color)
        # Centred next to a title and a path; at the title next to longer details.
        sticky = "n" if detail else ""
        top = theme.CARD_PADY - 1 if detail else 0
        self._banner_icon.grid(
            row=0, column=0, sticky=sticky, padx=(theme.CARD_PADX, 10), pady=(top, 0)
        )
        self._banner_title.configure(text=title, text_color=title_color)
        self._banner_content = "message"

        with_button = path is not None and kind != "error"
        width = _BANNER_TEXT_WIDTH_BUTTON if with_button else _BANNER_TEXT_WIDTH
        if with_button and path is not None:
            self._folder_button.configure(command=lambda: self._show_in_folder(path))
            self._folder_button.grid(
                row=0,
                column=2,
                sticky=sticky,
                padx=(10, theme.CARD_PADX),
                pady=(top - 4 if detail else 0, 0),
            )
        else:
            self._folder_button.grid_remove()

        self._banner_path = "" if path is None else str(path)
        self._banner_path_label.pack_forget()
        if path is not None:
            self._banner_path_label.configure(
                text=fit_text(str(path), self._fonts.small, width)
            )
            self._banner_path_label.pack(fill="x", pady=(2, 0))

        self._banner_detail_label.pack_forget()
        self._banner_detail_box.pack_forget()
        if detail and self._fits_in_label(detail, width):
            self._banner_detail_label.configure(text=detail, wraplength=width)
            self._banner_detail_label.pack(fill="x", pady=(2, 0))
        elif detail:
            box = self._banner_detail_box
            box.configure(state="normal", fg_color=background, width=width)
            box.delete("1.0", "end")
            box.insert("1.0", detail)
            box.configure(state="disabled")
            box.pack(fill="x", pady=(4, 0))

        self._banner.pack(fill="x", pady=(theme.GAP, 0), after=self._save_card)
        self._fit_height()

    def _hide_banner(self) -> None:
        self._banner.pack_forget()
        self._banner_content = "none"

    def _fits_in_label(self, text: str, width: int) -> bool:
        if "\n" in text:
            return False
        # Approximate: wrapping at word boundaries needs a little more room.
        lines = self._fonts.small.measure(text) / (width * 0.9)
        return lines <= _BANNER_MAX_LABEL_LINES

    # Window

    def _show_about(self) -> None:
        AboutDialog(
            self,
            self._fonts,
            __VERSION__,
            self._config.cleanup,
            paths.get_temp_dir(),
            config.get_config_path(),
        )

    def _close(self) -> None:
        if not self._busy:
            self.destroy()
            return
        # The dialog disappears at once; it is destroyed when the conversion has
        # stopped and cleaned up.
        self._closing = True
        self._cancel_conversion()
        self.withdraw()

    def _center(self) -> None:
        self.update_idletasks()
        scaling = ctk.ScalingTracker.get_window_scaling(self)
        width = round(theme.WINDOW_WIDTH * scaling)
        height = round(theme.WINDOW_HEIGHT * scaling)
        x = (self.winfo_screenwidth() - width) // 2
        y = (self.winfo_screenheight() - height) // 2
        self.wm_geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _fit_height(self) -> None:
        # The window has a fixed height, and the spacer above the buttons absorbs
        # the banner. Long texts or a large display scaling can need more room:
        # then the window grows instead of cutting off the buttons.
        self.update_idletasks()
        scaling = ctk.ScalingTracker.get_window_scaling(self)
        width = round(theme.WINDOW_WIDTH * scaling)
        height = round(theme.WINDOW_HEIGHT * scaling)
        needed = self.winfo_reqheight()
        if needed > max(height, self.winfo_height()):
            self.wm_geometry(f"{width}x{needed}")
