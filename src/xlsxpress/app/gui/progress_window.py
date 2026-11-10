"""Progress display of a running conversion (customtkinter).

`convert_with_progress` runs a conversion for the context menu entry that converts
without the dialog, and shows a small window while it runs. The conversion dialog
shows the same content in a `ProgressPopup` over itself.

The conversion runs in a worker thread, which never touches widgets. The window
polls the conversion's `Progress` with `after()`. Closing the window cancels the
conversion: the window disappears at once, and the conversion stops and cleans up
at its next check.
"""

from __future__ import annotations

import threading
import time
import tkinter as tk
from collections.abc import Callable
from typing import Any

import customtkinter as ctk

from xlsxpress.app.gui import theme
from xlsxpress.app.gui.dialog_logic import format_elapsed, progress_status
from xlsxpress.app.gui.theme import Fonts
from xlsxpress.app.gui.widgets import card, fit_text, text_label
from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import ConversionResult, run_conversion
from xlsxpress.core.progress import Progress

#: A conversion that is faster than this shows no progress window.
SHOW_DELAY_MS = 400
_POLL_MS = 100
_WIDTH = 420
_INNER_WIDTH = _WIDTH - 2 * theme.PAD - 2 * theme.CARD_PADX - 2


def convert_with_progress(options: ConversionOptions) -> ConversionResult:
    """Run a conversion and show a progress window while it takes long.

    The window appears only if the conversion takes longer than `SHOW_DELAY_MS`.
    Closing it cancels the conversion.

    Args:
        options: The conversion options.

    Returns:
        The output path and the collected warnings.

    Raises:
        errors.ConversionCancelled: If the user closed the window.
        errors.ConversionError: See `run_conversion`.
        errors.OpenOutputError: See `run_conversion`.
    """
    ctk.set_appearance_mode("light")
    window = _ConversionWindow(options)
    window.mainloop()
    if isinstance(window.outcome, ConversionResult):
        return window.outcome
    if window.outcome is None:
        raise RuntimeError("The conversion ended without a result.")
    raise window.outcome


class ProgressPanel(ctk.CTkFrame):
    """Shows the file name, the elapsed time, a progress bar and the stage."""

    def __init__(
        self,
        parent: Any,
        fonts: Fonts,
        filename: str,
        progress: Progress,
        started: float,
    ) -> None:
        """Create the panel.

        Args:
            parent: The window that shows the panel.
            fonts: The dialog fonts.
            filename: The name of the file that is converted.
            progress: The progress of the running conversion.
            started: The `time.monotonic()` value of the conversion's start.
        """
        super().__init__(parent, fg_color="transparent")
        self._progress = progress
        self._started = started
        self._determinate = False

        box = card(self)
        box.pack(fill="x")
        padx = theme.CARD_PADX
        top = ctk.CTkFrame(box, fg_color="transparent")
        top.pack(fill="x", padx=padx, pady=(theme.CARD_PADY, 3))
        text_label(top, "CONVERTING", fonts.caption, theme.TEXT_MUTED).pack(side="left")
        self._elapsed = text_label(top, "", fonts.small, theme.TEXT_MUTED)
        self._elapsed.pack(side="right")
        text_label(
            box,
            fit_text(filename, fonts.h_file, _INNER_WIDTH),
            fonts.h_file,
            theme.TEXT,
        ).pack(anchor="w", padx=padx)
        self._bar = ctk.CTkProgressBar(
            box,
            width=_INNER_WIDTH,
            height=6,
            corner_radius=3,
            fg_color=theme.SECONDARY,
            progress_color=theme.TEAL,
            mode="indeterminate",
        )
        self._bar.pack(padx=padx, pady=(10, 8))
        self._bar.start()
        self._status = text_label(box, "", fonts.small, theme.TEXT_MUTED)
        self._status.pack(anchor="w", padx=padx, pady=(0, theme.CARD_PADY))
        text_label(
            self,
            "Close this window to cancel. Large files can take a few minutes.",
            fonts.small,
            theme.TEXT_MUTED,
        ).pack(anchor="w", pady=(theme.GAP, 0))
        self.refresh()

    def refresh(self) -> None:
        """Show the current stage, progress and elapsed time."""
        stage, fraction = self._progress.stage, self._progress.fraction
        self._elapsed.configure(text=format_elapsed(time.monotonic() - self._started))
        self._status.configure(text=progress_status(stage, fraction))
        if fraction is None:
            if self._determinate:
                self._determinate = False
                self._bar.configure(mode="indeterminate")
                self._bar.start()
        else:
            if not self._determinate:
                self._determinate = True
                self._bar.stop()
                self._bar.configure(mode="determinate")
            self._bar.set(fraction)


class ProgressPopup(ctk.CTkToplevel):
    """Modal progress window over the conversion dialog."""

    def __init__(
        self,
        parent: ctk.CTk,
        fonts: Fonts,
        filename: str,
        progress: Progress,
        started: float,
        on_cancel: Callable[[], None],
    ) -> None:
        """Create the window over `parent`, hidden until `show` is called.

        Create it before the conversion starts: while the conversion runs, building
        the widgets takes more than a second, because the main thread has to wait
        for the worker thread at every widget call.

        Args:
            parent: The conversion dialog.
            fonts: The dialog fonts.
            filename: The name of the file that is converted.
            progress: The progress of the running conversion.
            started: The `time.monotonic()` value of the conversion's start.
            on_cancel: Called when the user closes the window.
        """
        super().__init__(parent, fg_color=theme.BG)
        self.withdraw()
        self.title("XlsxPress")
        self.resizable(False, False)
        self.transient(parent)
        self._panel = ProgressPanel(self, fonts, filename, progress, started)
        self._panel.pack(padx=theme.PAD, pady=theme.PAD)
        self.protocol("WM_DELETE_WINDOW", on_cancel)
        self.bind("<Escape>", lambda _event: on_cancel())

        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_reqwidth()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_reqheight()) // 2
        self.wm_geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def show(self) -> None:
        """Show the window and make it modal."""
        self.deiconify()
        # customtkinter may show a new toplevel behind its parent; raise it later.
        self.after(10, self._raise)

    def refresh(self) -> None:
        """Show the current stage, progress and elapsed time."""
        self._panel.refresh()

    def _raise(self) -> None:
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except tk.TclError:  # Not viewable yet; the window still works.
            pass


class _ConversionWindow(ctk.CTk):
    """Runs one conversion and shows its progress; closes itself when it ends."""

    def __init__(self, options: ConversionOptions) -> None:
        super().__init__(fg_color=theme.BG)
        self.withdraw()
        self.title("XlsxPress")
        self.resizable(False, False)
        # Explorer keeps the focus, so the window would open behind it.
        self.attributes("-topmost", True)
        #: The result or the error of the conversion, once it has ended.
        self.outcome: ConversionResult | BaseException | None = None
        self._progress = Progress()
        self._panel = ProgressPanel(
            self, Fonts(), options.input_path.name, self._progress, time.monotonic()
        )
        self._panel.pack(padx=theme.PAD, pady=theme.PAD)
        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.bind("<Escape>", lambda _event: self._cancel())

        self._worker = threading.Thread(
            target=self._convert, args=(options,), daemon=True
        )
        self._worker.start()
        self.after(SHOW_DELAY_MS, self._show)
        self.after(_POLL_MS, self._poll)

    def _convert(self, options: ConversionOptions) -> None:
        # Runs in the worker thread: no widget calls here.
        try:
            self.outcome = run_conversion(options, self._progress)
        except BaseException as error:  # noqa: BLE001 - raised again in the main thread
            self.outcome = error

    def _show(self) -> None:
        if self._progress.cancelled or not self._worker.is_alive():
            return
        self.update_idletasks()
        x = (self.winfo_screenwidth() - self.winfo_reqwidth()) // 2
        y = (self.winfo_screenheight() - self.winfo_reqheight()) // 2
        self.wm_geometry(f"+{max(x, 0)}+{max(y, 0)}")
        self.deiconify()

    def _poll(self) -> None:
        if self._worker.is_alive():
            self._panel.refresh()
            self.after(_POLL_MS, self._poll)
        else:
            self.destroy()

    def _cancel(self) -> None:
        # The window disappears at once; the process ends when the conversion has
        # stopped and cleaned up.
        self._progress.cancel()
        self.withdraw()
