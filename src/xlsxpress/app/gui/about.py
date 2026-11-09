"""The About dialog of the conversion dialog."""

from __future__ import annotations

import logging
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import messagebox

import customtkinter as ctk

from xlsxpress.app.gui import theme
from xlsxpress.app.gui.dialog_logic import about_sections, truncate_start
from xlsxpress.app.gui.theme import Fonts
from xlsxpress.app.gui.widgets import (
    card,
    fit_text,
    link_label,
    secondary_button,
    text_label,
    tooltip,
)
from xlsxpress.core import excel
from xlsxpress.core.cleanup import CleanupPolicy

logger = logging.getLogger(__name__)

REPOSITORY_URL = "https://github.com/maxperutzlabs-ms/xlsxpress"

_PATH_LABEL_WIDTH = 76
_PATH_LINK_WIDTH = theme.ABOUT_INNER_WIDTH - _PATH_LABEL_WIDTH - 8


class AboutDialog(ctk.CTkToplevel):
    """Modal window with the version, what the buttons do, and where files are."""

    def __init__(
        self,
        parent: ctk.CTk,
        fonts: Fonts,
        version: str,
        cleanup_policy: CleanupPolicy,
        temp_dir: Path,
        config_path: Path,
    ) -> None:
        """Create and show the dialog over `parent`.

        Args:
            parent: The conversion dialog.
            fonts: The dialog fonts.
            version: The installed XlsxPress version.
            cleanup_policy: When files in the temp folder are deleted.
            temp_dir: The temp folder.
            config_path: The file with the saved defaults.
        """
        super().__init__(parent, fg_color=theme.BG)
        self.title("About XlsxPress")
        self.resizable(False, False)
        self.geometry(f"{theme.ABOUT_WIDTH}x{theme.ABOUT_HEIGHT}")
        self.transient(parent)

        content = ctk.CTkFrame(self, fg_color="transparent")
        content.pack(fill="both", expand=True, padx=theme.PAD, pady=theme.PAD)

        header = ctk.CTkFrame(
            content, fg_color=theme.NAVY, corner_radius=theme.CARD_RADIUS
        )
        header.pack(fill="x")
        text_label(header, f"XlsxPress {version}", fonts.h_about, theme.WHITE).pack(
            anchor="w", padx=theme.CARD_PADX, pady=(theme.CARD_PADY, 0)
        )
        text_label(
            header,
            "Mass Spectrometry Facility · Max Perutz Labs",
            fonts.small,
            theme.ON_NAVY_MUTED,
        ).pack(anchor="w", padx=theme.CARD_PADX, pady=(2, theme.CARD_PADY))

        explanations = card(content)
        explanations.pack(fill="x", pady=(theme.GAP, 0))
        sections = about_sections(cleanup_policy)
        for index, (title, text) in enumerate(sections):
            top = theme.CARD_PADY if index == 0 else theme.GAP
            bottom = theme.CARD_PADY if index == len(sections) - 1 else 0
            text_label(explanations, title, fonts.body_bold, theme.TEXT).pack(
                anchor="w", padx=theme.CARD_PADX, pady=(top, 2)
            )
            text_label(
                explanations,
                text,
                fonts.body,
                theme.TEXT_2,
                wraplength=theme.ABOUT_INNER_WIDTH,
            ).pack(anchor="w", padx=theme.CARD_PADX, pady=(0, bottom))

        paths = card(content)
        paths.pack(fill="x", pady=(theme.GAP, 0))
        paths.grid_columnconfigure(1, weight=1)
        for row, (name, path) in enumerate(
            (("Temp folder", temp_dir), ("Defaults", config_path))
        ):
            pady = (
                theme.CARD_PADY if row == 0 else 4,
                theme.CARD_PADY if row == 1 else 0,
            )
            text_label(
                paths, name, fonts.small, theme.TEXT_MUTED, width=_PATH_LABEL_WIDTH
            ).grid(row=row, column=0, sticky="w", padx=(theme.CARD_PADX, 8), pady=pady)
            link = link_label(
                paths,
                fit_text(str(path), fonts.link, _PATH_LINK_WIDTH, truncate_start),
                lambda path=path: self._open_path(path),
                fonts.link,
            )
            link.grid(
                row=row, column=1, sticky="w", padx=(0, theme.CARD_PADX), pady=pady
            )
            tooltip(link, lambda path=path: str(path))

        ctk.CTkFrame(content, fg_color="transparent", width=1, height=1).pack(
            fill="both", expand=True
        )

        bottom = ctk.CTkFrame(content, fg_color="transparent")
        bottom.pack(fill="x")
        link_label(
            bottom,
            REPOSITORY_URL.removeprefix("https://"),
            lambda: webbrowser.open(REPOSITORY_URL),
            fonts.link_bold,
        ).pack(side="left")
        close_button = secondary_button(bottom, "Close", self.destroy, fonts, width=100)
        close_button.pack(side="right")

        self.bind("<Escape>", lambda _event: self.destroy())
        self.bind("<Return>", lambda _event: self.destroy())
        self._center_over(parent)
        # customtkinter may show a new toplevel behind its parent; raise it later.
        self.after(10, self._raise)

    def _raise(self) -> None:
        self.lift()
        self.focus_force()
        try:
            self.grab_set()
        except tk.TclError:  # Not viewable yet; the dialog still works.
            logger.debug("Could not grab the About dialog", exc_info=True)

    def _center_over(self, parent: ctk.CTk) -> None:
        self.update_idletasks()
        x = parent.winfo_rootx() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_rooty() + (parent.winfo_height() - self.winfo_height()) // 2
        self.wm_geometry(f"+{max(x, 0)}+{max(y, 0)}")

    def _open_path(self, path: Path) -> None:
        # A file or folder that doesn't exist yet: open the closest existing folder.
        target = path
        while not target.exists() and target.parent != target:
            target = target.parent
        try:
            excel.open_in_default_app(target)
        except OSError as error:
            logger.exception("Opening %s failed", target)
            messagebox.showerror(
                "XlsxPress",
                f"Could not open '{target}': {error.strerror or error}.",
                parent=self,
            )
