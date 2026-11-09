"""Colours, fonts and sizes of the conversion dialog and the About dialog.

Every widget gets its colours from here, so the dialogs look the same whatever the
Windows colour mode is. Sizes are in customtkinter units, which customtkinter
scales with the display scaling.
"""

from __future__ import annotations

import customtkinter as ctk

# Colours

BG = "#F3F5F7"
CARD = "#FFFFFF"
CARD_BORDER = "#DCE3E8"
NAVY = "#1D3C51"
NAVY_BTN = "#24475E"
NAVY_BTN_BORDER = "#4A6578"
NAVY_BTN_HOVER = "#2E5670"
ON_NAVY_MUTED = "#A9BCC8"
TEAL = "#2C7A73"
TEAL_HOVER = "#24665F"
TEAL_DARK = "#1F5A55"
RED = "#B8232F"
RED_HOVER = "#9E1D28"
SECONDARY = "#E8EEF1"
SECONDARY_HOVER = "#DCE5EA"
SECONDARY_BORDER = "#C9D2D9"
TEXT = "#1D2A33"
TEXT_2 = "#33424D"
TEXT_MUTED = "#5E6E79"
DISABLED_TEXT = "#8696A0"
WHITE = "#FFFFFF"
OK_BG = "#E6F2F0"
OK_BORDER = "#B9DAD5"
WARN_BG = "#FDF1E7"
WARN_BORDER = "#F0CFA8"
WARN_TEXT = "#8A4B0F"
WARN_ICON = "#C9741C"
ERR_BG = "#FBEAEA"
ERR_BORDER = "#EBB8BC"
ERR_TEXT = "#8E1B24"

# Geometry

WINDOW_WIDTH = 480
WINDOW_HEIGHT = 640
ABOUT_WIDTH = 440
ABOUT_HEIGHT = 520
#: Padding around the window content.
PAD = 14
#: Vertical gap between blocks.
GAP = 10
CARD_PADX = 14
CARD_PADY = 12
CARD_RADIUS = 10
#: Width of the text inside a card of the conversion dialog.
CARD_INNER_WIDTH = WINDOW_WIDTH - 2 * PAD - 2 * CARD_PADX - 2
#: Width of the text inside a card of the About dialog.
ABOUT_INNER_WIDTH = ABOUT_WIDTH - 2 * PAD - 2 * CARD_PADX - 2
#: Indent of a checkbox's help text, to the start of the checkbox text.
CHECKBOX_TEXT_INDENT = 24

FONT_FAMILY = "Segoe UI"


class Fonts:
    """The fonts of the dialogs.

    Create it once, after the root window exists, and pass it to every dialog.
    """

    def __init__(self) -> None:
        self.h_app = ctk.CTkFont(FONT_FAMILY, 18, "bold")
        self.h_about = ctk.CTkFont(FONT_FAMILY, 16, "bold")
        self.h_file = ctk.CTkFont(FONT_FAMILY, 15, "bold")
        self.h_card = ctk.CTkFont(FONT_FAMILY, 14, "bold")
        self.body = ctk.CTkFont(FONT_FAMILY, 13)
        self.body_bold = ctk.CTkFont(FONT_FAMILY, 13, "bold")
        self.small = ctk.CTkFont(FONT_FAMILY, 12)
        self.small_bold = ctk.CTkFont(FONT_FAMILY, 12, "bold")
        self.badge = ctk.CTkFont(FONT_FAMILY, 12, "bold")
        self.icon = ctk.CTkFont(FONT_FAMILY, 9, "bold")
        self.link = ctk.CTkFont(FONT_FAMILY, 12, underline=True)
        self.link_bold = ctk.CTkFont(FONT_FAMILY, 12, "bold", underline=True)
        self.caption = ctk.CTkFont(FONT_FAMILY, 11, "bold")
        self.btn_primary = ctk.CTkFont(FONT_FAMILY, 14, "bold")
