"""Factories for the widgets that the dialogs repeat, styled with `theme`.

customtkinter buttons and checkboxes can't be reached with the keyboard, so the
factories make them focusable with Tab, activatable with Return and Space, and give
them a visible focus indicator.
"""

from __future__ import annotations

import tkinter as tk
from collections.abc import Callable, Sequence
from typing import Any

import customtkinter as ctk

from xlsxpress.app.gui import theme
from xlsxpress.app.gui.dialog_logic import truncate_middle
from xlsxpress.app.gui.theme import Fonts


def card(parent: Any) -> ctk.CTkFrame:
    """Return a white card with a thin border."""
    return ctk.CTkFrame(
        parent,
        fg_color=theme.CARD,
        border_width=1,
        border_color=theme.CARD_BORDER,
        corner_radius=theme.CARD_RADIUS,
    )


def text_label(
    parent: Any, text: str, font: ctk.CTkFont, color: str, **kwargs: Any
) -> ctk.CTkLabel:
    """Return a label that is only as high as its text, aligned left."""
    return ctk.CTkLabel(
        parent,
        text=text,
        font=font,
        text_color=color,
        fg_color="transparent",
        height=1,
        anchor="w",
        justify="left",
        **kwargs,
    )


def card_title(parent: Any, number: int, text: str, fonts: Fonts) -> ctk.CTkFrame:
    """Return the title row of a card: a round step badge and the title."""
    row = ctk.CTkFrame(parent, fg_color="transparent")
    RoundBadge(row, 22, fonts.badge, str(number), theme.TEAL).pack(side="left")
    text_label(row, text, fonts.h_card, theme.TEXT).pack(side="left", padx=(8, 0))
    return row


def secondary_button(
    parent: Any, text: str, command: Callable[[], None], fonts: Fonts, **kwargs: Any
) -> ctk.CTkButton:
    """Return a light grey button with navy text."""
    options: dict[str, Any] = {
        "height": 36,
        "corner_radius": 6,
        "font": fonts.body,
        "fg_color": theme.SECONDARY,
        "hover_color": theme.SECONDARY_HOVER,
        "border_width": 1,
        "border_color": theme.SECONDARY_BORDER,
        "text_color": theme.NAVY,
        "text_color_disabled": theme.DISABLED_TEXT,
    }
    options.update(kwargs)
    button = ctk.CTkButton(parent, text=text, command=command, **options)
    make_keyboard_button(button, ring_color=theme.TEAL)
    return button


def primary_button(
    parent: Any, text: str, command: Callable[[], None], fonts: Fonts, **kwargs: Any
) -> ctk.CTkButton:
    """Return the red button of the dialog's main action."""
    options: dict[str, Any] = {
        "height": 40,
        "corner_radius": 8,
        "font": fonts.btn_primary,
        "fg_color": theme.RED,
        "hover_color": theme.RED_HOVER,
        "border_width": 0,
        "text_color": theme.WHITE,
        "text_color_disabled": theme.WHITE,
    }
    options.update(kwargs)
    button = ctk.CTkButton(parent, text=text, command=command, **options)
    make_keyboard_button(button, ring_color=theme.NAVY)
    return button


def link_label(
    parent: Any, text: str, command: Callable[[], object], font: ctk.CTkFont
) -> ctk.CTkLabel:
    """Return an underlined teal label that runs `command` when clicked."""
    label = text_label(parent, text, font, theme.TEAL, cursor="hand2")
    label.bind("<Button-1>", lambda _event: command())
    _make_focusable(
        label,
        command,
        on_focus=lambda: label.configure(
            text_color=theme.TEAL_DARK, fg_color=theme.OK_BG
        ),
        on_blur=lambda: label.configure(text_color=theme.TEAL, fg_color="transparent"),
    )
    return label


def checkbox(
    parent: Any,
    text: str,
    variable: ctk.BooleanVar,
    fonts: Fonts,
    command: Callable[[], None] | None = None,
) -> ctk.CTkCheckBox:
    """Return a checkbox with a teal fill when checked."""
    box = ctk.CTkCheckBox(
        parent,
        text=text,
        variable=variable,
        command=command,
        fg_color=theme.TEAL,
        hover_color=theme.TEAL_HOVER,
        border_color=theme.SECONDARY_BORDER,
        checkmark_color=theme.WHITE,
        corner_radius=4,
        border_width=2,
        checkbox_width=18,
        checkbox_height=18,
        text_color=theme.TEXT,
        text_color_disabled=theme.DISABLED_TEXT,
        font=fonts.body,
    )
    _make_focusable(
        box,
        box.toggle,
        on_focus=lambda: box.configure(
            border_color=theme.TEAL, text_color=theme.TEAL_DARK
        ),
        on_blur=lambda: box.configure(
            border_color=theme.SECONDARY_BORDER, text_color=theme.TEXT
        ),
    )
    return box


def make_keyboard_button(button: ctk.CTkButton, ring_color: str) -> None:
    """Make `button` reachable with Tab and show a ring while it has the focus."""
    width = button.cget("border_width")
    color = button.cget("border_color")
    _make_focusable(
        button,
        button.invoke,
        on_focus=lambda: button.configure(border_width=2, border_color=ring_color),
        on_blur=lambda: button.configure(border_width=width, border_color=color),
    )


class RoundBadge(ctk.CTkFrame):
    """A coloured circle with a short white text, such as a step number or an icon.

    A `CTkLabel` with a corner radius pads its text by that radius and turns into a
    pill, so the circle is a frame with the text placed in its centre.
    """

    def __init__(
        self, parent: Any, size: int, font: ctk.CTkFont, text: str = "", color: str = ""
    ) -> None:
        """Create the badge.

        Args:
            parent: The parent widget.
            size: The diameter.
            font: The font of the text.
            text: The text.
            color: The colour of the circle.
        """
        super().__init__(
            parent,
            width=size,
            height=size,
            corner_radius=size // 2,
            fg_color=color or theme.TEAL,
        )
        self._label = ctk.CTkLabel(
            self,
            text=text,
            width=1,
            height=1,
            fg_color="transparent",
            text_color=theme.WHITE,
            font=font,
        )
        self._label.place(relx=0.5, rely=0.5, anchor="center")

    def set(self, text: str, color: str) -> None:
        """Change the text and the colour of the circle."""
        self.configure(fg_color=color)
        self._label.configure(text=text)


class SegmentedChoice(ctk.CTkFrame):
    """A row of buttons on a grey track, of which exactly one is selected.

    Used instead of `CTkSegmentedButton`, which has one text colour for selected and
    unselected segments and can't be reached with the keyboard.
    """

    def __init__(
        self,
        parent: Any,
        values: Sequence[str],
        fonts: Fonts,
        command: Callable[[str], None] | None = None,
    ) -> None:
        """Create the choice with the first value selected.

        Args:
            parent: The parent widget.
            values: The texts of the segments.
            fonts: The dialog fonts.
            command: Called with the value when the user selects a segment.
        """
        super().__init__(parent, fg_color=theme.SECONDARY, corner_radius=8)
        self._command = command
        self._value = values[0]
        self._buttons: dict[str, ctk.CTkButton] = {}
        for column, value in enumerate(values):
            self.grid_columnconfigure(column, weight=1, uniform="segment")
            button = ctk.CTkButton(
                self,
                text=value,
                height=30,
                corner_radius=6,
                border_width=0,
                font=fonts.body_bold,
                command=lambda value=value: self._select(value),
            )
            button.grid(row=0, column=column, sticky="ew", padx=3, pady=3)
            make_keyboard_button(button, ring_color=theme.NAVY)
            self._buttons[value] = button
        self._paint()

    def get(self) -> str:
        """Return the selected value."""
        return self._value

    def set(self, value: str) -> None:
        """Select `value` without calling the command."""
        self._value = value
        self._paint()

    def set_enabled(self, enabled: bool) -> None:
        """Enable or disable all segments; the selection stays visible."""
        for button in self._buttons.values():
            button.configure(state="normal" if enabled else "disabled")

    def _select(self, value: str) -> None:
        self.set(value)
        if self._command is not None:
            self._command(value)

    def _paint(self) -> None:
        for value, button in self._buttons.items():
            if value == self._value:
                button.configure(
                    fg_color=theme.TEAL,
                    hover_color=theme.TEAL_HOVER,
                    text_color=theme.WHITE,
                    text_color_disabled=theme.WHITE,
                )
            else:
                button.configure(
                    fg_color=theme.SECONDARY,
                    hover_color=theme.SECONDARY_HOVER,
                    text_color=theme.NAVY,
                    text_color_disabled=theme.DISABLED_TEXT,
                )


def fit_text(
    text: str,
    font: ctk.CTkFont,
    width: int,
    truncate: Callable[[str, int], str] = truncate_middle,
) -> str:
    """Shorten `text` with `truncate` until it is at most `width` pixels wide in `font`.

    Widths are in customtkinter units, before display scaling, like the fonts.
    """
    if font.measure(text) <= width:
        return text
    for max_chars in range(len(text) - 1, 0, -1):
        shortened = truncate(text, max_chars)
        if font.measure(shortened) <= width:
            return shortened
    return truncate(text, 1)


def tooltip(widget: Any, text: Callable[[], str]) -> None:
    """Show `text()` in a small box below `widget` while the mouse is over it."""
    tip: list[tk.Toplevel | None] = [None]

    def show(_event: tk.Event) -> None:
        if tip[0] or not text():
            return
        x = widget.winfo_rootx() + 20
        y = widget.winfo_rooty() + widget.winfo_height() + 4
        window = tk.Toplevel(widget)
        window.wm_overrideredirect(True)
        window.wm_geometry(f"+{x}+{y}")
        tk.Label(
            window,
            text=text(),
            background=theme.CARD,
            foreground=theme.TEXT_2,
            highlightthickness=1,
            highlightbackground=theme.CARD_BORDER,
            borderwidth=0,
            padx=6,
            pady=3,
            justify="left",
            font=(theme.FONT_FAMILY, 9),
        ).pack()
        tip[0] = window

    def hide(_event: tk.Event) -> None:
        if tip[0]:
            tip[0].destroy()
            tip[0] = None

    widget.bind("<Enter>", show, add="+")
    widget.bind("<Leave>", hide, add="+")


def _make_focusable(
    widget: Any,
    activate: Callable[[], object],
    on_focus: Callable[[], None],
    on_blur: Callable[[], None],
) -> None:
    # customtkinter widgets are frames that draw on a canvas and route bind() to
    # the canvas. Key events go to the focused frame itself, so the frame takes the
    # focus and gets the bindings, through the tkinter methods.
    def on_key(_event: tk.Event) -> str:
        activate()
        return "break"  # Keeps Return from also triggering the window's binding.

    tk.Frame.configure(widget, takefocus=1)
    for sequence in ("<Return>", "<KP_Enter>", "<space>"):
        tk.Misc.bind(widget, sequence, on_key)
    tk.Misc.bind(widget, "<FocusIn>", lambda _event: on_focus(), "+")
    tk.Misc.bind(widget, "<FocusOut>", lambda _event: on_blur(), "+")
