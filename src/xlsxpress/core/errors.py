"""Exceptions and warnings of the conversion."""

from __future__ import annotations

from pathlib import Path


class ConversionError(ValueError):
    """Base class for all errors that stop a conversion."""


class UnsupportedFormatError(ConversionError):
    """Raised when a file extension is not supported."""


class ReadError(ConversionError):
    """Raised when an input file cannot be read as a table."""


class WriteError(ConversionError):
    """Raised when the xlsx file cannot be written."""


class ExcelLimitError(ConversionError):
    """Raised when a table does not fit into an Excel worksheet."""


class OpenOutputError(Exception):
    """Raised when the xlsx file was written but could not be opened.

    The conversion itself succeeded, so this is not a `ConversionError`. Frontends
    report the written file together with this error.

    Attributes:
        output_path: Path of the written xlsx file.
        warnings: Messages of the conversion's `ConversionWarning`s.
    """

    def __init__(
        self, message: str, output_path: Path, warnings: tuple[str, ...] = ()
    ) -> None:
        super().__init__(message)
        self.output_path = output_path
        self.warnings = warnings


class ConversionWarning(UserWarning):
    """Warns that values or names were changed, or may be read incorrectly."""
