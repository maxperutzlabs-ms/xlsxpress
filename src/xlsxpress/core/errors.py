"""Exceptions and warnings of the conversion."""

from __future__ import annotations


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


class ConversionWarning(UserWarning):
    """Warns that values or names were changed, or may be read incorrectly."""
