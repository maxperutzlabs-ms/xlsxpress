"""Exceptions and warnings of the conversion."""

from __future__ import annotations


class ConversionWarning(UserWarning):
    """Warns that values had to be changed to fit into an Excel workbook."""


class ExcelLimitError(ValueError):
    """Raised when a table does not fit into an Excel worksheet."""
