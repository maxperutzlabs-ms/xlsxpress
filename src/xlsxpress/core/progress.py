"""Progress reporting and cancelling of a running conversion."""

from __future__ import annotations

import threading
from typing import Literal

from xlsxpress.core.errors import ConversionCancelled

#: The steps of a conversion, in order. `saving` is the completion of the workbook
#: after all rows are written, and its move or copy to the output folder.
Stage = Literal["reading", "writing", "saving", "opening"]


class Progress:
    """Connects a running conversion with the frontend that shows it.

    The conversion reports its stage and checks whether it was cancelled. The
    frontend reads the stage and calls `cancel`, usually from another thread.

    Attributes:
        stage: The current step of the conversion.
        fraction: The completed part of the current stage, from 0 to 1, or None if
            the stage cannot tell.
    """

    def __init__(self) -> None:
        self.stage: Stage = "reading"
        self.fraction: float | None = None
        self._cancelled = threading.Event()

    @property
    def cancelled(self) -> bool:
        """True once `cancel` was called."""
        return self._cancelled.is_set()

    def cancel(self) -> None:
        """Ask the conversion to stop at its next check."""
        self._cancelled.set()

    def report(self, stage: Stage, fraction: float | None = None) -> None:
        """Set the current stage and its completed part.

        Args:
            stage: The current step of the conversion.
            fraction: The completed part of the stage, from 0 to 1. Defaults to None.
        """
        self.stage = stage
        self.fraction = fraction

    def raise_if_cancelled(self) -> None:
        """Stop the conversion if it was cancelled.

        Raises:
            ConversionCancelled: If `cancel` was called.
        """
        if self.cancelled:
            raise ConversionCancelled("The conversion was cancelled.")
