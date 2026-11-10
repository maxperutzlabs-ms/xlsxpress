from __future__ import annotations

import pytest

from xlsxpress.core.errors import ConversionCancelled, ConversionError
from xlsxpress.core.progress import Progress

pytestmark = pytest.mark.unit


class TestProgress:
    def test_progress_starts_with_reading_and_is_not_cancelled(self) -> None:
        progress = Progress()

        assert (progress.stage, progress.fraction, progress.cancelled) == ("reading", None, False)
        progress.raise_if_cancelled()

    def test_report_sets_stage_and_fraction_and_resets_the_fraction_by_default(self) -> None:
        progress = Progress()

        progress.report("writing", 0.5)
        assert (progress.stage, progress.fraction) == ("writing", 0.5)
        progress.report("saving")
        assert (progress.stage, progress.fraction) == ("saving", None)

    def test_raise_if_cancelled_raises_conversion_cancelled_after_cancel(self) -> None:
        progress = Progress()

        progress.cancel()

        assert progress.cancelled
        with pytest.raises(ConversionCancelled):
            progress.raise_if_cancelled()

    def test_conversion_cancelled_is_not_a_conversion_error(self) -> None:
        assert not issubclass(ConversionCancelled, ConversionError)
