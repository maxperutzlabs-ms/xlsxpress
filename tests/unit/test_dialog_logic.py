from __future__ import annotations

from pathlib import Path

import pytest

from xlsxpress.app.gui.dialog_logic import (
    OPEN_ENTRY_LABEL,
    FormState,
    about_text,
    build_options,
    defaults_from_form,
    form_from_defaults,
    outcome_of_error,
    outcome_of_result,
    parse_nrows,
)
from xlsxpress.config import DialogDefaults
from xlsxpress.core.cleanup import CleanupPolicy
from xlsxpress.core.errors import OpenOutputError, ReadError
from xlsxpress.core.pipeline import ConversionResult

pytestmark = pytest.mark.unit

_INPUT = Path("data.tsv")
_OUTPUT = Path("data.xlsx")
_LOG = Path("xlsxpress.log")


def _form(preview: bool = False, nrows_text: str = "100") -> FormState:
    return FormState(open_after=False, to_temp=True, format_header=False, preview=preview, nrows_text=nrows_text)


class TestFormFromDefaults:
    def test_form_from_defaults_round_trips_through_defaults_from_form(self) -> None:
        defaults = DialogDefaults(open_after=False, preview=True, nrows=25, to_temp=False, format_header=False)

        assert defaults_from_form(form_from_defaults(defaults), DialogDefaults()) == defaults


class TestParseNrows:
    @pytest.mark.parametrize(("text", "expected"), [("1", 1), ("250", 250), (" 30 ", 30)])
    def test_parse_nrows_accepts_positive_whole_numbers(self, text: str, expected: int) -> None:
        assert parse_nrows(text) == expected

    @pytest.mark.parametrize("text", ["", "0", "-5", "1.5", "abc"])
    def test_parse_nrows_rejects_other_text_and_names_it(self, text: str) -> None:
        with pytest.raises(ValueError, match=f"positive whole number, not {text!r}"):
            parse_nrows(text)


class TestBuildOptions:
    def test_build_options_takes_the_switches_from_the_form(self) -> None:
        policy = CleanupPolicy(max_age_days=3)

        options = build_options(_INPUT, _form(), policy)

        assert (options.input_path, options.open_after, options.to_temp, options.format_header) == (
            _INPUT,
            False,
            True,
            False,
        )
        assert options.cleanup_policy is policy

    def test_build_options_limits_rows_only_with_preview(self) -> None:
        assert build_options(_INPUT, _form(preview=True, nrows_text="7"), CleanupPolicy()).nrows == 7
        assert build_options(_INPUT, _form(preview=False, nrows_text="7"), CleanupPolicy()).nrows is None

    def test_build_options_ignores_an_invalid_rows_field_without_preview(self) -> None:
        assert build_options(_INPUT, _form(preview=False, nrows_text="abc"), CleanupPolicy()).nrows is None

    def test_build_options_raises_for_an_invalid_rows_field_with_preview(self) -> None:
        with pytest.raises(ValueError, match="positive whole number"):
            build_options(_INPUT, _form(preview=True, nrows_text="0"), CleanupPolicy())


class TestDefaultsFromForm:
    def test_defaults_from_form_saves_the_rows_field_without_preview_when_it_is_valid(self) -> None:
        assert defaults_from_form(_form(preview=False, nrows_text="40"), DialogDefaults(nrows=10)).nrows == 40

    def test_defaults_from_form_keeps_the_saved_rows_without_preview_when_the_field_is_invalid(self) -> None:
        assert defaults_from_form(_form(preview=False, nrows_text="abc"), DialogDefaults(nrows=10)).nrows == 10

    def test_defaults_from_form_raises_for_an_invalid_rows_field_with_preview(self) -> None:
        with pytest.raises(ValueError, match="positive whole number"):
            defaults_from_form(_form(preview=True, nrows_text="abc"), DialogDefaults())


class TestOutcomeOfResult:
    @pytest.mark.parametrize(
        ("opened", "warnings", "close_automatically"),
        [(True, (), True), (True, ("Text was truncated.",), False), (False, (), False)],
        ids=["opened", "opened-with-warnings", "not-opened"],
    )
    def test_outcome_of_result_closes_automatically_only_when_opened_without_warnings(
        self, opened: bool, warnings: tuple[str, ...], close_automatically: bool
    ) -> None:
        outcome = outcome_of_result(ConversionResult(_OUTPUT, warnings), opened=opened)

        assert outcome.kind == "success"
        assert outcome.output_path == _OUTPUT
        assert outcome.warnings == warnings
        assert outcome.close_automatically is close_automatically


class TestOutcomeOfError:
    def test_outcome_of_error_keeps_the_output_path_when_only_opening_failed(self) -> None:
        error = OpenOutputError("The xlsx file was written, but could not be opened.", _OUTPUT, ("Warned.",))

        outcome = outcome_of_error(error, _LOG)

        assert outcome.kind == "open_failed"
        assert outcome.message == str(error)
        assert outcome.output_path == _OUTPUT
        assert outcome.warnings == ("Warned.",)
        assert not outcome.close_automatically

    def test_outcome_of_error_shows_the_full_conversion_error_message(self) -> None:
        message = "Could not read 'data.tsv': " + "a long explanation " * 10

        outcome = outcome_of_error(ReadError(message), _LOG)

        assert outcome.kind == "error"
        assert outcome.message == f"Conversion failed: {message}"
        assert outcome.output_path is None

    def test_outcome_of_error_names_the_log_file_for_unexpected_errors(self) -> None:
        outcome = outcome_of_error(RuntimeError("boom"), _LOG)

        assert outcome.kind == "error"
        assert "unexpected error: RuntimeError('boom')" in outcome.message
        assert str(_LOG) in outcome.message


class TestAboutText:
    def test_about_text_names_the_version_the_open_entry_and_the_files(self) -> None:
        text = about_text("0.4.0", Path("config.toml"), Path("temp"), CleanupPolicy(max_age_days=14, max_files=100))

        assert text.startswith("XlsxPress 0.4.0\n")
        assert f'"{OPEN_ENTRY_LABEL}"' in text
        assert "always opens" in text
        assert "after 14 days, or when there are more than 100" in text
        assert "Temp folder: temp" in text
        assert "Defaults file: config.toml" in text
