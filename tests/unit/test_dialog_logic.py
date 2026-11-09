from __future__ import annotations

from pathlib import Path

import pytest

from xlsxpress.app.gui.dialog_logic import (
    ELLIPSIS,
    OPEN_ENTRY_LABEL,
    FormState,
    about_sections,
    build_options,
    defaults_from_form,
    form_from_defaults,
    outcome_of_error,
    outcome_of_result,
    parse_nrows,
    save_location_help,
    truncate_middle,
    truncate_start,
)
from xlsxpress.config import DialogDefaults
from xlsxpress.core.cleanup import CleanupPolicy
from xlsxpress.core.errors import OpenOutputError, ReadError
from xlsxpress.core.pipeline import ConversionResult

pytestmark = pytest.mark.unit

_INPUT = Path("data.tsv")
_OUTPUT = Path("data.xlsx")
_LOG = Path("xlsxpress.log")
_TEMP_DIR = r"C:\Users\me\AppData\Local\xlsxpress\Cache\temp"


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

    @pytest.mark.parametrize("text", ["", " ", "0", "-5", "1.5", "1e3", "abc"])
    def test_parse_nrows_rejects_other_text_and_names_it(self, text: str) -> None:
        with pytest.raises(ValueError) as excinfo:
            parse_nrows(text)

        assert str(excinfo.value) == f"Invalid number of rows {text!r}. Enter a whole number of rows greater than 0."


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
        with pytest.raises(ValueError, match="whole number of rows greater than 0"):
            build_options(_INPUT, _form(preview=True, nrows_text="0"), CleanupPolicy())


class TestDefaultsFromForm:
    def test_defaults_from_form_saves_the_rows_field_without_preview_when_it_is_valid(self) -> None:
        assert defaults_from_form(_form(preview=False, nrows_text="40"), DialogDefaults(nrows=10)).nrows == 40

    def test_defaults_from_form_keeps_the_saved_rows_without_preview_when_the_field_is_invalid(self) -> None:
        assert defaults_from_form(_form(preview=False, nrows_text="abc"), DialogDefaults(nrows=10)).nrows == 10

    def test_defaults_from_form_raises_for_an_invalid_rows_field_with_preview(self) -> None:
        with pytest.raises(ValueError, match="whole number of rows greater than 0"):
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

    @pytest.mark.parametrize(
        ("opened", "warnings", "title", "banner"),
        [
            (False, (), "Converted", "success"),
            (True, (), "Converted and opened in Excel", "success"),
            (False, ("Warned.",), "Converted with warnings", "warning"),
            (True, ("Warned.",), "Converted with warnings", "warning"),
        ],
    )
    def test_outcome_of_result_sets_the_banner_title_and_style(
        self, opened: bool, warnings: tuple[str, ...], title: str, banner: str
    ) -> None:
        outcome = outcome_of_result(ConversionResult(_OUTPUT, warnings), opened=opened)

        assert (outcome.title, outcome.banner, outcome.message) == (title, banner, "")


class TestOutcomeOfError:
    def test_outcome_of_error_keeps_the_output_path_when_only_opening_failed(self) -> None:
        error = OpenOutputError("The xlsx file was written, but could not be opened.", _OUTPUT, ("Warned.",))

        outcome = outcome_of_error(error, _LOG)

        assert outcome.kind == "open_failed"
        assert outcome.banner == "warning"
        assert outcome.title == "Converted, but not opened"
        assert outcome.message == str(error)
        assert outcome.output_path == _OUTPUT
        assert outcome.warnings == ("Warned.",)
        assert not outcome.close_automatically

    def test_outcome_of_error_shows_the_full_conversion_error_message(self) -> None:
        message = "Could not read 'data.tsv': " + "a long explanation " * 10

        outcome = outcome_of_error(ReadError(message), _LOG)

        assert outcome.kind == "error"
        assert outcome.banner == "error"
        assert outcome.title == "Conversion failed"
        assert outcome.message == message
        assert outcome.output_path is None

    def test_outcome_of_error_names_the_log_file_for_unexpected_errors(self) -> None:
        outcome = outcome_of_error(RuntimeError("boom"), _LOG)

        assert outcome.kind == "error"
        assert outcome.title == "Conversion failed"
        assert "Unexpected error: RuntimeError('boom')" in outcome.message
        assert str(_LOG) in outcome.message


class TestSaveLocationHelp:
    def test_save_location_help_for_the_temp_folder_names_the_cleanup_limits(self) -> None:
        text = save_location_help(True, "demo.csv", CleanupPolicy(max_age_days=7, max_files=30), "demo.xlsx")

        assert (
            text == "Deleted after 7 days or when there are more than 30 workbooks. Save a copy elsewhere to keep it."
        )

    def test_save_location_help_next_to_the_original_names_the_source_file(self) -> None:
        text = save_location_help(False, "demo.csv", CleanupPolicy())

        assert text == "Saved next to demo.csv. Existing files are never overwritten."

    def test_save_location_help_next_to_the_original_names_the_written_workbook(self) -> None:
        text = save_location_help(False, "demo.csv", CleanupPolicy(), "demo-02.xlsx")

        assert text == "Saved as demo-02.xlsx next to demo.csv. Existing files are never overwritten."


class TestAboutSections:
    def test_about_sections_explain_save_reset_and_the_temp_folder(self) -> None:
        sections = dict(about_sections(CleanupPolicy(max_age_days=14, max_files=100)))

        assert list(sections) == ["Save as default", "Reset", "Temp folder"]
        assert f"\u201c{OPEN_ENTRY_LABEL}\u201d" in sections["Save as default"]
        assert "always opens Excel" in sections["Save as default"]
        assert "Nothing is saved until you click Save as default." in sections["Reset"]
        assert sections["Temp folder"] == "Workbooks are deleted after 14 days or beyond 100 files, even if edited."

    def test_about_sections_use_the_configured_cleanup_limits(self) -> None:
        sections = dict(about_sections(CleanupPolicy(max_age_days=3, max_files=9)))

        assert "after 3 days or beyond 9 files" in sections["Temp folder"]


class TestTruncateMiddle:
    @pytest.mark.parametrize(
        ("text", "max_chars", "expected"),
        [
            (r"C:\Users\me\Projects\demo\demo.xlsx", 40, r"C:\Users\me\Projects\demo\demo.xlsx"),
            (r"C:\Users\me\Projects\demo\demo.xlsx", 35, r"C:\Users\me\Projects\demo\demo.xlsx"),
            (r"C:\Users\me\Projects\demo\demo.xlsx", 34, rf"C:\Users\{ELLIPSIS}\demo\demo.xlsx"),
            (r"C:\Users\me\Projects\demo\demo.xlsx", 22, rf"C:\Users\{ELLIPSIS}\demo.xlsx"),
            (r"\\server\share\lab\2026\run\data.tsv", 32, rf"\\server\share\{ELLIPSIS}\run\data.tsv"),
            ("/home/me/projects/demo/demo.csv", 25, f"/home/{ELLIPSIS}/demo/demo.csv"),
        ],
        ids=["fits", "exact-fit", "two-tail-components", "one-tail-component", "unc", "posix"],
    )
    def test_truncate_middle_keeps_the_root_and_the_last_components(
        self, text: str, max_chars: int, expected: str
    ) -> None:
        assert truncate_middle(text, max_chars) == expected

    @pytest.mark.parametrize("max_chars", [1, 2, 5, 12, 21])
    def test_truncate_middle_falls_back_to_characters_from_both_ends(self, max_chars: int) -> None:
        text = r"C:\Users\me\Projects\demo\a_very_long_file_name.xlsx"

        shortened = truncate_middle(text, max_chars)

        assert len(shortened) == max_chars
        start, ellipsis, end = shortened.partition(ELLIPSIS)
        assert ellipsis == ELLIPSIS
        assert text.startswith(start)
        assert text.endswith(end)

    def test_truncate_middle_shortens_a_name_without_separators(self) -> None:
        assert truncate_middle("abcdefghij", 5) == f"ab{ELLIPSIS}ij"

    def test_truncate_middle_falls_back_to_characters_when_leaving_out_folders_is_not_enough(self) -> None:
        assert truncate_middle(r"C:\Users\x\b.txt", 15) == f"C:\\User{ELLIPSIS}x\\b.txt"

    @pytest.mark.parametrize("max_chars", [0, -1])
    def test_truncate_middle_raises_for_a_max_chars_below_one(self, max_chars: int) -> None:
        with pytest.raises(ValueError, match=f"Invalid max_chars {max_chars!r}"):
            truncate_middle("abc", max_chars)


class TestTruncateStart:
    @pytest.mark.parametrize(
        ("max_chars", "expected"),
        [
            (60, _TEMP_DIR),
            (30, rf"{ELLIPSIS}\Local\xlsxpress\Cache\temp"),
            (12, rf"{ELLIPSIS}\Cache\temp"),
            (6, rf"{ELLIPSIS}\temp"),
            (4, f"{ELLIPSIS}emp"),
        ],
    )
    def test_truncate_start_keeps_as_many_last_components_as_fit(self, max_chars: int, expected: str) -> None:
        shortened = truncate_start(_TEMP_DIR, max_chars)

        assert shortened == expected
        assert len(shortened) <= max_chars

    def test_truncate_start_raises_for_a_max_chars_below_one(self) -> None:
        with pytest.raises(ValueError, match="Invalid max_chars 0"):
            truncate_start("abc", 0)
