# XlsxPress development plan

Status: 2026-11-06. Based on the project review of the same date.

## Overview of changes

1. **Distribution: from PyApp to uv.** `uv tool install xlsxpress` becomes the recommended install, and `pip install xlsxpress` is supported. A new `xlsxpress context-menu register | unregister | status` command replaces the PyApp launcher. The PyApp build script, installer, self-update logic and launcher window are removed. Updates and uninstall go through `uv tool upgrade` and `uv tool uninstall`.
2. **Excel writing: our own xlsxwriter code.** Polars is used only for reading. The writer streams rows (`constant_memory`), so memory use stays flat for large files (measured: 740 MB → 131 MB and 12.1 s → 8.4 s for 100k × 20 cells).
3. **Text-safe reading.** Values in the input file must not change on their way into Excel, for example through leading zeros or E-notation IDs. Columns become numeric only when that's lossless.
4. **Bug fixes.** See [Bugs to fix](#bugs-to-fix).
5. **Clear errors.** Errors in reading and writing become one conversion error with an actionable message, shown by both the CLI and the GUI.
6. **CLI: from typer to cyclopts.** Consistent with rawtrace-qc.
7. **Tooling aligned with rawtrace-qc.** ruff, pyrefly, pytest and pytest-cov are pinned in the `dev` group, the tool config lives in `pyproject.toml`, and `tests/ruff.toml` allows 120-character lines.
8. **Testing as we go.** Unit and integration tests are added with each change. See [Testing approach](#testing-approach).
9. **CI early, CD at the end.** GitHub Actions run lint, format, type checks and the test matrix from the start. Publishing to PyPI on a tag comes last.
10. **Python support.** `requires-python` goes from 3.14 down to 3.12, like rawtrace-qc.
11. **Documentation.** The README is rewritten: install, context menu, conversion rules, limitations. A `CHANGELOG.md` is added.

## Decisions

| Topic | Decision | Notes |
|---|---|---|
| Excel writer | Own xlsxwriter code, `constant_memory` mode | Full control per data type; no surprises from polars' Excel defaults (e.g. its old 3-decimal default). |
| Strings | Always written as text | `strings_to_formulas`, `strings_to_urls` and `strings_to_numbers` are off. `=1+1` stays text. |
| Integers > 15 digits | Written as numbers, like Excel itself | Excel keeps only 15 significant digits and replaces the rest with zeros. Documented as a limitation. Keep the writer's type mapping flexible, so this can later become text or an option. |
| `inf`, `-inf` | Text `inf`, `-inf` | Excel has no infinity value. |
| `NaN`, null | Blank cell | Current behaviour. |
| Dates and datetimes (Parquet) | Excel dates with a date or datetime format | Parquet types are explicit, not guessed, so this is safe and keeps sorting and date calculations working. Excel's resolution is 1 ms; finer precision is lost. Documented. Text files are never parsed as dates. |
| Timezone-aware datetimes | Text in ISO 8601 with offset | Documented; Excel has no timezones. |
| Time, duration, list, struct, binary columns | Text | e.g. `[1, 2]`. |
| Header formatting | Bold, wrapped, frozen first row, plain autofilter | No Excel Table. `--plain` turns all of this off. |
| Column widths | Estimated from header and first rows, capped | Replaces polars autofit and the per-cell width loop. |
| Excel limits | Too many rows or columns: clear error before writing. Strings > 32,767 characters: truncated with a warning. | |
| Install | uv tool (recommended), pip (supported) | PyApp exe dropped. |
| Existing PyApp installs | Low priority: a short migration note | Only one or two users. Re-registering overwrites the same registry keys; the old exe and its Python environment have to be removed by hand. |
| Context menu | HKCU `SystemFileAssociations`, command runs `pythonw.exe -m xlsxpress.app ...` from the install environment | `pythonw` avoids a console window. Warn when registering from an environment that isn't a tool install. |
| Python | `>= 3.12` | Matches rawtrace-qc. |
| CI | GitHub Actions like rawtrace-qc: lint job plus test matrix on Linux, Windows, macOS | Registry tests run on Windows only. |
| CD | Last step: publish to PyPI on a version tag (trusted publishing) | |

## Bugs to fix

Bugs in code that will be removed are marked *(removed with PyApp)* and won't be fixed separately.

- Context menu "Open in Excel" cuts the data to `nrows` rows even when preview is off
- Leading zeros, decimals like `1.10` and E-notation IDs are changed by type inference
- Strings starting with `=` become formulas
- `inf` and `-inf` crash the conversion
- UTF-16 files produce garbage without an error; Latin-1 files fail
- Strings over 32,767 characters are truncated without a warning
- Empty and duplicate headers are renamed without a warning
- Polars errors are not caught (CLI traceback, generic GUI error box)
- A type change after row 10,000 fails the conversion
- Temp files can clash with an open workbook of the same name
- CLI temp folder is never cleaned and differs from the app's folder
- Temp files that the user edited and saved get deleted by cleanup
- Cleanup can crash when several conversions run at once
- Dialog gives no output location after converting
- Dialog row field stays active when preview is off
- CLI has no `--version`
- `xlsxpress-app` crashes without the `gui` extra *(removed with PyApp)*
- Uninstall leaves the PyApp environment behind *(removed with PyApp)*
- Downloading a newer exe doesn't update the installed one *(removed with PyApp)*
- Build script picks the wrong cached PyApp version *(removed with PyApp)*
- README: wrong install commands, missing install step, broken GIF URL on PyPI, label mismatch, CI claim without a workflow
- Outdated docstrings and dead code (`registered_exe_path`)

## Implementation order

Every step ends with green checks (ruff, pyrefly, pytest) and its own commit or commits. For a bug, the test is written first and must fail; the fix follows in the same commit, so the branch never carries failing tests. Work happens on the `dev` branch.

1. **Tooling.** Dev group and tool config as in rawtrace-qc, `tests/ruff.toml`, pytest markers, `requires-python >= 3.12`, current `uv_build` pin. Fix the existing ruff findings and pyrefly errors.
2. **CI.** `.github/workflows/ci.yml` modelled on rawtrace-qc. From here on every push is checked.
3. **Hotfix for the `nrows` bug** in the current PyApp app. Not released right away; few users are affected.
4. **Writer.** New `writers.py` with the type mapping, Excel limits, column widths and header formatting from the decisions table. Thorough unit tests per data type and limit. Remove `polars.write_excel`.
5. **Reader.** Text-safe type inference, encoding detection, error wrapping and the late type change. Thorough unit tests.
6. **Pipeline and paths.** Conversion error type, unique temp names, one temp folder for CLI and app, cleanup fixes, read-only recommendation for temp files. First integration tests: input file → `.xlsx`, checked cell by cell.
7. **CLI.** Switch from typer to cyclopts, add `--version`, consistent messages and help texts.
8. **Distribution switch.** `context-menu` subcommand, a generalized `context_menu.py`, `xlsxpress.app` entry point for `open` and `dialog`. Remove the PyApp code. Short migration note for existing PyApp installs.
9. **Dialog.** Show the output location, enable the row field only with preview, show full error messages.
10. **Documentation.** README rewrite, conversion rules and limitations, CHANGELOG, migration note.
11. **CD and release 0.4.0.** Release workflow to PyPI, classifiers (Beta), dependency lower bounds checked.

The writer comes before the reader because its decisions are settled and it can be tested on its own. The distribution switch comes before the app-level fixes, so no time goes into code that gets removed. The CLI switches to cyclopts before the distribution switch, so the new `context-menu` subcommand is written in cyclopts right away.

## Testing approach

- **Unit tests** check one function or class on its own. Keep them focused: main behaviour, documented errors, realistic edge cases. No test for every trivial variation; use `pytest.mark.parametrize` instead of near-duplicate tests.
- **Exception: data conversion.** Reading and writing decide whether values survive, so they're tested thoroughly: every data type, every Excel limit, every decision in the table above.
- **Integration tests** check high-level behaviour across modules: a file goes in, a workbook comes out, CLI and GUI paths give the same result, context-menu registration round trips (Windows only).
- Coverage is reported in CI but isn't a target.
