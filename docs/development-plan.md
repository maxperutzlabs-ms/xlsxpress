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
| Floats | Numbers in Excel's `General` format, at most 16 significant digits | xlsxwriter writes numbers with 16 significant digits; Excel shows 15. A value that needs 17 digits for an exact round trip can differ in its last binary digit (relative difference below 1e-15), e.g. `0.1 + 0.2` is written as `0.3`. Float32 values are converted through their shortest decimal text, so `0.1` stays `0.1`. Decimals are written as numbers. |
| `inf`, `-inf` | Text `inf`, `-inf` | Excel has no infinity value. |
| `NaN`, null | Blank cell | Current behaviour. |
| Booleans | Excel booleans (`TRUE`, `FALSE`) | |
| Dates and datetimes (Parquet) | Excel dates with a date or datetime format | Parquet types are explicit, not guessed, so this is safe and keeps sorting and date calculations working. Format `yyyy-mm-dd`, `yyyy-mm-dd hh:mm:ss`, or with `.000` when a column has fractional seconds. Excel's resolution is 1 ms; finer precision is lost. Dates before 1900-03-01 are written as ISO 8601 text, because Excel's 1900 leap year bug makes earlier serial dates unreliable. Documented. Text files are never parsed as dates. |
| Timezone-aware datetimes | Text in ISO 8601 with offset, e.g. `2024-01-02T10:00:00+01:00` | Documented; Excel has no timezones. Converted by polars, which needs no `tzdata` package on Windows. |
| Time, duration, list, struct, binary columns | Text | Time as ISO 8601 (`10:30:05`), duration like `1 day, 0:00:05`, lists and structs like `[1, 2]` and `{'a': 1}`, binary as hexadecimal. |
| Header formatting | Bold, wrapped, frozen first row, plain autofilter | No Excel Table. `--plain` turns all of this off. Column names are always written as text, and empty names stay empty. |
| Column widths | Estimated from header and first 1,000 rows, between 40 and 125 pixels | Replaces polars autofit and the per-cell width loop. Date columns are always wide enough for the full date, because Excel shows `###` for a date that doesn't fit. |
| Excel limits | Too many rows or columns: `ExcelLimitError` before writing, no file is created. Strings > 32,767 characters: truncated with a `ConversionWarning`. | |
| Reading text files: types | Every column is read as text first. It becomes integers if every value is a plain integer (`-12`), floats if every value is a plain integer or decimal (`1.5`, `.5`, `1.5e-3`, `1e+05`) or infinity (`inf`, `-Infinity`, any case). Otherwise the whole column stays text. | Keeps leading zeros (`007`), `+5`, spaces, decimal commas, dates, `TRUE`/`FALSE`, gene names and IDs like `2310009E13` unchanged. An integer mantissa needs a signed exponent (`1e+05` is a number, `1E5` is text); R and .NET write signed exponents. Trailing zeros are allowed (`1.10` becomes 1.1, the value is unchanged; Excel does the same). Integers beyond the 64-bit range become floats, consistent with the 15-digit decision. Reviewed on 2026-11-06; making full-precision decimals text was rejected, as most pandas float columns would become text. |
| Reading text files: missing values | In an otherwise numeric column, `NA`, `NaN`, `nan`, `N/A`, `#N/A`, `null`, `NULL` and empty fields become blank cells | A column of only such tokens stays text. `-` and `--` are not missing values. No warning, since tools like MaxQuant write `NaN` everywhere. |
| Reading text files: encoding | UTF-8, UTF-16 or UTF-32 byte order mark decides; otherwise UTF-8. Not valid UTF-8: Windows-1252, then Latin-1, with a warning. NUL bytes without a byte order mark: error. | |
| Reading text files: format | Separator from the extension (`.tsv`) or sniffed from the first lines (`.csv`, `.txt`), preferring a separator with the same count on every line. Standard `"` quoting; if that fails, read again without quote handling and warn. Repeated column names are renamed by polars (`a_duplicated_0`), with a warning. | Rows with more fields than the header, e.g. from comment lines at the top, give an error. |
| Output files | Next to the input file by default, named after it (`data.xlsx`). An existing file is never overwritten; a counter is added instead (`data-02.xlsx`). The name is reserved by creating the file exclusively, so parallel conversions never collide. | If a conversion fails, the reserved file is deleted again. A file that can't be read creates no output at all. |
| Temp files | One temp folder for CLI, context menu and dialog: `platformdirs.user_cache_path("xlsxpress", appauthor=False) / "temp"`, e.g. `%LOCALAPPDATA%\xlsxpress\Cache\temp`. Names get a timestamp (`data_2026-11-06_142530.xlsx`), because Excel can't open two workbooks with the same name, even from different folders. Temp files are disposable, a quick way to look at a file in Excel, and open normally without a read-only prompt. | The old folders (`%TEMP%\xlsxpress`, `<appdata>\temp`) are no longer used. A read-only recommendation was dropped on 2026-11-06, because Excel asked about it on every "Open in Excel". To keep edits, save the workbook somewhere else. |
| Temp cleanup | Runs before each conversion to the temp folder, from the CLI too: deletes files older than `max_age_days` (14), then the oldest above `max_files` (100). The app config can override both. Age is the file's modification time. All `*.xlsx` files directly in the temp folder count, edited or not; Excel's lock files (`~$name.xlsx`), other files and subfolders are left alone. | Files that vanish or are locked (open in Excel) during cleanup are skipped. The README must say that temp files are deleted, edited or not. |
| Conversion warnings | Collected by the pipeline and returned with the result (`ConversionResult.warnings`), message text only. The CLI prints `Warning: ...` to stderr; the context menu and the dialog show them in a message box. | Conversion errors show their message in the context menu's error box; the log keeps the traceback. |
| Install | uv tool (recommended), pip (supported) | PyApp exe dropped. |
| Existing PyApp installs | Low priority: a short migration note | Only one or two users. Re-registering overwrites the same registry keys; the old exe and its Python environment have to be removed by hand. |
| Context menu | HKCU `SystemFileAssociations`, command runs `pythonw.exe -m xlsxpress.app ...` from the install environment | `pythonw` avoids a console window. Warn when registering from an environment that isn't a tool install. |
| Python | `>= 3.12` | Matches rawtrace-qc. |
| CI | GitHub Actions like rawtrace-qc: lint job plus test matrix on Linux, Windows, macOS | Registry tests run on Windows only. |
| CD | Last step: publish to PyPI on a version tag (trusted publishing) | |

## Bugs to fix

Bugs in code that will be removed are marked *(removed with PyApp)* and won't be fixed separately. Fixed bugs are ticked in the commit that fixes them.

- [x] Context menu "Open in Excel" cuts the data to `nrows` rows even when preview is off
- [x] Leading zeros, decimals like `1.10` and E-notation IDs are changed by type inference (`1.10` stays a number by decision)
- [x] Strings starting with `=` become formulas
- [x] `inf` and `-inf` crash the conversion
- [x] Timezone-aware datetime columns crash the conversion on Windows (missing `tzdata`)
- [x] UTF-16 files produce garbage without an error; Latin-1 files fail
- [x] Strings over 32,767 characters are truncated without a warning
- [x] Duplicate headers are renamed without a warning (empty headers now stay empty)
- [x] Read errors show a traceback in the CLI
- [x] Read errors show only a generic error box in the GUI (context menu; the dialog part is step 9)
- [x] A type change after row 10,000 fails the conversion
- [x] Temp files can clash with an open workbook of the same name
- [x] CLI temp folder is never cleaned and differs from the app's folder
- [x] ~~Temp files that the user edited and saved get deleted by cleanup~~ Not a bug: temp files are disposable (decided 2026-11-06)
- [x] Cleanup can crash when several conversions run at once
- [ ] Dialog gives no output location after converting
- [ ] Dialog row field stays active when preview is off
- [x] CLI has no `--version`
- [ ] `xlsxpress-app` crashes without the `gui` extra *(removed with PyApp)*
- [ ] Uninstall leaves the PyApp environment behind *(removed with PyApp)*
- [ ] Downloading a newer exe doesn't update the installed one *(removed with PyApp)*
- [ ] Build script picks the wrong cached PyApp version *(removed with PyApp)*
- [ ] README: wrong install commands, missing install step, broken GIF URL on PyPI, label mismatch, CI claim without a workflow
- [ ] Outdated docstrings
- [x] Dead code (`registered_exe_path`)

## Implementation order

Every step ends with green checks (ruff, pyrefly, pytest) and its own commit or commits. For a bug, the test is written first and must fail; the fix follows in the same commit, so the branch never carries failing tests. Work happens on the `dev` branch. Tick fixed bugs in the [bug list](#bugs-to-fix) in the same commit, and update the decisions table when a step settles a decision.

Steps 1–7 are done. Steps 8–11 below list the scope, the suggested design and what "done" means. A *suggested* design is a starting point: if the code shows a better way, take it and record the decision in the table.

1. **Tooling.** Done.
2. **CI.** Done.
3. **Hotfix for the `nrows` bug.** Done, not released.
4. **Writer.** Done.
5. **Reader.** Done.
6. **Pipeline, paths and temp files.** Done. Its decisions are in the table (output files, temp files, temp cleanup, conversion warnings).
7. **CLI: typer → cyclopts.** Done. `xlsxpress.cli:main` returns the exit code, as in rawtrace-qc: 0 for success, 1 for a conversion error (including a missing input file), cyclopts' code for invalid arguments. `Error: ...` and `Warning: ...` go to stderr as plain text, without color, and `Wrote <path>` to stdout.

### 8. Distribution switch: PyApp → uv tool

Goal: `uv tool install "xlsxpress[gui]"` plus `xlsxpress context-menu register` gives the full feature set, and all PyApp code is gone.

Scope:

- **App entry point.** Add `xlsxpress/app/__main__.py`, so `python -m xlsxpress.app open|dialog FILE` works. The launcher route (no arguments) goes away.
- **Generalize `context_menu.py`.** `register()` takes the command to run instead of an exe path, building `"<python>\pythonw.exe" -m xlsxpress.app open "%1"` from `sys.executable` (`pythonw.exe` sits next to `python.exe` in a uv venv's `Scripts` folder; checked). The registry key root stays a module constant, so tests can point it at a test key.
- **CLI subcommands:**
  - `xlsxpress context-menu register`: register both entries (Open in Excel, Open conversion dialog) for `.csv`, `.tsv`, `.txt` and `.parquet`. Warn if `sys.prefix` is not inside the uv tool directory (default `%APPDATA%\uv\tools`, or `UV_TOOL_DIR`), since a project or conda environment may be deleted later. Check that the `gui` extra is installed, and if not, say how to install it.
  - `xlsxpress context-menu unregister`: remove all entries; idempotent.
  - `xlsxpress context-menu status`: show whether entries exist, the registered command, whether its Python still exists, and the config file path.
  - On macOS and Linux the subcommands exit with a clear "only available on Windows" message.
- **Remove PyApp code.** `app/installation.py`, `app/gui/launcher.py`, `scripts/build_app.py`, the `xlsxpress-app` console script, the install manifest (`install.toml`) and the exe path helpers in `appdata.py`. Keep `config.toml` and the log file in the appdata folder.
- **Migration.** If `status` or `register` finds entries pointing at an old `xlsxpress.exe`, say so: `register` overwrites them, and the old exe (`%LOCALAPPDATA%\perutzms\xlsxpress`) and PyApp's environment (under `%LOCALAPPDATA%\pyapp`) can be deleted by hand. Check both paths on a real PyApp install before writing them into messages. Low priority; one or two users.
- **Open question for David:** keep `customtkinter` in the `gui` extra, or make it a required dependency, since the context menu dialog needs it? A required dependency means one less install mistake.

Tests: registry round trip (register, status, unregister) against a test key under `HKCU\Software\xlsxpress-tests`, never the real `SystemFileAssociations` key; Windows only (`skipif`). Unit tests for building the command string and for the tool-environment check.

Bugs: the four *(removed with PyApp)* items are closed by deleting the code; tick them.

### 9. Conversion dialog

Goal: the dialog shows what happened and where the result is.

Scope:

- After a successful conversion, show the output path, a "Show in folder" button and any warnings. Close automatically only when the file was opened in Excel and there are no warnings.
- Enable the Rows field only while "Preview only" is switched on.
- Show full error messages (wrapped, not cut at 80 characters), and log the traceback.
- Update the GUI only from the Tk main thread: the worker thread hands its result over through a queue that `after()` polls, instead of calling widget methods from the thread.

Tests: GUI code isn't unit tested. Move the decisions (options from widget state, which message to show for a result) into plain functions and test those; check the window itself by hand.

Bugs: dialog gives no output location, row field stays active, generic error (dialog part).

### 10. Documentation

Goal: a non-developer can install, use, update and remove XlsxPress from the README alone.

Scope:

- **README:** what it does and why (the Excel conversion problem in two sentences); install with uv (installing uv itself, `uv tool install "xlsxpress[gui]"`, `xlsxpress context-menu register`), with pip as an alternative; usage from the CLI and from File Explorer (including Windows 11's "Show more options"); where temp files go and that they are deleted after 14 days or beyond 100 files, edited or not, so edits must be saved elsewhere; update (`uv tool upgrade xlsxpress`); uninstall (unregister first, then `uv tool uninstall`); manual registry cleanup commands for recovery.
- **Conversion rules and limitations:** a user-facing summary of the decisions table: what stays text, what becomes a number or date, the 15-digit limit, missing values, `inf`, timezones, encodings, Excel limits. Suggested: `docs/conversion-rules.md`, linked from the README.
- Fix the README bugs: install commands, the GIF URL (a raw URL, so PyPI shows it), context menu labels matching the code, and only claim CI/CD that exists.
- `CHANGELOG.md` in the format used by rawtrace-qc, starting with 0.4.0.
- Migration note for PyApp users.
- Check all docstrings against the code (bug: outdated docstrings).

### 11. CD and release 0.4.0

Goal: tagging a version publishes it to PyPI without manual steps.

Scope:

- `.github/workflows/release.yml` on tags `v*`: run the CI checks, `uv build`, check that the tag matches the version in `pyproject.toml`, publish with PyPI trusted publishing (`pypa/gh-action-pypi-publish`), create a GitHub release with the changelog section.
- **Manual step for David:** set up the trusted publisher for `xlsxpress` on PyPI (repository, workflow name, environment).
- Check the dependency lower bounds (`polars>=1.0` is far below the tested 1.41): raise them to versions that work, or add a CI job with `uv sync --resolution lowest-direct`.
- Classifier `Development Status :: 4 - Beta`; version 0.4.0; add `.coverage` and `htmlcov/` to `.gitignore`.
- Merge `dev` into `main` (ask David first), then tag.

The writer came before the reader because its decisions were settled and it could be tested on its own. The distribution switch comes before the dialog fixes, so no time goes into code that gets removed. The CLI switches to cyclopts before the distribution switch, so the new `context-menu` subcommand is written in cyclopts right away.

## Ideas for later

- Optional decimal-comma parsing (`1,5`) for German-locale exports. Not automatic: `1,234` means 1.234 in German and 1234 in English.
- Perseus matrices: their `#!{Type}` annotation rows below the header make every column text.
- An option to write integers over 15 digits as text.

## Testing approach

- **Unit tests** check one function or class on its own. Keep them focused: main behaviour, documented errors, realistic edge cases. No test for every trivial variation; use `pytest.mark.parametrize` instead of near-duplicate tests.
- **Exception: data conversion.** Reading and writing decide whether values survive, so they're tested thoroughly: every data type, every Excel limit, every decision in the table above.
- **Integration tests** check high-level behaviour across modules: a file goes in, a workbook comes out, CLI and GUI paths give the same result, context-menu registration round trips (Windows only).
- Coverage is reported in CI but isn't a target.

## Working notes

- **Test speed.** The whole suite runs in about 2 s. Keep each test well under 1 s; test limits with lowered constants instead of huge inputs.
- **Commits.** Commit the plan together with the change it describes. `AGENTS.md`, `CLAUDE.md` and `.claude/` are not committed.
