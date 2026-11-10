# XlsxPress development plan

Status: 2026-11-09. Based on the project review of 2026-11-06.

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
| Reading text files: encoding | UTF-8, UTF-16 or UTF-32 byte order mark decides; otherwise UTF-8. Not valid UTF-8: Windows-1252, then Latin-1, with a warning. NUL bytes without a byte order mark: error. Not valid text in the encoding of the byte order mark (e.g. cut-off UTF-16): error. | |
| Reading text files: format | Separator from the extension (`.tsv`) or sniffed from the first lines (`.csv`, `.txt`). The sniffer removes fully quoted fields (`"a,b"` between separators or line ends) before counting, so separators inside them don't count, while a quote inside a field (`5" cable`) is plain data. It prefers a separator with the same count on every line, then the most frequent on the header line. Standard `"` quoting; if that fails, read again without quote handling and warn if a name or value starts with a quote. The header row is checked first, because polars silently strips quotes from a header like `"5 inch" cable` (a quote inside a field is plain data). Repeated column names are renamed by polars (`a_duplicated_0`), with a warning. | Rows with more fields than the header, e.g. from comment lines at the top, give an error. |
| Output files | Next to the input file by default, named after it (`data.xlsx`). An existing file is never overwritten; a counter is added instead (`data-02.xlsx`). The workbook is first written to a work folder (`%LOCALAPPDATA%\xlsxpress\Cache\work`) and moved (temp folder) or copied (next to the input) to the output name only when complete, so a file under the output name is never unfinished, even if the process is killed (e.g. the dialog is closed during a conversion). The output name is then reserved by creating the file exclusively, so parallel conversions never collide. A copy goes to `<output name>.part` in the output folder and replaces the reserved file with `os.replace` when complete; a process killed during the copy leaves the empty reserved file and the `.part` file, never a truncated workbook. | If a conversion fails, the reserved file and the work file are deleted again. A file that can't be read creates no output at all. Moved to the temp folder, which sits next to the work folder (same drive, same permissions), because a copy is a second new file that Windows Defender checks again when Excel opens it: measured +0.6 s for a 2.5 MB workbook. Copied next to the input, because a moved file keeps the permissions of the cache folder on Windows, and a move fails across drives (network shares). xlsxwriter's own temp files (the uncompressed rows) go to the work folder too. Decided 2026-11-06. |
| Temp files | One temp folder for CLI, context menu and dialog: `platformdirs.user_cache_path("xlsxpress", appauthor=False) / "temp"`, e.g. `%LOCALAPPDATA%\xlsxpress\Cache\temp`. Names get a timestamp (`data_2026-11-06_142530.xlsx`), because Excel can't open two workbooks with the same name, even from different folders. Temp files are disposable, a quick way to look at a file in Excel, and open normally without a read-only prompt. | The old folders (`%TEMP%\xlsxpress`, `<appdata>\temp`) are no longer used. A read-only recommendation was dropped on 2026-11-06, because Excel asked about it on every "Open in Excel". To keep edits, save the workbook somewhere else. |
| Temp cleanup | Runs before each conversion to the temp folder, from the CLI too: deletes files older than `max_age_days` (14), then the oldest above `max_files` (100). The app config can override both. Age is the file's modification time. All `*.xlsx` files directly in the temp folder count, edited or not; Excel's lock files (`~$name.xlsx`), other files and subfolders are left alone. | Files that vanish or are locked (open in Excel) during cleanup are skipped. The README must say that temp files are deleted, edited or not. Before each conversion, files in the work folder older than one day are deleted too; they only exist when a conversion was killed, and the age limit protects conversions still running. |
| Conversion warnings | Collected by the pipeline and returned with the result (`ConversionResult.warnings`), message text only. The CLI prints `Warning: ...` to stderr; the context menu and the dialog show them in a message box. | Conversion errors show their message in the context menu's error box; the log keeps the traceback. |
| Opening the written file fails | `OpenOutputError` (not a `ConversionError`) with the output path and the warnings. The message says the file was written and where, that it couldn't be opened, and that maybe no program is set up for `.xlsx` files. | The CLI prints the warnings and `Wrote <path>`, then `Error: ...`, and exits with 1, because the requested action failed. The context menu shows the warnings and the message in its boxes. Decided 2026-11-06. |
| Install | uv tool (recommended), pip (supported) | PyApp exe dropped. `customtkinter` is a required dependency; there is no `gui` extra, so `uv tool install xlsxpress` gives the full feature set. |
| Existing PyApp installs | Not supported: no migration note and no detection in the code (decided 2026-11-09) | Only one or two users, handled by hand. Re-registering overwrites the same registry keys (command and label). An old `Icon` value is left in place; remove it by hand if needed. |
| Context menu | HKCU `SystemFileAssociations`, command runs `pythonw.exe -m xlsxpress.app ...` from the install environment | `pythonw` avoids a console window; it is found next to `sys.executable`. Registering from an environment that isn't a uv tool install gives a warning, not an error, so a development venv can still register. A uv tool environment is recognised by its `uv-receipt.toml`, which works with any `UV_TOOL_DIR`. No menu icon. |
| Context menu entries | "Open in Excel (XlsxPress)" converts with the saved defaults and always opens the result, whatever the saved "Open in Excel after conversion" says; that option only sets the dialog's default. "Open conversion dialog (XlsxPress)" shows the dialog. | Decided 2026-11-06. |
| Conversion dialog | The row limit field ("Convert only the first [N] rows", config key `preview`) is enabled only when its checkbox is checked. "Reset" sets the form to the built-in defaults without saving; "Save as default" saves. Without the row limit, an invalid value in its field is ignored when converting and keeps the saved value when saving; with it, the value must be a whole number greater than 0, otherwise nothing is converted or saved and the field gets a red border until it is valid. After a conversion that wrote a file, a banner shows the path, "Show in folder" and the warnings, the options are disabled and Convert turns into Close. It closes by itself only when the file was opened in Excel and there are no warnings. Errors are shown in full in the banner and logged with the traceback; Convert stays available. "About" is a modal window with the version, what "Save as default" and "Reset" do, the cleanup rules, and links to the temp folder and the config file. Colours, fonts and sizes live in `app/gui/theme.py`; the window uses light mode and explicit colours, so it looks the same in Windows dark mode. Buttons and checkboxes are reachable with Tab (customtkinter's are not by default). | Closing the window during a conversion ends it; the work folder makes this safe, so it isn't blocked. Decided 2026-11-06. Redesign (card layout, result banner, keyboard access) decided 2026-11-09; converting again after a successful conversion was considered and rejected. |
| Settings | `xlsxpress.config`: config file and log file, shared by CLI and app | Both stay in `platformdirs.user_data_dir("xlsxpress", "perutzms")`, e.g. `%LOCALAPPDATA%\perutzms\xlsxpress`. The CLI no longer imports its settings from the app package. |
| Python | `>= 3.12` | Matches rawtrace-qc. |
| CI | GitHub Actions like rawtrace-qc: lint job plus test matrix on Linux, Windows, macOS | Registry tests run on Windows only. They use an in-memory fake of `winreg` and never touch the real registry, except one round trip through it (`register`, `read_entries`, `unregister` below `Software\xlsxpress-tests`) that runs only with `XLSXPRESS_REGISTRY_TESTS=1`, which only the Windows CI job sets (decided 2026-11-09). |
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
- [x] Read errors show only a generic error box in the GUI
- [x] A type change after row 10,000 fails the conversion
- [x] Temp files can clash with an open workbook of the same name
- [x] CLI temp folder is never cleaned and differs from the app's folder
- [x] ~~Temp files that the user edited and saved get deleted by cleanup~~ Not a bug: temp files are disposable (decided 2026-11-06)
- [x] Cleanup can crash when several conversions run at once
- [x] Opening the written file fails with a traceback (CLI) or "XlsxPress failed" (context menu), as if the conversion had failed
- [x] Dialog gives no output location after converting
- [x] Dialog row field stays active when preview is off
- [x] CLI has no `--version`
- [x] `xlsxpress-app` crashes without the `gui` extra *(removed with PyApp)*
- [x] Uninstall leaves the PyApp environment behind *(removed with PyApp)*
- [x] Downloading a newer exe doesn't update the installed one *(removed with PyApp)*
- [x] Build script picks the wrong cached PyApp version *(removed with PyApp)*
- [x] README: wrong install commands, missing install step, broken GIF URL on PyPI, label mismatch, CI claim without a workflow
- [x] Outdated docstrings
- [x] Dead code (`registered_exe_path`)
- [x] Separator detection counts separators inside quoted fields and silently picks the wrong separator, e.g. `id;"protein,description,info"` becomes three columns
- [x] A process killed while copying the workbook next to the input leaves a truncated `.xlsx` under the output name
- [x] A cut-off UTF-16 file fails with an uncaught `UnicodeDecodeError` instead of a `ReadError`
- [x] The context menu app dies silently if the log file can't be set up
- [x] A config section that isn't a table (`dialog = 1`) makes every conversion fail; `true`, negative numbers and an `nrows` of 0 are accepted as settings

## Implementation order

Every step ends with green checks (ruff, pyrefly, pytest) and its own commit or commits. For a bug, the test is written first and must fail; the fix follows in the same commit, so the branch never carries failing tests. Work happens on the `dev` branch. Tick fixed bugs in the [bug list](#bugs-to-fix) in the same commit, and update the decisions table when a step settles a decision.

Steps 1–10 are done. Step 11 below lists the scope, what is done and what is still missing. A *suggested* design is a starting point: if the code shows a better way, take it and record the decision in the table.

1. **Tooling.** Done.
2. **CI.** Done.
3. **Hotfix for the `nrows` bug.** Done, not released.
4. **Writer.** Done.
5. **Reader.** Done.
6. **Pipeline, paths and temp files.** Done. Its decisions are in the table (output files, temp files, temp cleanup, conversion warnings).
7. **CLI: typer → cyclopts.** Done. `xlsxpress.cli:main` returns the exit code, as in rawtrace-qc: 0 for success, 1 for a conversion error (including a missing input file), cyclopts' code for invalid arguments. `Error: ...` and `Warning: ...` go to stderr as plain text, without color, and `Wrote <path>` to stdout.
8. **Distribution switch: PyApp → uv tool.** Done. `uv tool install xlsxpress` plus `xlsxpress context-menu register` gives the full feature set. The context menu runs `pythonw.exe -m xlsxpress.app open|dialog FILE` (`xlsxpress/app/__main__.py`); the app has no launcher route anymore. `xlsxpress context-menu register | unregister | status` manage the entries; on macOS and Linux they exit with an error. The PyApp code is gone: installer, launcher window, self-update, build script, install manifest, `xlsxpress-app` script and exe path helpers. `customtkinter` became a required dependency and the `gui` extra is gone, so the planned check for the extra wasn't needed. Config and log moved from `app/appdata.py` and `app/config.py` to `xlsxpress/config.py`, so the CLI reads its settings without the app package; the CLI still imports `app.context_menu` for its `context-menu` commands, which is where that code belongs. Checked with a real `uv tool install .`: `register` writes the expected commands and `status` reads them back. The right-click entries themselves still need a look by hand.
9. **Conversion dialog.** Done. Its decisions are in the table (context menu entries, conversion dialog). The decisions live in `app/gui/dialog_logic.py` and are unit tested; the window in `app/gui/dialog.py` only reads widgets and shows outcomes. The worker thread hands its outcome to the Tk main thread through a queue polled with `after()`. All states were checked by driving the real window from a script and taking screenshots; "Show in folder" (`explorer /select,`) and the window on a second monitor still need a look by hand. Two changes came first, each in its own commit: unfinished workbooks are written to a work folder (table: output files, temp cleanup), and a file that was written but couldn't be opened raises `OpenOutputError` (table: opening the written file fails).

10. **Documentation.** Done. A non-developer can install, use, update and remove XlsxPress from the README alone. The README covers what XlsxPress does and why, install with uv (and pip), usage from File Explorer (including Windows 11's "Show more options") and the CLI, the temp folder and its cleanup, conversion rules and limitations, update, uninstall and manual registry cleanup. The conversion rules live in the README itself; a separate `docs/conversion-rules.md` and a migration note for PyApp users were dropped on 2026-11-06. `CHANGELOG.md` uses the format of rawtrace-qc and starts with 0.4.0. Docstrings were checked against the code. Leftover PyApp build files (`build/pyapp/`, old exe files and wheels in `dist/`) were deleted locally on 2026-11-09.

### 11. CD and release 0.4.0

Goal: tagging a version publishes it to PyPI without manual steps.

Scope:

- [x] `.github/workflows/release.yml` on tags `v*`. Written on 2026-11-09 and merged into `main`; it runs for the first time with the `v0.4.0` tag. Jobs in order: validate (tag matches `version` in `pyproject.toml`, newest changelog entry is `## [X.Y.Z] - YYYY-MM-DD`), the CI checks (`ci.yml` is called as a reusable workflow), `uv build` with a smoke test of the wheel (`xlsxpress --version`), publish to PyPI with trusted publishing (`pypa/gh-action-pypi-publish`, environment `pypi`), then a GitHub release with the newest changelog entry and the built files. PyPI comes before the GitHub release because a version can never be uploaded to PyPI again, while a GitHub release can be redone.
- [x] Release tags start with `v`. The older tags `0.1.0` to `0.3.0` were renamed to `v0.1.0` to `v0.3.0` on 2026-11-09, with the same tagger, date and message.
- [x] Trusted publisher for `xlsxpress` on PyPI (repository, workflow name, environment), set up by David on 2026-11-09, together with the GitHub environment `pypi`, which allows only `v*` tags.
- [x] Classifier `Development Status :: 4 - Beta` (done 2026-11-09).
- [x] `.coverage` and `htmlcov/` in `.gitignore` (done 2026-11-06).
- [x] Polars lower bound: raised to `polars>=1.18` on 2026-11-06. The full test suite passes on the latest patch of every minor version from 1.18 to 1.41. Below 1.18 the tests can't be collected because `pl.Int128` is missing. With that attribute stubbed, 1.10 to 1.17 pass all other tests, while 1.9 and older fail real reader and writer tests (e.g. no warning for irregular quotes, `OverflowError` for large integers).
- [x] Lower bounds of the other dependencies, checked on 2026-11-09 with the full test suite on Python 3.12 and 3.14 (other dependencies at their locked versions), going down through the latest patch of every minor version and then the first patch of the lowest passing minor. Versions below a declared bound couldn't be resolved, so the search stopped there. `xlsxwriter` raised to `>=3.2.1`: 3.2.0 lacks `xlsxwriter.utility.cell_autofit_width`. `customtkinter` raised to `>=5.2.1`, checked with a script that drives the real dialog through a conversion, since the test suite doesn't import customtkinter: 5.2.0 imports `distutils`, which Python 3.12 removed, so the dialog fails. `platformdirs>=4.0` (4.0.0 to 4.14.5) and `cyclopts>=4.8` (4.8.0 to 5.4.0) stay. All minimums together (`uv pip install --resolution lowest-direct`, including polars 1.18.0) pass the suite and the dialog check on 3.12 and 3.14.
- [x] Lower bounds rechecked on 2026-11-09 below the declared bounds, with checks beyond the test suite, which misses both changes. `platformdirs` lowered to `>=2.1`: 2.0.x lacks `user_cache_path` (the tests monkeypatch `paths.get_cache_dir`, so they pass anyway), and 2.1.0 to 4.10.0 return the same data and cache paths on Windows. `cyclopts` raised to `>=5.0`: 4.x exits with 1 for invalid arguments, the same code as a conversion error, and 5.0.0 exits with 2. `xlsxwriter>=3.2.1` and `customtkinter>=5.2.1` confirmed: no older release works. xlsxwriter 3.2.1, platformdirs 2.1.0, cyclopts 5.0.0 and customtkinter 5.2.1 together pass the suite, the dialog check, the path check and the exit-code check on Python 3.12 and 3.14 (Windows only). Two tests now cover these gaps: `TestGetCacheDir` calls the real `get_cache_dir()`, and the CLI test asserts exit code 2 for invalid arguments; they fail with platformdirs 2.0.2 and cyclopts 4.25.3. The latest releases (polars 1.44.2 with `<2`, xlsxwriter 3.3.1, platformdirs 4.14.5, cyclopts 5.4.0, customtkinter 6.0.0) pass the same checks on 3.12 and 3.14, the dialog check without deprecation warnings.
- [x] polars 2: polars 2.1.0 no longer parses a quote inside an unquoted field (`5" cable`) with quote handling, so the reader fell back and warned. The fallback now warns only if a column name or value starts with a quote, which gives the same result on polars 1.x and 2.x. The full suite passes on both.
- [x] The registry tests no longer write to the real registry on a developer's machine: they use an in-memory fake of `winreg`, and only the Windows CI job runs one round trip through the real registry (done 2026-11-09).
- [ ] Manual checks left over from steps 8 and 9: the right-click entries themselves, "Show in folder" and the dialog on a second monitor.
- [ ] Version 0.4.0 in `pyproject.toml` (with `uv lock`) and the release date in `CHANGELOG.md`, in one commit; the release workflow checks both against the tag.
- [ ] CI is green on `main` for the release commit.
- [ ] Tag `v0.4.0` on `main` and push the tag (ask David first).

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
