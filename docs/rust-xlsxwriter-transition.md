# Possible transition to rust_xlsxwriter

This document records an investigation from 2026-11-10 into why large conversions are slow and how a Rust-based Excel writer could speed them up. Nothing here is decided or implemented. It is the basis for the implementation, should we decide to do it.

## Summary

- Almost all conversion time is spent writing the workbook with xlsxwriter, which is pure Python and costs about 4 µs per cell. Reading is negligible.
- Tuning within xlsxwriter gives at most 10–15 %.
- [rust_xlsxwriter](https://github.com/jmcnamara/rust_xlsxwriter) is a Rust crate by the author of xlsxwriter. It has no official Python package.
- The most likely route is [rustpy-xlsxwriter](https://github.com/rahmadafandi/rustpy-xlsxwriter), a third-party Python wrapper around that crate. It made the write step 5–7x faster on our test files and wrote every number exactly.
- It takes a whole DataFrame instead of single cells, so several conversion decisions in `docs/design.md` must be reproduced by preparing the columns in polars first. One decision (dates before 1900) has to change.
- The open points with the highest risk are progress reporting and cancelling, which the current writer does per row, and the maturity of the wrapper.

## Where the time goes

Measured on commit `376bbac` (before progress reporting was added), Windows 11, Python 3.12, polars 1.41.2, xlsxwriter 3.2.9. The input files were Spectronaut reports.

| Stage | 1.6 MB, 0.24 M cells | 15 MB, 1.1 M cells | 48 MB, 3.7 M cells |
|---|---|---|---|
| Imports (mostly polars) | 0.39 s | 0.34 s | 0.35 s |
| Read: encoding check, header check, `read_csv` | 0.01 s | 0.02 s | 0.06 s |
| Read: numeric conversion | 0.05 s | 0.11 s | 0.32 s |
| Write: column widths, header | 0.01 s | 0.01 s | 0.02 s |
| Write: cell loop | 0.78 s | 4.47 s | 16.56 s |
| Write: `workbook.close()` (zip) | 0.22 s | 1.30 s | 3.08 s |
| Total | 1.46 s | 6.25 s | 20.40 s |

Not included: interpreter start, loading the config, temp cleanup, the final move or copy, and Excel opening the file.

The cell loop scales linearly with the number of non-empty cells. The profile shows the time spread thinly over xlsxwriter's per-cell Python code; getting the rows out of polars takes under 2 % of the loop.

Variants tried on the 15 MB file (current: 5.70 s):

| Variant | Total | Note |
|---|---|---|
| In-memory mode instead of `constant_memory` | 5.74 s | Only moves the work into `close()`. |
| Calling xlsxwriter's undecorated private cell writers | 5.20 s | Private API. |
| Zip compression level 1 | 4.93 s | File 12.2 MB instead of 10.1 MB; xlsxwriter has no option for it. |

## Options

| Option | Assessment |
|---|---|
| rustpy-xlsxwriter | Most likely route. Prebuilt wheels, no build step for us, measured below. |
| Own thin binding over rust_xlsxwriter (PyO3, maturin) | Keeps every per-cell decision, progress and cancelling in our hands. Adds a Rust toolchain, wheel builds for three platforms and a compiled package to the release process. The fallback if rustpy-xlsxwriter turns out unsuitable. |
| Generating the sheet XML ourselves with polars expressions | No dependency, likely fast, but replaces the whole writer with our own xlsx implementation. Not investigated further. |
| pyaccelsx, xlsxturbo | Dropped: little adoption. pyaccelsx still makes one Python call per cell and claims only 2.2x. xlsxturbo writes integers above 2^53 as text. |

## rustpy-xlsxwriter

Evaluated version: 0.9.1 (released 2026-11-07), built on rust_xlsxwriter 0.102.0. It installed from a prebuilt Windows wheel without a Rust toolchain. The README states wheels for Linux, macOS and Windows; the Windows wheels cover CPython 3.11 to 3.15.

### Speed

Write step only, best of two runs, same DataFrame for both writers. The rustpy-xlsxwriter call used `write_worksheet(df, path, freeze_row=1, autofilter=True, header_format=..., column_width=12, inf_value="inf")`, which runs in constant-memory mode by default.

| File | xlsxwriter | rustpy-xlsxwriter | With `compression_level=1` |
|---|---|---|---|
| 1.6 MB | 1.10 s | 0.17 s (6.6x) | 0.08 s (13.3x), file 2.2 MB instead of 1.4 MB |
| 15 MB | 5.68 s | 1.06 s (5.4x) | 0.67 s (8.5x), file 15.1 MB instead of 9.9 MB |
| 48 MB | 18.90 s | 2.78 s (6.8x) | 1.78 s (10.6x), file 34.2 MB instead of 21.1 MB |

With default compression the files have the same size as today. Whether a larger temp file is worth the faster write (it also has to be scanned and opened by Excel) was not measured.

### Safety review

Done statically on 0.9.1 before running it:

- The Python files in the PyPI wheel are identical to the GitHub source apart from line endings, and the wheel hash matches PyPI.
- The Rust and Python sources contain no network or process code. File access is limited to writing the output. The compiled module imports only Windows runtime DLLs and `python312.dll`, and the only URLs in it are XML namespaces.
- Dependencies: `rust_xlsxwriter`, `pyo3`, `arrow-array`, `arrow-schema`, `indexmap`, `ryu`, `itoa`. MIT licence. The wheel ships an SBOM.
- Not verified: that the compiled module was built from exactly that source.

Maintenance risk:

- 33 GitHub stars, two maintainers, first release in January 2025, pre-1.0 with 26 releases.
- Its own error messages say that earlier versions wrote unsupported columns and out-of-range dates as empty cells without an error. Behaviour has changed between releases and may change again.
- The PyPI download numbers (about 4,500 in April 2026, 205,000 in October 2026) do not fit 33 stars and 10 external issues, so they say little about real adoption.
- Releases are uploaded with a long-lived PyPI token, not trusted publishing.

Consequence: pin an exact version, review the changelog and rerun the conversion tests before every bump, and never rely on its defaults for a conversion decision.

## Behaviour compared with our decisions

Tested by writing the same DataFrame with both writers and comparing the cells in the sheet XML (type, value, number format).

Identical without any preparation:

- Strings stay text: `=1+1`, `@SUM(A1)`, `+5`, `007`, `1e5`, `2310009E13`, URLs, `mailto:` addresses, padded strings, tabs, line breaks, control characters, `_x000D_`, Unicode, `TRUE`, and a string of exactly 32,767 characters.
- Booleans, categoricals, null columns.
- Null and `NaN` give no cell.
- `inf` and `-inf` as text, with `inf_value="inf"`.
- Integers, including the loss of precision above 2^53, and decimals as numbers.
- On the three real files, every text cell was identical between the two writers.

Different:

| Case | Our decision (xlsxwriter today) | rustpy-xlsxwriter 0.9.1 | What to do |
|---|---|---|---|
| Text over 32,767 characters | Truncated with a `ConversionWarning` | Raises `RuntimeError` | Truncate in polars and warn before writing. Test how both count characters outside the Basic Multilingual Plane (e.g. emoji) at the limit. |
| Dates and datetimes before 1900-03-01 | ISO 8601 text, mixed with real dates in the same column | Raises `ValueError` for values before 1900-01-01; a column has one type | Decision change needed, see below. Our cutoff of 1900-03-01 must be applied by us; the wrapper accepts January and February 1900. |
| Time | Text `10:30:05` | Excel time number with format `hh:mm:ss` | Convert to text first. |
| Datetime with time zone | ISO 8601 text with offset | Wall-clock number, offset dropped | Nothing new: `_prepare_columns` already converts these to text. |
| Duration, list, struct, binary | Text (`1 day, 0:00:05`, `[1, 2]`, `{'a': 1}`, hexadecimal) | Raises `ValueError` | Convert to text first. The text must match today's output, which polars' own casts will not all produce, so these rare types may need the slow route through Python (`str`). |
| Float32 | `0.1` stays `0.1` | `0.10000000149011612` | Nothing new: `_prepare_columns` already converts through decimal text. |
| Empty string | Empty text cell | No cell | Decision needed, see below. |
| Float precision | 16 significant digits; a value can differ in its last binary digit | Shortest exact representation; every number bit-identical to the DataFrame | Improvement. Update the "Floats" row in `docs/design.md`. On the 15 MB file 159,140 of 636,009 numbers, and on the 48 MB file 134,787 numbers, differ in the last digit today (relative error up to 5.7e-16). |
| Default datetime format | `yyyy-mm-dd hh:mm:ss`, with `.000` for fractional seconds | `yyyy-mm-ddThh:mm:ss` | Keep ours and set it explicitly. Ours shows fractional seconds, and the space form is easier to read. Excel probably re-reads the space form as a datetime but not the `T` form (not tested). |

Integers and whole floats are written as `0.0` instead of `0` in the XML. That is the same number for Excel.

### Decisions to make

- **Dates before 1900-03-01.** A column can no longer mix text and date cells. Proposal: if any value in a date or datetime column is before 1900-03-01, write the whole column as ISO 8601 text with a `ConversionWarning`. Such dates should be very rare in our data.
- **Empty strings.** "No cell" means the position has no entry in the sheet and Excel shows an ordinary blank cell, exactly as for null today. It is fully valid. The differences: `ISBLANK` is true and `COUNTA` does not count it, and "empty text" can no longer be told apart from "missing". Proposal: accept blank cells. This probably only affects Parquet input, as polars seems to read empty fields in text files as null (not checked).
- **Compression level.** Default (same file size as today) or level 1 (about 1.5x faster write, files about 1.5x larger).

## Open points to verify before implementing

These were not tested. Each can turn out to be a blocker.

- **Progress and cancelling.** Today the writer reports progress and checks for a cancel about every 10,000 cells (see "Progress and cancelling" in `docs/design.md`). `write_worksheet` is one blocking call for the whole sheet, so there is no percentage and no cancel while it runs. Writing would be short (an estimated 7 s instead of 44 s for the 132 MB file mentioned there), but not instant. To clarify: whether an indeterminate "writing" stage is acceptable, how a cancel takes effect (waiting for the call to end and deleting the file, or writing in a child process that can be ended), and what a cancel leaves behind.
- **GIL.** The progress window runs in the Tk main thread while the conversion runs in a worker thread. If the Rust call holds the GIL for the whole write, the window freezes. The README says the save phase releases the GIL; the row phase for DataFrames was not checked.
- **Temp files.** In constant-memory mode rust_xlsxwriter writes temporary row files. Check where they go, whether the folder can be set to our work folder (as `tmpdir` today), and that nothing is left behind after an error or a kill.
- **Header and layout.** Check on the written file: bold, centred, wrapped header; frozen first row; plain autofilter over the right range without an Excel Table; `--plain` output with none of these.
- **Column widths.** We set widths in pixels (40 to 125) from `xlsxwriter.utility.cell_autofit_width`. Check which unit `column_widths` takes and whether the result matches. If xlsxwriter stays only for that helper, replace or copy it so the dependency can go.
- **Number formats per column.** `yyyy-mm-dd` for dates, and the datetime format with or without `.000` per column, through `column_formats` or `datetime_format`.
- **Column names.** Always text, empty names stay empty, names that look like numbers or formulas, and whatever polars allows that a DataFrame-level API may reject.
- **Excel limits.** We check rows and columns before writing; confirm the wrapper is never reached with too large a table.
- **Errors.** Map its `RuntimeError` and `ValueError` to `WriteError` with our messages, including a folder that is not writable and a full disk, and confirm no partial file stays.
- **Memory.** Confirm that memory use stays flat for large tables, as with xlsxwriter's `constant_memory` mode.
- **Platforms.** Install and run the test suite on the CI matrix (Linux, Windows, macOS) and with the oldest supported Python (3.12).
- **Start-up time.** Measure the import time of the module; it adds to every "Open in Excel".

## Implementation outline

1. Work through the open points above with small experiments. Stop if progress and cancelling or the temp files cannot be solved cleanly, and reconsider the own binding.
2. Take the decisions listed under "Decisions to make" and update `docs/design.md`.
3. Extend `_prepare_columns` so that every column reaches the writer in a type the wrapper writes the way we want: truncation with warning, date columns with early dates, time, duration, list, struct and binary as text.
4. Replace `_SheetWriter` with one call to the wrapper, with every relevant option set explicitly (formats, widths, freeze, autofilter, `inf_value`), so a changed default in a later release cannot change our output.
5. Adapt progress reporting and cancelling in `writers.py` and `pipeline.py`, and the description in `docs/design.md`.
6. Keep the existing writer tests, which check the written xlsx file, as the acceptance test. They must pass unchanged except for the decisions that were changed on purpose. Add tests for the new preparation steps.
7. Add the dependency with an exact pin, remove xlsxwriter if nothing uses it any more, and update the "Excel writer" and "Dependency bounds" rows in `docs/design.md`.
8. Repeat the timing on real files of 1.5 MB to 280 MB, including the full "Open in Excel" path.
