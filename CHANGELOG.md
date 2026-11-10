# Changelog

## [0.4.0] - 2026-11-10

### Added
- `xlsxpress context-menu register`, `unregister` and `status` commands, which add, remove and check the File Explorer context menu entries on Windows.
- `xlsxpress --version` option.
- Warnings for values or names that XlsxPress changed or may have read incorrectly: text longer than Excel's limit of 32,767 characters, repeated column names, files that are not valid UTF-8, and files with irregular quotes. The command line, the context menu and the conversion dialog show these warnings.
- Support for text files in UTF-16 and UTF-32 with a byte order mark, and in Windows-1252 or Latin-1.
- Support for Python 3.12 and 3.13.
- A progress window for conversions that take longer than a moment, from **Open in Excel (XlsxPress)** and from the conversion dialog. It shows the elapsed time and how much of the workbook is written. Close the window to cancel the conversion.
- The conversion dialog shows where XlsxPress saved the workbook, with a **Show in folder** button and the warnings of the conversion.

### Changed
- Install with `uv tool install xlsxpress` or `pip install xlsxpress`. The desktop parts are included, so the `gui` extra is no longer necessary.
- A column of a text file becomes a number column only if every value in it is a plain number. Otherwise the whole column stays text, so leading zeros (`007`), signs (`+5`), decimal commas, dates and IDs such as `2310009E13` stay exactly as in the file.
- Missing values such as `NA`, `NaN` and `#N/A` become empty cells in number columns.
- The separator of `.csv` and `.txt` files is detected from the first lines and prefers a separator that occurs equally often on every line. Separators inside quoted fields, such as the commas in `id;"protein,description"`, are ignored.
- Dates and datetimes from Parquet files become Excel dates. Datetimes with a time zone, times, durations, lists and binary values become text.
- XlsxPress writes the workbook row by row, so large files need less memory.
- Column widths are estimated from the header and the first 1,000 rows, between 40 and 125 pixels. Date columns are always wide enough for the full date.
- XlsxPress no longer writes the data as an Excel table. Instead, the formatted header row has plain filter buttons, and `--plain` turns off all header formatting, including the filter buttons.
- XlsxPress never overwrites a file and adds a number to the name instead (`data-02.xlsx`). It writes each workbook to a work folder first, so a file under the output name is always complete.
- The command line, the context menu and the conversion dialog share one temp folder, `%LOCALAPPDATA%\xlsxpress\Cache\temp` on Windows. Names in the temp folder contain the date and time, and XlsxPress deletes workbooks there after 14 days or when there are more than 100.
- **Open in Excel (XlsxPress)** always opens the workbook, whatever the default of **Open in Excel after conversion** is.
- The conversion dialog and its About window have a new, card-based design that looks the same in light and dark Windows mode. The switches became checkboxes, **Save to temp directory** became a choice between **Temp folder** and **Next to original**, **Preview only (limit rows)** and the **Rows** field became **Convert only the first [N] rows**, and **Update default options** became **Save as default**. The row limit field is active only when its checkbox is checked.
- The conversion dialog shows the result, the warnings or the error of a conversion in a banner below the options.
- The buttons and checkboxes of the conversion dialog can be used with the keyboard: **Tab** moves between them, **Space** or **Enter** activates them, **Enter** elsewhere converts, and **Esc** closes the dialog.
- The command line shows errors as short messages without a traceback and returns exit code 1 when a conversion fails.

### Removed
- The standalone Windows app (`xlsxpress.exe`), with its launcher window, installer and self-update. Use `uv tool upgrade xlsxpress` to update and `uv tool uninstall xlsxpress` to uninstall.
- The `xlsxpress-app` command and the `gui` extra.

### Fixed
- **Open in Excel (XlsxPress)** converted only the preview rows, even when preview was off.
- Floating-point numbers were rounded to three decimal places.
- Text that starts with `=` became a formula.
- `inf` and `-inf` values stopped the conversion.
- Datetime columns with a time zone stopped the conversion on Windows.
- UTF-16 files gave unreadable text without an error, and Latin-1 files could not be read.
- A text file whose values changed type after row 10,000 could not be converted.
- A workbook in the temp folder could have the same name as a workbook that was open in Excel.
- Temp files of the command line were never deleted.
- Cleanup of the temp folder could fail when several conversions ran at the same time.
- A workbook that was written but could not be opened was reported as a failed conversion.
- Read errors showed a traceback on the command line and a general error message in the context menu.

### Dependencies
- `cyclopts>=5.0` replaces `typer` for the command line.
- `customtkinter>=5.2.1` is now a required dependency instead of part of the `gui` extra. 5.2.0 and older don't work on Python 3.12 and newer.
- `polars>=1.18` instead of `>=1.0`.
- `xlsxwriter>=3.2.1` instead of `>=3.2`.
- `platformdirs>=2.1` instead of `>=4.0`.
