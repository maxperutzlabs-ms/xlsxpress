# XlsxPress

Open tabular data files directly as safe, readable Excel workbooks.

XlsxPress converts files such as `.csv`, `.tsv`, `.txt`, and `.parquet` into `.xlsx` files and optionally opens them in Excel immediately. It is designed for exploratory data inspection and reporting workflows where opening raw text files in Excel is inconvenient or risky.


## Table of Contents

- [Features](#features)
- [Installation and Usage](#installation-and-usage)
- [Command Line Usage](#command-line-usage)
- [Standalone Windows App](#standalone-windows-app)
- [Releases](#releases)
- [Supported formats](#supported-formats)
- [Feedback and Issues](#feedback-and-issues)


## Features

- Open **tabular files** (CSV, TSV, TXT) directly in Excel
- Open **Parquet** files directly in Excel
- Add XlsxPress actions to the **Windows file context menu** for opening and converting files directly from File Explorer (only available via the standalone Windows app)
- Convert into temporary `.xlsx` files to **avoid modifying the original** text files
- **Preserve text values** that Excel commonly auto-converts (for example gene names becoming dates)
- Optional **preview mode** to only load the first *N* rows of very large datasets
- Automatic **header formatting** for improved readability
- Safe output naming (never overwrites existing Excel files)

### Automatic header formatting

By default, XlsxPress makes the first row easier to work with in Excel:

- header cells are emphasized visually
- text wrapping is enabled for long column names
- the header row remains visible while scrolling (freeze panes)
- column widths are adjusted automatically to fit content (within reasonable limits)

This behavior can be disabled with `--plain` in the command line interface, or by disabling `format header row` in the standalone app.


## Installation and Usage

### Installation

Install from PyPI with pip

```bash
pip install xlsxpress
```

or with [uv](https://docs.astral.sh/uv/)

```bash
uv add xlsxpress
```

### Run as a command line tool

```bash
xlsxpress open data.tsv
```

or

```bash
python -m xlsxpress open data.tsv
```

### Run directly with `uv` without installing it first

```bash
uvx run xlsxpress open data.tsv
```


## Command Line Usage

### Available commands

- `open` creates a temporary `.xlsx` file and opens it in Excel
    - `--nrows` - preview only the first *N* rows of a large dataset
    - `--plain` - disable automatic header formatting
- `convert` creates an `.xlsx` file next to the original file
    - `--nrows` - preview only the first *N* rows of a large dataset
    - `--plain` - disable automatic header formatting
    - `--temp` - write to a temporary file instead of next to the original
    - `--open` - open the converted file in Excel immediately

### Examples

Open in Excel using a temporary file:

```bash
xlsxpress open input.tsv
```

Convert next to the original file:

```bash
xlsxpress convert input.tsv
```

Preview only the first rows:

```bash
xlsxpress open input.tsv --nrows 100
```

Convert and immediately open:

```bash
xlsxpress convert input.tsv --open
```


## Standalone Windows App

The Windows app extends XlsxPress with **Windows File Context Menu integration**, allowing supported files to be opened or converted directly from File Explorer without using the command line.

Installation:

- Download the latest `.exe` from GitHub Releases
- Run the application to open the launcher
- Select `Register context menu` to add XlsxPress actions to the Windows file context menu

After installation you can:

- Right click supported files (`.csv`, `.tsv`, `.txt`, `.parquet`) to:
    - `Open in Excel (xlsxpress)` - convert and open in Excel using your saved default options
    - `Open conversion dialog (xlsxpress)` - open a GUI and select conversion options before opening
- Run the application to
    - Uninstall the XlsxPress app
    - Register or remove the context menu actions
    - Update XlsxPress to the latest version
    - Update the default conversion settings used by context menu actions

The conversion dialog includes an **Update defaults** action. Saved defaults are automatically reused by `Open in Excel (xlsxpress)` so frequently used options only need to be configured once.

On Windows 11, XlsxPress adds entries to the classic context menu. Right-click a file and select **Show more options** to access the XlsxPress actions.

> [!TIP]
> Hold `Shift + Right Click` on a file to open the classic context menu directly.

The Windows app installs per-user into local application data and does not require administrator privileges.


## Releases

- Python package: published on [PyPI](https://pypi.org/project/xlsxpress) as `xlsxpress`
- Standalone application: automatically built with [PyApp](https://ofek.dev/pyapp/latest/) and published as GitHub release artifacts through the CI/CD pipeline


## Supported formats

Current supported inputs:

- `.csv`
- `.tsv`
- `.txt`
- `.parquet`

Output:

- `.xlsx`


## Feedback and Issues

Bug reports, feature requests, and improvement suggestions are welcome.

Please use the GitHub issue tracker to report problems, propose new features, or discuss potential enhancements.