"""Command line interface: a thin routing layer without conversion logic."""

from __future__ import annotations

import sys
from collections.abc import Sequence
from pathlib import Path
from typing import Annotated

from cyclopts import App, Parameter, validators

from xlsxpress import __VERSION__
from xlsxpress.app.config import load_config
from xlsxpress.core.errors import ConversionError
from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import run_conversion

app = App(
    name="xlsxpress",
    help="Open tabular files (csv, tsv, txt, parquet) as safely formatted Excel files.",
    version=__VERSION__,
)

_Flag = Parameter(negative_bool=(), show_default=False)
_Nrows = Annotated[
    int | None, Parameter(name=["--nrows", "-n"], validator=validators.Number(gte=1))
]


def main(argv: Sequence[str] | None = None) -> int:
    """Run the command line application.

    Args:
        argv: Command line arguments. If None (default), uses the process arguments.

    Returns:
        The exit code: 0 on success, 1 if the conversion failed, and cyclopts' exit
        code for invalid arguments.
    """
    result = app(argv, result_action="return_value")
    return result if isinstance(result, int) else 0


@app.command(name="open")
def open_file(
    file: Path,
    /,
    *,
    nrows: _Nrows = None,
    plain: Annotated[bool, _Flag] = False,
) -> int:
    """Convert to a temporary xlsx file and open it in Excel.

    The xlsx file is written to the XlsxPress temp folder, which is cleaned up
    automatically. Save the workbook somewhere else to keep your changes.

    Args:
        file: Tabular input file (csv, tsv, txt or parquet).
        nrows: Only include the first N data rows (preview).
        plain: Skip header formatting (bold header, frozen first row, filter).
    """
    options = ConversionOptions(
        file, nrows=nrows, to_temp=True, open_after=True, format_header=not plain
    )
    return _run(options)


@app.command
def convert(
    file: Path,
    /,
    *,
    nrows: _Nrows = None,
    plain: Annotated[bool, _Flag] = False,
    temp: Annotated[bool, _Flag] = False,
    open_after: Annotated[
        bool, Parameter(name="--open", negative_bool=(), show_default=False)
    ] = False,
) -> int:
    """Convert to an xlsx file next to the input file.

    An existing xlsx file is never overwritten; a numbered name is used instead.

    Args:
        file: Tabular input file (csv, tsv, txt or parquet).
        nrows: Only include the first N data rows (preview).
        plain: Skip header formatting (bold header, frozen first row, filter).
        temp: Write to the XlsxPress temp folder instead, which is cleaned up
            automatically.
        open_after: Open the xlsx file in Excel after converting.
    """
    options = ConversionOptions(
        file,
        nrows=nrows,
        to_temp=temp,
        open_after=open_after,
        format_header=not plain,
    )
    return _run(options)


def _run(options: ConversionOptions) -> int:
    """Run the conversion and report the result, without tracebacks."""
    options.cleanup_policy = load_config().cleanup
    try:
        result = run_conversion(options)
    except ConversionError as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    for message in result.warnings:
        print(f"Warning: {message}", file=sys.stderr)
    print(f"Wrote {result.output_path}")
    return 0
