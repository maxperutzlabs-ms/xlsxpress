"""Command line interface — thin routing layer, no domain logic."""

from __future__ import annotations

from pathlib import Path
from typing import Annotated, Optional

import typer

from xlsxpress.core.options import ConversionOptions
from xlsxpress.core.pipeline import run_conversion
from xlsxpress.core.readers import UnsupportedFormatError

app = typer.Typer(
    name="xlsxpress",
    help="Open tabular files as safely formatted Excel files.",
    no_args_is_help=True,
)

FileArgument = Annotated[Path, typer.Argument(help="Tabular input file.", exists=True, dir_okay=False)]
NrowsOption = Annotated[
    Optional[int], typer.Option("--nrows", "-n", min=1, help="Only include the first N rows (preview).")
]
PlainOption = Annotated[bool, typer.Option("--plain", help="Skip header formatting.")]


@app.command()
def open(filename: FileArgument, nrows: NrowsOption = None, plain: PlainOption = False) -> None:
    """Convert to a temporary xlsx file and open it in Excel."""
    _run(ConversionOptions(filename, nrows=nrows, to_temp=True, open_after=True, format_header=not plain))


@app.command()
def convert(
    filename: FileArgument,
    nrows: NrowsOption = None,
    plain: PlainOption = False,
    temp: Annotated[bool, typer.Option("--temp", help="Write to the temp directory instead of next to the input.")] = False,
    open_after: Annotated[bool, typer.Option("--open", help="Open the result in Excel after converting.")] = False,
) -> None:
    """Convert to an xlsx file next to the original (default) or in temp."""
    _run(ConversionOptions(filename, nrows=nrows, to_temp=temp, open_after=open_after, format_header=not plain))


@app.command()
def dialog(filename: FileArgument) -> None:
    """Open the conversion options dialog for a file (phase two)."""
    typer.echo(f"[xlsxpress] options dialog for: {filename} (not implemented yet)")


def _run(options: ConversionOptions) -> None:
    """Run the pipeline, reporting errors as exit codes instead of tracebacks."""
    try:
        output_path = run_conversion(options)
    except (FileNotFoundError, UnsupportedFormatError) as error:
        typer.secho(f"Error: {error}", fg=typer.colors.RED, err=True)
        raise typer.Exit(code=1)
    typer.echo(f"Written '.xlsx' file to: {output_path}")


if __name__ == "__main__":
    app()
