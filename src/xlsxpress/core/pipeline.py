"""The main conversion pipeline: read -> write xlsx -> optionally open."""

from __future__ import annotations

from pathlib import Path

from xlsxpress.core import excel, paths, readers, writers
from xlsxpress.core.options import ConversionOptions


def run_conversion(options: ConversionOptions) -> Path:
    """Execute a full conversion according to `options`.

    Returns:
        Path of the written xlsx file.

    Raises:
        FileNotFoundError: If the input file does not exist.
        readers.UnsupportedFormatError: If the input format is unsupported.
    """
    input_path = options.input_path.resolve()
    if not input_path.is_file():
        raise FileNotFoundError(f"Input file not found: {input_path}")

    output_dir = paths.get_temp_dir() if options.to_temp else input_path.parent
    output_path = paths.resolve_output_path(input_path, output_dir)

    df = readers.read_table(input_path, nrows=options.nrows)
    writers.write_excel(df, output_path, format_header=options.format_header)

    if options.open_after:
        excel.open_in_default_app(output_path)
    return output_path
