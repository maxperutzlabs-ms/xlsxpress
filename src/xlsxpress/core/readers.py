"""Reading tabular files into polars DataFrames."""

from __future__ import annotations

from pathlib import Path

import polars as pl

#: Extensions read as character-separated text, with their default separator.
#: For ".txt" the separator is sniffed from the file content.
_SEPARATED_TEXT_EXTENSIONS = {".csv": ",", ".tsv": "\t", ".txt": None}

SUPPORTED_EXTENSIONS = frozenset(_SEPARATED_TEXT_EXTENSIONS) | {".parquet"}


class UnsupportedFormatError(ValueError):
    """Raised when a file extension is not supported."""


def read_table(path: Path, nrows: int | None = None) -> pl.DataFrame:
    """Read a tabular file into a DataFrame.

    Args:
        path: Source file; format is inferred from the extension.
        nrows: If set, read at most this many data rows.

    Raises:
        UnsupportedFormatError: If the file extension is not supported.
    """
    extension = path.suffix.lower()
    if extension == ".parquet":
        return pl.read_parquet(path, n_rows=nrows)
    if extension in _SEPARATED_TEXT_EXTENSIONS:
        separator = _SEPARATED_TEXT_EXTENSIONS[extension] or _sniff_separator(path)
        return pl.read_csv(
            path,
            separator=separator,
            n_rows=nrows,
            infer_schema_length=10_000,
            try_parse_dates=False,
        )
    raise UnsupportedFormatError(
        f"Unsupported file extension {extension!r}. "
        f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
    )


def _sniff_separator(path: Path) -> str:
    """Guess the separator of a text file from its first line.

    Picks the most frequent candidate separator; defaults to tab, which is
    the most common case in our domain.
    """
    with path.open("r", encoding="utf-8", errors="replace") as file:
        first_line = file.readline()
    candidates = {
        "\t": first_line.count("\t"),
        ",": first_line.count(","),
        ";": first_line.count(";"),
    }
    best = max(candidates, key=candidates.get)  # type: ignore[arg-type]
    return best if candidates[best] > 0 else "\t"
