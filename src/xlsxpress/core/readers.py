"""Reading tabular files into polars DataFrames without changing their values.

Text files are read with every column as text first. A column only becomes numeric
when every value in it is an unambiguous number, so identifiers such as `007` or
`2310009E13` and gene names stay exactly as they are in the file.
"""

from __future__ import annotations

import codecs
import csv
import warnings
from pathlib import Path

import polars as pl

from xlsxpress.core.errors import ConversionWarning, ReadError, UnsupportedFormatError

#: Extensions read as character-separated text, with their default separator.
#: For ".csv" and ".txt" the separator is sniffed from the first lines.
_SEPARATED_TEXT_EXTENSIONS = {".csv": None, ".tsv": "\t", ".txt": None}

SUPPORTED_EXTENSIONS = frozenset(_SEPARATED_TEXT_EXTENSIONS) | {".parquet"}

#: Values that mean "missing" in an otherwise numeric column; they become blank.
MISSING_VALUE_TOKENS = ("NA", "NaN", "nan", "N/A", "#N/A", "null", "NULL")

_INTEGER = r"-?(?:0|[1-9][0-9]*)"
_DECIMAL = r"-?(?:0|[1-9][0-9]*)?\.[0-9]+(?:[eE][+-]?[0-9]+)?"
#: An integer mantissa needs a signed exponent: `1e+05` is a number, while
#: `2310009E13` (a mouse gene ID) and `1E5` stay text.
_INTEGER_WITH_EXPONENT = r"-?(?:0|[1-9][0-9]*)[eE][+-][0-9]+"
_INTEGER_PATTERN = f"^{_INTEGER}$"
_FLOAT_PATTERN = f"^(?:{_INTEGER}|{_DECIMAL}|{_INTEGER_WITH_EXPONENT})$"
_INFINITY_PATTERN = r"(?i)^-?(?:inf|infinity)$"

_SEPARATOR_NAMES = {"\t": "tab", ",": "comma", ";": "semicolon"}
_ENCODING_NAMES = {"cp1252": "Windows-1252 (Western European)", "latin-1": "Latin-1"}
_SNIFF_BYTES = 64 * 1024
_SNIFF_LINES = 20
_VALIDATION_CHUNK_BYTES = 1024 * 1024


def read_table(path: Path, nrows: int | None = None) -> pl.DataFrame:
    """Read a tabular file into a DataFrame, keeping every value as in the file.

    Parquet files keep their own data types. Text files (`.csv`, `.tsv`, `.txt`)
    are read with every column as text, and a column is then converted:

    - To integers if every value is a plain integer such as `-12` or `0`. Integers
      too large for 64 bits become floats.
    - To floats if every value is a plain integer or decimal number such as `1.5`,
      `.5`, `1.5e-3` or `1e+05`, or infinity (`inf`, `-Infinity`, any case).
    - Otherwise the column stays text. This keeps leading zeros (`007`), signs
      (`+5`), spaces, decimal commas (`1,5`), dates, booleans, gene names and
      identifiers like `2310009E13` exactly as they are.

    In a numeric column, the missing-value tokens in `MISSING_VALUE_TOKENS` (such as
    `NA` and `NaN`) become null. A column of only such tokens stays text.

    Text files that start with a byte order mark are decoded as UTF-8, UTF-16 or
    UTF-32, as the mark says, and other text files as UTF-8. A file that is not
    valid UTF-8 is read as Windows-1252, or as Latin-1 if that fails, with a
    warning.

    The separator is tab for `.tsv` files. For `.csv` and `.txt` files it is
    guessed from the first lines. If standard `"` quoting fails, the file is read
    again without quote handling, with a warning.

    Args:
        path: Source file. The format is inferred from the extension.
        nrows: If set, read at most this many data rows. Defaults to None.

    Returns:
        The table.

    Raises:
        UnsupportedFormatError: If the file extension is not supported.
        ReadError: If the file does not exist, cannot be opened, is empty, contains
            NUL bytes without a byte order mark, or cannot be read as a table.

    Warns:
        ConversionWarning: If a text file is not valid UTF-8, was read without
            quote handling, or has repeated column names that were renamed.
    """
    extension = path.suffix.lower()
    if extension not in SUPPORTED_EXTENSIONS:
        raise UnsupportedFormatError(
            f"Unsupported file extension {extension!r}. "
            f"Supported: {', '.join(sorted(SUPPORTED_EXTENSIONS))}"
        )
    if not path.is_file():
        raise ReadError(
            f"Input file '{path}' does not exist or is not a file. Check the path."
        )
    try:
        if extension == ".parquet":
            return _read_parquet(path, nrows)
        return _read_separated_text(path, nrows, _SEPARATED_TEXT_EXTENSIONS[extension])
    except OSError as error:
        raise ReadError(
            f"Could not open {path.name!r}: {error.strerror or error}. Check that "
            "you have permission to read the file and that no other program locks it."
        ) from error


def _read_parquet(path: Path, nrows: int | None) -> pl.DataFrame:
    try:
        return pl.read_parquet(path, n_rows=nrows)
    except pl.exceptions.PolarsError as error:
        raise ReadError(
            f"Could not read {path.name!r} as a Parquet file: {_first_line(error)}"
        ) from error


def _read_separated_text(
    path: Path, nrows: int | None, separator: str | None
) -> pl.DataFrame:
    encoding = _detect_encoding(path)
    separator = separator or _sniff_separator(path, encoding)
    quote_char: str | None = '"'
    try:
        if _header_has_irregular_quotes(path, separator, encoding):
            # Polars would silently strip quotes from such a header, so skip to
            # the fallback below.
            raise pl.exceptions.ComputeError("irregular quotes in the header")
        df = _read_csv(path, separator, encoding, quote_char, nrows)
    except pl.exceptions.NoDataError as error:
        raise ReadError(f"{path.name!r} is empty.") from error
    except pl.exceptions.PolarsError as error:
        # A field that starts with a quote but isn't fully quoted, e.g.
        # `"5 inch" cable`, breaks quote parsing. Reading without quote handling
        # keeps such values exactly as in the file.
        quote_char = None
        try:
            df = _read_csv(path, separator, encoding, quote_char, nrows)
        except pl.exceptions.PolarsError:
            raise ReadError(_describe_text_error(path, separator, error)) from error
        # Polars 1.x reads a quote inside a field (`5" cable`) with quote handling,
        # polars 2.x does not. Warn only if a field starts with a quote, so the
        # behaviour is the same on both: that is where quotes are really kept.
        if _has_leading_quote(df):
            warnings.warn(
                f"{path.name!r} has quotes that don't follow the usual quoting rules, "
                "so it was read without quote handling. Quote characters are kept as "
                "part of the values.",
                ConversionWarning,
                stacklevel=3,
            )
    _warn_about_repeated_column_names(path, df, separator, encoding, quote_char)
    conversions = [_numeric_conversion(df[name]) for name in df.columns]
    return df.with_columns(
        conversion for conversion in conversions if conversion is not None
    )


def _header_has_irregular_quotes(path: Path, separator: str, encoding: str) -> bool:
    """Whether the header row has a quoted field that isn't fully quoted.

    Polars silently alters such a header, e.g. `"5 inch" cable` becomes
    `5 inch" cabl`. Only the first record is parsed, so this is cheap.
    """
    codec = "utf-8-sig" if encoding == "utf-8" else encoding
    with path.open(newline="", encoding=codec) as file:
        try:
            next(csv.reader(file, delimiter=separator, strict=True), None)
        except csv.Error:
            return True
    return False


def _has_leading_quote(df: pl.DataFrame) -> bool:
    """Whether a column name or text value starts with a double quote."""
    if any(name.startswith('"') for name in df.columns):
        return True
    return any(df[name].str.starts_with('"').any() for name in df.columns)


def _read_csv(
    path: Path,
    separator: str,
    encoding: str,
    quote_char: str | None,
    nrows: int | None,
    has_header: bool = True,
) -> pl.DataFrame:
    """Read all columns as text; polars decodes non-UTF-8 encodings with Python."""
    return pl.read_csv(
        path,
        separator=separator,
        quote_char=quote_char,
        has_header=has_header,
        n_rows=nrows,
        infer_schema=False,
        encoding="utf8" if encoding == "utf-8" else encoding,
    )


def _numeric_conversion(column: pl.Series) -> pl.Expr | None:
    """Return the expression converting a text column to numbers, if that is safe."""
    missing = column.is_null() | column.is_in(MISSING_VALUE_TOKENS)
    infinite = column.str.contains(_INFINITY_PATTERN).fill_null(False)
    values = column.filter(~missing & ~infinite)
    if values.is_empty():
        return None
    if not infinite.any() and values.str.contains(_INTEGER_PATTERN).all():
        fits_int64 = values.cast(pl.Int64, strict=False).null_count() == 0
        dtype = pl.Int64 if fits_int64 else pl.Float64
    elif values.str.contains(_FLOAT_PATTERN).all():
        dtype = pl.Float64
    else:
        return None
    text = pl.col(column.name)
    return (
        pl.when(text.is_in(MISSING_VALUE_TOKENS))
        .then(None)
        .otherwise(text.str.to_lowercase())
        .cast(dtype)
        .alias(column.name)
    )


def _detect_encoding(path: Path) -> str:
    with path.open("rb") as file:
        head = file.read(_SNIFF_BYTES)
    if head.startswith(codecs.BOM_UTF8):
        return "utf-8"
    if head.startswith((codecs.BOM_UTF32_LE, codecs.BOM_UTF32_BE)):
        return "utf-32"
    if head.startswith((codecs.BOM_UTF16_LE, codecs.BOM_UTF16_BE)):
        return "utf-16"
    if b"\x00" in head:
        raise ReadError(
            f"{path.name!r} contains NUL bytes, so it is not a UTF-8 or Windows-1252 "
            "text file. It may be UTF-16 text without a byte order mark, or not a "
            "text file at all. Save it as UTF-8 text and try again."
        )
    for encoding in ("utf-8", "cp1252"):
        if _is_valid_text(path, encoding):
            break
    else:
        encoding = "latin-1"
    if encoding != "utf-8":
        warnings.warn(
            f"{path.name!r} is not valid UTF-8 text and was read as "
            f"{_ENCODING_NAMES[encoding]}. Check that special characters such as "
            "umlauts look right.",
            ConversionWarning,
            stacklevel=2,
        )
    return encoding


def _is_valid_text(path: Path, encoding: str) -> bool:
    """Return True if the whole file decodes with `encoding`, reading in chunks."""
    decoder = codecs.getincrementaldecoder(encoding)()
    try:
        with path.open("rb") as file:
            while chunk := file.read(_VALIDATION_CHUNK_BYTES):
                decoder.decode(chunk)
        decoder.decode(b"", final=True)
    except UnicodeDecodeError:
        return False
    return True


def _sniff_separator(path: Path, encoding: str) -> str:
    """Guess the separator of a text file from its first lines.

    Prefers separators that occur equally often on the header line and the
    following lines, so a header name like `Protein names, short` in a
    tab-separated file doesn't win. Among those, or among all candidates if none is
    consistent, picks the most frequent on the header line. Defaults to tab, the
    most common case in our domain.
    """
    with path.open("rb") as file:
        head = file.read(_SNIFF_BYTES)
    lines = [
        line for line in head.decode(encoding, errors="replace").splitlines() if line
    ]
    if len(head) == _SNIFF_BYTES and len(lines) > 1:
        lines.pop()  # The last line may be cut off by the sample size.
    sample = lines[: _SNIFF_LINES + 1]
    if not sample:
        return "\t"
    header_counts = {
        separator: sample[0].count(separator) for separator in _SEPARATOR_NAMES
    }
    consistent = [
        separator
        for separator, count in header_counts.items()
        if count > 0 and all(line.count(separator) == count for line in sample[1:])
    ]
    candidates = consistent or [
        separator for separator, count in header_counts.items() if count > 0
    ]
    if not candidates:
        return "\t"
    return max(candidates, key=lambda separator: header_counts[separator])


def _warn_about_repeated_column_names(
    path: Path, df: pl.DataFrame, separator: str, encoding: str, quote_char: str | None
) -> None:
    """Warn if polars renamed repeated column names (`a` -> `a_duplicated_0`)."""
    if not any("_duplicated_" in name for name in df.columns):
        return
    header = _read_csv(path, separator, encoding, quote_char, 1, has_header=False).row(
        0
    )
    names = ["" if name is None else name for name in header]
    repeated = sorted({name for name in names if names.count(name) > 1})
    if repeated:
        warnings.warn(
            f"Column names appear more than once in {path.name!r}: "
            f"{', '.join(map(repr, repeated))}. Excel shows the repeats renamed, "
            "e.g. 'a' as 'a_duplicated_0'.",
            ConversionWarning,
            stacklevel=4,
        )


def _describe_text_error(
    path: Path, separator: str, error: pl.exceptions.PolarsError
) -> str:
    separator_name = _SEPARATOR_NAMES.get(separator, repr(separator))
    if "more fields" in str(error):
        return (
            f"Could not read {path.name!r}: a row has more {separator_name}-separated "
            "fields than the header row. Check that the file starts with a single "
            "header line, without comment or metadata lines above it."
        )
    return (
        f"Could not read {path.name!r} as {separator_name}-separated text: "
        f"{_first_line(error)}"
    )


def _first_line(error: Exception) -> str:
    return (
        str(error).strip().splitlines()[0]
        if str(error).strip()
        else type(error).__name__
    )
