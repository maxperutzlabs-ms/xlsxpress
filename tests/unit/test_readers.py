from __future__ import annotations

import codecs
import datetime as dt
import math
import warnings
from pathlib import Path

import polars as pl
import pytest

from xlsxpress.core.errors import ConversionWarning, ReadError, UnsupportedFormatError
from xlsxpress.core.readers import read_table

pytestmark = pytest.mark.unit


def _column_file(tmp_path: Path, values: list[str], name: str = "data.tsv") -> Path:
    path = tmp_path / name
    path.write_text("value\n" + "".join(f"{value}\n" for value in values), encoding="utf-8")
    return path


def _write_bytes(tmp_path: Path, content: bytes, name: str = "data.tsv") -> Path:
    path = tmp_path / name
    path.write_bytes(content)
    return path


@pytest.fixture
def no_warnings():
    with warnings.catch_warnings():
        warnings.simplefilter("error")
        yield


class TestReadTableTypeInference:
    @pytest.mark.parametrize(
        ("values", "expected"),
        [
            (["1", "2", "-3", "0"], [1, 2, -3, 0]),
            (["-0"], [0]),
            (["1234567890123456789"], [1_234_567_890_123_456_789]),
        ],
    )
    def test_read_table_reads_plain_integers_as_integers(
        self, tmp_path: Path, values: list[str], expected: list
    ) -> None:
        column = read_table(_column_file(tmp_path, values))["value"]

        assert column.dtype == pl.Int64
        assert column.to_list() == expected

    def test_read_table_reads_integers_beyond_int64_as_floats(self, tmp_path: Path) -> None:
        column = read_table(_column_file(tmp_path, ["12345678901234567890", "1"]))["value"]

        assert column.dtype == pl.Float64
        assert column.to_list() == [12345678901234567890.0, 1.0]

    @pytest.mark.parametrize(
        ("values", "expected"),
        [
            (["1.5", "2", "-0.25"], [1.5, 2.0, -0.25]),
            ([".5", "-.5"], [0.5, -0.5]),
            (["1.10", "2.00"], [1.1, 2.0]),
            (["1e+05", "2.5E-3", "1.5e3", "3.2e-10"], [1e5, 2.5e-3, 1.5e3, 3.2e-10]),
            (["0.30000000000000004"], [0.30000000000000004]),
        ],
    )
    def test_read_table_reads_plain_decimals_as_floats(self, tmp_path: Path, values: list[str], expected: list) -> None:
        column = read_table(_column_file(tmp_path, values))["value"]

        assert column.dtype == pl.Float64
        assert column.to_list() == expected

    @pytest.mark.parametrize(
        "values",
        [
            ["007", "12"],
            ["0.5", "00.5"],
            ["2310009E13", "4930481E13"],
            ["1E5"],
            ["+5", "6"],
            [" 5", "6"],
            ["5 ", "6"],
            ["1,5", "2,5"],
            ["1,234"],
            ["1."],
            ["0x1F"],
            ["1_000"],
            ["SEPT2", "MARCH1", "1-Mar", "DEC1"],
            ["TRUE", "FALSE"],
            ["2024-01-02", "2024-01-03"],
            ["5%", "10%"],
            ["P12345;Q67890", "P11111"],
            ["1", "2", "abc"],
            ["NA", "NaN"],
            ["-", "1"],
        ],
    )
    def test_read_table_keeps_columns_with_any_non_number_as_unchanged_text(
        self, tmp_path: Path, values: list[str]
    ) -> None:
        column = read_table(_column_file(tmp_path, values))["value"]

        assert column.dtype == pl.String
        assert column.to_list() == values

    def test_read_table_keeps_a_late_text_value_as_text_in_the_whole_column(self, tmp_path: Path) -> None:
        values = [str(i) for i in range(10_001)] + ["abc"]

        column = read_table(_column_file(tmp_path, values))["value"]

        assert column.dtype == pl.String
        assert column.to_list() == values

    @pytest.mark.parametrize("token", ["NA", "NaN", "nan", "N/A", "#N/A", "null", "NULL", ""])
    def test_read_table_reads_missing_value_tokens_in_numeric_columns_as_null(self, tmp_path: Path, token: str) -> None:
        column = read_table(_column_file(tmp_path, ["1.5", token, "2"]))["value"]

        assert column.dtype == pl.Float64
        assert column.to_list() == [1.5, None, 2.0]

    def test_read_table_reads_integer_columns_with_missing_values_as_integers(self, tmp_path: Path) -> None:
        column = read_table(_column_file(tmp_path, ["1", "NA", "3"]))["value"]

        assert column.dtype == pl.Int64
        assert column.to_list() == [1, None, 3]

    @pytest.mark.parametrize(
        ("token", "expected"),
        [("inf", math.inf), ("-inf", -math.inf), ("Inf", math.inf), ("-Infinity", -math.inf), ("INF", math.inf)],
    )
    def test_read_table_reads_infinity_tokens_in_numeric_columns_as_infinity(
        self, tmp_path: Path, token: str, expected: float
    ) -> None:
        column = read_table(_column_file(tmp_path, ["1", token]))["value"]

        assert column.dtype == pl.Float64
        assert column.to_list() == [1.0, expected]

    def test_read_table_keeps_columns_of_only_infinity_tokens_as_text(self, tmp_path: Path) -> None:
        column = read_table(_column_file(tmp_path, ["inf", "-Infinity"]))["value"]

        assert column.dtype == pl.String

    def test_read_table_keeps_empty_columns_as_text(self, tmp_path: Path) -> None:
        path = tmp_path / "data.tsv"
        path.write_text("a\tb\n1\t\n2\t\n", encoding="utf-8")

        df = read_table(path)

        assert df.schema == pl.Schema({"a": pl.Int64, "b": pl.String})
        assert df["b"].to_list() == [None, None]

    def test_read_table_infers_each_column_independently(self, tmp_path: Path) -> None:
        path = tmp_path / "data.tsv"
        path.write_text("id\tgene\tintensity\n007\tSEPT2\t1.5e+08\n012\tMARCH1\tNaN\n", encoding="utf-8")

        df = read_table(path)

        assert df.schema == pl.Schema({"id": pl.String, "gene": pl.String, "intensity": pl.Float64})
        assert df.rows() == [("007", "SEPT2", 1.5e8), ("012", "MARCH1", None)]


class TestReadTableEncoding:
    def test_read_table_reads_utf8_without_warning(self, tmp_path: Path, no_warnings: None) -> None:
        df = read_table(_write_bytes(tmp_path, "name\nMüller αβ 中文\n".encode()))

        assert df["name"].to_list() == ["Müller αβ 中文"]

    def test_read_table_removes_the_utf8_byte_order_mark(self, tmp_path: Path, no_warnings: None) -> None:
        df = read_table(_write_bytes(tmp_path, codecs.BOM_UTF8 + "name\nMüller\n".encode()))

        assert df.columns == ["name"]
        assert df["name"].to_list() == ["Müller"]

    @pytest.mark.parametrize("encoding", ["utf-16", "utf-16-be", "utf-32"])
    def test_read_table_reads_utf16_and_utf32_files_with_byte_order_mark(
        self, tmp_path: Path, encoding: str, no_warnings: None
    ) -> None:
        content = "name\tvalue\nMüller\t1\n".encode(encoding)
        if encoding == "utf-16-be":
            content = codecs.BOM_UTF16_BE + content

        df = read_table(_write_bytes(tmp_path, content))

        assert df.rows() == [("Müller", 1)]

    def test_read_table_reads_windows_1252_files_with_warning(self, tmp_path: Path) -> None:
        path = _write_bytes(tmp_path, "name\nMüller – 5 €\n".encode("cp1252"))

        with pytest.warns(ConversionWarning, match="Windows-1252"):
            df = read_table(path)

        assert df["name"].to_list() == ["Müller – 5 €"]

    def test_read_table_falls_back_to_latin1_for_bytes_undefined_in_windows_1252(self, tmp_path: Path) -> None:
        path = _write_bytes(tmp_path, b"name\nM\xfcller\x81\n")

        with pytest.warns(ConversionWarning, match="Latin-1"):
            df = read_table(path)

        assert df["name"].to_list() == ["Müller\x81"]

    def test_read_table_raises_read_error_for_truncated_utf16(self, tmp_path: Path) -> None:
        path = _write_bytes(tmp_path, codecs.BOM_UTF16_LE + "a\tb\n1\t2\n".encode("utf-16-le") + b"x")

        with pytest.raises(ReadError, match="not valid UTF-16-LE text"):
            read_table(path)

    def test_read_table_raises_for_nul_bytes_without_byte_order_mark(self, tmp_path: Path) -> None:
        path = _write_bytes(tmp_path, "name\nvalue\n".encode("utf-16-le"))

        with pytest.raises(ReadError, match="NUL bytes"):
            read_table(path)


class TestReadTableSeparator:
    @pytest.mark.parametrize(
        ("name", "content"),
        [
            ("data.tsv", "a\tb\n1\t2\n"),
            ("data.csv", "a,b\n1,2\n"),
            ("data.csv", "a;b\n1;2\n"),
            ("data.txt", "a\tb\n1\t2\n"),
            ("data.txt", "a,b\n1,2\n"),
            ("data.txt", "a;b\n1;2\n"),
        ],
    )
    def test_read_table_uses_the_separator_of_the_file(self, tmp_path: Path, name: str, content: str) -> None:
        path = tmp_path / name
        path.write_text(content, encoding="utf-8")

        assert read_table(path).rows() == [(1, 2)]

    @pytest.mark.parametrize(
        "content",
        [
            'id;"protein,description,info"\n1;"P1,albumin,serum"\n2;"P2,keratin,skin"\n',
            '"id;alias;note",value\n"a;b;c",1\n"d;e;f",2\n',
        ],
    )
    def test_read_table_ignores_separators_inside_quoted_fields_of_the_header(
        self, tmp_path: Path, content: str, no_warnings: None
    ) -> None:
        path = tmp_path / "data.csv"
        path.write_text(content, encoding="utf-8")

        assert read_table(path).width == 2

    def test_read_table_prefers_the_consistent_separator_over_commas_in_the_header(self, tmp_path: Path) -> None:
        path = tmp_path / "data.txt"
        path.write_text("Protein names, short, long\tGene\nAlbumin\tALB\nKeratin\tKRT1\n", encoding="utf-8")

        df = read_table(path)

        assert df.columns == ["Protein names, short, long", "Gene"]

    def test_read_table_reads_single_column_text_files(self, tmp_path: Path) -> None:
        path = tmp_path / "data.txt"
        path.write_text("gene\nALB\nKRT1\n", encoding="utf-8")

        assert read_table(path)["gene"].to_list() == ["ALB", "KRT1"]


class TestReadTableQuoting:
    def test_read_table_removes_quotes_of_quoted_fields(self, tmp_path: Path, no_warnings: None) -> None:
        path = tmp_path / "data.csv"
        path.write_text('name,value\n"Smith, J.",1\n"P12345",2\n', encoding="utf-8")

        assert read_table(path).rows() == [("Smith, J.", 1), ("P12345", 2)]

    def test_read_table_keeps_quotes_inside_fields(self, tmp_path: Path, no_warnings: None) -> None:
        path = tmp_path / "data.tsv"
        path.write_text('name\tvalue\n5" cable\t1\n', encoding="utf-8")

        assert read_table(path).rows() == [('5" cable', 1)]

    def test_read_table_reads_irregular_quotes_without_quote_handling_with_warning(self, tmp_path: Path) -> None:
        path = tmp_path / "data.tsv"
        path.write_text('name\tvalue\n"5 inch" cable\t1\n', encoding="utf-8")

        with pytest.warns(ConversionWarning, match="without quote handling"):
            df = read_table(path)

        assert df.rows() == [('"5 inch" cable', 1)]

    def test_read_table_keeps_quotes_inside_fields_of_several_rows(self, tmp_path: Path, no_warnings: None) -> None:
        path = tmp_path / "data.tsv"
        path.write_text('name\tvalue\nab"c\t1\nx\t2\n', encoding="utf-8")

        assert read_table(path).rows() == [('ab"c', 1), ("x", 2)]

    def test_read_table_keeps_irregular_quotes_in_the_header_with_warning(self, tmp_path: Path) -> None:
        path = tmp_path / "data.tsv"
        path.write_text('"5 inch" cable\tvalue\nx\t1\n', encoding="utf-8")

        with pytest.warns(ConversionWarning, match="without quote handling"):
            df = read_table(path)

        assert df.columns == ['"5 inch" cable', "value"]


class TestReadTableHeader:
    def test_read_table_warns_about_repeated_column_names(self, tmp_path: Path) -> None:
        path = tmp_path / "data.tsv"
        path.write_text("a\ta\tb\n1\t2\t3\n", encoding="utf-8")

        with pytest.warns(ConversionWarning, match="more than once.*'a'"):
            df = read_table(path)

        assert df.columns == ["a", "a_duplicated_0", "b"]

    def test_read_table_keeps_a_single_empty_column_name_without_warning(
        self, tmp_path: Path, no_warnings: None
    ) -> None:
        path = tmp_path / "data.tsv"
        path.write_text("a\t\n1\t2\n", encoding="utf-8")

        assert read_table(path).columns == ["a", ""]

    def test_read_table_reads_header_only_files_as_empty_tables(self, tmp_path: Path) -> None:
        path = tmp_path / "data.tsv"
        path.write_text("a\tb\n", encoding="utf-8")

        df = read_table(path)

        assert df.columns == ["a", "b"]
        assert df.height == 0


class TestReadTableRows:
    @pytest.mark.parametrize("name", ["data.tsv", "data.parquet"])
    def test_read_table_reads_only_the_first_nrows(self, tmp_path: Path, name: str) -> None:
        path = tmp_path / name
        df = pl.DataFrame({"a": list(range(10))})
        if name.endswith(".parquet"):
            df.write_parquet(path)
        else:
            df.write_csv(path, separator="\t")

        assert read_table(path, nrows=3)["a"].to_list() == [0, 1, 2]


class TestReadTableParquet:
    def test_read_table_keeps_parquet_data_types(self, tmp_path: Path) -> None:
        path = tmp_path / "data.parquet"
        df = pl.DataFrame({"id": ["007"], "n": [1], "x": [1.5], "d": [dt.date(2024, 1, 2)], "b": [True]})
        df.write_parquet(path)

        assert read_table(path).equals(df)


class TestReadTableErrors:
    def test_read_table_raises_for_unsupported_extensions(self, tmp_path: Path) -> None:
        path = tmp_path / "data.xlsx"
        path.write_bytes(b"")

        with pytest.raises(UnsupportedFormatError, match="'.xlsx'"):
            read_table(path)

    def test_read_table_raises_for_empty_files(self, tmp_path: Path) -> None:
        with pytest.raises(ReadError, match="is empty"):
            read_table(_write_bytes(tmp_path, b""))

    @pytest.mark.parametrize(
        "content",
        ["a\tb\n1\t2\n3\t4\t5\n", "# exported by instrument\n# version 1\na\tb\n1\t2\n"],
    )
    def test_read_table_raises_for_rows_with_more_fields_than_the_header(self, tmp_path: Path, content: str) -> None:
        with pytest.raises(ReadError, match="more tab-separated fields than the header row"):
            read_table(_write_bytes(tmp_path, content.encode()))

    def test_read_table_raises_for_invalid_parquet_files(self, tmp_path: Path) -> None:
        with pytest.raises(ReadError, match="as a Parquet file"):
            read_table(_write_bytes(tmp_path, b"not parquet", name="data.parquet"))

    @pytest.mark.parametrize("name", ["missing.tsv", "folder.csv"])
    def test_read_table_raises_read_error_for_missing_files_and_folders(self, tmp_path: Path, name: str) -> None:
        (tmp_path / "folder.csv").mkdir()

        with pytest.raises(ReadError, match="does not exist or is not a file"):
            read_table(tmp_path / name)

    def test_read_table_raises_read_error_when_the_file_cannot_be_opened(
        self, tmp_path: Path, monkeypatch: pytest.MonkeyPatch
    ) -> None:
        path = _column_file(tmp_path, ["1"])

        def deny(*args: object, **kwargs: object) -> None:
            raise PermissionError(13, "Permission denied")

        monkeypatch.setattr(Path, "open", deny)

        with pytest.raises(ReadError, match="Could not open 'data.tsv': Permission denied"):
            read_table(path)
