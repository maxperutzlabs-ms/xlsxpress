from __future__ import annotations

import xml.etree.ElementTree as ET
from pathlib import Path
from zipfile import ZipFile

import polars as pl
import pytest

from xlsxpress.core.writers import write_excel

pytestmark = pytest.mark.unit


class TestWriteExcel:
    def test_write_excel_preserves_float_precision_in_general_format(self, tmp_path: Path) -> None:
        output_path = tmp_path / "precision.xlsx"
        write_excel(pl.DataFrame({"measurement": [1.23456789]}), output_path)

        namespaces = {"x": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
        with ZipFile(output_path) as workbook:
            worksheet = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
            styles = ET.fromstring(workbook.read("xl/styles.xml"))

        cell = worksheet.find(".//x:c[@r='A2']", namespaces)
        assert cell is not None
        assert cell.findtext("x:v", namespaces=namespaces) == "1.23456789"

        style_index = int(cell.get("s", "0"))
        style = styles.find("x:cellXfs", namespaces)[style_index]
        assert style.get("numFmtId", "0") == "0"
