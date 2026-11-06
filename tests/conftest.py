from __future__ import annotations

import re
import xml.etree.ElementTree as ET
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path
from zipfile import ZipFile

import pytest

from xlsxpress.app import appdata
from xlsxpress.core import paths

_NAMESPACE = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_BUILTIN_NUMBER_FORMATS = {0: "General"}
_ESCAPED_CHARACTER = re.compile(r"_x([0-9A-Fa-f]{4})_")


@dataclass(frozen=True)
class Cell:
    kind: str
    value: str
    number_format: str
    bold: bool


@dataclass
class Sheet:
    cells: dict[str, Cell] = field(default_factory=dict)
    frozen_rows: int = 0
    autofilter: str | None = None
    column_widths: list[float] = field(default_factory=list)


def read_sheet(path: Path) -> Sheet:
    with ZipFile(path) as workbook:
        worksheet = ET.fromstring(workbook.read("xl/worksheets/sheet1.xml"))
        styles = ET.fromstring(workbook.read("xl/styles.xml"))
        shared_strings = []
        if "xl/sharedStrings.xml" in workbook.namelist():
            root = ET.fromstring(workbook.read("xl/sharedStrings.xml"))
            shared_strings = [_text_of(item) for item in root.iter(f"{_NAMESPACE}si")]

    number_formats = dict(_BUILTIN_NUMBER_FORMATS)
    for number_format in styles.iter(f"{_NAMESPACE}numFmt"):
        number_formats[int(number_format.get("numFmtId", "0"))] = number_format.get("formatCode", "")
    fonts = [font.find(f"{_NAMESPACE}b") is not None for font in styles.iter(f"{_NAMESPACE}font")]
    cell_xfs = styles.find(f"{_NAMESPACE}cellXfs")
    cell_styles = [
        (number_formats.get(int(xf.get("numFmtId", "0")), "unknown"), fonts[int(xf.get("fontId", "0"))])
        for xf in (cell_xfs if cell_xfs is not None else [])
    ]

    sheet = Sheet()
    for cell in worksheet.iter(f"{_NAMESPACE}c"):
        number_format, bold = cell_styles[int(cell.get("s", "0"))]
        parsed = _parse_cell(cell, shared_strings)
        if parsed is not None:
            kind, value = parsed
            sheet.cells[cell.get("r", "")] = Cell(kind, value, number_format, bold)
    pane = worksheet.find(f".//{_NAMESPACE}pane")
    if pane is not None and pane.get("state") == "frozen":
        sheet.frozen_rows = int(float(pane.get("ySplit", "0")))
    autofilter = worksheet.find(f"{_NAMESPACE}autoFilter")
    if autofilter is not None:
        sheet.autofilter = autofilter.get("ref")
    for col in worksheet.iter(f"{_NAMESPACE}col"):
        span = int(col.get("max", "1")) - int(col.get("min", "1")) + 1
        sheet.column_widths.extend([float(col.get("width", "0"))] * span)
    return sheet


def _parse_cell(cell: ET.Element, shared_strings: list[str]) -> tuple[str, str] | None:
    cell_type = cell.get("t")
    if cell.find(f"{_NAMESPACE}f") is not None:
        return "formula", cell.findtext(f"{_NAMESPACE}f", "")
    if cell_type == "inlineStr":
        return "text", _text_of(cell)
    value = cell.findtext(f"{_NAMESPACE}v")
    if value is None:
        return None
    if cell_type == "s":
        return "text", shared_strings[int(value)]
    if cell_type == "b":
        return "boolean", value
    return "number", value


def _text_of(element: ET.Element) -> str:
    text = "".join(node.text or "" for node in element.iter(f"{_NAMESPACE}t"))
    return _ESCAPED_CHARACTER.sub(lambda match: chr(int(match.group(1), 16)), text)


@pytest.fixture
def xlsx_reader() -> Callable[[Path], Sheet]:
    return read_sheet


@pytest.fixture(autouse=True)
def isolated_cache_dir(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    directory = tmp_path / "user-cache"
    monkeypatch.setattr(paths, "get_cache_dir", lambda: directory)
    monkeypatch.setattr(appdata, "get_appdata_dir", lambda: tmp_path / "user-data")
    return directory
