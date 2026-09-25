"""Small, dependency-free XLSX reader for source-data workbooks.

It reads values from standard XLSX XML, including files that some versions of
openpyxl treat as empty because worksheet dimensions or relationships are odd.
"""
from __future__ import annotations

import posixpath
import re
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
OFFICE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _column_index(address: str) -> int:
    match = re.match(r"([A-Z]+)", address)
    if not match:
        return 0
    value = 0
    for char in match.group(1):
        value = value * 26 + ord(char) - 64
    return value - 1


def _target_path(target: str) -> str:
    target = target.lstrip("/")
    if target.startswith("xl/"):
        return target
    return posixpath.normpath("xl/" + target)


def read_xlsx(path: str | Path) -> dict[str, list[list[object]]]:
    """Return {sheet_name: rows}, with formulas represented by cached values."""
    with zipfile.ZipFile(path) as archive:
        shared: list[str] = []
        if "xl/sharedStrings.xml" in archive.namelist():
            root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
            for item in root.findall(f"{{{MAIN}}}si"):
                shared.append("".join(node.text or "" for node in item.iter() if node.tag == f"{{{MAIN}}}t"))

        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        relations = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        relation_map = {node.attrib["Id"]: node.attrib["Target"] for node in relations}
        result: dict[str, list[list[object]]] = {}

        sheets = workbook.find(f"{{{MAIN}}}sheets")
        if sheets is None:
            return result
        for sheet in sheets:
            name = sheet.attrib["name"]
            relation_id = sheet.attrib[f"{{{OFFICE_REL}}}id"]
            root = ET.fromstring(archive.read(_target_path(relation_map[relation_id])))
            rows: list[list[object]] = []
            for row in root.findall(f".//{{{MAIN}}}sheetData/{{{MAIN}}}row"):
                cells: dict[int, object] = {}
                for cell in row.findall(f"{{{MAIN}}}c"):
                    index = _column_index(cell.attrib.get("r", "A1"))
                    cell_type = cell.attrib.get("t")
                    value_node = cell.find(f"{{{MAIN}}}v")
                    inline = cell.find(f"{{{MAIN}}}is")
                    value: object = None
                    if cell_type == "s" and value_node is not None:
                        value = shared[int(value_node.text or 0)]
                    elif cell_type == "inlineStr" and inline is not None:
                        value = "".join(node.text or "" for node in inline.iter() if node.tag == f"{{{MAIN}}}t")
                    elif cell_type == "b" and value_node is not None:
                        value = value_node.text == "1"
                    elif value_node is not None:
                        text = value_node.text or ""
                        try:
                            number = float(text)
                            value = int(number) if number.is_integer() else number
                        except ValueError:
                            value = text
                    cells[index] = value
                if cells:
                    rows.append([cells.get(index) for index in range(max(cells) + 1)])
            result[name] = rows
        return result


def first_nonempty_sheet(path: str | Path) -> list[list[object]]:
    sheets = read_xlsx(path)
    return next((rows for rows in sheets.values() if rows), [])
