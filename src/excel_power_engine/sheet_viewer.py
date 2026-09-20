from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": MAIN_NS}


def col_to_num(col: str) -> int:
    n = 0
    for ch in col.upper():
        n = n * 26 + ord(ch) - 64
    return n


def num_to_col(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def split_ref(ref: str) -> tuple[str, int]:
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref.upper())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref}")
    return m.group(1), int(m.group(2))


@dataclass(slots=True)
class SheetCell:
    ref: str
    value: str = ""
    formula: str | None = None
    kind: str | None = None


def _shared_strings(z: ZipFile) -> list[str]:
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out: list[str] = []
    for si in root.findall("m:si", NS):
        parts = []
        for t in si.findall(".//m:t", NS):
            parts.append(t.text or "")
        out.append("".join(parts))
    return out


def workbook_sheets(path: str | Path) -> dict[str, str]:
    path = Path(path)
    with ZipFile(path, "r") as z:
        wb = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{REL_NS}}}Relationship")}
        out = {}
        for sh in wb.find("m:sheets", NS):
            name = sh.get("name") or ""
            rid = sh.get(f"{{{DOC_REL_NS}}}id")
            target = relmap.get(rid)
            if not target:
                continue
            part = target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")
            out[name] = part
        return out


def read_region(path: str | Path, sheet: str, *, center_cell: str | None = None,
                row_radius: int = 12, col_radius: int = 8, max_cells: int = 1000):
    path = Path(path)
    with ZipFile(path, "r") as z:
        sheets = workbook_sheets(path)
        if sheet not in sheets:
            raise KeyError(f"Worksheet not found: {sheet}")
        root = ET.fromstring(z.read(sheets[sheet]))
        shared = _shared_strings(z)
        cells: dict[tuple[int, int], SheetCell] = {}
        used_rows: list[int] = []
        used_cols: list[int] = []
        for c in root.findall(".//m:c", NS):
            ref = (c.get("r") or "").upper()
            if not ref:
                continue
            try:
                col, row = split_ref(ref)
            except ValueError:
                continue
            formula_el = c.find("m:f", NS)
            v_el = c.find("m:v", NS)
            is_el = c.find("m:is", NS)
            typ = c.get("t")
            value = ""
            if typ == "s" and v_el is not None and v_el.text is not None:
                try:
                    value = shared[int(v_el.text)]
                except Exception:
                    value = v_el.text
            elif typ == "inlineStr" and is_el is not None:
                value = "".join((t.text or "") for t in is_el.findall(".//m:t", NS))
            elif v_el is not None and v_el.text is not None:
                value = v_el.text
            formula = None if formula_el is None else (formula_el.text or "")
            cells[(row, col_to_num(col))] = SheetCell(ref, value, formula, typ)
            used_rows.append(row)
            used_cols.append(col_to_num(col))

        if not cells:
            return [], [], None
        if center_cell:
            cc, rr = split_ref(center_cell)
            center_row, center_col = rr, col_to_num(cc)
        else:
            center_row = min(used_rows)
            center_col = min(used_cols)
        min_row = max(1, center_row - row_radius)
        max_row = min(max(used_rows), center_row + row_radius)
        min_col = max(1, center_col - col_radius)
        max_col = min(max(used_cols), center_col + col_radius)
        if (max_row - min_row + 1) * (max_col - min_col + 1) > max_cells:
            max_col = min_col + max(1, max_cells // max(1, max_row - min_row + 1)) - 1
        columns = [num_to_col(c) for c in range(min_col, max_col + 1)]
        rows = []
        for r in range(min_row, max_row + 1):
            row_data = []
            for c in range(min_col, max_col + 1):
                item = cells.get((r, c))
                row_data.append(item)
            rows.append((r, row_data))
        target = center_cell.upper() if center_cell else None
        return columns, rows, target
