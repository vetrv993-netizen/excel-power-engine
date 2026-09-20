from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re
from typing import Any

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": MAIN_NS}

CELL_RE = r"\$?[A-Z]{1,3}\$?\d+"
RANGE_RE = rf"({CELL_RE})(?::({CELL_RE}))?"

@dataclass(slots=True)
class CellSnapshot:
    sheet: str
    cell: str
    cell_type: str
    formula: str | None
    value: Any = None
    cached_value: Any = None
    data_type: str | None = None
    style_id: int | None = None
    sheet_protected: bool = False
    error: str | None = None

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


def _strip_abs(ref: str) -> str:
    return ref.replace("$", "").upper()


def _split_cell(ref: str) -> tuple[str, int]:
    ref = _strip_abs(ref)
    m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
    if not m:
        raise ValueError(f"Invalid cell reference: {ref}")
    return m.group(1), int(m.group(2))


def _col_to_num(col: str) -> int:
    n = 0
    for ch in col:
        n = n * 26 + ord(ch) - 64
    return n


def _num_to_col(n: int) -> str:
    out = ""
    while n:
        n, r = divmod(n - 1, 26)
        out = chr(65 + r) + out
    return out


def _expand_range(start: str, end: str, limit: int = 2000) -> list[str]:
    sc, sr = _split_cell(start)
    ec, er = _split_cell(end)
    scn, ecn = _col_to_num(sc), _col_to_num(ec)
    count = (abs(ern := er - sr) + 1) * (abs(ecn - scn) + 1)
    if count > limit:
        return [f"{_strip_abs(start)}:{_strip_abs(end)}"]
    out = []
    for r in range(min(sr, er), max(sr, er) + 1):
        for c in range(min(scn, ecn), max(scn, ecn) + 1):
            out.append(f"{_num_to_col(c)}{r}")
    return out


def _mask_strings(formula: str) -> str:
    # Replace Excel string literals with spaces so A1-like text inside strings
    # is not treated as a cell reference.
    return re.sub(r'"(?:""|[^"])*"', lambda m: " " * len(m.group(0)), formula)


def extract_references(formula: str, current_sheet: str, *, expand_ranges: bool = True) -> list[str]:
    if not formula:
        return []
    text = _mask_strings(formula)
    refs: list[str] = []
    occupied: list[tuple[int, int]] = []

    # Cross-sheet references: 'Sheet Name'!A1 or Sheet1!A1.
    cross_pat = re.compile(r"(?P<sheet>'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_. ]*)!(?P<ref>" + RANGE_RE + r")")
    for m in cross_pat.finditer(text):
        raw_sheet = m.group("sheet")
        sheet = raw_sheet[1:-1].replace("''", "'") if raw_sheet.startswith("'") else raw_sheet
        ref1, ref2 = m.group("ref").split(":") if ":" in m.group("ref") else (m.group("ref"), None)
        refs.extend(_qualified_refs(sheet, ref1, ref2, expand_ranges))
        occupied.append((m.start(), m.end()))

    # Local references.
    cell_pat = re.compile(RANGE_RE)
    for m in cell_pat.finditer(text):
        if any(a <= m.start() < b for a, b in occupied):
            continue
        ref1 = m.group(1)
        ref2 = m.group(2)
        if _looks_like_function_or_name_prefix(text, m.start()):
            continue
        if ref2 and expand_ranges:
            expanded = _expand_range(ref1, ref2)
            refs.extend([f"{current_sheet}!{x}" for x in expanded])
        elif ref2:
            refs.append(f"{current_sheet}!{_strip_abs(ref1)}:{_strip_abs(ref2)}")
        else:
            refs.append(f"{current_sheet}!{_strip_abs(ref1)}")
    # Preserve order, remove duplicates.
    seen: set[str] = set()
    out: list[str] = []
    for ref in refs:
        if ref not in seen:
            seen.add(ref)
            out.append(ref)
    return out


def _looks_like_function_or_name_prefix(text: str, pos: int) -> bool:
    # Avoid treating the row number in a token such as Sheet1 or a named token
    # as a cell reference. Function calls normally don't match CELL_RE anyway;
    # this guard handles alphanumeric names immediately before a cell token.
    return pos > 0 and text[pos - 1] not in " (+-*/^=,<>&:[{;" and text[pos - 1].isalnum()


def _qualified_refs(sheet: str, start: str, end: str | None, expand_ranges: bool) -> list[str]:
    if end and expand_ranges:
        cells = _expand_range(start, end)
        return [f"{sheet}!{c}" for c in cells]
    if end:
        return [f"{sheet}!{_strip_abs(start)}:{_strip_abs(end)}"]
    return [f"{sheet}!{_strip_abs(start)}"]


def _shared_strings(z: ZipFile) -> list[str]:
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out = []
    for si in root.findall("m:si", NS):
        out.append("".join((t.text or "") for t in si.findall(".//m:t", NS)))
    return out


def _workbook_sheet_parts(z: ZipFile) -> dict[str, str]:
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    relmap = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{REL_NS}}}Relationship")}
    out: dict[str, str] = {}
    for sh in wb.findall("m:sheets/m:sheet", NS):
        name = sh.get("name") or ""
        rid = sh.get(f"{{{DOC_REL_NS}}}id")
        target = relmap.get(rid)
        if not target:
            continue
        out[name] = target.lstrip("/") if target.startswith("/") else "xl/" + target.lstrip("/")
    return out


def _decode_cell_value(c: ET.Element, shared: list[str]) -> Any:
    typ = c.get("t")
    v = c.find("m:v", NS)
    inline = c.find("m:is", NS)
    if typ == "s" and v is not None and v.text is not None:
        try:
            return shared[int(v.text)]
        except Exception:
            return v.text
    if typ == "inlineStr" and inline is not None:
        return "".join((t.text or "") for t in inline.findall(".//m:t", NS))
    if v is None or v.text is None:
        return None
    raw = v.text
    if typ == "b":
        return raw == "1"
    if typ == "e":
        return raw
    try:
        if re.fullmatch(r"-?\d+", raw):
            return int(raw)
        return float(raw)
    except ValueError:
        return raw


def cell_info(path: str | Path, sheet: str, cell: str) -> CellSnapshot:
    path = Path(path)
    cell = _strip_abs(cell)
    with ZipFile(path, "r") as z:
        parts = _workbook_sheet_parts(z)
        if sheet not in parts:
            raise KeyError(f"Worksheet not found: {sheet}")
        root = ET.fromstring(z.read(parts[sheet]))
        protected = root.find("m:sheetProtection", NS) is not None
        shared = _shared_strings(z)
        found = None
        for c in root.findall(".//m:c", NS):
            if _strip_abs(c.get("r") or "") == cell:
                found = c
                break
        if found is None:
            return CellSnapshot(sheet, cell, "blank", None, None, None, None, None, protected, None)
        f = found.find("m:f", NS)
        formula = f.text if f is not None else None
        value = _decode_cell_value(found, shared)
        error = value if found.get("t") == "e" or (isinstance(value, str) and value.startswith("#")) else None
        cell_type = "formula" if formula is not None else ("error" if error else ("value" if value is not None else "blank"))
        return CellSnapshot(
            sheet, cell, cell_type, formula, value, value, found.get("t"),
            int(found.get("s")) if found.get("s", "").isdigit() else None,
            protected, error,
        )


def formula_map(path: str | Path) -> dict[str, dict[str, Any]]:
    path = Path(path)
    out: dict[str, dict[str, Any]] = {}
    with ZipFile(path, "r") as z:
        parts = _workbook_sheet_parts(z)
        shared = _shared_strings(z)
        for sheet, part in parts.items():
            root = ET.fromstring(z.read(part))
            for c in root.findall(".//m:c", NS):
                ref = _strip_abs(c.get("r") or "")
                if not ref:
                    continue
                f = c.find("m:f", NS)
                if f is None or not f.text:
                    continue
                out[f"{sheet}!{ref}"] = {
                    "sheet": sheet,
                    "cell": ref,
                    "formula": f.text,
                    "cached_value": _decode_cell_value(c, shared),
                    "dependencies": extract_references(f.text, sheet),
                }
    return out


def trace_formula(path: str | Path, sheet: str, cell: str, *, max_depth: int = 6) -> dict[str, Any]:
    formulas = formula_map(path)
    root = cell_info(path, sheet, cell)
    key = f"{sheet}!{_strip_abs(cell)}"
    dependents: dict[str, list[str]] = {}
    for owner, data in formulas.items():
        for dep in data["dependencies"]:
            dependents.setdefault(dep, []).append(owner)

    def walk_dependencies(node: str, depth: int, seen: set[str]) -> list[dict[str, Any]]:
        if depth > max_depth or node in seen:
            return []
        data = formulas.get(node)
        if not data:
            return [{"cell": node, "formula": None, "dependencies": []}]
        nxt = []
        new_seen = seen | {node}
        for dep in data["dependencies"]:
            nxt.append({"cell": dep, "formula": formulas.get(dep, {}).get("formula"), "cached_value": formulas.get(dep, {}).get("cached_value"), "children": walk_dependencies(dep, depth + 1, new_seen)})
        return nxt

    return {
        "target": key,
        "cell": root.as_dict(),
        "dependencies": walk_dependencies(key, 0, set()),
        "dependents": dependents.get(key, []),
    }


def formula_issues(path: str | Path) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    for key, data in formula_map(path).items():
        formula = data["formula"] or ""
        cached = data.get("cached_value")
        if "#REF!" in formula.upper():
            issues.append({"cell": key, "kind": "broken-reference", "message": "الصيغة تحتوي على #REF!", "formula": formula})
        if isinstance(cached, str) and cached.startswith("#"):
            issues.append({"cell": key, "kind": "cached-error", "message": f"القيمة المخزنة خطأ Excel: {cached}", "formula": formula, "cached_value": cached})
    return issues
