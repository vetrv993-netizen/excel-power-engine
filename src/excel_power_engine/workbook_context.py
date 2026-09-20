from __future__ import annotations

from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from zipfile import ZipFile
from xml.etree import ElementTree as ET
import re

MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"m": MAIN}

CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")
RANGE_RE = re.compile(r"^([A-Z]+\d+):([A-Z]+\d+)$")


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
    m = CELL_RE.fullmatch(ref.upper())
    if not m:
        raise ValueError(ref)
    return m.group(1), int(m.group(2))


def _norm_target(target: str) -> str:
    target = target.replace("\\", "/")
    if target.startswith("/"):
        return target[1:]
    if target.startswith("xl/"):
        return target
    return "xl/" + target.lstrip("/")


def _shared_strings(z: ZipFile) -> list[str]:
    try:
        root = ET.fromstring(z.read("xl/sharedStrings.xml"))
    except KeyError:
        return []
    out: list[str] = []
    for si in root.findall("m:si", NS):
        out.append("".join((t.text or "") for t in si.findall(".//m:t", NS)))
    return out


def _color(el: ET.Element | None) -> str | None:
    if el is None:
        return None
    rgb = el.get("rgb")
    if rgb:
        return "#" + rgb[-6:]
    indexed = el.get("indexed")
    if indexed:
        return f"indexed:{indexed}"
    theme = el.get("theme")
    if theme:
        return f"theme:{theme}"
    return None


def _style_maps(z: ZipFile) -> dict[int, dict[str, Any]]:
    """Read a lightweight style map without loading/saving the workbook."""
    try:
        root = ET.fromstring(z.read("xl/styles.xml"))
    except KeyError:
        return {}
    fills = []
    fill_parent = root.find("m:fills", NS)
    if fill_parent is not None:
        for fill in fill_parent.findall("m:fill", NS):
            pat = fill.find("m:patternFill", NS)
            fills.append(_color(pat.find("m:fgColor", NS) if pat is not None else None))
    fonts = []
    font_parent = root.find("m:fonts", NS)
    if font_parent is not None:
        for font in font_parent.findall("m:font", NS):
            fonts.append(_color(font.find("m:color", NS)))
    numfmts: dict[int, str] = {}
    nf_parent = root.find("m:numFmts", NS)
    if nf_parent is not None:
        for n in nf_parent.findall("m:numFmt", NS):
            try:
                numfmts[int(n.get("numFmtId", "0"))] = n.get("formatCode", "")
            except ValueError:
                pass
    out: dict[int, dict[str, Any]] = {}
    cellxfs = root.find("m:cellXfs", NS)
    if cellxfs is None:
        return out
    for idx, xf in enumerate(cellxfs.findall("m:xf", NS)):
        fill_id = int(xf.get("fillId", "0")) if xf.get("fillId") else 0
        font_id = int(xf.get("fontId", "0")) if xf.get("fontId") else 0
        num_id = int(xf.get("numFmtId", "0")) if xf.get("numFmtId") else 0
        out[idx] = {
            "fill": fills[fill_id] if fill_id < len(fills) else None,
            "font_color": fonts[font_id] if font_id < len(fonts) else None,
            "num_format": numfmts.get(num_id),
            "fill_id": fill_id,
            "font_id": font_id,
            "num_fmt_id": num_id,
        }
    return out


@dataclass(slots=True)
class SheetContext:
    name: str
    part: str
    used_range: str | None = None
    header_row: int | None = None
    data_start_row: int | None = None
    data_end_row: int | None = None
    columns: list[dict[str, Any]] = field(default_factory=list)
    suggested_targets: list[str] = field(default_factory=list)
    dominant_fills: list[str] = field(default_factory=list)
    dominant_font_colors: list[str] = field(default_factory=list)
    formula_cells: int = 0
    value_cells: int = 0
    merged_count: int = 0
    protected: bool = False
    sample_rows: list[dict[str, Any]] = field(default_factory=list)

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(slots=True)
class WorkbookContext:
    path: Path
    sheet_count: int
    xlsm: bool
    has_vba: bool
    sheets: list[SheetContext]
    selected_sheet: str | None = None
    selected_cell: str | None = None
    suggestions: dict[str, Any] = field(default_factory=dict)

    def as_dict(self) -> dict[str, Any]:
        return {
            "path": str(self.path),
            "sheet_count": self.sheet_count,
            "xlsm": self.xlsm,
            "has_vba": self.has_vba,
            "selected_sheet": self.selected_sheet,
            "selected_cell": self.selected_cell,
            "suggestions": self.suggestions,
            "sheets": [s.as_dict() for s in self.sheets],
        }

    def sheet(self, name: str | None) -> SheetContext | None:
        if not name:
            return None
        for s in self.sheets:
            if s.name == name:
                return s
        return None


def _cell_value(c: ET.Element, shared: list[str]) -> tuple[str, str | None]:
    typ = c.get("t")
    v = c.find("m:v", NS)
    formula = c.find("m:f", NS)
    if formula is not None:
        return (v.text if v is not None and v.text is not None else "", "=" + (formula.text or ""))
    if typ == "s" and v is not None and v.text:
        try:
            return shared[int(v.text)], None
        except Exception:
            return v.text, None
    if typ == "inlineStr":
        return "".join((t.text or "") for t in c.findall(".//m:t", NS)), None
    if v is not None and v.text is not None:
        return v.text, None
    return "", None


def _detect_header(rows: dict[int, list[tuple[str, str, str | None, int]]], max_scan: int = 20) -> int | None:
    scores: list[tuple[float, int]] = []
    for row_no in sorted(rows)[:max_scan]:
        cells = rows[row_no]
        vals = [str(v).strip() for _, v, _, _ in cells if str(v).strip()]
        if not vals:
            continue
        textish = sum(not re.fullmatch(r"[-+]?\d+(?:\.\d+)?", v) for v in vals)
        unique = len(set(v.casefold() for v in vals))
        score = (len(vals) * 1.2) + (textish * 1.5) + (unique * 0.3)
        if row_no == min(rows):
            score += 0.5
        scores.append((score, row_no))
    return max(scores)[1] if scores else None


def _build_sheet_context(name: str, part: str, z: ZipFile, shared: list[str], styles: dict[int, dict[str, Any]]) -> SheetContext:
    root = ET.fromstring(z.read(part))
    rows: dict[int, list[tuple[str, str, str | None, int]]] = defaultdict(list)
    used_rows: list[int] = []
    used_cols: list[int] = []
    fills: Counter[str] = Counter()
    font_colors: Counter[str] = Counter()
    formula_cells = value_cells = 0
    merged_count = len(root.findall(".//m:mergeCell", NS))
    protected = root.find("m:sheetProtection", NS) is not None

    for c in root.findall(".//m:c", NS):
        ref = (c.get("r") or "").upper()
        if not ref:
            continue
        try:
            col, row = split_ref(ref)
        except ValueError:
            continue
        val, formula = _cell_value(c, shared)
        style_id = int(c.get("s", "0")) if c.get("s") else 0
        rows[row].append((col, val, formula, style_id))
        used_rows.append(row); used_cols.append(col_to_num(col))
        if formula:
            formula_cells += 1
        elif val != "":
            value_cells += 1
        sty = styles.get(style_id, {})
        if sty.get("fill"):
            fills[sty["fill"]] += 1
        if sty.get("font_color"):
            font_colors[sty["font_color"]] += 1

    if not rows:
        return SheetContext(name, part, merged_count=merged_count, protected=protected)

    min_r, max_r = min(used_rows), max(used_rows)
    min_c, max_c = min(used_cols), max(used_cols)
    used_range = f"{num_to_col(min_c)}{min_r}:{num_to_col(max_c)}{max_r}"
    header_row = _detect_header(rows)
    data_start = header_row + 1 if header_row else min_r
    data_end = max_r
    columns: list[dict[str, Any]] = []
    if header_row and header_row in rows:
        headers = {col_to_num(col): val.strip() for col, val, _, _ in rows[header_row] if val.strip()}
        for col_num in range(min_c, max_c + 1):
            title = headers.get(col_num, num_to_col(col_num))
            sample_vals: list[str] = []
            for rr in range(data_start, min(max_r, data_start + 8) + 1):
                for cc, val, formula, _style in rows.get(rr, []):
                    if col_to_num(cc) == col_num and (val or formula):
                        sample_vals.append("=" + formula if formula else val)
            kind = "blank"
            joined = " ".join(sample_vals)
            if not sample_vals:
                kind = "blank"
            elif any(s.startswith("=") for s in sample_vals):
                kind = "formula"
            elif all(re.fullmatch(r"[-+]?\d+(?:\.\d+)?", s.strip()) for s in sample_vals):
                kind = "number"
            elif all(re.fullmatch(r"\d{1,4}[-/]\d{1,2}[-/]\d{1,4}", s.strip()) for s in sample_vals):
                kind = "date"
            else:
                kind = "text"
            columns.append({"column": num_to_col(col_num), "title": title, "kind": kind, "samples": sample_vals[:5]})

    suggested: list[str] = []
    if header_row:
        suggested.append(f"{num_to_col(min_c)}{data_start}:{num_to_col(max_c)}{max_r}")
        for col in columns:
            if col["kind"] in {"text", "number", "date"} and col["title"] != col["column"]:
                suggested.append(f"{col['column']}{data_start}:{col['column']}{max_r}")
        if max_r > data_start:
            suggested.append(f"{num_to_col(min_c)}{data_start}:{num_to_col(max_c)}{min(max_r, data_start + 29)}")
    else:
        suggested.append(used_range)

    samples: list[dict[str, Any]] = []
    for rr in sorted(rows)[: min(len(rows), 6)]:
        vals = {cc: ("=" + fm if fm else val) for cc, val, fm, _sty in rows[rr] if val or fm}
        if vals:
            samples.append({"row": rr, "values": vals})

    return SheetContext(
        name=name, part=part, used_range=used_range, header_row=header_row,
        data_start_row=data_start, data_end_row=data_end, columns=columns,
        suggested_targets=list(dict.fromkeys(suggested))[:12],
        dominant_fills=[c for c, _ in fills.most_common(6)],
        dominant_font_colors=[c for c, _ in font_colors.most_common(6)],
        formula_cells=formula_cells, value_cells=value_cells,
        merged_count=merged_count, protected=protected, sample_rows=samples,
    )


def analyze_workbook(path: str | Path, selected_sheet: str | None = None, selected_cell: str | None = None) -> WorkbookContext:
    p = Path(path).resolve()
    with ZipFile(p, "r") as z:
        root = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        relmap = {r.get("Id"): r.get("Target") for r in rels.findall(f"{{{REL}}}Relationship")}
        sheet_parts: list[tuple[str, str]] = []
        for sh in root.findall("m:sheets/m:sheet", NS):
            name = sh.get("name", "")
            rid = sh.get(f"{{{DOC}}}id", "")
            target = relmap.get(rid)
            if target:
                sheet_parts.append((name, _norm_target(target)))
        shared = _shared_strings(z)
        styles = _style_maps(z)
        sheets = [_build_sheet_context(name, part, z, shared, styles) for name, part in sheet_parts]
        xlsm = p.suffix.lower() == ".xlsm"
        has_vba = "xl/vbaProject.bin" in z.namelist()

    if selected_sheet not in {s.name for s in sheets}:
        selected_sheet = sheets[0].name if sheets else None
    selected = next((s for s in sheets if s.name == selected_sheet), None)
    ranked = sorted(
        sheets,
        key=lambda sh: (
            1 if sh.header_row else 0,
            (sh.value_cells + sh.formula_cells),
            len(sh.columns),
        ),
        reverse=True,
    )
    suggestions: dict[str, Any] = {
        "sheet": selected_sheet,
        "cell": selected_cell or (selected.columns[0]["column"] + str(selected.data_start_row or 1) if selected and selected.columns else "A1"),
        "target": (selected.suggested_targets[0] if selected and selected.suggested_targets else "A1"),
        "ranges": ";".join(selected.suggested_targets[:5]) if selected else "",
        "search_scope": "all",
        "search_fields": [c["title"] for c in (selected.columns if selected else []) if c["title"] and c["title"] != c["column"]][:30],
        "fill": selected.dominant_fills[0] if selected and selected.dominant_fills else "",
        "font_color": selected.dominant_font_colors[0] if selected and selected.dominant_font_colors else "",
        "header_row": selected.header_row if selected else None,
        "field_map": {c["title"]: c["column"] for c in (selected.columns if selected else []) if c["title"] != c["column"]},
        "data_start_row": selected.data_start_row if selected else None,
        "recommended_sheets": [
            {"sheet": sh.name, "used_range": sh.used_range, "header_row": sh.header_row, "value_cells": sh.value_cells, "formula_cells": sh.formula_cells}
            for sh in ranked[:8]
        ],
    }
    return WorkbookContext(p, len(sheets), xlsm, has_vba, sheets, selected_sheet, selected_cell, suggestions)
