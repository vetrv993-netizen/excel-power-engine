from __future__ import annotations

import hashlib
import shutil
import time
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

from .target_engine import parse_target_specs, num_to_col


def _sha256(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def _read_package_parts(path: Path) -> dict[str, bytes]:
    with zipfile.ZipFile(path, "r") as z:
        return {n: z.read(n) for n in z.namelist()}


def _has_part(path: Path, name: str) -> bool:
    with zipfile.ZipFile(path, "r") as z:
        return name in z.namelist()


def _restore_vba(source: Path, output: Path, original_vba: bytes) -> None:
    tmp = output.with_suffix(output.suffix + ".vba_tmp")
    with zipfile.ZipFile(output, "r") as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = original_vba if item.filename == "xl/vbaProject.bin" else zin.read(item.filename)
            zout.writestr(item, data)
    if output.exists():
        output.unlink()
    tmp.replace(output)


def _vba_hash(path: Path) -> str | None:
    try:
        with zipfile.ZipFile(path, "r") as z:
            if "xl/vbaProject.bin" not in z.namelist():
                return None
            return _sha256(z.read("xl/vbaProject.bin"))
    except Exception:
        return None


@dataclass
class OperationResult:
    operation: str
    output: str
    visible: bool
    vba_preserved: bool | None
    x14_present: bool
    changed: bool
    warnings: list[str]

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


class ExcelOperations:
    """Excel-native formatting, merge, print and PDF operations.

    Existing XLSM workbooks are first copied to the requested output. Excel performs
    the native operation, then the original VBA binary is restored when necessary.
    """

    def _open(self, output: Path, visible: bool):
        try:
            import xlwings as xw
        except Exception as exc:
            raise RuntimeError("xlwings/Excel غير متاح؛ ثبّت دعم Windows أو شغّل العملية على Windows مع Excel.") from exc
        app = xw.App(visible=visible, add_book=False)
        app.display_alerts = False
        app.screen_updating = True
        book = app.books.open(str(output), update_links=False, read_only=False)
        return app, book

    def _prepare_copy(self, source: str | Path, output: str | Path | None, *, backup: bool = True) -> tuple[Path, Path, bytes | None, str | None]:
        src = Path(source).resolve()
        if not src.exists():
            raise FileNotFoundError(src)
        if output:
            out = Path(output).resolve()
        else:
            out = src.with_name(src.stem + "_OPERATED" + src.suffix)
        if out == src:
            raise ValueError("Output must be different from the source file")
        original_vba = None
        original_vba_hash = None
        with zipfile.ZipFile(src, "r") as z:
            if "xl/vbaProject.bin" in z.namelist():
                original_vba = z.read("xl/vbaProject.bin")
                original_vba_hash = _sha256(original_vba)
        shutil.copy2(src, out)
        if backup:
            backup_path = src.with_name(src.stem + "_BACKUP" + src.suffix)
            if backup_path != out:
                shutil.copy2(src, backup_path)
        return out, src, original_vba, original_vba_hash

    def _finish(self, operation: str, out: Path, original_vba: bytes | None, original_vba_hash: str | None, visible: bool, warnings: list[str], *, changed: bool = True) -> OperationResult:
        if original_vba is not None:
            current = _vba_hash(out)
            if current != original_vba_hash:
                _restore_vba(out, out, original_vba)  # type: ignore[arg-type]
            current = _vba_hash(out)
            vba_ok = current == original_vba_hash
        else:
            vba_ok = None
        with zipfile.ZipFile(out, "r") as z:
            x14 = any("x14" in z.read(n).decode("utf-8", "ignore") for n in z.namelist() if n.startswith("xl/worksheets/") and n.endswith(".xml"))
        return OperationResult(operation, str(out), visible, vba_ok, x14, changed, warnings)

    def _run(self, source, output, operation, action, *, visible=False, backup=True) -> OperationResult:
        out, _, original_vba, original_vba_hash = self._prepare_copy(source, output, backup=backup)
        app = book = None
        warnings: list[str] = []
        try:
            app, book = self._open(out, visible)
            action(book)
            book.save()
            time.sleep(0.3)
        finally:
            try:
                if book is not None:
                    book.close()
            finally:
                if app is not None:
                    app.quit()
        return self._finish(operation, out, original_vba, original_vba_hash, visible, warnings)

    @staticmethod
    def _sheet(book, sheet: str):
        return book.sheets[sheet]

    @classmethod
    def _native_ranges(cls, sht, specs: list[str]) -> list[str]:
        """Convert cell/range/row/column targets into safe A1 ranges for Excel COM."""
        used = sht.used_range
        last_row = max(1, int(used.last_cell.row))
        last_col = max(1, int(used.last_cell.column))
        out=[]
        for raw in specs:
            for spec in parse_target_specs(raw):
                if spec.kind in {"cell", "range"}:
                    out.append(spec.raw)
                elif spec.kind == "row":
                    out.append(f"A{spec.start_row}:{num_to_col(last_col)}{spec.end_row}")
                elif spec.kind == "column":
                    out.append(f"{num_to_col(spec.start_col)}1:{num_to_col(spec.end_col)}{last_row}")
        return out

    def format(self, source, sheet, ranges, *, output=None, fill=None, font_color=None, bold=None, italic=None, underline=None, font_size=None, horizontal=None, vertical=None, wrap=None, number_format=None, visible=False, backup=True):
        def action(book):
            sht = self._sheet(book, sheet)
            for rng in self._native_ranges(sht, ranges):
                r = sht.range(rng)
                if fill:
                    r.color = fill
                if font_color:
                    r.font.color = font_color
                if bold is not None:
                    r.font.bold = bool(bold)
                if italic is not None:
                    r.font.italic = bool(italic)
                if underline is not None:
                    r.font.underline = underline
                if font_size is not None:
                    r.font.size = float(font_size)
                if horizontal:
                    r.api.HorizontalAlignment = _XL_ALIGN.get(horizontal, -4108)
                if vertical:
                    r.api.VerticalAlignment = _XL_ALIGN.get(vertical, -4108)
                if wrap is not None:
                    r.wrap_text = bool(wrap)
                if number_format:
                    r.number_format = number_format
        return self._run(source, output, "format", action, visible=visible, backup=backup)

    def borders(self, source, sheet, ranges, *, output=None, line_style=None, weight=None, color=None, visible=False, backup=True):
        def action(book):
            sht = self._sheet(book, sheet)
            for rng in self._native_ranges(sht, ranges):
                r = sht.range(rng)
                for idx in (7, 8, 9, 10, 11, 12):
                    try:
                        b = r.api.Borders(idx)
                        if line_style is not None: b.LineStyle = line_style
                        if weight is not None: b.Weight = weight
                        if color is not None: b.Color = color
                    except Exception:
                        continue
        return self._run(source, output, "borders", action, visible=visible, backup=backup)

    def merge(self, source, sheet, ranges, *, output=None, unmerge=False, visible=False, backup=True):
        def action(book):
            sht = self._sheet(book, sheet)
            for rng in self._native_ranges(sht, ranges):
                target = sht.range(rng)
                if unmerge:
                    target.unmerge()
                else:
                    target.merge()
        return self._run(source, output, "unmerge" if unmerge else "merge", action, visible=visible, backup=backup)

    def print_setup(self, source, sheet, *, output=None, print_area=None, orientation=None, paper_size=None, fit_width=None, fit_height=None, repeat_rows=None, repeat_columns=None, visible=False, backup=True):
        def action(book):
            sht = self._sheet(book, sheet)
            ps = sht.api.PageSetup
            if print_area:
                ps.PrintArea = print_area
            if orientation:
                ps.Orientation = 2 if orientation.lower().startswith("land") else 1
            if paper_size:
                ps.PaperSize = int(paper_size)
            if fit_width is not None or fit_height is not None:
                ps.Zoom = False
                if fit_width is not None: ps.FitToPagesWide = int(fit_width)
                if fit_height is not None: ps.FitToPagesTall = int(fit_height)
            if repeat_rows:
                ps.PrintTitleRows = repeat_rows
            if repeat_columns:
                ps.PrintTitleColumns = repeat_columns
        return self._run(source, output, "print-setup", action, visible=visible, backup=backup)

    def export_pdf(self, source, *, output_pdf=None, sheet=None, ranges=None, visible=False, open_after=False):
        src = Path(source).resolve()
        pdf = Path(output_pdf).resolve() if output_pdf else src.with_suffix(".pdf")
        if pdf.exists():
            pdf.unlink()
        try:
            import xlwings as xw
        except Exception as exc:
            raise RuntimeError("xlwings/Excel غير متاح؛ لا يمكن تصدير PDF عبر Excel Native.") from exc
        app = xw.App(visible=visible, add_book=False)
        app.display_alerts = False
        book = None
        try:
            book = app.books.open(str(src), update_links=False, read_only=True)
            if sheet:
                sht = book.sheets[sheet]
                if ranges:
                    # Configure print area temporarily; Excel exports the sheet/book using current setup.
                    sht.api.PageSetup.PrintArea = ",".join(ranges)
                sht.api.ExportAsFixedFormat(0, str(pdf))
            else:
                book.api.ExportAsFixedFormat(0, str(pdf))
        finally:
            if book is not None:
                book.close(SaveChanges=False)
            app.quit()
        return {"operation": "pdf", "output_pdf": str(pdf), "exists": pdf.exists(), "sheet": sheet, "ranges": ranges or []}

    def workflow(self, source, operations: list[dict[str, Any]], *, output=None, backup=True, visible=False) -> dict[str, Any]:
        """Apply multiple native Excel operations in one transaction/session."""
        out, _, original_vba, original_vba_hash = self._prepare_copy(source, output, backup=backup)
        app = book = None
        pdf_outputs: list[str] = []
        applied: list[dict[str, Any]] = []
        try:
            app, book = self._open(out, visible)
            for op in operations:
                kind = str(op.get("kind", "")).strip().lower()
                sheet = op.get("sheet")
                ranges = op.get("ranges") or []
                sht = self._sheet(book, sheet) if sheet else None
                if kind == "format":
                    payload = op.get("payload", {})
                    for rng in self._native_ranges(sht, ranges):
                        r=sht.range(rng)
                        if payload.get("fill"): r.color=payload["fill"]
                        if payload.get("font_color"): r.font.color=payload["font_color"]
                        if payload.get("bold") is not None: r.font.bold=bool(payload["bold"])
                        if payload.get("italic") is not None: r.font.italic=bool(payload["italic"])
                        if payload.get("font_size") is not None: r.font.size=float(payload["font_size"])
                        if payload.get("wrap") is not None: r.wrap_text=bool(payload["wrap"])
                        if payload.get("number_format"): r.number_format=payload["number_format"]
                elif kind == "borders":
                    payload=op.get("payload", {})
                    for rng in self._native_ranges(sht, ranges):
                        r=sht.range(rng)
                        for idx in (7,8,9,10,11,12):
                            try:
                                b=r.api.Borders(idx)
                                if payload.get("line_style") is not None: b.LineStyle=payload["line_style"]
                                if payload.get("weight") is not None: b.Weight=payload["weight"]
                                if payload.get("color") is not None: b.Color=payload["color"]
                            except Exception:
                                pass
                elif kind in {"merge","unmerge"}:
                    for rng in self._native_ranges(sht, ranges):
                        t=sht.range(rng)
                        (t.unmerge() if kind=="unmerge" else t.merge())
                elif kind == "print-setup":
                    payload=op.get("payload", {})
                    ps=sht.api.PageSetup
                    if payload.get("print_area"): ps.PrintArea=payload["print_area"]
                    if payload.get("orientation"): ps.Orientation=2 if str(payload["orientation"]).lower().startswith("land") else 1
                    if payload.get("paper_size") is not None: ps.PaperSize=int(payload["paper_size"])
                    if payload.get("fit_width") is not None or payload.get("fit_height") is not None:
                        ps.Zoom=False
                        if payload.get("fit_width") is not None: ps.FitToPagesWide=int(payload["fit_width"])
                        if payload.get("fit_height") is not None: ps.FitToPagesTall=int(payload["fit_height"])
                elif kind == "pdf":
                    pdf=Path(op["output_pdf"]).resolve()
                    if pdf.exists(): pdf.unlink()
                    if sheet:
                        sht.api.ExportAsFixedFormat(0,str(pdf))
                    else:
                        book.api.ExportAsFixedFormat(0,str(pdf))
                    pdf_outputs.append(str(pdf))
                else:
                    raise ValueError(f"Unsupported workflow operation: {kind}")
                applied.append(op)
            book.save()
            time.sleep(0.3)
        finally:
            try:
                if book is not None: book.close()
            finally:
                if app is not None: app.quit()
        result=self._finish("workflow", out, original_vba, original_vba_hash, visible, [], True).as_dict()
        result["applied_operations"]=applied
        result["pdf_outputs"]=pdf_outputs
        return result


_XL_ALIGN = {
    "left": -4131,
    "center": -4108,
    "right": -4152,
    "top": -4160,
    "middle": -4108,
    "bottom": -4107,
}
