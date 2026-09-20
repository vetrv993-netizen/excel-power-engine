from __future__ import annotations

from dataclasses import dataclass, asdict
from pathlib import Path
from shutil import copy2
from typing import Any, Iterable
from time import monotonic, sleep

from .inspect import inspect_workbook
from .smart_engine import SmartPlan


ERROR_TOKENS = ("#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NUM!", "#NULL!")


@dataclass(slots=True)
class NativeCalcReport:
    source: Path
    output: Path
    backup: Path | None
    engine: str
    excel_version: str | None
    calculation_state_before: str | None
    calculation_state_after: str | None
    calculation_mode: str | None
    full_rebuild_requested: bool
    saved: bool
    vba_preserved: bool | None
    x14_preserved: bool | None
    formula_errors: list[dict[str, Any]]
    warnings: list[str]

    def as_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["source"] = str(self.source)
        data["output"] = str(self.output)
        data["backup"] = str(self.backup) if self.backup else None
        return data


class NativeExcelBridge:
    """Windows-only Excel Object Model bridge.

    Uses xlwings first and pywin32 as a fallback. The bridge is intentionally
    isolated from the normal OOXML engine so Linux/macOS users can still use
    all non-native functionality.
    """

    def recalculate_file(
        self,
        source: str | Path,
        output: str | Path | None = None,
        *,
        create_backup: bool = True,
        full_rebuild: bool = True,
        visible: bool = False,
    ) -> dict[str, Any]:
        source = Path(source)
        output = Path(output) if output else source.with_name(source.stem + "_RECALCULATED" + source.suffix)
        if source.resolve() == output.resolve():
            raise ValueError("Native recalculation requires a different output path")
        output.parent.mkdir(parents=True, exist_ok=True)
        backup = source.with_name(source.stem + "_BACKUP" + source.suffix) if create_backup else None
        if backup:
            copy2(source, backup)

        before = inspect_workbook(source, deep=True)
        original_vba = before.integrity.get("vba_sha256") if before.has_vba else None
        original_x14 = "x14" in before.features
        formula_errors_before = self._scan_formula_errors(source)

        last_exc: Exception | None = None
        try:
            import xlwings as xw
            result = self._recalculate_xlwings(
                xw, source, output, backup,
                full_rebuild=full_rebuild, visible=visible,
                original_vba=original_vba, original_x14=original_x14,
                formula_errors_before=formula_errors_before,
            )
            return result
        except ImportError:
            pass
        except Exception as exc:
            last_exc = exc

        try:
            import win32com.client as win32
            return self._recalculate_pywin32(
                win32, source, output, backup,
                full_rebuild=full_rebuild, visible=visible,
                original_vba=original_vba, original_x14=original_x14,
                formula_errors_before=formula_errors_before,
            )
        except ImportError as exc:
            if last_exc:
                raise RuntimeError(f"Excel Native via xlwings failed: {last_exc}; pywin32 unavailable: {exc}") from last_exc
            raise RuntimeError("Excel Native requires xlwings or pywin32 on Windows.") from exc

    def edit_cells(
        self,
        source,
        sheet,
        edits: Iterable[Any],
        output=None,
        *,
        create_backup=True,
        recalculate=False,
        plan: SmartPlan | None = None,
    ):
        source = Path(source)
        output = Path(output) if output else source.with_name(source.stem + "_EXCEL_NATIVE" + source.suffix)
        if source.resolve() == output.resolve():
            raise ValueError("Native Excel edit requires a different output path")
        backup = source.with_name(source.stem + "_BACKUP" + source.suffix) if create_backup else None
        if backup:
            copy2(source, backup)

        ops = list(edits)
        before = inspect_workbook(source, deep=True)
        original_vba = before.integrity.get("vba_sha256") if before.has_vba else None
        original_x14 = "x14" in before.features
        last_exc: Exception | None = None

        try:
            import xlwings as xw
            result = self._edit_xlwings(
                xw, source, output, sheet, ops, recalculate, plan, backup,
                original_vba, original_x14,
            )
            return result
        except ImportError:
            pass
        except Exception as exc:
            last_exc = exc

        try:
            import win32com.client as win32
            return self._edit_pywin32(
                win32, source, output, sheet, ops, recalculate, plan, backup,
                original_vba, original_x14,
            )
        except ImportError as exc:
            if last_exc:
                raise RuntimeError(f"Excel Native via xlwings failed: {last_exc}; pywin32 unavailable: {exc}") from last_exc
            raise RuntimeError("Excel Native requires xlwings or pywin32 on Windows.") from exc

    def _recalculate_xlwings(self, xw, source, output, backup, *, full_rebuild, visible, original_vba, original_x14, formula_errors_before):
        app = xw.App(visible=visible, add_book=False)
        app.display_alerts = False
        book = None
        staging = self._prepare_staging_path(output)
        try:
            book = app.books.open(str(source), update_links=False, read_only=False)
            version = str(getattr(app, "version", "unknown"))
            state_before = self._calc_state_xlwings(app)
            mode = self._calc_mode_xlwings(app)
            self._calculate_xlwings(app, full_rebuild)
            self._wait_for_calculation_xlwings(app)
            state_after = self._calc_state_xlwings(app)
            book.save(str(staging))
        finally:
            if book is not None:
                book.close()
            app.quit()

        self._preserve_vba_binary(source, staging, original_vba)
        self._commit_staging(staging, output)
        return self._build_report(source, output, backup, "xlwings", version, state_before, state_after, mode, full_rebuild, original_vba, original_x14, formula_errors_before)

    def _edit_xlwings(self, xw, source, output, sheet, ops, recalculate, plan, backup, original_vba, original_x14):
        app = xw.App(visible=False, add_book=False)
        app.display_alerts = False
        book = None
        staging = self._prepare_staging_path(output)
        try:
            book = app.books.open(str(source), update_links=False, read_only=False)
            sht = book.sheets[sheet]
            for op in ops:
                rng = sht.range(op.cell.upper())
                rng.formula = op.formula if op.formula is not None else op.value
            state_before = self._calc_state_xlwings(app)
            if recalculate:
                self._calculate_xlwings(app, True)
                self._wait_for_calculation_xlwings(app)
            state_after = self._calc_state_xlwings(app)
            book.save(str(staging))
        finally:
            if book is not None:
                book.close()
            app.quit()

        self._preserve_vba_binary(source, staging, original_vba)
        self._commit_staging(staging, output)
        return self._edit_report(source, output, backup, sheet, ops, "xlwings", plan, original_vba, original_x14, state_before, state_after)

    def _recalculate_pywin32(self, win32, source, output, backup, *, full_rebuild, visible, original_vba, original_x14, formula_errors_before):
        app = win32.DispatchEx("Excel.Application")
        app.Visible = bool(visible)
        app.DisplayAlerts = False
        book = None
        staging = self._prepare_staging_path(output)
        try:
            book = app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=False)
            version = str(getattr(app, "Version", "unknown"))
            state_before = self._calc_state_com(app)
            mode = self._calc_mode_com(app)
            if full_rebuild:
                app.CalculateFullRebuild()
            else:
                app.CalculateFull()
            self._wait_for_calculation_com(app)
            state_after = self._calc_state_com(app)
            book.SaveAs(str(staging), FileFormat=book.FileFormat)
        finally:
            if book is not None:
                book.Close(SaveChanges=False)
            app.Quit()

        self._preserve_vba_binary(source, staging, original_vba)
        self._commit_staging(staging, output)
        return self._build_report(source, output, backup, "pywin32", version, state_before, state_after, mode, full_rebuild, original_vba, original_x14, formula_errors_before)

    def _edit_pywin32(self, win32, source, output, sheet, ops, recalculate, plan, backup, original_vba, original_x14):
        app = win32.DispatchEx("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        book = None
        staging = self._prepare_staging_path(output)
        try:
            book = app.Workbooks.Open(str(source), UpdateLinks=0, ReadOnly=False)
            ws = book.Worksheets(sheet)
            for op in ops:
                cell = ws.Range(op.cell.upper())
                if op.formula is not None:
                    cell.Formula = op.formula
                else:
                    cell.Value = op.value
            state_before = self._calc_state_com(app)
            if recalculate:
                app.CalculateFullRebuild()
                self._wait_for_calculation_com(app)
            state_after = self._calc_state_com(app)
            book.SaveAs(str(staging), FileFormat=book.FileFormat)
        finally:
            if book is not None:
                book.Close(SaveChanges=False)
            app.Quit()

        self._preserve_vba_binary(source, staging, original_vba)
        self._commit_staging(staging, output)
        return self._edit_report(source, output, backup, sheet, ops, "pywin32", plan, original_vba, original_x14, state_before, state_after)

    @staticmethod
    def _calculate_xlwings(app, full_rebuild: bool) -> None:
        api = app.api
        if full_rebuild:
            api.CalculateFullRebuild()
        else:
            try:
                api.CalculateFull()
            except Exception:
                app.calculate()

    @staticmethod
    def _calc_state_xlwings(app) -> str | None:
        try:
            return str(app.api.CalculationState)
        except Exception:
            return None

    @staticmethod
    def _calc_mode_xlwings(app) -> str | None:
        try:
            return str(app.api.Calculation)
        except Exception:
            return None

    @staticmethod
    def _calc_state_com(app) -> str | None:
        try:
            return str(app.CalculationState)
        except Exception:
            return None

    @staticmethod
    def _calc_mode_com(app) -> str | None:
        try:
            return str(app.Calculation)
        except Exception:
            return None

    def _build_report(self, source, output, backup, engine, version, state_before, state_after, mode, full_rebuild, original_vba, original_x14, formula_errors_before):
        after = inspect_workbook(output, deep=True)
        vba_ok = None if original_vba is None else after.integrity.get("vba_sha256") == original_vba
        x14_ok = None if not original_x14 else "x14" in after.features
        formula_errors_after = self._scan_formula_errors(output)
        before_keys = {(e.get("sheet"), e.get("cell"), e.get("error")) for e in formula_errors_before}
        after_keys = {(e.get("sheet"), e.get("cell"), e.get("error")) for e in formula_errors_after}
        new_formula_errors = [e for e in formula_errors_after if (e.get("sheet"), e.get("cell"), e.get("error")) not in before_keys]
        warnings = []
        if vba_ok is False:
            warnings.append("تغيرت بصمة VBA بعد الحفظ Native ولم يمكن الحفاظ عليها تلقائيًا؛ راجع الملف قبل اعتماده.")
        if x14_ok is False:
            warnings.append("لم تعد امتدادات x14 موجودة بعد الحفظ Native؛ راجع الملف.")
        if state_after not in (None, "0"):
            warnings.append("ظل Excel في حالة حساب غير مكتملة بعد الانتظار؛ راجع الحساب قبل الاعتماد.")
        if new_formula_errors:
            warnings.append(f"ظهرت أخطاء صيغ جديدة بعد إعادة الحساب: {len(new_formula_errors)}.")
        report = NativeCalcReport(
            source=Path(source), output=Path(output), backup=Path(backup) if backup else None,
            engine=engine, excel_version=version, calculation_state_before=state_before,
            calculation_state_after=state_after, calculation_mode=mode,
            full_rebuild_requested=full_rebuild, saved=Path(output).exists(),
            vba_preserved=vba_ok, x14_preserved=x14_ok,
            formula_errors=formula_errors_after, warnings=warnings,
        ).as_dict()
        report["formula_errors_before"] = formula_errors_before
        report["new_formula_errors"] = new_formula_errors
        report["preexisting_formula_errors"] = [e for e in formula_errors_after if (e.get("sheet"), e.get("cell"), e.get("error")) in before_keys and (e.get("sheet"), e.get("cell"), e.get("error")) in after_keys]
        report["safe_to_accept"] = bool(report["saved"] and report["vba_preserved"] is not False and report["x14_preserved"] is not False and not new_formula_errors and state_after in (None, "0"))
        return report

    def _edit_report(self, source, output, backup, sheet, ops, engine, plan, original_vba, original_x14, state_before, state_after):
        after = inspect_workbook(output, deep=True)
        vba_ok = None if original_vba is None else after.integrity.get("vba_sha256") == original_vba
        x14_ok = None if not original_x14 else "x14" in after.features
        warnings = []
        if vba_ok is False:
            warnings.append("تغيرت بصمة VBA بعد الحفظ عبر Excel Native؛ راجع الملف قبل اعتماده.")
        if x14_ok is False:
            warnings.append("لم تعد امتدادات x14 موجودة بعد الحفظ عبر Excel Native؛ راجع الملف.")
        return {
            "source": str(source), "output": str(output), "backup": str(backup) if backup else None,
            "sheet": sheet, "edited_cells": [op.cell.upper() for op in ops],
            "changed_parts": ["excel-native-save"],
            "protected_parts_unchanged": None,
            "vba_preserved": vba_ok,
            "x14_preserved": x14_ok,
            "formula_errors": self._scan_formula_errors(output),
            "calculation_state_before": state_before,
            "calculation_state_after": state_after,
            "warnings": warnings,
            "smart_plan": plan.as_dict() if plan else None,
            "executed_engine": "excel-native",
        }

    @staticmethod
    def _wait_for_calculation_xlwings(app, timeout: float = 30.0) -> str | None:
        deadline = monotonic() + timeout
        api = app.api
        try:
            waiter = getattr(api, "CalculateUntilAsyncQueriesDone", None)
            if callable(waiter):
                waiter()
        except Exception:
            pass
        while monotonic() < deadline:
            try:
                state = int(api.CalculationState)
            except Exception:
                return None
            if state == 0:
                return "0"
            sleep(0.15)
        try:
            return str(api.CalculationState)
        except Exception:
            return None

    @staticmethod
    def _wait_for_calculation_com(app, timeout: float = 30.0) -> str | None:
        deadline = monotonic() + timeout
        try:
            waiter = getattr(app, "CalculateUntilAsyncQueriesDone", None)
            if callable(waiter):
                waiter()
        except Exception:
            pass
        while monotonic() < deadline:
            try:
                state = int(app.CalculationState)
            except Exception:
                return None
            if state == 0:
                return "0"
            sleep(0.15)
        try:
            return str(app.CalculationState)
        except Exception:
            return None

    @staticmethod
    def _prepare_staging_path(output: str | Path) -> Path:
        """Return a temporary XLSM path that Excel can save without owning the final output."""
        output = Path(output)
        staging = output.with_name(f"{output.stem}.excel_native_tmp{output.suffix}")
        staging.unlink(missing_ok=True)
        return staging

    @staticmethod
    def _commit_staging(staging: str | Path, output: str | Path) -> None:
        """Move the unlocked staging workbook to the requested output path."""
        from os import replace
        staging = Path(staging)
        output = Path(output)
        try:
            replace(staging, output)
        except PermissionError as exc:
            raise PermissionError(
                f"لا يمكن استبدال ملف الإخراج '{output}'. أغلق الملف إذا كان مفتوحًا في Excel ثم أعد المحاولة."
            ) from exc

    @staticmethod
    def _preserve_vba_binary(source: str | Path, output: str | Path, original_vba: str | None) -> bool | None:
        """Restore the original VBA binary after an Excel Native save when Excel rewrites it.

        Recalculation is not intended to modify VBA. Restoring only xl/vbaProject.bin
        keeps Excel's recalculated worksheet/package parts while retaining the exact
        original VBA payload. The final hash is verified by the report layer.
        """
        if original_vba is None:
            return None
        from zipfile import ZipFile, ZIP_DEFLATED
        import tempfile
        source = Path(source)
        output = Path(output)
        with ZipFile(source, "r") as src_zip:
            original_bin = src_zip.read("xl/vbaProject.bin")
        with ZipFile(output, "r") as out_zip:
            infos = out_zip.infolist()
            entries = {info.filename: out_zip.read(info.filename) for info in infos}
        if entries.get("xl/vbaProject.bin") == original_bin:
            return True
        entries["xl/vbaProject.bin"] = original_bin
        temp = output.with_suffix(output.suffix + ".vba_tmp")
        with ZipFile(temp, "w") as z:
            for info in infos:
                data = entries[info.filename]
                if info.filename == "xl/vbaProject.bin":
                    zi = info
                    zi.compress_type = ZIP_DEFLATED
                    z.writestr(zi, data)
                else:
                    z.writestr(info, data)
        temp.replace(output)
        return True

    @staticmethod
    def _scan_formula_errors(path: str | Path) -> list[dict[str, Any]]:
        """Scan cached Excel formula results directly from OOXML.

        This deliberately avoids OpenPyXL because merely loading XLSM files with
        unsupported extensions such as x14 can emit a warning and, on save, can
        strip those extensions. Formula-error detection is therefore read-only
        and performed against the worksheet XML inside the ZIP package.
        """
        from zipfile import ZipFile
        from xml.etree import ElementTree as ET

        errors: list[dict[str, Any]] = []
        path = Path(path)
        if not path.exists():
            return []

        NS = {
            "main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
            "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
        }
        REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
        DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        SHEET_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"

        try:
            with ZipFile(path, "r") as zf:
                workbook_root = ET.fromstring(zf.read("xl/workbook.xml"))
                rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

                rel_targets: dict[str, str] = {}
                for rel in rels_root.findall(f"{{{REL_NS}}}Relationship"):
                    rid = rel.get("Id")
                    target = rel.get("Target")
                    rel_type = rel.get("Type")
                    if rid and target and rel_type == SHEET_REL_TYPE:
                        if target.startswith("/"):
                            part = target.lstrip("/")
                        else:
                            part = "xl/" + target.lstrip("/")
                        rel_targets[rid] = part

                sheets = workbook_root.find("main:sheets", NS)
                if sheets is None:
                    return []

                for sheet in sheets.findall("main:sheet", NS):
                    name = sheet.get("name") or ""
                    rid = sheet.get(f"{{{DOC_REL}}}id")
                    part = rel_targets.get(rid or "")
                    if not part or part not in zf.namelist():
                        continue
                    root = ET.fromstring(zf.read(part))
                    for cell in root.findall(".//main:c", NS):
                        formula = cell.find("main:f", NS)
                        if formula is None:
                            continue
                        value = cell.find("main:v", NS)
                        cached = (value.text or "").strip() if value is not None else ""
                        cell_type = cell.get("t")
                        if cell_type == "e" or cached in ERROR_TOKENS:
                            if cached in ERROR_TOKENS:
                                errors.append({
                                    "sheet": name,
                                    "cell": cell.get("r"),
                                    "error": cached,
                                })
                                if len(errors) >= 500:
                                    return errors
        except Exception:
            # Formula-error scanning is best-effort and must never turn a successful
            # Excel-native save into a hard failure.
            return []
        return errors
