from __future__ import annotations

import json
import shutil
import tempfile
import time
import uuid
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Any

from .audit import log_event
from .bulk import (
    execute_matrix,
    execute_multi_ranges,
    parse_matrix,
    parse_target_ranges,
    preview_matrix,
    preview_multi_ranges,
    read_data_file,
)
from .format_engine import ExcelOperations
from .operations import preview_operations
from .search_engine import search_workbook
from .sheet_viewer import workbook_sheets
from .target_engine import parse_target_specs, resolve_target_specs


@dataclass
class WorkspaceOperation:
    kind: str
    sheet: str | None = None
    ranges: list[str] | None = None
    payload: dict[str, Any] | None = None

    def as_dict(self) -> dict[str, Any]:
        out = asdict(self)
        out["ranges"] = out.get("ranges") or []
        out["payload"] = out.get("payload") or {}
        return out


def _load_ops(operations: list[dict[str, Any]] | dict[str, Any] | str | Path) -> list[dict[str, Any]]:
    if isinstance(operations, (str, Path)):
        data = json.loads(Path(operations).read_text(encoding="utf-8"))
    else:
        data = operations
    if isinstance(data, dict):
        data = data.get("operations", data.get("workflow", []))
    if not isinstance(data, list):
        raise ValueError("Workspace operations must be a list or {operations:[...]}")
    return [dict(x) for x in data]


def validate_workspace_operations(operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    allowed = {
        "search", "bulk", "format", "borders", "merge", "unmerge",
        "print-setup", "pdf", "open-excel"
    }
    out=[]
    for i, op in enumerate(operations, 1):
        if not isinstance(op, dict):
            raise ValueError(f"Operation #{i} must be an object")
        kind=str(op.get("kind", "")).strip().lower()
        if kind not in allowed:
            raise ValueError(f"Unsupported workspace operation: {kind}")
        if kind not in {"search", "pdf"} and not op.get("sheet"):
            raise ValueError(f"Operation {kind} requires sheet")
        if kind in {"format", "borders", "merge", "unmerge"}:
            ranges=op.get("ranges") or []
            if isinstance(ranges, str):
                ranges=[x.strip() for x in ranges.split(";") if x.strip()]
            if not ranges:
                raise ValueError(f"Operation {kind} requires ranges")
            for token in ranges:
                parse_target_specs(str(token))
            op["ranges"]=ranges
        payload = op.get("payload") or {}
        if kind == "bulk":
            if not payload.get("data") and not payload.get("data_file"):
                raise ValueError("bulk requires payload.data or payload.data_file")
            if not payload.get("start_cell") and not op.get("ranges"):
                raise ValueError("bulk requires payload.start_cell or ranges")
        if kind == "search" and not payload.get("term"):
            raise ValueError("search requires payload.term")
        if kind == "pdf" and not payload.get("output_pdf"):
            raise ValueError("pdf requires payload.output_pdf")
        op["kind"]=kind
        op["payload"]=payload
        out.append(op)
    return out


def preview_workspace(source: str | Path, operations: list[dict[str, Any]] | dict[str, Any] | str | Path) -> dict[str, Any]:
    src=Path(source).resolve()
    if not src.exists():
        raise FileNotFoundError(src)
    ops=validate_workspace_operations(_load_ops(operations))
    summary=[]
    estimated_cells=0
    searches=[]
    for i,op in enumerate(ops,1):
        kind=op["kind"]
        sheet=op.get("sheet")
        ranges=op.get("ranges") or []
        payload=op.get("payload") or {}
        item={"step":i,"kind":kind,"sheet":sheet,"ranges":ranges}
        if kind in {"format","borders","merge","unmerge"} and sheet:
            count=0
            for target in ranges:
                try:
                    count += len(resolve_target_specs(src, sheet, target, max_cells=200000))
                except Exception:
                    count += 0
            item["target_cell_count"]=count
            estimated_cells += count
        elif kind == "bulk":
            data = read_data_file(payload["data_file"], sheet=payload.get("data_sheet")) if payload.get("data_file") else parse_matrix(str(payload.get("data", "")))
            if ranges:
                changes=preview_multi_ranges(src, sheet, ";".join(ranges), data, interpret_formulas=payload.get("interpret_formulas", True), clear_empty=payload.get("clear_empty", False))
            else:
                changes=preview_matrix(src, sheet, payload["start_cell"], data, interpret_formulas=payload.get("interpret_formulas", True), clear_empty=payload.get("clear_empty", False))
            item["rows"]=len(data); item["columns"]=max((len(r) for r in data), default=0); item["change_count"]=sum(c.before != c.after for c in changes)
            item["sample_changes"]= [c.as_dict() for c in changes[:20]]
            estimated_cells += item["change_count"]
        elif kind == "search":
            hits=search_workbook(src, str(payload["term"]), sheets=[sheet] if sheet else None,
                                 mode=payload.get("mode", "contains"), case_sensitive=payload.get("case_sensitive", False),
                                 search_formulas=payload.get("include_formulas", True), search_values=payload.get("include_values", True),
                                 max_results=int(payload.get("max_results", 2000)))
            item["match_count"]=len(hits); item["sample_hits"]= [h.as_dict() for h in hits[:20]]
            searches.extend(h.as_dict() for h in hits)
        elif kind in {"pdf","print-setup"}:
            item["requires_excel"]=True
        summary.append(item)
    return {
        "source": str(src),
        "operation_count": len(ops),
        "estimated_cells_touched": estimated_cells,
        "operations": summary,
        "search_hits": searches,
        "safe_to_execute": True,
    }


def _copy(src: Path, dst: Path) -> None:
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)


def _transaction_backup(output: Path) -> Path:
    return output.with_name(output.stem + "_TRANSACTION_BACKUP" + output.suffix)


def _next_tmp(dir_path: Path, suffix: str) -> Path:
    return dir_path / f"step_{uuid.uuid4().hex}{suffix}"


def run_workspace(source: str | Path, operations: list[dict[str, Any]] | dict[str, Any] | str | Path, output: str | Path, *, backup: bool=True, visible: bool=False, open_after: bool=True) -> dict[str, Any]:
    src=Path(source).resolve(); out=Path(output).resolve()
    if src == out:
        raise ValueError("Output must be different from source")
    ops=validate_workspace_operations(_load_ops(operations))
    if not src.exists(): raise FileNotFoundError(src)
    out.parent.mkdir(parents=True, exist_ok=True)
    preview=preview_workspace(src, ops)
    backup_path=_transaction_backup(out)
    if backup:
        _copy(src, backup_path)
    temp_dir=Path(tempfile.mkdtemp(prefix="excel_power_engine_v09_"))
    working=temp_dir / out.name
    _copy(src, working)
    current_sheet=ops[0].get("sheet") if ops else None
    current_cell=None
    applied=[]
    pdf_outputs=[]
    search_results=[]
    try:
        for idx,op in enumerate(ops,1):
            kind=op["kind"]; sheet=op.get("sheet") or current_sheet; payload=op.get("payload") or {}; ranges=op.get("ranges") or []
            if kind == "search":
                hits=search_workbook(working, str(payload["term"]), sheets=[sheet] if sheet else None, mode=payload.get("mode","contains"), case_sensitive=payload.get("case_sensitive",False), include_formulas=payload.get("include_formulas",True), include_values=payload.get("include_values",True), max_results=int(payload.get("max_results",2000)))
                search_results.extend(h.as_dict() for h in hits)
                if hits:
                    current_sheet=hits[0].sheet; current_cell=hits[0].cell
                applied.append({"step":idx,"kind":kind,"matches":len(hits)})
                continue
            next_path=_next_tmp(temp_dir, working.suffix)
            if kind == "bulk":
                data=read_data_file(payload["data_file"], sheet=payload.get("data_sheet")) if payload.get("data_file") else parse_matrix(str(payload.get("data","")))
                if ranges:
                    result=execute_multi_ranges(working, sheet, ";".join(ranges), data, next_path, interpret_formulas=payload.get("interpret_formulas",True), clear_empty=payload.get("clear_empty",False), create_backup=False)
                    first=ranges[0].split(":",1)[0]
                else:
                    result=execute_matrix(working, sheet, payload["start_cell"], data, next_path, interpret_formulas=payload.get("interpret_formulas",True), clear_empty=payload.get("clear_empty",False), create_backup=False)
                    first=payload["start_cell"]
                current_cell=first
                current_sheet=sheet
            elif kind in {"format","borders","merge","unmerge","print-setup"}:
                xo=ExcelOperations()
                if kind=="format": result=xo.format(working,sheet,ranges,output=next_path,fill=payload.get("fill"),font_color=payload.get("font_color"),bold=payload.get("bold"),italic=payload.get("italic"),font_size=payload.get("font_size"),horizontal=payload.get("horizontal"),vertical=payload.get("vertical"),wrap=payload.get("wrap"),number_format=payload.get("number_format"),visible=visible,backup=False)
                elif kind=="borders": result=xo.borders(working,sheet,ranges,output=next_path,line_style=payload.get("line_style"),weight=payload.get("weight"),color=payload.get("color"),visible=visible,backup=False)
                elif kind in {"merge","unmerge"}: result=xo.merge(working,sheet,ranges,output=next_path,unmerge=(kind=="unmerge"),visible=visible,backup=False)
                else: result=xo.print_setup(working,sheet,output=next_path,print_area=payload.get("print_area"),orientation=payload.get("orientation"),paper_size=payload.get("paper_size"),fit_width=payload.get("fit_width"),fit_height=payload.get("fit_height"),repeat_rows=payload.get("repeat_rows"),repeat_columns=payload.get("repeat_columns"),visible=visible,backup=False)
                current_sheet=sheet; current_cell=(ranges[0].split(":",1)[0] if ranges else current_cell)
            elif kind=="pdf":
                pdf=Path(payload["output_pdf"]).resolve()
                result=ExcelOperations().export_pdf(working,output_pdf=pdf,sheet=sheet,ranges=ranges or None,visible=visible)
                pdf_outputs.append(str(pdf)); applied.append({"step":idx,"kind":kind,"result":result}); continue
            elif kind=="open-excel":
                from .open_excel import open_in_excel
                current_sheet=sheet or current_sheet; current_cell=payload.get("cell") or current_cell
                result=open_in_excel(working,current_sheet,current_cell)
                applied.append({"step":idx,"kind":kind,"result":result}); continue
            else:
                raise ValueError(kind)
            _copy(next_path, working)
            applied.append({"step":idx,"kind":kind,"result":result.as_dict() if hasattr(result,"as_dict") else result})
            try: next_path.unlink()
            except OSError: pass
        _copy(working,out)
        if open_after and current_sheet and current_cell:
            try:
                from .open_excel import open_in_excel
                visual=open_in_excel(out,current_sheet,current_cell)
            except Exception as exc:
                visual={"opened":False,"error":str(exc)}
        else:
            visual={"opened":False,"reason":"no_target"}
        result={"source":str(src),"output":str(out),"backup":str(backup_path) if backup else None,"transaction_id":backup_path.stem,"preview":preview,"applied_operations":applied,"search_results":search_results,"pdf_outputs":pdf_outputs,"visual_verification":visual,"rollback_available":backup_path.exists()}
        log_event("workspace-run",path=str(src),details={"output":str(out),"transaction_id":result["transaction_id"],"operation_count":len(ops)})
        return result
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def rollback_workspace(output: str | Path, *, destination: str | Path | None = None) -> dict[str, Any]:
    out=Path(output).resolve(); backup=_transaction_backup(out)
    if not backup.exists():
        raise FileNotFoundError(f"Transaction backup not found: {backup}")
    dest=Path(destination).resolve() if destination else out
    if dest.exists():
        tmp=dest.with_suffix(dest.suffix+".rollback_tmp")
        _copy(dest,tmp)
        _copy(backup,dest)
        try: tmp.unlink()
        except OSError: pass
    else:
        _copy(backup,dest)
    result={"rolled_back":True,"output":str(dest),"backup":str(backup),"visual_verification":{"opened":False}}
    log_event("workspace-rollback",path=str(dest),details=result)
    return result
