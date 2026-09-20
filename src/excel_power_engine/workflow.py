from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from .format_engine import ExcelOperations
from .target_engine import parse_target_specs


def load_workflow(path: str | Path) -> list[dict[str, Any]]:
    data=json.loads(Path(path).read_text(encoding="utf-8"))
    if isinstance(data, dict):
        data=data.get("operations", [])
    if not isinstance(data, list):
        raise ValueError("Workflow JSON must contain a list or {operations:[...]}")
    return data


def validate_workflow(operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out=[]
    allowed={"format","borders","merge","unmerge","print-setup","pdf"}
    for op in operations:
        if not isinstance(op, dict):
            raise ValueError("Each workflow operation must be an object")
        kind=str(op.get("kind", "")).lower()
        if kind not in allowed:
            raise ValueError(f"Unsupported workflow operation: {kind}")
        if kind != "pdf" and not op.get("sheet"):
            raise ValueError(f"Operation {kind} requires sheet")
        if kind in {"format","borders","merge","unmerge"}:
            if not op.get("ranges"):
                raise ValueError(f"Operation {kind} requires ranges")
            # Validate ranges without expanding; supports cells, ranges, rows and columns.
            for token in op["ranges"]:
                parse_target_specs(str(token))
        out.append(op)
    return out


def run_workflow(source: str | Path, operations: list[dict[str, Any]], output: str | Path, *, backup: bool=True, visible: bool=False) -> dict[str, Any]:
    ops=validate_workflow(operations)
    return ExcelOperations().workflow(source, ops, output=output, backup=backup, visible=visible)
