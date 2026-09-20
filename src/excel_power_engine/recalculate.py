from __future__ import annotations

from pathlib import Path
from typing import Any

from .native_excel import NativeExcelBridge
from .smart_engine import build_plan


def smart_recalculate(source: str | Path, output: str | Path | None = None, *, create_backup: bool = True, full_rebuild: bool = True, visible: bool = False) -> dict[str, Any]:
    plan = build_plan(source, "recalculate", recalculate=True)
    if plan.engine != "excel-native":
        return {
            "source": str(Path(source)),
            "output": str(Path(output)) if output else None,
            "executed_engine": plan.engine,
            "smart_plan": plan.as_dict(),
            "errors": ["Excel Native غير متاح؛ لم يتم تنفيذ إعادة الحساب الفعلية."],
        }
    result = NativeExcelBridge().recalculate_file(
        source, output,
        create_backup=create_backup,
        full_rebuild=full_rebuild,
        visible=visible,
    )
    result["smart_plan"] = plan.as_dict()
    result["executed_engine"] = "excel-native"
    return result
