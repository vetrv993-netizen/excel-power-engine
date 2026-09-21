from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any

from .risk import classify_operation


@dataclass
class OperationPreview:
    kind: str
    target: str
    description: str
    requires_excel: bool = True
    destructive: bool = False
    risk: str = "SAFE"

    def as_dict(self):
        return asdict(self)


def preview_operations(operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out=[]
    descriptions={
        "format":"تطبيق تنسيق على نطاقات محددة",
        "borders":"تطبيق حدود على نطاقات محددة",
        "merge":"دمج النطاقات المحددة",
        "unmerge":"إلغاء دمج النطاقات المحددة",
        "print-setup":"تعديل إعدادات الطباعة",
        "pdf":"تصدير PDF",
        "column-width":"ضبط عرض الأعمدة",
        "row-height":"ضبط ارتفاع الصفوف",
        "visibility":"إخفاء/إظهار صف أو عمود",
        "alignment":"تعديل المحاذاة والتفاف النص",
    }
    for op in operations:
        kind=str(op.get("kind", "")).lower()
        target=op.get("sheet") or "workbook"
        ranges=",".join(map(str, op.get("ranges", [])))
        risk=classify_operation(kind, op.get("payload") or {})
        out.append(OperationPreview(kind, f"{target}:{ranges}" if ranges else str(target), descriptions.get(kind, kind), kind not in {"search"}, kind in {"merge","unmerge"}, risk).as_dict())
    return out
