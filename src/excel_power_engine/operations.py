from __future__ import annotations

from dataclasses import dataclass, asdict
from typing import Any


@dataclass
class OperationPreview:
    kind: str
    target: str
    description: str
    requires_excel: bool = True
    destructive: bool = False

    def as_dict(self):
        return asdict(self)


def preview_operations(operations: list[dict[str, Any]]) -> list[dict[str, Any]]:
    out=[]
    for op in operations:
        kind=str(op.get("kind", "")).lower()
        target=op.get("sheet") or "workbook"
        ranges=",".join(map(str, op.get("ranges", [])))
        desc={
            "format":"تطبيق تنسيق على نطاقات محددة",
            "borders":"تطبيق حدود على نطاقات محددة",
            "merge":"دمج النطاقات المحددة",
            "unmerge":"إلغاء دمج النطاقات المحددة",
            "print-setup":"تعديل إعدادات الطباعة",
            "pdf":"تصدير PDF",
        }.get(kind, kind)
        out.append(OperationPreview(kind, f"{target}:{ranges}" if ranges else str(target), desc, True, kind in {"merge","unmerge"}).as_dict())
    return out
