from __future__ import annotations

from typing import Any

from .workbook_context import WorkbookContext


def smart_suggestions(context: WorkbookContext, sheet: str | None = None) -> list[dict[str, Any]]:
    sh = context.sheet(sheet or context.selected_sheet)
    if sh is None:
        return []
    out: list[dict[str, Any]] = []
    if sh.header_row:
        out.append({"id": "header", "title": "تحسين صف العناوين", "reason": f"تم اكتشاف صف العناوين رقم {sh.header_row}.", "action": "format_bold", "preview": True})
    for col in sh.columns:
        title = str(col.get("title", ""))
        if col.get("kind") == "text" and len(title) >= 4:
            out.append({"id": f"width-{col['column']}", "title": f"ضبط عرض {title}", "reason": "هذا العمود نصي وقد يحتاج عرضًا أوضح.", "action": "column_width", "target": title, "preview": True})
    if sh.value_cells and sh.data_end_row and sh.data_start_row and sh.data_end_row > sh.data_start_row + 1:
        out.append({"id": "table-borders", "title": "إضافة حدود لمنطقة البيانات", "reason": f"البيانات تمتد إلى الصف {sh.data_end_row}.", "action": "borders", "target": sh.used_range, "preview": True})
    out.append({"id": "print", "title": "مراجعة إعدادات الطباعة", "reason": "تحقق من الاتجاه والملاءمة قبل التصدير.", "action": "print-review", "preview": True})
    return out[:12]
