from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from typing import Any

from .semantic_engine import SemanticWorkbook, normalize_text

COLORS = {
    "احمر": "#FF0000", "أحمر": "#FF0000", "الاحمر": "#FF0000", "الأحمر": "#FF0000", "red": "#FF0000",
    "اخضر": "#00A651", "أخضر": "#00A651", "الاخضر": "#00A651", "الأخضر": "#00A651", "green": "#00A651",
    "ازرق": "#0070C0", "أزرق": "#0070C0", "الازرق": "#0070C0", "الأزرق": "#0070C0", "blue": "#0070C0",
    "اصفر": "#FFFF00", "أصفر": "#FFFF00", "الاصفر": "#FFFF00", "الأصفر": "#FFFF00", "yellow": "#FFFF00",
    "برتقالي": "#F4B183", "بنفسجي": "#7030A0", "ابيض": "#FFFFFF", "أبيض": "#FFFFFF", "اسود": "#000000", "أسود": "#000000",
    "رمادي": "#D9E1F2", "وردي": "#F4CCCC",
}

COLOR_NAMES = {normalize_text(k): v for k, v in COLORS.items()}


@dataclass(frozen=True)
class CommandIntent:
    action: str
    target_text: str | None = None
    parameters: dict[str, Any] | None = None
    confidence: float = 0.0
    source_text: str = ""

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class CommandPlan:
    command: str
    intent: CommandIntent | None
    target: dict[str, Any]
    operations: list[dict[str, Any]]
    risk: str
    status: str
    explanation: str
    warnings: list[str]

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


def _extract_target(text: str, prefixes: list[str]) -> str:
    work = str(text).strip()
    for prefix in prefixes:
        work = re.sub(rf"^\s*{re.escape(prefix)}\s+", "", work, flags=re.IGNORECASE)
    work = re.sub(r"^(هذا|هذه|ذا|ذي)\s+", "", work)
    return work.strip(" .،")


def parse_command(command: str) -> CommandIntent | None:
    text = str(command or "").strip()
    n = normalize_text(text)
    if not n:
        return None

    color_match = re.search(r"(?:لون|لوّن|لَوِّن)\s+(?:عمود\s+|العمود\s+|الخلايا\s+|النطاق\s+)?(.+?)\s+(?:(?:باللون|باللون)\s+)?(الأحمر|الاحمر|أحمر|احمر|الأخضر|الاخضر|أخضر|اخضر|الأزرق|الازرق|أزرق|ازرق|الأصفر|الاصفر|أصفر|اصفر|برتقالي|بنفسجي|ابيض|أبيض|اسود|أسود|رمادي|وردي|[\wأ-ي]+)$", text, flags=re.IGNORECASE)
    if color_match:
        color_word = normalize_text(color_match.group(2))
        if color_word in COLOR_NAMES:
            target = _extract_target(color_match.group(1), ["عمود", "العمود", "الخلايا", "النطاق"])
            return CommandIntent("format_fill", target, {"fill": COLOR_NAMES[color_word], "color_name": color_match.group(2)}, 0.94, text)

    attached_color = re.search(r"(?:لون|لوّن|لَوِّن)\s+(?:عمود\s+|العمود\s+)?(.+?)\s+بال(احمر|أحمر|الاحمر|الأحمر|اخضر|أخضر|الاخضر|الأخضر|ازرق|أزرق|الازرق|الأزرق|اصفر|أصفر|الاصفر|الأصفر)$", text, flags=re.IGNORECASE)
    if attached_color:
        color_word = normalize_text(attached_color.group(2))
        if color_word in COLOR_NAMES:
            return CommandIntent("format_fill", _extract_target(attached_color.group(1), ["عمود", "العمود"]), {"fill": COLOR_NAMES[color_word], "color_name": attached_color.group(2)}, 0.94, text)

    m = re.search(r"(?:اجعل|اجعل|حوّل)\s+(.+?)\s+(?:بالخط\s+)?العريض$", text)
    if m:
        return CommandIntent("format_bold", _extract_target(m.group(1), []), {"bold": True}, 0.93, text)

    m = re.search(r"(?:وسّع|وسع)\s+(?:عمود\s+|العمود\s+)?(.+)$", text)
    if m and "الطباعة" not in n:
        return CommandIntent("column_width", _extract_target(m.group(1), ["عمود", "العمود"]), {"width": 20}, 0.90, text)

    m = re.search(r"(?:أخف|اخف)\s+(?:عمود\s+|العمود\s+)?(.+)$", text)
    if m:
        return CommandIntent("hide_column", _extract_target(m.group(1), ["عمود", "العمود"]), {"hidden": True}, 0.90, text)

    m = re.search(r"(?:أظهر|اظهر)\s+(?:عمود\s+|العمود\s+)?(.+)$", text)
    if m:
        return CommandIntent("show_column", _extract_target(m.group(1), ["عمود", "العمود"]), {"hidden": False}, 0.90, text)

    if re.search(r"(?:حدود|حدودا|حدوداً)", n) and re.search(r"(?:جدول|نطاق|الخلايا|هذا)", n):
        target_match = re.search(r"(?:حول|على|ل)\s+(.+?)(?:$)", text)
        target = _extract_target(target_match.group(1) if target_match else "الجدول", [])
        return CommandIntent("borders", target, {"line_style": 1, "weight": 2}, 0.91, text)

    m = re.search(r"(?:محاذاة|حاذِ)\s*(?:وسط|توسط)\s*(.+)?$", text)
    if m:
        return CommandIntent("alignment", _extract_target(m.group(1) or "", ["هذا"]), {"horizontal": "center"}, 0.88, text)

    m = re.search(r"(?:التفاف\s+النص|لف\s+النص)\s*(.+)?$", text)
    if m:
        return CommandIntent("wrap_text", _extract_target(m.group(1) or "", ["هذا"]), {"wrap": True}, 0.88, text)

    if "الطباعة" in n and ("افقي" in n or "افقية" in n or "أفقية" in n):
        return CommandIntent("print_landscape", None, {"orientation": "landscape"}, 0.96, text)

    if ("يتناسب" in n and "صفحه واحده" in n) or "صفحة واحدة" in text or "صفحه واحده" in n:
        return CommandIntent("print_fit_one_page", None, {"fit_width": 1, "fit_height": 1}, 0.90, text)

    if re.search(r"(?:صدّر|صدر|تصدير).*pdf", text, flags=re.IGNORECASE):
        return CommandIntent("export_pdf", None, {}, 0.94, text)

    m = re.search(r"(?:غيّر|غير)\s+(?:حجم\s+)?الخط\s+(?:إلى|الى)\s*(\d+(?:\.\d+)?)", text)
    if m:
        return CommandIntent("font_size", None, {"font_size": float(m.group(1))}, 0.94, text)

    m = re.search(r"(?:غيّر|غير)\s+لون\s+الخط\s+(?:إلى|الى)?\s*(\w+)", text)
    if m:
        c = COLOR_NAMES.get(normalize_text(m.group(1)))
        if c:
            return CommandIntent("font_color", None, {"font_color": c, "color_name": m.group(1)}, 0.91, text)

    if re.search(r"^\s*ادمج\s+", text):
        target = _extract_target(text, ["ادمج"])
        return CommandIntent("merge", target, {}, 0.88, text)
    if re.search(r"^\s*(?:الغ|الغِ|الغاء)\s+الدمج", text) or "الغ الدمج" in n:
        target = re.sub(r"^\s*(?:الغ|الغِ|الغاء)\s+الدمج\s*", "", text).strip() or None
        return CommandIntent("unmerge", target, {}, 0.88, text)

    if "الخلايا الفارغه" in n or "الخلايا الفارغة" in text:
        return CommandIntent("find_empty_cells", None, {}, 0.84, text)

    if re.search(r"ابحث\s+عن", text):
        return CommandIntent("search", None, {"term": re.sub(r"^.*?ابحث\s+عن\s*", "", text)}, 0.82, text)

    if re.search(r"اجعل\s+الصف\s+الأول\s+عنوان", text):
        return CommandIntent("set_header_row", None, {"row": 1}, 0.89, text)

    return CommandIntent("unknown", None, {}, 0.30, text)


def build_command_plan(command: str, context: Any, *, sheet: str | None = None, explicit_target: str | None = None) -> CommandPlan:
    intent = parse_command(command)
    if intent is None:
        return CommandPlan(command, None, {}, [], "SAFE", "REJECTED", "لم يتم إدخال أمر.", ["أدخل أمرًا واضحًا."])
    semantic = SemanticWorkbook(context)
    target_text = explicit_target or intent.target_text
    target: dict[str, Any] = {"resolved": True, "sheet": sheet or context.selected_sheet}
    if intent.action == "borders" and target_text and normalize_text(target_text).startswith("جدول"):
        selected_sheet = context.sheet(sheet or context.selected_sheet)
        if selected_sheet and selected_sheet.used_range:
            target = {"resolved": True, "ambiguous": False, "target_type": "table", "sheet": selected_sheet.name, "range": selected_sheet.used_range, "title": "جدول البيانات", "score": 0.80}
    warnings: list[str] = []

    if intent.action in {"format_fill", "format_bold", "column_width", "hide_column", "show_column", "borders", "merge", "unmerge", "alignment", "wrap_text"}:
        if not target_text:
            return CommandPlan(command, intent, {}, [], "SAFE", "NEEDS_TARGET", "حدد الهدف من العملية.", ["لم يتم تحديد عمود/نطاق."])
        if not (intent.action == "borders" and target.get("range")):
            target = semantic.resolve_target(target_text, sheet or context.selected_sheet)
        if not target.get("resolved"):
            candidates = target.get("candidates", [])
            reason = "وجد النظام أكثر من هدف محتمل." if target.get("ambiguous") else "لم يتم العثور على الهدف داخل المصنف."
            if candidates:
                warnings.append("المرشحون: " + "، ".join(f'{c["title"]} ({c["column"]})' for c in candidates[:5]))
            return CommandPlan(command, intent, target, [], "SAFE", "AMBIGUOUS" if target.get("ambiguous") else "NOT_FOUND", reason, warnings)

    resolved_sheet = target.get("sheet") or sheet or context.selected_sheet
    rng = target.get("range")
    payload: dict[str, Any] = dict(intent.parameters or {})
    operations: list[dict[str, Any]] = []

    if intent.action == "format_fill":
        operations = [{"kind": "format", "sheet": resolved_sheet, "ranges": [rng], "payload": {"fill": payload["fill"]}}]
        explanation = f"تعبئة {target.get('title', rng)} باللون {payload.get('color_name', payload['fill'])}."
        risk = "SAFE"
    elif intent.action == "format_bold":
        operations = [{"kind": "format", "sheet": resolved_sheet, "ranges": [rng], "payload": {"bold": True}}]
        explanation = f"جعل {target.get('title', rng)} بالخط العريض."
        risk = "SAFE"
    elif intent.action == "column_width":
        operations = [{"kind": "column-width", "sheet": resolved_sheet, "ranges": [target.get("column") or rng], "payload": {"width": payload.get("width", 20)}}]
        explanation = f"ضبط عرض العمود {target.get('title', rng)}."
        risk = "SAFE"
    elif intent.action in {"hide_column", "show_column"}:
        operations = [{"kind": "visibility", "sheet": resolved_sheet, "ranges": [target.get("column") or rng], "payload": {"axis": "column", "hidden": payload.get("hidden", True)}}]
        explanation = ("إخفاء" if payload.get("hidden", True) else "إظهار") + f" {target.get('title', rng)}."
        risk = "WARNING"
    elif intent.action == "borders":
        operations = [{"kind": "borders", "sheet": resolved_sheet, "ranges": [rng], "payload": payload}]
        explanation = f"إضافة حدود إلى {target.get('title', rng)}."
        risk = "SAFE"
    elif intent.action in {"merge", "unmerge"}:
        operations = [{"kind": intent.action, "sheet": resolved_sheet, "ranges": [rng], "payload": payload}]
        explanation = ("دمج" if intent.action == "merge" else "إلغاء دمج") + f" {rng}."
        risk = "WARNING"
    elif intent.action in {"alignment", "wrap_text"}:
        operations = [{"kind": "alignment", "sheet": resolved_sheet, "ranges": [rng], "payload": payload}]
        explanation = "تعديل المحاذاة/التفاف النص."
        risk = "SAFE"
    elif intent.action == "print_landscape":
        operations = [{"kind": "print-setup", "sheet": resolved_sheet, "ranges": [], "payload": {"orientation": "landscape"}}]
        explanation = "ضبط اتجاه الطباعة إلى أفقي."
        risk = "SAFE"
    elif intent.action == "print_fit_one_page":
        operations = [{"kind": "print-setup", "sheet": resolved_sheet, "ranges": [], "payload": {"fit_width": 1, "fit_height": 1}}]
        explanation = "ضبط الطباعة لتناسب صفحة واحدة."
        risk = "SAFE"
    elif intent.action == "export_pdf":
        operations = [{"kind": "pdf", "sheet": resolved_sheet, "ranges": [], "payload": {"output_pdf": str(context.path.with_suffix(".pdf"))}}]
        explanation = "تصدير الورقة الحالية إلى PDF."
        risk = "SAFE"
    elif intent.action in {"font_size", "font_color"}:
        operations = [{"kind": "format", "sheet": resolved_sheet, "ranges": [rng] if rng else [context.suggestions.get("target", "A1")], "payload": payload}]
        explanation = "تعديل تنسيق الخط."
        risk = "SAFE"
    elif intent.action == "search":
        operations = [{"kind": "search", "sheet": resolved_sheet, "payload": {"term": payload.get("term", ""), "result_id": "command-search"}}]
        explanation = "البحث داخل المصنف عن النص المطلوب."
        risk = "SAFE"
    elif intent.action in {"find_empty_cells", "set_header_row"}:
        return CommandPlan(command, intent, target, [], "SAFE", "PLANNED", "تم فهم الأمر، لكن التنفيذ المباشر لهذه العملية يحتاج مسارًا تحليليًا مخصصًا.", ["يمكن عرض النتائج أو تحويلها إلى TargetSet أولاً."])
    else:
        return CommandPlan(command, intent, target, [], "SAFE", "UNSUPPORTED", "فهمت بعض الكلمات، لكن هذا الأمر غير مدعوم للتنفيذ الآلي بعد.", ["استخدم أدوات التعديل اليدوي أو عدّل صياغة الأمر."])

    return CommandPlan(command, intent, target, operations, risk, "READY", explanation, warnings)
