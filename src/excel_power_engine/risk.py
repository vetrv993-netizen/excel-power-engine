from __future__ import annotations

SAFE = "SAFE"
WARNING = "WARNING"
HIGH_RISK = "HIGH_RISK"


def classify_operation(kind: str, payload: dict | None = None) -> str:
    k = str(kind or "").lower()
    if k in {"format", "borders", "column-width", "row-height", "alignment", "print-setup"}:
        return SAFE
    if k in {"visibility", "merge", "unmerge", "bulk", "keyed-bulk", "search"}:
        return WARNING
    if k in {"delete-row", "delete-column", "delete-sheet", "vba", "wide-workbook-edit"}:
        return HIGH_RISK
    if k == "pdf":
        return SAFE
    return WARNING
