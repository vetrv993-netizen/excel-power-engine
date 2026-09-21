from __future__ import annotations

import re
from dataclasses import dataclass, asdict
from difflib import SequenceMatcher
from typing import Any

from .workbook_context import WorkbookContext, SheetContext

_ARABIC_DIACRITICS = re.compile(r"[\u0610-\u061A\u064B-\u065F\u0670\u06D6-\u06ED]")


def normalize_text(value: str) -> str:
    text = str(value or "").strip().casefold()
    text = _ARABIC_DIACRITICS.sub("", text)
    text = text.replace("ـ", "")
    text = text.translate(str.maketrans({
        "أ": "ا", "إ": "ا", "آ": "ا", "ٱ": "ا",
        "ى": "ي", "ؤ": "و", "ئ": "ي",
    }))
    text = re.sub(r"[\u200f\u200e]", "", text)
    text = re.sub(r"[^\w\s]+", " ", text, flags=re.UNICODE)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _tokens(value: str) -> set[str]:
    return {x for x in normalize_text(value).split() if len(x) > 1}


def _aliases(title: str) -> set[str]:
    n = normalize_text(title)
    aliases = {n}
    tokens = _tokens(title)
    if "ال" in tokens:
        pass
    aliases.update(tokens)
    aliases.update({t.removeprefix("ال") for t in tokens if t.startswith("ال") and len(t) > 3})
    # Common Arabic plural/singular hints without pretending to be a language model.
    extras = {
        "طلاب": "طالب", "الطلاب": "طالب", "اسماء": "اسم", "اسماء الطلاب": "اسم الطالب", "اسم الطلاب": "اسم الطالب", "الاسم": "اسم الطالب",
        "درجات": "درجة", "الدرجات": "درجة", "درجة": "درجات", "معلمات": "معلمة", "معلمين": "معلم",
    }
    if n in extras:
        aliases.add(normalize_text(extras[n]))
    return aliases


@dataclass(frozen=True)
class SemanticCandidate:
    sheet: str
    column: str
    title: str
    score: float
    kind: str = "column"
    data_range: str = ""

    def as_dict(self) -> dict[str, Any]:
        return asdict(self)


@dataclass(frozen=True)
class SemanticResolution:
    query: str
    resolved: bool
    ambiguous: bool
    candidates: list[SemanticCandidate]
    reason: str = ""

    @property
    def best(self) -> SemanticCandidate | None:
        return self.candidates[0] if self.candidates else None

    def as_dict(self) -> dict[str, Any]:
        return {
            "query": self.query,
            "resolved": self.resolved,
            "ambiguous": self.ambiguous,
            "reason": self.reason,
            "candidates": [c.as_dict() for c in self.candidates],
        }


class SemanticWorkbook:
    """Small deterministic semantic layer over the v0.10 WorkbookContext.

    It intentionally resolves against real headers/ranges discovered from the workbook;
    it does not invent columns that are not present in the workbook.
    """

    def __init__(self, context: WorkbookContext):
        self.context = context

    def choices(self, sheet: str | None = None) -> list[dict[str, Any]]:
        result: list[dict[str, Any]] = []
        sheets = [self.context.sheet(sheet)] if sheet else self.context.sheets
        for sh in sheets:
            if sh is None:
                continue
            for col in sh.columns:
                result.append({
                    "sheet": sh.name,
                    "column": col["column"],
                    "title": col["title"],
                    "kind": col["kind"],
                    "range": f'{col["column"]}{sh.data_start_row or 1}:{col["column"]}{sh.data_end_row or (sh.data_start_row or 1)}',
                    "samples": col.get("samples", []),
                })
        return result

    def resolve_column(self, query: str, sheet: str | None = None, *, ambiguity_margin: float = 0.08) -> SemanticResolution:
        q = normalize_text(re.sub(r"^(عمود|العمود|حقل|الـعمود)\s+", "", str(query or "")))
        if not q:
            return SemanticResolution(str(query), False, False, [], "query-empty")
        candidates: list[SemanticCandidate] = []
        query_tokens = _tokens(q)
        sheets = [self.context.sheet(sheet)] if sheet else self.context.sheets
        for sh in sheets:
            if sh is None:
                continue
            for col in sh.columns:
                title = str(col.get("title", ""))
                cn = normalize_text(title)
                aliases = _aliases(title)
                score = 0.0
                if q == cn:
                    score = 1.0
                elif q in aliases:
                    score = 0.96
                elif cn in q:
                    score = 0.90
                elif q in cn:
                    score = 0.88
                else:
                    token_overlap = len(query_tokens & aliases)
                    if query_tokens and token_overlap:
                        score = 0.65 + min(0.25, token_overlap / max(len(query_tokens), 1) * 0.25)
                    for qt in query_tokens:
                        for alias in aliases:
                            if qt == alias or qt in alias or alias in qt:
                                score = max(score, 0.74)
                            else:
                                ratio = SequenceMatcher(None, qt, alias).ratio()
                                if ratio >= 0.72:
                                    score = max(score, 0.68 + (ratio - 0.72) * 0.5)
                    score = max(score, SequenceMatcher(None, q, cn).ratio() * 0.82)
                if score >= 0.56:
                    data_start = sh.data_start_row or 1
                    data_end = sh.data_end_row or data_start
                    candidates.append(SemanticCandidate(sh.name, col["column"], title, round(score, 3), "column", f'{col["column"]}{data_start}:{col["column"]}{data_end}'))
        candidates.sort(key=lambda c: (c.score, len(c.title)), reverse=True)
        best = candidates[0].score if candidates else 0.0
        ambiguous = bool(len(candidates) > 1 and best - candidates[1].score < ambiguity_margin)
        return SemanticResolution(str(query), bool(candidates) and not ambiguous, ambiguous, candidates[:8], "ambiguous" if ambiguous else ("resolved" if candidates else "not-found"))

    def resolve_target(self, target_text: str, sheet: str | None = None) -> dict[str, Any]:
        raw = str(target_text or "").strip()
        if not raw:
            return {"resolved": False, "reason": "target-empty"}
        resolution = self.resolve_column(raw, sheet)
        if resolution.resolved and resolution.best:
            best = resolution.best
            return {
                "resolved": True,
                "ambiguous": False,
                "target_type": "semantic_column",
                "sheet": best.sheet,
                "column": best.column,
                "title": best.title,
                "range": best.data_range,
                "score": best.score,
            }
        if resolution.ambiguous:
            return {"resolved": False, "ambiguous": True, "reason": "ambiguous", "candidates": [c.as_dict() for c in resolution.candidates]}
        # Direct A1/range fallback is intentionally conservative and requires an existing sheet.
        if re.fullmatch(r"[A-Za-z]+\d+(?::[A-Za-z]+\d+)?", raw) or re.fullmatch(r"[A-Za-z]+(?::[A-Za-z]+)?|\d+(?::\d+)?", raw):
            return {"resolved": True, "ambiguous": False, "target_type": "direct", "sheet": sheet, "range": raw.upper(), "score": 1.0}
        return {"resolved": False, "ambiguous": False, "reason": "not-found", "candidates": [c.as_dict() for c in resolution.candidates]}
