"""Shared, read-first workbook state used by the CLI, GUI, and workspace runner."""
from __future__ import annotations

import hashlib
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from .formula_intelligence import formula_issues
from .workbook_context import WorkbookContext, analyze_workbook


def fingerprint(path: str | Path) -> str:
    h = hashlib.sha256()
    with Path(path).open("rb") as f:
        for block in iter(lambda: f.read(1024 * 1024), b""):
            h.update(block)
    return h.hexdigest()


@dataclass(slots=True)
class TargetSet:
    id: str
    targets: dict[str, list[str]] = field(default_factory=dict)
    provenance: str = "manual"
    source_result_set: str | None = None

    def add(self, sheet: str, cells: list[str]) -> None:
        existing = self.targets.setdefault(sheet, [])
        existing.extend(c.upper() for c in cells)
        self.targets[sheet] = list(dict.fromkeys(existing))

    @property
    def count(self) -> int:
        return sum(len(cells) for cells in self.targets.values())

    def first(self) -> tuple[str | None, str | None]:
        for sheet, cells in self.targets.items():
            if cells:
                return sheet, cells[0]
        return None, None

    def as_dict(self) -> dict[str, Any]:
        return {"id": self.id, "targets": self.targets, "count": self.count,
                "provenance": self.provenance, "source_result_set": self.source_result_set}


@dataclass(slots=True)
class SearchResultSet:
    id: str
    hits: list[dict[str, Any]]
    provenance: str = "search"

    def to_target_set(self, target_id: str | None = None, selected: list[int] | None = None) -> TargetSet:
        chosen = self.hits if selected is None else [self.hits[i] for i in selected]
        targets = TargetSet(target_id or f"{self.id}-targets", provenance=self.provenance, source_result_set=self.id)
        for hit in chosen:
            targets.add(str(hit["sheet"]), [str(hit["cell"])])
        return targets


@dataclass(slots=True)
class WorkbookSession:
    source_path: Path
    source_fingerprint: str
    workbook_context: WorkbookContext
    selected_sheet: str | None
    selected_targets: TargetSet = field(default_factory=lambda: TargetSet("selected"))
    suggestions: dict[str, Any] = field(default_factory=dict)
    formula_baseline: list[dict[str, Any]] = field(default_factory=list)
    sensitive_parts_baseline: dict[str, str] = field(default_factory=dict)
    transaction_manifest: dict[str, Any] = field(default_factory=dict)
    audit_context: dict[str, Any] = field(default_factory=dict)
    result_sets: dict[str, SearchResultSet] = field(default_factory=dict)

    @classmethod
    def open(cls, path: str | Path, *, selected_sheet: str | None = None, selected_cell: str | None = None) -> "WorkbookSession":
        source = Path(path).resolve()
        context = analyze_workbook(source, selected_sheet, selected_cell)
        return cls(source, fingerprint(source), context, context.selected_sheet,
                   suggestions=context.suggestions, formula_baseline=formula_issues(source))

    def ensure_current(self) -> None:
        if fingerprint(self.source_path) != self.source_fingerprint:
            raise RuntimeError("Workbook changed after session analysis; create a new WorkbookSession.")

    def register_search(self, hits: list[dict[str, Any]], result_id: str = "last-search") -> SearchResultSet:
        result = SearchResultSet(result_id, hits)
        self.result_sets[result_id] = result
        return result

    def select_search_hits(self, result_id: str = "last-search", selected: list[int] | None = None) -> TargetSet:
        targets = self.result_sets[result_id].to_target_set("selected", selected)
        self.selected_targets = targets
        return targets
