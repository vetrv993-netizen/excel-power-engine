from __future__ import annotations
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

@dataclass(slots=True)
class WorkbookProfile:
    path: Path
    suffix: str
    is_ooxml: bool
    is_xlsm: bool
    has_vba: bool = False
    internal_parts: list[str] = field(default_factory=list)
    sheet_parts: dict[str, str] = field(default_factory=dict)
    features: set[str] = field(default_factory=set)
    part_sizes: dict[str, int] = field(default_factory=dict)
    sheet_profiles: dict[str, dict[str, Any]] = field(default_factory=dict)
    defined_names: list[str] = field(default_factory=list)
    integrity: dict[str, Any] = field(default_factory=dict)

@dataclass(slots=True)
class DiffItem:
    sheet: str
    cell: str
    before: Any
    after: Any
    kind: str = "value"

@dataclass(slots=True)
class SafeEditReport:
    source: Path
    output: Path
    backup: Path | None
    sheet: str
    edited_cells: list[str]
    changed_parts: list[str]
    protected_parts_unchanged: bool
    vba_preserved: bool
    x14_preserved: bool
    errors: list[str] = field(default_factory=list)
    def as_dict(self) -> dict[str, Any]:
        return {
            "source": str(self.source), "output": str(self.output),
            "backup": str(self.backup) if self.backup else None,
            "sheet": self.sheet, "edited_cells": self.edited_cells,
            "changed_parts": self.changed_parts,
            "protected_parts_unchanged": self.protected_parts_unchanged,
            "vba_preserved": self.vba_preserved, "x14_preserved": self.x14_preserved,
            "errors": self.errors,
        }
