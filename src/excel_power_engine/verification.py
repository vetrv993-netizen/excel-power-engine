"""Verification gate shared by all staged workbook modifications."""
from __future__ import annotations

import hashlib
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any
from zipfile import ZipFile

from .formula_intelligence import formula_issues


@dataclass(slots=True)
class VerificationResult:
    status: str
    output_exists: bool
    target_values_verified: bool
    formula_errors_before: list[dict[str, Any]] = field(default_factory=list)
    formula_errors_after: list[dict[str, Any]] = field(default_factory=list)
    new_formula_errors: list[dict[str, Any]] = field(default_factory=list)
    vba_preserved: bool | None = None
    x14_preserved: bool | None = None
    changed_parts: list[str] = field(default_factory=list)
    unexpected_changed_parts: list[str] = field(default_factory=list)
    visual_policy: str = "optional"
    visual_verification: dict[str, Any] = field(default_factory=dict)
    warnings: list[str] = field(default_factory=list)

    @property
    def accepted(self) -> bool:
        return self.status == "PASS"

    def as_dict(self) -> dict[str, Any]:
        data = asdict(self)
        data["accepted"] = self.accepted
        return data


def _parts(path: Path) -> dict[str, bytes]:
    with ZipFile(path, "r") as archive:
        return {name: archive.read(name) for name in archive.namelist()}


def _x14(parts: dict[str, bytes]) -> bool:
    return any(b"x14" in content or b"x15" in content for content in parts.values())


def verify_workbook_change(before: str | Path, after: str | Path, *, expected_cells: list[tuple[str, str, Any]] | None = None,
                           allowed_changed_parts: set[str] | None = None, visual_policy: str = "optional") -> VerificationResult:
    source, output = Path(before), Path(after)
    before_issues = formula_issues(source)
    if not output.exists():
        return VerificationResult("FAILED", False, False, formula_errors_before=before_issues,
                                  warnings=["Output workbook does not exist."], visual_policy=visual_policy)
    try:
        old_parts, new_parts = _parts(source), _parts(output)
    except Exception as exc:
        return VerificationResult("FAILED", True, False, formula_errors_before=before_issues,
                                  warnings=[f"Cannot read output OOXML package: {exc}"], visual_policy=visual_policy)
    changed = sorted(name for name in set(old_parts) | set(new_parts) if old_parts.get(name) != new_parts.get(name))
    allowed = allowed_changed_parts or set()
    unexpected = [part for part in changed if allowed and part not in allowed]
    vba_before = old_parts.get("xl/vbaProject.bin")
    vba_ok = None if vba_before is None else new_parts.get("xl/vbaProject.bin") == vba_before
    x14_before = _x14(old_parts)
    x14_ok = None if not x14_before else _x14(new_parts)
    target_ok = True
    if expected_cells:
        from .formula_intelligence import cell_info
        for sheet, cell, value in expected_cells:
            snap = cell_info(output, sheet, cell)
            observed = snap.formula if snap.formula is not None else snap.value
            if isinstance(observed, str) and isinstance(value, str) and observed.lstrip("=") == value.lstrip("="):
                continue
            if observed != value:
                target_ok = False
                break
    after_issues = formula_issues(output)
    before_keys = {(x.get("cell"), x.get("kind"), x.get("formula"), x.get("cached_value")) for x in before_issues}
    new_issues = [x for x in after_issues if (x.get("cell"), x.get("kind"), x.get("formula"), x.get("cached_value")) not in before_keys]
    failed = not target_ok or vba_ok is False or x14_ok is False or bool(unexpected) or bool(new_issues)
    warnings: list[str] = []
    if unexpected: warnings.append("Unexpected OOXML parts changed: " + ", ".join(unexpected))
    if new_issues: warnings.append(f"New formula errors: {len(new_issues)}")
    return VerificationResult("FAILED" if failed else "PASS", True, target_ok, before_issues, after_issues, new_issues,
                              vba_ok, x14_ok, changed, unexpected, visual_policy, {}, warnings)
