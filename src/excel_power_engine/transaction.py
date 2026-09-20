from __future__ import annotations

from dataclasses import dataclass, field, asdict
from pathlib import Path
from typing import Any


@dataclass
class Operation:
    kind: str
    sheet: str | None = None
    target: str | None = None
    payload: dict[str, Any] = field(default_factory=dict)

    def as_dict(self):
        return asdict(self)


@dataclass
class TransactionPlan:
    source: str
    operations: list[Operation]
    output: str
    backup: bool = True

    def as_dict(self):
        return {"source": self.source, "output": self.output, "backup": self.backup, "operations": [x.as_dict() for x in self.operations]}


def build_plan(source: str | Path, output: str | Path, operations: list[Operation], *, backup: bool = True) -> TransactionPlan:
    if not operations:
        raise ValueError("Transaction must contain at least one operation")
    return TransactionPlan(str(Path(source).resolve()), operations, str(Path(output).resolve()), backup)
