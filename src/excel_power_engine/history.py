from __future__ import annotations

import json
import time
from pathlib import Path
from typing import Any


def default_history_path() -> Path:
    root = Path.home() / ".excel_power_engine"
    root.mkdir(parents=True, exist_ok=True)
    return root / "operation_history.jsonl"


class OperationHistory:
    def __init__(self, path: str | Path | None = None):
        self.path = Path(path).expanduser().resolve() if path else default_history_path()
        self.path.parent.mkdir(parents=True, exist_ok=True)

    def add(self, command: str, *, status: str, sheet: str | None = None, ranges: list[str] | None = None, details: dict[str, Any] | None = None) -> dict[str, Any]:
        entry = {"time": time.strftime("%Y-%m-%d %H:%M:%S"), "command": command, "status": status, "sheet": sheet, "ranges": ranges or [], "details": details or {}}
        with self.path.open("a", encoding="utf-8") as fh:
            fh.write(json.dumps(entry, ensure_ascii=False, default=str) + "\n")
        return entry

    def recent(self, limit: int = 50) -> list[dict[str, Any]]:
        if not self.path.exists():
            return []
        lines = self.path.read_text(encoding="utf-8", errors="replace").splitlines()
        result = []
        for line in lines[-limit:]:
            try:
                result.append(json.loads(line))
            except Exception:
                continue
        return list(reversed(result))
