from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import json
import os


def default_audit_path() -> Path:
    base = Path(os.environ.get("LOCALAPPDATA") or Path.home()) / "ExcelPowerEngine"
    base.mkdir(parents=True, exist_ok=True)
    return base / "audit.log.jsonl"


def log_event(action: str, *, status: str = "ok", details: dict | None = None, path: str | None = None) -> None:
    event = {
        "timestamp": datetime.now(timezone.utc).isoformat(),
        "action": action,
        "status": status,
        "path": path,
        "details": details or {},
    }
    p = default_audit_path()
    with p.open("a", encoding="utf-8") as f:
        f.write(json.dumps(event, ensure_ascii=False) + "\n")


def read_recent(limit: int = 200) -> list[dict]:
    p = default_audit_path()
    if not p.exists():
        return []
    lines = p.read_text(encoding="utf-8").splitlines()[-limit:]
    out = []
    for line in reversed(lines):
        try:
            out.append(json.loads(line))
        except json.JSONDecodeError:
            continue
    return out
