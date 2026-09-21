from __future__ import annotations

import json
from pathlib import Path
from typing import Any


def save_workflow(path: str | Path, operations: list[dict[str, Any]], name: str = "workflow") -> Path:
    p = Path(path).resolve()
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(json.dumps({"name": name, "operations": operations}, ensure_ascii=False, indent=2), encoding="utf-8")
    return p
