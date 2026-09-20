from __future__ import annotations

from pathlib import Path
from typing import Any


def read_with_openpyxl(path: str | Path, sheet: str | None = None, data_only: bool = False):
    from openpyxl import load_workbook

    wb = load_workbook(path, data_only=data_only, read_only=False, keep_vba=True)
    if sheet:
        return wb[sheet]
    return wb


def read_dataframe(path: str | Path, sheet: str = "Sheet1", engine: str = "auto"):
    import pandas as pd
    p = Path(path)
    if engine == "openpyxl" or (engine == "auto" and p.suffix.lower() in {".xlsx", ".xlsm"}):
        return pd.read_excel(p, sheet_name=sheet, engine="openpyxl")
    return pd.read_excel(p, sheet_name=sheet, engine=engine)


def read_fast(path: str | Path, sheet_name: str | None = None):
    try:
        import fastexcel
    except ImportError as exc:
        raise RuntimeError("fastexcel is not installed. Install with: pip install -e \".[fast]\"") from exc
    wb = fastexcel.read_excel(path)
    if sheet_name:
        return wb.load_sheet(sheet_name)
    return wb
