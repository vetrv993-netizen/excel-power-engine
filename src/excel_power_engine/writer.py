from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
import shutil


def save_with_openpyxl(workbook, path: str | Path) -> None:
    workbook.save(path)


def create_new_xlsx_from_dataframe(df, path: str | Path, sheet_name: str = "Sheet1") -> None:
    import pandas as pd
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)


def surgical_replace_part(source: str | Path, output: str | Path, part_name: str, new_bytes: bytes) -> None:
    source = Path(source)
    output = Path(output)
    if source.resolve() == output.resolve():
        raise ValueError("Source and output must differ")
    if output.exists():
        output.unlink()
    with ZipFile(source, "r") as zin, ZipFile(output, "w", ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = new_bytes if info.filename == part_name else zin.read(info.filename)
            zout.writestr(info, data)
