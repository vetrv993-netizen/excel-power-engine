from __future__ import annotations

from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED


def list_parts(path: str | Path) -> list[str]:
    with ZipFile(path) as zf:
        return zf.namelist()


def read_part(path: str | Path, part_name: str) -> bytes:
    with ZipFile(path) as zf:
        return zf.read(part_name)


def write_part(source: str | Path, output: str | Path, part_name: str, data: bytes) -> None:
    source = Path(source)
    output = Path(output)
    if source.resolve() == output.resolve():
        raise ValueError("source and output must differ")
    with ZipFile(source, "r") as zin, ZipFile(output, "w", ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, data if info.filename == part_name else zin.read(info.filename))
