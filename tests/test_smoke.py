from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from excel_power_engine.inspect import inspect_workbook
from excel_power_engine.ooxml import list_parts


def make_xlsx(path: Path) -> None:
    import openpyxl
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data"
    ws["A1"] = "hello"
    wb.save(path)


def test_inspection(tmp_path):
    p = tmp_path / "sample.xlsx"
    make_xlsx(p)
    profile = inspect_workbook(p)
    assert profile.is_ooxml
    assert "Data" in profile.sheet_parts
    assert len(list_parts(p)) > 0
