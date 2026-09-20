from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED

from excel_power_engine.native_excel import NativeExcelBridge


def make_formula_error_xlsx(path: Path):
    files = {
        "[Content_Types].xml": """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'>
<Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/>
<Default Extension='xml' ContentType='application/xml'/>
<Override PartName='/xl/workbook.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'/>
<Override PartName='/xl/worksheets/sheet1.xml' ContentType='application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'/>
</Types>""",
        "_rels/.rels": """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'></Relationships>""",
        "xl/workbook.xml": """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<workbook xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'><sheets><sheet name='Data' sheetId='1' r:id='rId1'/></sheets></workbook>""",
        "xl/_rels/workbook.xml.rels": """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet' Target='worksheets/sheet1.xml'/></Relationships>""",
        "xl/worksheets/sheet1.xml": """<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<worksheet xmlns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'><sheetData><row r='1'><c r='A1'><f>1/0</f><v t='e'>#DIV/0!</v></c><c r='B1'><f>NOPE()</f><v t='e'>#NAME?</v></c><c r='C1'><v>#REF!</v></c></row></sheetData></worksheet>""",
    }
    with ZipFile(path, "w", ZIP_DEFLATED) as zf:
        for name, text in files.items():
            zf.writestr(name, text)


def test_formula_scan_uses_cached_ooxml_values(tmp_path):
    p = tmp_path / "sample.xlsx"
    make_formula_error_xlsx(p)
    errors = NativeExcelBridge._scan_formula_errors(p)
    assert errors == [
        {"sheet": "Data", "cell": "A1", "error": "#DIV/0!"},
        {"sheet": "Data", "cell": "B1", "error": "#NAME?"},
    ]
