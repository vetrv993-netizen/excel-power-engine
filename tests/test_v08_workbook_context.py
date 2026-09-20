from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from excel_power_engine.workbook_context import analyze_workbook


def make_xlsx(path: Path):
    files = {
        '[Content_Types].xml': '''<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>''',
        'xl/workbook.xml': '''<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>''',
        'xl/_rels/workbook.xml.rels': '''<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>''',
        'xl/worksheets/sheet1.xml': '''<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Name</t></is></c><c r="B1" t="inlineStr"><is><t>Count</t></is></c><c r="C1" t="inlineStr"><is><t>Status</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Ahmed</t></is></c><c r="B2"><v>3</v></c><c r="C2" t="inlineStr"><is><t>Open</t></is></c></row><row r="3"><c r="A3" t="inlineStr"><is><t>Ali</t></is></c><c r="B3"><f>B2+1</f><v>4</v></c><c r="C3" t="inlineStr"><is><t>Closed</t></is></c></row></sheetData></worksheet>''',
    }
    with ZipFile(path, 'w', ZIP_DEFLATED) as z:
        for name, data in files.items(): z.writestr(name, data)


def test_context_detects_headers_columns_and_targets(tmp_path):
    p = tmp_path / 'sample.xlsx'
    make_xlsx(p)
    ctx = analyze_workbook(p, selected_sheet='Data', selected_cell='B3')
    sheet = ctx.sheet('Data')
    assert ctx.sheet_count == 1
    assert sheet.used_range == 'A1:C3'
    assert sheet.header_row == 1
    assert sheet.data_start_row == 2
    assert sheet.columns[1]['title'] == 'Count'
    assert sheet.columns[1]['kind'] == 'formula'
    assert 'A2:C3' in sheet.suggested_targets
    assert ctx.suggestions['sheet'] == 'Data'
    assert ctx.suggestions['cell'] == 'B3'
    assert ctx.suggestions['field_map']['Name'] == 'A'


def test_context_is_read_only_and_reports_formulas(tmp_path):
    p = tmp_path / 'sample.xlsx'
    make_xlsx(p)
    before = p.read_bytes()
    ctx = analyze_workbook(p, selected_sheet='Data')
    after = p.read_bytes()
    assert before == after
    assert ctx.sheet('Data').formula_cells == 1
