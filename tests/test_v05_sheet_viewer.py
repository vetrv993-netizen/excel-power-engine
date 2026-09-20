from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from excel_power_engine.sheet_viewer import workbook_sheets, read_region


def test_ooxml_sheet_viewer_reads_formula_and_inline_value(tmp_path):
    p = Path(tmp_path) / "demo.xlsx"
    workbook = '''<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'''
    rels = '''<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>'''
    sheet = '''<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Hello</t></is></c><c r="B1"><f>SUM(C1:C2)</f><v>3</v></c></row><row r="2"><c r="C2"><v>2</v></c></row></sheetData></worksheet>'''
    with ZipFile(p, 'w', ZIP_DEFLATED) as z:
        z.writestr('xl/workbook.xml', workbook)
        z.writestr('xl/_rels/workbook.xml.rels', rels)
        z.writestr('xl/worksheets/sheet1.xml', sheet)
    assert workbook_sheets(p)['Sheet1'] == 'xl/worksheets/sheet1.xml'
    cols, rows, target = read_region(p, 'Sheet1', center_cell='B1', row_radius=1, col_radius=2)
    assert target == 'B1'
    assert 'B' in cols
    cells = {(r, c.ref): c for r, row in rows for c in row if c is not None}
    assert cells[(1, 'A1')].value == 'Hello'
    assert cells[(1, 'B1')].formula == 'SUM(C1:C2)'
