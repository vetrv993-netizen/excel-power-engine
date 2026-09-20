from excel_power_engine.operations import preview_operations


def test_preview_operations():
    out=preview_operations([
        {'kind':'format','sheet':'Data','ranges':['A1:C2']},
        {'kind':'merge','sheet':'Data','ranges':['A1:C1']},
        {'kind':'pdf','output_pdf':'x.pdf'},
    ])
    assert len(out)==3
    assert out[1]['destructive'] is True
    assert out[2]['target']=='workbook'
