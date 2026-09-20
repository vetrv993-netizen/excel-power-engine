import importlib

def test_vbs_string_uses_valid_double_quote_literals():
    mod = importlib.import_module("excel_power_engine.open_excel")
    assert mod._vbs_string("C:/A B/test.xlsm") == '"C:/A B/test.xlsm"'
    assert mod._vbs_string('A "quote"') == '"A ""quote"""'
