import importlib

def test_open_excel_module_exposes_navigation_helpers():
    mod = importlib.import_module("excel_power_engine.open_excel")
    assert callable(mod.open_in_excel)
    assert callable(mod._open_with_pywin32)
    assert callable(mod._open_with_vbs)

def test_vbs_string_escaping():
    mod = importlib.import_module("excel_power_engine.open_excel")
    assert mod._vbs_string('A "quote"') == '"A ""quote"""'
