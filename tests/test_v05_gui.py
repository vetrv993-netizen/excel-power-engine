import os
os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


def test_gui_module_and_parser():
    import excel_power_engine.gui as gui
    from excel_power_engine.sheet_viewer import num_to_col, col_to_num
    assert num_to_col(22) == "V"
    assert col_to_num("V") == 22
    assert hasattr(gui, "MainWindow")
    assert callable(gui.main)
