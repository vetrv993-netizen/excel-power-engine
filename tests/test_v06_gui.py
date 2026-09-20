import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')


def test_gui_v06_symbols():
    import excel_power_engine.gui as gui
    assert hasattr(gui.MainWindow, '_build_intelligence_tab')
    assert hasattr(gui.MainWindow, '_inspect_cell_intelligence')
    assert hasattr(gui.MainWindow, '_bulk_preview')
    assert hasattr(gui.MainWindow, '_bulk_execute')
    assert hasattr(gui.MainWindow, '_open_excel')
