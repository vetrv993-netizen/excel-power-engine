import os

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


def test_theme_manager_applies_and_persists_preferences():
    from PySide6.QtWidgets import QApplication
    from excel_power_engine.theme import ThemeManager

    app = QApplication.instance() or QApplication([])
    app.setApplicationName("Excel Power Engine Theme Test")
    app.setOrganizationName("Excel Power Engine Tests")
    manager = ThemeManager(app)
    manager.settings.clear()

    assert manager.apply("dark") == "dark"
    assert manager.preference() == "dark"
    assert "#111827" in app.styleSheet()

    assert manager.apply("light") == "light"
    assert manager.preference() == "light"
    assert "#f7f8fa" in app.styleSheet()


def test_main_window_exposes_settings_tab():
    from PySide6.QtWidgets import QApplication
    from excel_power_engine.gui import MainWindow

    app = QApplication.instance() or QApplication([])
    window = MainWindow()
    assert hasattr(window, "theme_combo")
    assert window.tabs.count() == 9
    window.close()
