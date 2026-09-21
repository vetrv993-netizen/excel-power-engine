from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class ThemePalette:
    window: str
    surface: str
    surface_alt: str
    text: str
    text_muted: str
    input: str
    border: str
    accent: str
    accent_hover: str
    accent_pressed: str
    selection: str
    selection_text: str
    header: str
    danger: str
    warning: str


PALETTES = {
    "light": ThemePalette(
        window="#f7f8fa", surface="#ffffff", surface_alt="#f1f5f9",
        text="#172033", text_muted="#475569", input="#ffffff",
        border="#cbd5e1", accent="#2563eb", accent_hover="#1d4ed8",
        accent_pressed="#1e40af", selection="#bfdbfe", selection_text="#172033",
        header="#e2e8f0", danger="#b91c1c", warning="#a16207",
    ),
    "dark": ThemePalette(
        window="#111827", surface="#1f2937", surface_alt="#273449",
        text="#f8fafc", text_muted="#cbd5e1", input="#182231",
        border="#475569", accent="#60a5fa", accent_hover="#93c5fd",
        accent_pressed="#3b82f6", selection="#1d4ed8", selection_text="#ffffff",
        header="#334155", danger="#fca5a5", warning="#fde68a",
    ),
}


def stylesheet(palette: ThemePalette) -> str:
    return f"""
    QWidget {{ font-family: "Segoe UI"; font-size: 10pt; color: {palette.text}; }}
    QMainWindow, QDialog {{ background: {palette.window}; }}
    QGroupBox, QFrame {{ background: {palette.surface}; color: {palette.text}; }}
    QLineEdit, QComboBox, QSpinBox, QTableWidget, QTextEdit, QListWidget {{
        background: {palette.input}; color: {palette.text};
        border: 1px solid {palette.border}; border-radius: 6px; padding: 5px;
        selection-background-color: {palette.selection};
        selection-color: {palette.selection_text};
    }}
    QLineEdit:focus, QComboBox:focus, QSpinBox:focus, QTextEdit:focus {{
        border: 2px solid {palette.accent};
    }}
    QPushButton {{
        background: {palette.accent}; color: #ffffff; border: none;
        border-radius: 6px; padding: 8px 14px;
    }}
    QPushButton:hover {{ background: {palette.accent_hover}; }}
    QPushButton:pressed {{ background: {palette.accent_pressed}; }}
    QPushButton:disabled {{ background: {palette.border}; color: {palette.text_muted}; }}
    QTabWidget::pane {{ border: 1px solid {palette.border}; background: {palette.surface}; }}
    QTabBar::tab {{ color: {palette.text}; background: {palette.surface_alt}; padding: 8px 14px; }}
    QTabBar::tab:hover {{ background: {palette.selection}; }}
    QTabBar::tab:selected {{ background: {palette.surface}; border-bottom: 2px solid {palette.accent}; }}
    QGroupBox {{ border: 1px solid {palette.border}; border-radius: 8px; margin-top: 12px; padding: 10px; }}
    QGroupBox::title {{ subcontrol-origin: margin; left: 10px; padding: 0 5px; }}
    QHeaderView::section {{ background: {palette.header}; color: {palette.text}; border: none; padding: 5px; }}
    QTableCornerButton::section {{ background: {palette.header}; border: none; }}
    QToolTip {{ background: {palette.surface}; color: {palette.text}; border: 1px solid {palette.border}; }}
    QFrame#hero {{ background: {palette.surface}; border: 1px solid {palette.border}; border-radius: 14px; }}
    QLabel#heroTitle {{ font-size: 20pt; font-weight: 700; color: {palette.text}; }}
    QLabel#heroSubtitle {{ color: {palette.text_muted}; font-size: 10pt; }}
    QLabel#statusPill {{ background: {palette.surface_alt}; border: 1px solid {palette.border}; border-radius: 12px; padding: 6px 12px; color: {palette.text_muted}; }}
    QLineEdit#commandBox {{ min-height: 42px; border-radius: 10px; padding: 8px 14px; font-size: 11pt; }}
    QPushButton#primaryButton {{ min-height: 42px; border-radius: 10px; font-weight: 600; }}
    QTabBar::tab {{ min-width: 110px; }}
    QListWidget::item {{ padding: 6px; }}
    QSplitter::handle {{ background: {palette.border}; }}
    """


class ThemeManager:
    """Owns the GUI theme, its persisted preference, and live application."""

    VALID_THEMES = ("light", "dark", "auto")

    def __init__(self, app):
        from PySide6.QtCore import QSettings

        self.app = app
        self.settings = QSettings()

    def preference(self) -> str:
        value = self.settings.value("theme", "light")
        return value if value in self.VALID_THEMES else "light"

    def resolved_theme(self, preference: str | None = None) -> str:
        preference = preference or self.preference()
        if preference != "auto":
            return preference
        from PySide6.QtGui import QGuiApplication

        color = QGuiApplication.palette().window().color()
        return "dark" if color.lightness() < 128 else "light"

    def apply(self, preference: str | None = None) -> str:
        selected = preference or self.preference()
        if selected not in self.VALID_THEMES:
            selected = "light"
        self.settings.setValue("theme", selected)
        active = self.resolved_theme(selected)
        self.app.setStyleSheet(stylesheet(PALETTES[active]))
        return active
