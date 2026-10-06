"""Shared Qt tokens. All dimensions are logical pixels; Qt handles per-screen DPI."""
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QFont, QPalette
from PySide6.QtWidgets import QApplication, QStyleFactory
from dataclasses import replace


def studio_settings(settings):
    # The Studio and switcher use their own dark surface, independently of
    # the system-themed preferences window.
    return replace(settings,theme='dark')


def tokens(settings):
    app = QApplication.instance()
    dark = settings.theme == "dark"
    if settings.theme == "system" and app:
        dark = app.styleHints().colorScheme() == Qt.ColorScheme.Dark
    accent=QColor(settings.accent)
    def linear(value):
        return value/12.92 if value<=.04045 else ((value+.055)/1.055)**2.4
    luminance=sum(weight*linear(value) for weight,value in zip((.2126,.7152,.0722),(accent.redF(),accent.greenF(),accent.blueF())))
    return dict(accent_text='#10241c' if luminance>.179 else '#ffffff',bg="#171e26" if dark else "#f2f5f4",
                canvas="#10161d" if dark else "#e5ece9",
                card="#26343e" if dark else "#ffffff",
                text="#ecf4f2" if dark else "#172c26",
                muted="#a8b8be" if dark else "#486059",
                line="#435761" if dark else "#94aaa1",
                selected="#513a3f" if dark else "#ffe3dd",
                selection="#ff9b91" if dark else "#ae4237",
                accent=settings.accent, error="#f3bd72" if dark else "#8a4300")


def dark_shell(settings):
    return settings.theme == 'dark'


# Dark surface styles: chips, chooser, library and search fields of the dark
# Studio and switcher. The system-themed Preferences window does not use it.
SHELL_DARK = """
    QMainWindow, QDialog { background: #171e26; border: 1px solid #35454c; border-radius: 24px; }
    QPushButton, QToolButton { background: #242e38; border: 1px solid transparent; border-radius: 9px;
        padding: 7px 10px; font-size: 11px; color: #c1cdd1; }
    QPushButton:hover, QToolButton:hover, QPushButton:focus, QToolButton:focus { background: #35454c; color: white; border: 1px solid transparent; }
    QPushButton:checked, QToolButton:checked { background: #28483f; border: 1px solid #73cfa8; color: #b5f7db; }
    QPushButton:disabled, QToolButton:disabled { color: #667980; background: #1a232b; }
    QPushButton[unavailable=true] { background: #1a232b; border-color: #303e45; color: #82939a; }
    QPushButton[unavailable=true]:hover { background: #29343b; border-color: #f3bd72; color: #f3bd72; }
    QPushButton[danger=true]:hover, QPushButton[danger=true]:focus { background: #56323a; color: #ffd7d7; }
    QPushButton[primary=true] { background: #8ce8c3; color: #10241c; font-weight: bold; }
    QPushButton[primary=true]:hover { background: #b5f7db; color: #10241c; }
    QLineEdit { background: #202d35; border: 1px solid #526570; border-radius: 10px; padding: 9px 12px; color: #ecf4f2; }
    QLineEdit:focus { border: 1px solid #8ce8c3; }
    QListWidget { background: #10161d; border: 1px solid #2d3b44; border-radius: 12px; padding: 8px; }
    QListWidget::item { background: #26343e; border: 1px solid #435761; border-radius: 10px; padding: 10px; margin: 3px 0; color: #ecf4f2; }
    QListWidget::item:hover { background: #314d4b; border-color: #8ce8c3; }
    QListWidget::item:selected { background: #28483f; border-color: #73cfa8; color: #ecf4f2; }
    QListWidget[chooser=true]::item { background: #242e38; border: 1px solid transparent; border-radius: 9px; padding: 11px 12px; }
    QListWidget[chooser=true]::item:hover { background: #35454c; border-color: #73cfa8; }
    QListWidget[chooser=true]::item:selected { background: #28483f; border-color: #73cfa8; }
    QListWidget[chooser=true]::item:disabled { background: transparent; border: none; color: #8ce8c3; font-size: 11px; font-weight: 700; padding: 9px 5px 4px; }
    QLabel[legend=true] { color: #93a6ad; font-size: 10px; }
    QScrollBar:vertical { background: transparent; width: 10px; margin: 4px 2px; }
    QScrollBar::handle:vertical { background: #435761; border-radius: 3px; min-height: 30px; }
    QScrollBar::handle:vertical:hover { background: #73cfa8; }
    QScrollBar::add-line, QScrollBar::sub-line, QScrollBar::add-page, QScrollBar::sub-page { background: none; height: 0; }
    QDoubleSpinBox { background: #202b34; color: #edf5f2; border: 1px solid #435761; border-radius: 7px; padding: 5px 7px; font-size: 11px; }
    QDoubleSpinBox:focus { border-color: #73cfa8; }
    QDoubleSpinBox:disabled { color: #91a5a9; background: #171e26; border-color: #303e45; }
"""


def _arrow(color):
    """A small down arrow for combo boxes, written once per colour."""
    import tempfile
    from pathlib import Path
    path = Path(tempfile.gettempdir()) / f"smartgrid-arrow-{color.lstrip('#')}.svg"
    if not path.exists():
        path.write_text(f'<svg xmlns="http://www.w3.org/2000/svg" width="10" height="6" viewBox="0 0 10 6">'
                        f'<path d="M1 1l4 4 4-4" fill="none" stroke="{color}" stroke-width="1.6" '
                        f'stroke-linecap="round" stroke-linejoin="round"/></svg>', encoding='utf-8')
    return path.as_posix()


def apply_theme(settings,target=None):
    app = QApplication.instance()
    if not app:
        return
    t = tokens(settings)
    palette = QPalette(app.style().standardPalette())
    for role, color in [(QPalette.Window,t['bg']), (QPalette.Base,t['canvas']),
                        (QPalette.AlternateBase,t['card']), (QPalette.Text,t['text']),
                        (QPalette.WindowText,t['text']), (QPalette.Button,t['card']),
                        (QPalette.ButtonText,t['text']), (QPalette.Highlight,t['accent']),
                        (QPalette.HighlightedText,t['accent_text'])]:
        palette.setColor(role, QColor(color))
    surface=target or app
    surface.setPalette(palette)
    surface.setFont(QFont("Segoe UI", 10))
    surface.setStyleSheet(f"""
        QWidget {{ color: {t['text']}; }}
        QMainWindow, QDialog {{ background: {t['bg']}; border: 1px solid {t['line']}; border-radius: 22px; }}
        QPushButton, QToolButton, QComboBox {{ background: {t['card']}; border: 1px solid {t['line']};
            border-radius: 8px; padding: 7px 10px; min-height: 18px; }}
        QPushButton:hover, QToolButton:hover {{ border-color: {t['accent']}; }}
        QComboBox {{ padding-right: 26px; }}
        QComboBox::drop-down {{ subcontrol-origin: padding; subcontrol-position: center right; width: 24px;
            border: none; background: transparent; }}
        QComboBox::down-arrow {{ image: url({_arrow(t['muted'])}); width: 10px; height: 6px; }}
        QPushButton:focus, QToolButton:focus, QComboBox:focus, QLineEdit:focus {{ border: 2px solid {t['accent']}; }}
        QPushButton:checked, QToolButton:checked {{ background: {'#28483f' if settings.theme=='dark' else t['canvas']}; border-color: #73cfa8; }}
        QPushButton:disabled {{ color: {t['muted']}; background: {t['canvas']}; }}
        QPushButton[disclosure="true"], QPushButton[disclosure="true"]:checked {{ border: none; background: transparent;
            text-align: left; padding: 4px 0; color: {t['text']}; font-weight: 600; }}
        QPushButton[disclosure="true"]:hover {{ text-decoration: underline; }}
        QPushButton[disclosure="true"]:disabled {{ color: {t['muted']}; background: transparent; }}
        QPushButton[unavailable=true] {{ color: #82939a; background: #1a232b; border-color: #303e45; }}
        QPushButton[primary=true] {{ background: {t['accent']}; color: {t['accent_text']}; font-weight: bold; }}
        QLineEdit, QSpinBox, QDoubleSpinBox, QListWidget, QTreeWidget, QTextEdit {{
            background: {t['canvas']}; border: 1px solid {t['line']}; border-radius: 8px; padding: 5px; }}
        QComboBox:editable {{ background: {t['canvas']}; padding: 5px 26px 5px 5px; }}
        QComboBox:editable QLineEdit {{ background: transparent; border: none; padding: 0; }}
        QScrollArea {{ border: none; background: {t['bg']}; }}
        QScrollArea > QWidget > QWidget {{ background: {t['bg']}; }}
        QGroupBox {{ border: 1px solid {t['line']}; border-radius: 10px; margin-top: 14px; padding-top: 12px; }}
        QGroupBox::title {{ subcontrol-origin: margin; left: 12px; padding: 0 5px; }}
        QTabWidget::pane {{ border: none; }} QTabWidget::tab-bar {{ alignment: center; }}
        QTabBar::tab {{ background: {t['card']}; border-radius: 6px; padding: 8px 12px; margin: 2px; }}
        QTabBar::tab:selected {{ border-bottom: 2px solid {t['accent']}; }}
        QMenu {{ background: {t['card']}; border: 1px solid {t['line']}; padding: 5px; }}
        QMenu::item {{ padding: 7px 18px; color: {t['text']}; }}
        QMenu::item:selected {{ background: {t['canvas']}; color: {t['text']}; }}
        QMenu::item:disabled {{ color: #5d6d74; }}
        QMenu::item:disabled:selected {{ background: transparent; }}
        QMenu::separator {{ height: 1px; background: {t['line']}; margin: 5px 10px; }}
        QLabel[menu_status=true] {{ color: {t['accent']}; font-weight: bold; padding: 7px 18px; }}
        QLabel[menu_heading=true] {{ color: {t['muted']}; font-size: 8pt; font-weight: bold; letter-spacing: 1px; padding: 8px 18px 2px 18px; }}
        QLabel[muted=true] {{ color: {t['muted']}; }} QLabel[error=true] {{ color: {t['error']}; }}
        QLabel[eyebrow=true] {{ color: {t['accent']}; font-size: 9px; font-weight: 800; }}
        QGroupBox[preferences=true] {{ background: {t['card']}; border-color: {t['line']}; padding: 12px; margin-top: 18px; }}
        QGroupBox[collapsed=true] {{ border: none; background: transparent; }}
        QLabel[pref_heading=true] {{ color: {t['muted']}; font-size: 8pt; font-weight: bold; letter-spacing: 1px; padding-top: 8px; }}
        QGroupBox::indicator {{ width: 14px; height: 14px; border: 1px solid {t['line']}; border-radius: 3px; background: {t['canvas']}; }}
        QGroupBox::indicator:checked {{ background: {t['accent']}; border-color: {t['accent']}; }}
        QToolTip {{ background: {t['card']}; color: {t['text']}; border: 1px solid {t['accent']}; padding: 7px; }}
    """+(SHELL_DARK if dark_shell(settings) else ''))
