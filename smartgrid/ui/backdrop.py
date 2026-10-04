"""Dialog-only desktop dimming; never shown by the background tiling runtime."""
from PySide6.QtCore import Qt
from PySide6.QtWidgets import QWidget,QApplication


class Backdrop(QWidget):
    def __init__(self,dialog):
        super().__init__()
        self.dialog=dialog
        self.setWindowFlags(Qt.WindowType.Tool|Qt.WindowType.FramelessWindowHint|Qt.WindowType.WindowDoesNotAcceptFocus|Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setStyleSheet('background-color: rgba(16,22,29,150);')

    def paintEvent(self,event):
        from PySide6.QtGui import QPainter,QColor
        painter=QPainter(self);painter.fillRect(self.rect(),QColor(16,22,29,150))

    def mousePressEvent(self,event): self.dialog.close()


def show_backdrops(dialog):
    if getattr(dialog.controller.backend,'is_fake',False): return
    backdrops=getattr(dialog,'_desktop_backdrops',[])
    screens=QApplication.instance().screens()
    while len(backdrops)>len(screens): backdrops.pop().close()
    while len(backdrops)<len(screens): backdrops.append(Backdrop(dialog))
    for backdrop,screen in zip(backdrops,screens):
        backdrop.setGeometry(screen.geometry());backdrop.show()
    dialog._desktop_backdrops=backdrops
    dialog.raise_()


def hide_backdrops(dialog):
    for backdrop in getattr(dialog,'_desktop_backdrops',[]): backdrop.hide()
