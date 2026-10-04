"""Nonactivating focused-window palette with four actions.

A 28×40 handle with a chevron sits on the focused window's left edge; hovering
or clicking it reveals the 44×176 palette, which returns to the handle after
1.6 s, or 350 ms after the pointer leaves it.
"""
from PySide6.QtCore import Qt, Signal, QTimer, QPropertyAnimation, QPointF, QRectF
from PySide6.QtGui import QColor, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import QWidget, QVBoxLayout, QAbstractButton, QLabel

FLAGS = (Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint |
         Qt.WindowType.WindowDoesNotAcceptFocus)
SURFACE = QColor(27, 36, 45, 245)


def _glyph(painter, kind, box, color):
    """Monochrome 17 px glyphs standing in for GNOME's symbolic window icons."""
    pen = QPen(color, 1.6); pen.setCapStyle(Qt.PenCapStyle.RoundCap); pen.setJoinStyle(Qt.PenJoinStyle.RoundJoin)
    painter.setPen(pen); painter.setBrush(Qt.BrushStyle.NoBrush)
    x, y, s = box.x(), box.y(), box.width()
    if kind == 'float':          # window-pop-out: frame with an outgoing arrow
        painter.drawRoundedRect(QRectF(x + 1, y + 4, s - 5, s - 5), 2, 2)
        painter.drawLine(QPointF(x + s * .45, y + s * .55), QPointF(x + s - 1, y + 1))
        painter.drawLine(QPointF(x + s * .62, y + 1), QPointF(x + s - 1, y + 1))
        painter.drawLine(QPointF(x + s - 1, y + 1), QPointF(x + s - 1, y + s * .38))
    elif kind == 'tile':         # view-grid: four tiles
        half = (s - 3) / 2
        for dx in (0, half + 3):
            for dy in (0, half + 3):
                painter.drawRoundedRect(QRectF(x + dx, y + dy, half, half), 1.5, 1.5)
    elif kind == 'minimize':
        painter.drawLine(QPointF(x + 3, y + s - 4), QPointF(x + s - 3, y + s - 4))
    elif kind == 'maximize':
        painter.drawRoundedRect(QRectF(x + 2, y + 2, s - 4, s - 4), 2, 2)
    elif kind == 'restore':
        painter.drawRoundedRect(QRectF(x + 2, y + 5, s - 7, s - 7), 2, 2)
        painter.drawPolyline([QPointF(x + 5, y + 5), QPointF(x + 5, y + 2), QPointF(x + s - 2, y + 2),
                              QPointF(x + s - 2, y + s - 5), QPointF(x + s - 5, y + s - 5)])
    elif kind == 'close':
        painter.drawLine(QPointF(x + 3, y + 3), QPointF(x + s - 3, y + s - 3))
        painter.drawLine(QPointF(x + s - 3, y + 3), QPointF(x + 3, y + s - 3))
    elif kind == 'next':         # go-next chevron, 12 px
        painter.drawPolyline([QPointF(x + s * .35, y + s * .2), QPointF(x + s * .7, y + s * .5), QPointF(x + s * .35, y + s * .8)])


class ActionButton(QAbstractButton):
    hovered = Signal(object)

    def __init__(self, kind, label, danger=False):
        super().__init__()
        self.kind, self.danger, self.selected = kind, danger, False
        self.accent = QColor('#5a96d7')
        self.setFixedSize(36, 36)
        self.setAccessibleName(label); self.tooltip = label
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def enterEvent(self, event):
        self.hovered.emit(self); self.update(); super().enterEvent(event)

    def leaveEvent(self, event):
        self.hovered.emit(None); self.update(); super().leaveEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        hover = self.underMouse() or self.hasFocus()
        color = QColor('#d5e0e3')
        if self.selected:
            fill = QColor(self.accent); fill.setAlphaF(.24)
            painter.setPen(Qt.PenStyle.NoPen); painter.setBrush(fill)
            painter.drawRoundedRect(QRectF(self.rect()), 8, 8); color = QColor('#ecf5ff')
        if hover:
            painter.setPen(Qt.PenStyle.NoPen); painter.setBrush(QColor('#6b3038' if self.danger else '#38464d'))
            painter.drawRoundedRect(QRectF(self.rect()), 8, 8)
            color = QColor('#ffd7d7' if self.danger else '#ffffff')
        _glyph(painter, self.kind, QRectF(9.5, 9.5, 17, 17), color)


class ActionHandle(QAbstractButton):
    entered = Signal()

    def __init__(self):
        super().__init__()
        self.setWindowFlags(FLAGS)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setFixedSize(28, 40)
        self.setAccessibleName('Show window actions')

    def enterEvent(self, event):
        self.entered.emit(); self.update(); super().enterEvent(event)

    def leaveEvent(self, event):
        self.update(); super().leaveEvent(event)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        hover = self.underMouse() or self.hasFocus()
        painter.setPen(QPen(QColor('#b2d7ee' if hover else '#7893a3'), 1))
        painter.setBrush(QColor('#344d60') if hover else SURFACE)
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(.5, .5, -.5, -.5), 9, 9)
        _glyph(painter, 'next', QRectF(8, 14, 12, 12), QColor('#ffffff' if hover else '#edf5f9'))


class WindowActions(QWidget):
    action = Signal(str)

    def __init__(self):
        super().__init__()
        self.setWindowFlags(FLAGS)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setFixedSize(44, 176)
        layout = QVBoxLayout(self); layout.setContentsMargins(4, 7, 4, 7); layout.setSpacing(3)
        self.buttons = {}
        for key, label, danger in [('float', 'Float window', False), ('minimize', 'Minimize window', False),
                                   ('maximize', 'Maximize window', False), ('close', 'Close window', True)]:
            button = ActionButton(key, label, danger)
            button.clicked.connect(lambda checked=False, k=key: self.action.emit(k))
            button.hovered.connect(self._tooltip)
            self.buttons[key] = button; layout.addWidget(button)
        self.tip = QLabel(); self.tip.setWindowFlags(FLAGS | Qt.WindowType.WindowTransparentForInput)
        self.tip.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.tip.setStyleSheet('QLabel { background: rgba(24,30,37,245); border: 1px solid #55636b; border-radius: 8px;'
                               ' padding: 6px 9px; color: #f0f4f6; font-size: 11px; font-weight: bold; }')
        self.handle = ActionHandle()
        self.handle.entered.connect(self.reveal); self.handle.clicked.connect(self.reveal)
        self.hide_timer = QTimer(self); self.hide_timer.setSingleShot(True)
        self.hide_timer.timeout.connect(self.collapse)
        self.fade = QPropertyAnimation(self, b'windowOpacity', self)
        self.handle_fade = QPropertyAnimation(self.handle, b'windowOpacity', self)
        self.duration = 140

    def set_state(self, floating, maximized):
        float_button, max_button = self.buttons['float'], self.buttons['maximize']
        float_button.kind = 'tile' if floating else 'float'
        float_button.selected = floating
        float_button.tooltip = 'Tile window' if floating else 'Float window'
        max_button.kind = 'restore' if maximized else 'maximize'
        max_button.tooltip = 'Restore window' if maximized else 'Maximize window'
        for button in (float_button, max_button):
            button.setAccessibleName(button.tooltip); button.update()

    def set_accent(self, color):
        self.buttons['float'].accent = QColor(color)

    def show_handle(self, duration=0):
        self.handle_fade.stop()
        if duration:
            self.handle.setWindowOpacity(0.0); self.handle.show()
            self.handle_fade.setDuration(duration); self.handle_fade.setStartValue(0.0); self.handle_fade.setEndValue(1.0)
            self.handle_fade.start()
        else:
            self.handle.setWindowOpacity(1.0); self.handle.show()

    def reveal(self):
        self.handle.hide()
        appearing = not self.isVisible()
        self.fade.stop()
        if appearing and self.duration:
            self.setWindowOpacity(0.0); self.show()
            self.fade.setDuration(self.duration); self.fade.setStartValue(0.0); self.fade.setEndValue(1.0); self.fade.start()
        else:
            self.setWindowOpacity(1.0); self.show()
        self.hide_timer.start(1600)

    def collapse(self):
        self.tip.hide(); self.hide()
        if self.handle.property('available'): self.show_handle(min(self.duration, 120))

    def _tooltip(self, button):
        if button is None or not self.isVisible():
            self.tip.hide(); return
        self.hide_timer.stop()
        self.tip.setText(button.tooltip); self.tip.adjustSize()
        origin = button.mapToGlobal(button.rect().topRight())
        screen = self.screen().availableGeometry() if self.screen() else None
        x = origin.x() + 8
        if screen and x + self.tip.width() > screen.right() - 6:
            x = button.mapToGlobal(button.rect().topLeft()).x() - self.tip.width() - 8
        self.tip.move(x, origin.y() + (button.height() - self.tip.height()) // 2)
        self.tip.show()

    def enterEvent(self, event):
        self.hide_timer.stop(); super().enterEvent(event)

    def leaveEvent(self, event):
        self.tip.hide(); self.hide_timer.start(350); super().leaveEvent(event)

    def dismiss(self):
        self.hide_timer.stop(); self.handle.setProperty('available', False)
        self.handle.hide(); self.tip.hide(); self.hide()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor('#7893a3'), 1)); painter.setBrush(SURFACE)
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(.5, .5, -.5, -.5), 10, 10)

    def closeEvent(self, event):
        self.handle.close(); self.tip.close(); super().closeEvent(event)
