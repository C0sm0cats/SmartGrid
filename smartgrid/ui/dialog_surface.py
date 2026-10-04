"""Rounded shell surface and motion policy shared by the two shell dialogs."""
from PySide6.QtCore import Qt,QRectF,QPropertyAnimation,QEasingCurve
from PySide6.QtGui import QPainter,QColor,QPen
from PySide6.QtWidgets import QApplication


def paint_surface(widget):
    painter=QPainter(widget)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setBrush(QColor('#171e26'));painter.setPen(QPen(QColor('#435761'),1))
    painter.drawRoundedRect(QRectF(widget.rect()).adjusted(.5,.5,-.5,-.5),22,22)


def animate_open(widget):
    old=getattr(widget,'_open_animation',None)
    if old: old.stop()
    settings=widget.controller.settings
    duration=settings.visual_duration()
    if not duration or QApplication.platformName()=='offscreen':
        widget.setWindowOpacity(1);return
    animation=QPropertyAnimation(widget,b'windowOpacity',widget)
    animation.setDuration(duration);animation.setStartValue(.15);animation.setEndValue(1)
    curve={'ease-out':QEasingCurve.Type.OutCubic,'linear':QEasingCurve.Type.Linear,
           'ease-in-out':QEasingCurve.Type.InOutCubic}[settings.animation_curve]
    animation.setEasingCurve(curve);widget._open_animation=animation;animation.start()
