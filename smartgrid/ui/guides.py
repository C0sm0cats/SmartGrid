"""Desktop guides: drop target, space indicator, swap guides, pinned-tile cards."""
from PySide6.QtCore import Qt, QRectF, QTimer, QPropertyAnimation, Signal
from PySide6.QtGui import QColor, QFont, QPainter, QPen, QBrush
from PySide6.QtWidgets import QWidget

FLAGS = (Qt.WindowType.Tool | Qt.WindowType.FramelessWindowHint | Qt.WindowType.WindowStaysOnTopHint |
         Qt.WindowType.WindowDoesNotAcceptFocus)
SURFACE = QColor(24, 30, 37, 237)


class Overlay(QWidget):
    def __init__(self, interactive=False):
        super().__init__()
        self.setWindowFlags(FLAGS if interactive else FLAGS | Qt.WindowType.WindowTransparentForInput)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.color = '#78aef0'

    def font(self, size, bold=True):
        font = QFont('Segoe UI', size)
        font.setBold(bold)
        return font


class TargetGuide(Overlay):
    """Drop target with its label ("Swap · App" / "Move here")."""
    def __init__(self):
        super().__init__()
        self.label = None
        self.content = None  # (icon, name) of the dragged window's app
        self.setAccessibleName('Drop target')

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        accent = QColor(self.color)
        fill = QColor(accent); fill.setAlpha(31)
        painter.setPen(QPen(accent, 2)); painter.setBrush(fill)
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(1, 1, -1, -1), 12, 12)
        if self.content:
            # Icon and app name in a dark pill, centred.
            icon, name = self.content
            painter.setFont(self.font(9))
            width = min(self.width() - 24, 24 + 8 + painter.fontMetrics().horizontalAdvance(name) + 20)
            box = QRectF((self.width() - width) / 2, (self.height() - 38) / 2, width, 38)
            painter.setPen(Qt.PenStyle.NoPen); painter.setBrush(QColor(24, 30, 37, 133))
            painter.drawRoundedRect(box, 8, 8)
            if icon:
                icon.paint(painter, round(box.x() + 10), round(box.y() + 7), 24, 24)
            painter.setPen(QColor('#eef3f4'))
            painter.drawText(QRectF(box.x() + 42, box.y(), box.width() - 50, box.height()), Qt.AlignmentFlag.AlignVCenter,
                             painter.fontMetrics().elidedText(name, Qt.TextElideMode.ElideRight, round(box.width() - 50)))
        if self.label:
            painter.setFont(self.font(9))
            width = painter.fontMetrics().horizontalAdvance(self.label) + 22
            box = QRectF(12, 12, min(width, self.width() - 24), 28)
            painter.setPen(QPen(accent, 1)); painter.setBrush(SURFACE)
            painter.drawRoundedRect(box, 10, 10)
            painter.setPen(QColor('#f0f4f6'))
            painter.drawText(box, Qt.AlignmentFlag.AlignCenter, self.label)


class SpaceOsd(Overlay):
    """Space indicator: three dots and "Space N · Display N", fading out."""
    def __init__(self):
        super().__init__()
        self.space, self.text = 0, ''
        self.resize(180, 64)
        self.fade = QPropertyAnimation(self, b'windowOpacity', self)
        self.fade.finished.connect(lambda: self.hide() if self.windowOpacity() < .05 else None)
        self.timer = QTimer(self); self.timer.setSingleShot(True)
        self.timer.timeout.connect(self._fade_out)

    def show_space(self, center_x, top, space, text, duration):
        self.space, self.text = space, text
        self.move(round(center_x - 90), round(top + 72))
        self.fade.stop(); self.setWindowOpacity(1.0 if not duration else 0.0)
        self.show(); self.update()
        if duration:
            self.fade.setDuration(duration); self.fade.setStartValue(0.0); self.fade.setEndValue(1.0); self.fade.start()
        self.timer.start(900)

    def _fade_out(self):
        self.fade.stop(); self.fade.setDuration(220)
        self.fade.setStartValue(self.windowOpacity()); self.fade.setEndValue(0.0); self.fade.start()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(210, 220, 225, 71), 1)); painter.setBrush(QColor(24, 30, 37, 219))
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(.5, .5, -.5, -.5), 12, 12)
        for i in range(3):
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(QColor(self.color if i == self.space else '#68767d'))
            painter.drawEllipse(QRectF(self.width() / 2 - 20 + i * 16, 16, 8, 8))
        painter.setPen(QColor('#eef3f4')); painter.setFont(self.font(9))
        painter.drawText(QRectF(0, 30, self.width(), 24), Qt.AlignmentFlag.AlignCenter, self.text)


class PinPlaceholder(Overlay):
    """Pinned-tile card: icon, application name and PINNED · STATE."""
    clicked = Signal()

    def __init__(self):
        super().__init__(interactive=True)
        # Above the desktop, below real windows.
        self.setWindowFlags((FLAGS & ~Qt.WindowType.WindowStaysOnTopHint) | Qt.WindowType.WindowStaysOnBottomHint)
        self.icon = None
        self.name = self.state = ''
        self.actionable = True
        self.hover = False
        self.setCursor(Qt.CursorShape.PointingHandCursor)

    def describe(self, icon, name, state, actionable):
        self.icon, self.name, self.state, self.actionable = icon, name, state, actionable
        kind = state.split(' · ')[-1]
        verb = {'FLOATING': 'Tile floating window', 'MINIMIZED': 'Restore window', 'CLOSED': 'Open application'}.get(kind, kind)
        self.setAccessibleName(f"{verb}: {name}")
        self.setCursor(Qt.CursorShape.PointingHandCursor if actionable else Qt.CursorShape.ArrowCursor)
        self.update()

    def enterEvent(self, event):
        self.hover = True; self.update(); super().enterEvent(event)

    def leaveEvent(self, event):
        self.hover = False; self.update(); super().leaveEvent(event)

    def mouseReleaseEvent(self, event):
        if self.actionable and event.button() == Qt.MouseButton.LeftButton and self.rect().contains(event.position().toPoint()):
            self.clicked.emit()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setOpacity(1 if self.actionable else .55)
        hover = self.hover and self.actionable
        painter.setPen(QPen(QColor('#74bf8e') if hover else QColor(116, 191, 142, 140), 1))
        painter.setBrush(QColor(37, 68, 53, 189) if hover else QColor(27, 38, 43, 133))
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(.5, .5, -.5, -.5), 12, 12)
        compact = self.height() < 110 or self.width() < 150
        icon = 0 if self.height() < 60 else 20 if compact else 32
        text_height = (14 + 12) if not compact else (12 + 10)
        top = (self.height() - icon - (6 if icon else 0) - text_height) / 2
        if icon and self.icon:
            self.icon.paint(painter, round((self.width() - icon) / 2), round(top), icon, icon)
        y = top + icon + (6 if icon else 0)
        painter.setPen(QColor('#edf4ef')); painter.setFont(self.font(8 if compact else 9))
        name = painter.fontMetrics().elidedText(self.name, Qt.TextElideMode.ElideRight, max(20, self.width() - 16))
        painter.drawText(QRectF(8, y, self.width() - 16, 16), Qt.AlignmentFlag.AlignCenter, name)
        painter.setPen(QColor('#9bd6aa')); painter.setFont(self.font(6 if compact else 7))
        painter.drawText(QRectF(8, y + (14 if compact else 18), self.width() - 16, 12), Qt.AlignmentFlag.AlignCenter, self.state)


class PlacementGhost(Overlay):
    """Icon and name travelling from the old tile to the new one."""
    def __init__(self, icon, name, start, end, duration, curve):
        from PySide6.QtCore import QParallelAnimationGroup, QEasingCurve
        super().__init__()
        self.icon, self.name = icon, name
        self.setGeometry(start); self.setWindowOpacity(145 / 255)
        self.group = QParallelAnimationGroup(self)
        for prop, a, b in ((b'geometry', start, end), (b'windowOpacity', 145 / 255, 0.0)):
            animation = QPropertyAnimation(self, prop, self)
            animation.setDuration(duration); animation.setStartValue(a); animation.setEndValue(b)
            animation.setEasingCurve(curve)
            self.group.addAnimation(animation)
        self.group.finished.connect(self.close)
        self.group.finished.connect(self.deleteLater)

    def start(self):
        self.show(); self.group.start()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(205, 220, 225, 107), 1)); painter.setBrush(QColor(35, 45, 53, 148))
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(.5, .5, -.5, -.5), 12, 12)
        painter.setFont(self.font(9))
        width = 24 + 8 + painter.fontMetrics().horizontalAdvance(self.name)
        x = (self.width() - width) / 2; y = (self.height() - 24) / 2
        if self.icon:
            self.icon.paint(painter, round(x), round(y), 24, 24)
        painter.setPen(QColor('#eef3f4'))
        painter.drawText(QRectF(x + 32, y, self.width(), 24), Qt.AlignmentFlag.AlignVCenter, self.name)


class WindowGhost(QWidget):
    """Live image of a window travelling from its old tile to the new one.
    Windows draws the image (a DWM thumbnail), so the move is
    smooth however heavy the application. It lands exactly on the window,
    which is already in place: the end is seamless."""
    def __init__(self, api, source, frame, start, end, duration, curve):
        from PySide6.QtCore import QVariantAnimation
        super().__init__()
        # Opaque and not click-through: Windows does not draw a thumbnail in a
        # layered window. It only lives for the animation.
        self.setWindowFlags(FLAGS)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.api, self.source, self.frame = api, source, frame
        self.start_rect, self.end_rect, self.handle = start, end, None
        self.animation = QVariantAnimation(self)
        self.animation.setStartValue(0.0); self.animation.setEndValue(1.0)
        self.animation.setDuration(duration); self.animation.setEasingCurve(curve)
        self.animation.valueChanged.connect(self._step)
        self.animation.finished.connect(self._finish)

    def start(self):
        """False when Windows has no image for the window: the caller falls back."""
        import ctypes as C
        from smartgrid.platform.windows.api import HANDLE
        handle = HANDLE()
        self.setGeometry(self.start_rect)
        # The thumbnail is registered before the window shows: no empty frame.
        if self.api.DwmRegisterThumbnail(int(self.winId()), self.source, C.byref(handle)) != 0 or not handle.value:
            self.deleteLater()
            return False
        self.handle = handle.value
        # Rounded corners like the window it shows (Windows 11; ignored elsewhere).
        corner = C.c_int(2)
        self.api.DwmSetWindowAttribute(int(self.winId()), 33, C.byref(corner), C.sizeof(corner))
        if not self._step(0.0):
            self._finish()
            return False
        self.show(); self.animation.start()
        return True

    def _step(self, value):
        import ctypes as C
        from PySide6.QtCore import QRect
        from smartgrid.platform.windows.api import RECT, DWM_THUMBNAIL_PROPERTIES
        if not self.handle:
            return False
        t, a, b = float(value), self.start_rect, self.end_rect
        rect = QRect(round(a.x() + (b.x() - a.x()) * t), round(a.y() + (b.y() - a.y()) * t),
                     max(1, round(a.width() + (b.width() - a.width()) * t)),
                     max(1, round(a.height() + (b.height() - a.height()) * t)))
        self.setGeometry(rect)
        scale = self.devicePixelRatioF()
        properties = DWM_THUMBNAIL_PROPERTIES()
        # Destination, source, opacity, visibility, whole window.
        properties.dwFlags = 0x01 | 0x02 | 0x04 | 0x08 | 0x10
        properties.rcDestination = RECT(0, 0, round(rect.width() * scale), round(rect.height() * scale))
        properties.rcSource = RECT(*self.frame)
        properties.opacity = 255
        properties.fVisible = True
        properties.fSourceClientAreaOnly = False
        return self.api.DwmUpdateThumbnailProperties(self.handle, C.byref(properties)) == 0

    def _finish(self):
        if self.handle:
            self.api.DwmUnregisterThumbnail(self.handle)
            self.handle = None
        self.close(); self.deleteLater()


class LiveImage(QWidget):
    """Live image of a neighbour on its future tile while a window edge is
    dragged. Windows draws it (a DWM thumbnail), so it follows the pointer
    however heavy the application; the real window takes the tile at the
    release, then the image goes. Qt never paints it: moving it costs only
    the native move, batched with the other images by place_images."""
    def __init__(self, api):
        super().__init__()
        # Opaque and not click-through: Windows does not draw a thumbnail in a
        # layered window.
        self.setWindowFlags(FLAGS)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_PaintOnScreen)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_OpaquePaintEvent)
        self.api, self.source, self.handle = api, None, None
        # The thumbnail as drawn (source frame, size) and as wanted; the
        # rectangle the window was last moved to.
        self.shown = self.wanted = self.placed = None

    def paintEngine(self):
        return None

    def paintEvent(self, event):
        pass

    def show_window(self, source, frame, logical, physical):
        """Show source's visible frame stretched over the tile (physical
        rectangle; logical for Qt's first placement). False when Windows has
        no image for it. The move itself is done by place_images; the
        thumbnail is never smaller than the window meanwhile (a growing
        window would show a strip of nothing for a frame): it is grown here,
        before the move, and shrunk by settle after it."""
        if source != self.source and not self.prepare(source):
            return False
        size = (physical.width, physical.height)
        self.wanted = (frame, size)
        if self.shown is None or frame != self.shown[0] or size[0] > self.shown[1][0] or size[1] > self.shown[1][1]:
            grown = size if self.shown is None else (max(size[0], self.shown[1][0]), max(size[1], self.shown[1][1]))
            if not self._draw(frame, grown):
                self.release()
                return False
        if not self.isVisible():
            self.setGeometry(logical)
            self.show()
        return True

    def prepare(self, source):
        """Register source's thumbnail, the window still hidden: done for the
        neighbours when a gesture starts, so the first frame has only to show
        them. False when Windows has no image for it."""
        import ctypes as C
        from smartgrid.platform.windows.api import HANDLE
        self.release()
        handle = HANDLE()
        if self.api.DwmRegisterThumbnail(int(self.winId()), source, C.byref(handle)) != 0 or not handle.value:
            return False
        self.source, self.handle = source, handle.value
        # Rounded corners like the window it shows (Windows 11; ignored elsewhere).
        corner = C.c_int(2)
        self.api.DwmSetWindowAttribute(int(self.winId()), 33, C.byref(corner), C.sizeof(corner))
        return True

    def settle(self):
        """After the move: the thumbnail at the exact size of the tile."""
        if self.handle and self.wanted and self.wanted != self.shown:
            self._draw(*self.wanted)

    def _draw(self, frame, size):
        import ctypes as C
        from smartgrid.platform.windows.api import RECT, DWM_THUMBNAIL_PROPERTIES
        properties = DWM_THUMBNAIL_PROPERTIES()
        # Destination, source, opacity, visibility, whole window.
        properties.dwFlags = 0x01 | 0x02 | 0x04 | 0x08 | 0x10
        properties.rcDestination = RECT(0, 0, size[0], size[1])
        properties.rcSource = RECT(*frame)
        properties.opacity = 255
        properties.fVisible = True
        properties.fSourceClientAreaOnly = False
        if self.api.DwmUpdateThumbnailProperties(self.handle, C.byref(properties)) != 0:
            return False
        self.shown = (frame, size)
        return True

    def release(self):
        if self.handle:
            self.api.DwmUnregisterThumbnail(self.handle)
        self.source = self.handle = self.shown = self.wanted = self.placed = None
        self.hide()


class TileCover(QWidget):
    """Flat cover of what a linked resize leaves uncovered of the old tiles
    (a new gap over a window not placed yet, the old size of a late held
    window), in the colour of the gaps around: nothing shows at its old
    size. Opaque, painted by Qt; moved natively with the images."""
    def __init__(self):
        super().__init__()
        self.setWindowFlags(FLAGS | Qt.WindowType.WindowTransparentForInput)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)
        self.setAttribute(Qt.WidgetAttribute.WA_OpaquePaintEvent)
        self.setAttribute(Qt.WidgetAttribute.WA_NoSystemBackground)
        self.color = QColor(32, 32, 32)
        self.placed = None

    def paintEvent(self, event):
        QPainter(self).fillRect(self.rect(), self.color)


def place_images(api, placements):
    """Move every image (and cover) to its place in one native batch: the
    neighbours move together, in the same frame. placements: [(QWidget,
    physical Rect)]."""
    if not placements:
        return
    # SWP_NOZORDER|SWP_NOACTIVATE|SWP_NOREDRAW|SWP_NOCOPYBITS: Windows draws
    # the thumbnail; there is nothing of our own to repaint or copy.
    flags = 0x0004 | 0x0010 | 0x0008 | 0x0100
    batch = api.BeginDeferWindowPos(len(placements))
    for image, rect in placements:
        if not batch:
            break
        batch = api.DeferWindowPos(batch, int(image.winId()), None, rect.x, rect.y, rect.width, rect.height, flags)
    if batch and api.EndDeferWindowPos(batch):
        return
    # A failed batch moves nothing: one by one instead.
    for image, rect in placements:
        api.SetWindowPos(int(image.winId()), None, rect.x, rect.y, rect.width, rect.height, flags)


class SourceGuide(Overlay):
    """Faint outline of the slot a swapped window leaves."""
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor(210, 220, 225, 115), 1)); painter.setBrush(QColor(210, 220, 225, 10))
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(.5, .5, -.5, -.5), 14, 14)
