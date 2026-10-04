"""Card rows of the switcher and the library, painted from multi-line item text."""
from PySide6.QtCore import Qt, QRect, QRectF, QSize
from PySide6.QtGui import QColor, QFont, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QStyledItemDelegate, QStyle, QPushButton
from smartgrid.core.geometry import resolve_layout
from smartgrid.core.models import Rect


def font(pixels, weight=QFont.Weight.Normal):
    result = QFont('Segoe UI'); result.setPixelSize(pixels); result.setWeight(weight)
    return result


def quick_preview(preset, count, filled=(), pinned=(), tiles=None, ratio=.6):
    """An 84×54 miniature, filled cells in teal."""
    pixmap = QPixmap(84, 54); pixmap.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pixmap); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(QPen(QColor('#53656b'), 1)); painter.setBrush(QColor('#131f25'))
    painter.drawRoundedRect(QRectF(.5, .5, 83, 53), 7, 7)
    try:
        rects = resolve_layout(Rect(0, 0, 84, 54), max(1, count), preset=preset, gap=2, padding=3, ratio=ratio, tiles=tiles)
    except ValueError:
        rects = []
    for index, r in enumerate(rects):
        full = index in filled
        border = '#9be9b2' if index in pinned else '#8ce8c3' if full else '#6a7f83'
        painter.setPen(QPen(QColor(border), 1)); painter.setBrush(QColor('#538e8a' if full else '#34434a'))
        painter.drawRoundedRect(QRectF(r.x + .5, r.y + .5, r.width - 1, r.height - 1), 2, 2)
    painter.end()
    return QIcon(pixmap)


class CardDelegate(QStyledItemDelegate):
    """Rows: icon or preview, bold title, metadata (green "N TILES"), detail, trailing ›."""
    def __init__(self, parent=None, switcher=False):
        super().__init__(parent)
        self.switcher = switcher

    def sizeHint(self, option, index):
        if not index.flags() & Qt.ItemFlag.ItemIsSelectable:
            return QSize(0, 30)
        return QSize(0, 78 if self.switcher else 60)

    def paint(self, painter, option, index):
        painter.save(); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        rect = QRectF(option.rect).adjusted(0, 2, -1, -3)
        lines = (index.data(Qt.ItemDataRole.DisplayRole) or '').split('\n')
        if not index.flags() & Qt.ItemFlag.ItemIsSelectable:
            painter.setFont(font(9 if self.switcher else 11, QFont.Weight.ExtraBold)); painter.setPen(QColor('#8ce8c3'))
            painter.drawText(rect.adjusted(7, 6, 0, 0), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, lines[0])
            painter.restore(); return
        hover = option.state & (QStyle.StateFlag.State_MouseOver | QStyle.StateFlag.State_Selected)
        fill = ('#30443f' if hover else '#222e36') if self.switcher else ('#314d4b' if hover else '#26343e')
        border = '#8ce8c3' if hover else ('#222e36' if self.switcher else '#435761')
        painter.setPen(QPen(QColor(border), 1)); painter.setBrush(QColor(fill))
        painter.drawRoundedRect(rect, 11 if self.switcher else 10, 11 if self.switcher else 10)
        icon = index.data(Qt.ItemDataRole.DecorationRole)
        x = rect.x() + (9 if self.switcher else 14)
        if isinstance(icon, QIcon) and not icon.isNull():
            size = QSize(84, 54) if self.switcher else QSize(30, 30)
            icon.paint(painter, QRect(round(x), round(rect.center().y() - size.height() / 2), size.width(), size.height()))
            x += size.width() + 12
        right = rect.right() - (30 if self.switcher else 12)
        y = rect.y() + (10 if self.switcher else 11)
        painter.setFont(font(12, QFont.Weight.Bold)); painter.setPen(QColor('#edf5f2'))
        painter.drawText(QRectF(x, y, right - x, 18), Qt.AlignmentFlag.AlignVCenter,
                         painter.fontMetrics().elidedText(lines[0], Qt.TextElideMode.ElideRight, round(right - x)))
        y += 20
        if len(lines) > 1:
            meta = lines[1]
            if self.switcher and 'TILES' in meta:
                count, _, rest = meta.partition('  ·  ')
                painter.setFont(font(9, QFont.Weight.ExtraBold)); painter.setPen(QColor('#abefcb'))
                painter.drawText(QRectF(x, y, right - x, 15), Qt.AlignmentFlag.AlignVCenter, count)
                offset = painter.fontMetrics().horizontalAdvance(count) + 7
                painter.setFont(font(10)); painter.setPen(QColor('#a2b8b9'))
                painter.drawText(QRectF(x + offset, y, right - x - offset, 15), Qt.AlignmentFlag.AlignVCenter, rest)
            else:
                painter.setFont(font(10)); painter.setPen(QColor('#a2b8b9' if self.switcher else '#93a6ad'))
                painter.drawText(QRectF(x, y, right - x, 15), Qt.AlignmentFlag.AlignVCenter,
                                 painter.fontMetrics().elidedText(meta, Qt.TextElideMode.ElideRight, round(right - x)))
            y += 17
        if len(lines) > 2:
            painter.setFont(font(10)); painter.setPen(QColor('#c1d7d2'))
            painter.drawText(QRectF(x, y, right - x, 15), Qt.AlignmentFlag.AlignVCenter,
                             painter.fontMetrics().elidedText(lines[2], Qt.TextElideMode.ElideRight, round(right - x)))
        if self.switcher:
            active = index.data(Qt.ItemDataRole.UserRole + 3)
            painter.setFont(font(9 if active else 13, QFont.Weight.ExtraBold if active else QFont.Weight.Normal))
            painter.setPen(QColor('#aaf0cb' if active else '#a2b8b9'))
            painter.drawText(QRectF(right, rect.y(), rect.right() - right - (6 if active else 8), rect.height()),
                             Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight, 'ACTIVE' if active else '›')
        painter.restore()


class CardButton(QPushButton):
    """The switcher's primary "Arrange open windows" row."""
    def paintEvent(self, event):
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        hover = self.underMouse() or self.hasFocus()
        painter.setPen(QPen(QColor('#8ce8c3' if hover else '#69cfa4'), 1))
        painter.setBrush(QColor('#30443f' if hover else '#254039'))
        rect = QRectF(self.rect()).adjusted(.5, .5, -.5, -.5)
        painter.drawRoundedRect(rect, 11, 11)
        lines = self.text().split('\n')
        x = 9
        if not self.icon().isNull():
            self.icon().paint(painter, QRect(9, round(rect.center().y() - 27), 84, 54)); x += 96
        right = rect.right() - 30
        painter.setFont(font(12, QFont.Weight.Bold)); painter.setPen(QColor('#edf5f2'))
        painter.drawText(QRectF(x, 10, right - x, 18), Qt.AlignmentFlag.AlignVCenter, lines[0])
        if len(lines) > 1:
            painter.setFont(font(10)); painter.setPen(QColor('#a2b8b9'))
            painter.drawText(QRectF(x, 30, right - x, 15), Qt.AlignmentFlag.AlignVCenter, lines[1])
        if len(lines) > 2:
            painter.setPen(QColor('#c1d7d2'))
            painter.drawText(QRectF(x, 47, right - x, 15), Qt.AlignmentFlag.AlignVCenter, lines[2])
        painter.setFont(font(13)); painter.setPen(QColor('#a2b8b9'))
        painter.drawText(QRectF(right, 0, rect.right() - right - 8, rect.height()), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight, '›')
