"""Accessible tile cards, application assignment drops and shared-edge handles."""
import json
from PySide6.QtCore import Qt, Signal, QPoint, QRect, QRectF, QMimeData
from PySide6.QtGui import QColor, QDrag, QFont, QIcon, QPainter, QPen, QPixmap
from PySide6.QtWidgets import QAbstractButton, QApplication, QWidget, QStyle
from smartgrid.core.models import Rect,app_display_name
from smartgrid.core.geometry import capacity, auto_preset, resolve_layout, directional_neighbor, effective_layout, move_edge
from .theme import tokens

MIME = "application/x-smartgrid-assignment"


def layout_icon(preset, settings, tiles=None):
    """Small geometry previews, rather than names alone, for layout choices."""
    pixmap=QPixmap(64,40)
    pixmap.fill(Qt.GlobalColor.transparent)
    painter=QPainter(pixmap)
    painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(QColor(tokens(settings)['accent']))
    if preset=='custom' and not tiles:
        painter.drawRoundedRect(QRectF(4,4,34,32),3,3)
        painter.drawRoundedRect(QRectF(41,4,19,14),3,3)
        painter.drawRoundedRect(QRectF(41,21,19,15),3,3)
    else:
        choice='2x2' if preset=='auto' else preset
        for rect in resolve_layout(Rect(2,2,60,36),capacity(choice,tiles),choice,gap=2,padding=0,tiles=tiles):
            painter.drawRoundedRect(QRectF(rect.x,rect.y,rect.width,rect.height),2,2)
    painter.end()
    return QIcon(pixmap)


def application_icon(app):
    if app and getattr(app, 'icon_path', ''):
        icon = QIcon(app.icon_path)
        if not icon.isNull():
            return icon
    return QApplication.style().standardIcon(QStyle.StandardPixmap.SP_ComputerIcon)


class TileCard(QAbstractButton):
    selected = Signal(int)
    dropped = Signal(int, object)
    context_requested = Signal(int, object)
    keyboard_action = Signal(int, str)

    def __init__(self, index, settings, parent=None):
        super().__init__(parent)
        self.index, self.settings = index, settings
        self.app_name, self.title, self.badges = "", "", []
        self.icon = application_icon(None)
        self.assignment = None
        self.preview_active = False
        self.merge_target = False
        self.is_selected = self.drop_target = self.dragging = False
        self._press = QPoint()
        self.setAcceptDrops(True)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.clicked.connect(lambda: self.selected.emit(self.index))
        self.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.customContextMenuRequested.connect(lambda p: self.context_requested.emit(self.index, self.mapToGlobal(p)))

    def describe(self, assignment, app, window=None, shared=(), pending=False):
        self.assignment = assignment
        self.window=window
        self.app_name = (app.name if app else app_display_name(assignment.app_id)) if assignment else 'Choose an app'
        # Reserved-tile captions: a draft app opens on Apply; a pin
        # reports the state of its reserved application.
        if assignment and assignment.pinned and window and (window.state == 'minimized' or window.floating):
            # The Studio shows a pinned window that is away as its reserved tile.
            self.title = 'PINNED · ' + ('FLOATING' if window.floating else 'MINIMIZED')
            window = self.window = None
        elif window:
            self.title = window.title or 'Window'
        elif not assignment:
            self.title = ''
        elif assignment.pinned:
            self.title = 'PINNED · ' + ('OPENING' if pending else 'CLOSED' if app else 'UNAVAILABLE')
        else:
            self.title = 'OPENING' if pending else 'OPENS ON APPLY'
        self.badges = []
        if window and window.ref.hwnd == getattr(self.parent(), 'active_hwnd', None):
            self.badges.append('ACTIVE')
        if window and assignment and assignment.pinned:
            self.badges.append('PINNED')
        if shared:
            self.badges.append('SHARED')
        if window and (window.state != 'normal' or window.floating or not window.eligible):
            state = {'minimized': 'MINIMIZED', 'parked': 'PARKED', 'fullscreen': 'FULLSCREEN',
                     'maximized': 'MAXIMIZED', 'cloaked': 'OTHER DESKTOP'}.get(window.state, window.state.upper())
            self.badges.append('FLOATING' if window.floating else state)
        self.icon = application_icon(app)
        if not assignment:
            description = f"Tile {self.index + 1}: Choose an app"
        elif window:
            description = f"Tile {self.index + 1}: {self.app_name}"
            if shared:
                description += ". Shared: " + ', '.join(f'Space {s+1}' for s in shared)
        elif assignment.pinned:
            description = f"Tile {self.index + 1}: {self.app_name}, pinned, {self.title.split(' · ')[1].lower()}"
        else:
            description = f"Tile {self.index + 1}: {self.app_name}, opens on Apply"
        if window and window.exclusion_reason:
            description += ". " + window.exclusion_reason
        self.setAccessibleName(description)
        self.setAccessibleDescription("Enter assigns, Delete clears, Shift+F10 opens the menu. Arrows move the selection; Ctrl+arrow swaps.")
        self.setToolTip(description)
        self.update()

    BADGES = {'ACTIVE': ('#244b59', None, '#aeeaff'), 'SHARED': ('#284055', None, '#a8cfff'),
              'PINNED': ('#24553a', '#57e389', '#bfffd2')}

    @staticmethod
    def _font(pixels, weight=QFont.Weight.Normal):
        font = QFont("Segoe UI"); font.setPixelSize(pixels); font.setWeight(weight)
        return font

    def paintEvent(self, event):
        """Tile card: centred number, icon, app name, window title and badge pills."""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        empty = not self.assignment
        hover = self.underMouse() and self.isEnabled() and not self.is_selected
        fill, line, width, dashed = ('#182129', '#53646d', 1, True) if empty else ('#26343e', '#435761', 1, False)
        if hover:
            fill, line = ('#23403d' if empty else '#314d4b'), '#8ce8c3'
        if self.is_selected:
            fill, line, width, dashed = '#513a3f', '#ff9b91', 2, False
        if self.drop_target:
            fill, line, width, dashed = '#356358', '#b5f7db', 2, False
        if self.dragging:
            fill, line, width, dashed = '#182129', '#8ce8c3', 1, True
        if self.merge_target:
            fill, line, width, dashed = '#51412a', '#f3bd72', 3, True
        pen = QPen(QColor(line), width)
        if dashed: pen.setStyle(Qt.PenStyle.DashLine)
        painter.setPen(pen); painter.setBrush(QColor(fill))
        painter.drawRoundedRect(QRectF(self.rect()).adjusted(1.5, 1.5, -1.5, -1.5), 12, 12)
        if self.hasFocus() and not self.is_selected:
            painter.setPen(QPen(QColor('#8ce8c3'), 1, Qt.PenStyle.DotLine)); painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRoundedRect(QRectF(self.rect()).adjusted(5, 5, -5, -5), 9, 9)
        w, h = self.width(), self.height()
        number = f'{self.index + 1:02d}'
        if self.dragging:
            painter.setPen(QColor('#8ce8c3')); painter.setFont(self._font(10, QFont.Weight.ExtraBold))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, 'SOURCE'); return
        if self.drop_target:
            # Drop preview: the dragged app replaces the tile content.
            source = getattr(self, 'drop_source', None)
            if source:
                box = QRectF((w - min(w - 16, 180)) / 2, (h - 86) / 2, min(w - 16, 180), 86)
                painter.setPen(QPen(QColor('#8ce8c3'), 1)); painter.setBrush(QColor(24, 46, 43, 235))
                painter.drawRoundedRect(box, 9, 9)
                source[0].paint(painter, QRect(round(box.center().x() - 14), round(box.y() + 9), 28, 28))
                painter.setFont(self._font(11, QFont.Weight.Bold)); painter.setPen(QColor('#ecf4f2'))
                painter.drawText(QRectF(box.x() + 6, box.y() + 42, box.width() - 12, 18), Qt.AlignmentFlag.AlignCenter,
                                 painter.fontMetrics().elidedText(source[1], Qt.TextElideMode.ElideRight, round(box.width() - 12)))
                painter.setFont(self._font(8, QFont.Weight.ExtraBold)); painter.setPen(QColor('#b5f7db'))
                painter.drawText(QRectF(box.x(), box.y() + 62, box.width(), 14), Qt.AlignmentFlag.AlignCenter, 'DROP HERE')
                return
            painter.setPen(QColor('#b5f7db')); painter.setFont(self._font(8, QFont.Weight.ExtraBold))
            painter.drawText(self.rect().adjusted(0, 0, 0, -10), Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignBottom, 'DROP HERE')
        label = 'Empty' if empty and not self.isEnabled() else self.app_name
        reserved = bool(self.assignment) and not self.window
        dense = h < 56 or (reserved and h < 100)
        if dense:
            # Horizontal row: number, icon, name.
            painter.setFont(self._font(10, QFont.Weight.ExtraBold)); painter.setPen(QColor('#8ce8c3'))
            metrics = painter.fontMetrics(); x = 10
            painter.drawText(QRect(x, 0, 24, h), Qt.AlignmentFlag.AlignVCenter, number); x += metrics.horizontalAdvance(number) + 6
            if self.assignment and h >= 30:
                self.icon.paint(painter, QRect(x, (h - 16) // 2, 16, 16)); x += 21
            painter.setFont(self._font(10, QFont.Weight.Bold))
            painter.setPen(QColor('#83969e' if empty else '#dce9e5' if reserved else '#ecf4f2'))
            painter.drawText(QRect(x, 0, w - x - 8, h), Qt.AlignmentFlag.AlignVCenter,
                             painter.fontMetrics().elidedText(label, Qt.TextElideMode.ElideRight, max(10, w - x - 8)))
            return
        compact = h < 130
        minimal = w < 120 or h < 88
        items = [('number', 14 if not compact else 13)]
        icon_size = 0
        if self.assignment and not self.preview_active and not minimal:
            icon_size = (20 if h < 100 or w < 130 else 30) if reserved else (26 if w < 170 or h < 170 else 36)
            items.append(('icon', icon_size))
        if self.preview_active:
            items.append(('preview', max(0, h - 110)))
        items.append(('name', 16))
        if self.title and not (w < 170 or h < 170) and not reserved and not self.preview_active:
            items.append(('title', 14))
        if reserved and self.title:
            items.append(('status', 12))
        badges = [b for b in self.badges] if not minimal else []
        if badges:
            items.append(('badges', 17))
        spacing = 2 if compact else 7
        total = sum(size for _, size in items) + spacing * (len(items) - 1)
        y = max(4, (h - total) / 2)
        for kind, size in items:
            box = QRectF(8, y, w - 16, size)
            if kind == 'number':
                painter.setFont(self._font(11 if compact else 14, QFont.Weight.ExtraBold)); painter.setPen(QColor('#8ce8c3'))
                painter.drawText(box, Qt.AlignmentFlag.AlignCenter, number)
            elif kind == 'icon':
                self.icon.paint(painter, QRect(round((w - size) / 2), round(y), size, size))
            elif kind == 'name':
                if empty:
                    painter.setFont(self._font(11)); painter.setPen(QColor('#83969e'))
                else:
                    painter.setFont(self._font(11 if reserved else 12, QFont.Weight.Bold))
                    painter.setPen(QColor('#dce9e5' if reserved else '#ecf4f2'))
                painter.drawText(box, Qt.AlignmentFlag.AlignCenter,
                                 painter.fontMetrics().elidedText(label, Qt.TextElideMode.ElideRight, max(10, w - 28)))
            elif kind == 'title':
                painter.setFont(self._font(10)); painter.setPen(QColor('#a8b8be'))
                painter.drawText(box, Qt.AlignmentFlag.AlignCenter,
                                 painter.fontMetrics().elidedText(self.title, Qt.TextElideMode.ElideRight, max(10, w - 36)))
            elif kind == 'status':
                painter.setFont(self._font(9, QFont.Weight.ExtraBold)); painter.setPen(QColor('#8ce8c3'))
                painter.drawText(box, Qt.AlignmentFlag.AlignCenter,
                                 painter.fontMetrics().elidedText(self.title, Qt.TextElideMode.ElideRight, max(10, w - 28)))
            elif kind == 'badges':
                self._paint_badges(painter, badges, y, w)
            y += size + spacing

    def _paint_badges(self, painter, badges, y, width):
        painter.setFont(self._font(8, QFont.Weight.ExtraBold))
        metrics = painter.fontMetrics()
        sizes = [metrics.horizontalAdvance(b) + 12 for b in badges]
        x = (width - sum(sizes) - 4 * (len(badges) - 1)) / 2
        for badge, size in zip(badges, sizes):
            background, border, text = self.BADGES.get(badge, ('#3a3324', '#f3bd72', '#f3bd72'))
            painter.setPen(QPen(QColor(border), 1) if border else Qt.PenStyle.NoPen)
            painter.setBrush(QColor(background))
            box = QRectF(x, y, size, 17)
            painter.drawRoundedRect(box, 7, 7)
            painter.setPen(QColor(text)); painter.drawText(box, Qt.AlignmentFlag.AlignCenter, badge)
            x += size + 4

    def enterEvent(self, event):
        self.update(); super().enterEvent(event)

    def leaveEvent(self, event):
        self.update(); super().leaveEvent(event)

    def mousePressEvent(self, event):
        self._press = event.position().toPoint()
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.assignment and event.buttons() & Qt.MouseButton.LeftButton and (event.position().toPoint()-self._press).manhattanLength() >= QApplication.startDragDistance():
            drag = QDrag(self)
            mime = QMimeData()
            mime.setData(MIME, json.dumps({'tile': self.index}).encode())
            drag.setMimeData(mime)
            drag.setPixmap(self.drag_ghost())
            drag.setHotSpot(self._press)
            self.dragging = True
            self.update()
            drag.exec(Qt.DropAction.MoveAction)
            self.dragging = False
            self.setDown(False)
            self.update()
            return
        super().mouseMoveEvent(event)

    def drag_ghost(self):
        """Drag ghost: green pill with the app icon and name."""
        font = self._font(12, QFont.Weight.Bold)
        from PySide6.QtGui import QFontMetrics
        width = min(320, QFontMetrics(font).horizontalAdvance(self.app_name) + 24 + 9 + 32)
        pixmap = QPixmap(width, 48); pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(QPen(QColor('#8ce8c3'), 1)); painter.setBrush(QColor('#28483f'))
        painter.drawRoundedRect(QRectF(.5, .5, width - 1, 47), 12, 12)
        self.icon.paint(painter, QRect(16, 12, 24, 24))
        painter.setFont(font); painter.setPen(QColor('white'))
        painter.drawText(QRect(49, 0, width - 57, 48), Qt.AlignmentFlag.AlignVCenter,
                         painter.fontMetrics().elidedText(self.app_name, Qt.TextElideMode.ElideRight, width - 57))
        painter.end()
        self._press = QPoint(24, 24)
        return pixmap

    def dragEnterEvent(self, event):
        if event.mimeData().hasFormat(MIME):
            try:
                source = json.loads(bytes(event.mimeData().data(MIME))).get('tile')
            except (ValueError, TypeError, AttributeError):
                source = None
            canvas = self.parent()
            card = canvas.cards[source] if isinstance(source, int) and 0 <= source < len(getattr(canvas, 'cards', [])) else None
            self.drop_source = (card.icon, card.app_name) if card is not None and card is not self else None
            self.drop_target = True
            event.acceptProposedAction()
            self.update()

    def dragLeaveEvent(self, event):
        self.drop_target = False
        self.update()

    def dropEvent(self, event):
        self.drop_target = False
        try:
            payload = json.loads(bytes(event.mimeData().data(MIME)))
        except (ValueError, TypeError):
            event.ignore()
        else:
            self.dropped.emit(self.index, payload)
            event.acceptProposedAction()
        self.update()

    def keyPressEvent(self, event):
        directions = {Qt.Key.Key_Left:'left', Qt.Key.Key_Right:'right', Qt.Key.Key_Up:'up', Qt.Key.Key_Down:'down'}
        if event.key() in directions:
            self.keyboard_action.emit(self.index, ('swap_' if event.modifiers() & Qt.KeyboardModifier.ControlModifier else 'select_') + directions[event.key()])
            return
        if event.key() in (Qt.Key.Key_Delete,Qt.Key.Key_Backspace):
            self.keyboard_action.emit(self.index, 'clear')
            return
        if event.key() in (Qt.Key.Key_Return,Qt.Key.Key_Enter):
            self.selected.emit(self.index)
            self.keyboard_action.emit(self.index, 'assign')
            return
        if event.key() == Qt.Key.Key_F10 and event.modifiers() & Qt.KeyboardModifier.ShiftModifier:
            self.context_requested.emit(self.index,self.mapToGlobal(self.rect().center()))
            return
        super().keyPressEvent(event)


class EdgeHandle(QWidget):
    moved = Signal(int, str, float)
    def __init__(self,index,edge,canvas):
        super().__init__(canvas)
        self.index,self.edge,self.canvas=index,edge,canvas
        self.setCursor(Qt.CursorShape.SizeHorCursor if edge in ('left','right') else Qt.CursorShape.SizeVerCursor)
        self.setToolTip("Drag this shared divider. Neighboring tiles follow; Escape cancels.")
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.setAccessibleName(f"{edge.title()} divider of tile {index+1}")
        self.origin=None

    def enterEvent(self,event):
        self.update();super().enterEvent(event)

    def leaveEvent(self,event):
        self.update();super().leaveEvent(event)

    def paintEvent(self,event):
        # Divider: a rounded green grip, lighter on hover or drag.
        painter=QPainter(self);painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor('#b5f7db' if self.underMouse() or self.origin is not None or self.hasFocus() else '#73cfa8'))
        if self.edge in ('left','right'):
            length=min(44,self.height()-4);painter.drawRoundedRect(QRectF(self.width()/2-3,(self.height()-length)/2,6,length),3,3)
        else:
            length=min(44,self.width()-4);painter.drawRoundedRect(QRectF((self.width()-length)/2,self.height()/2-3,length,6),3,3)

    def mousePressEvent(self,event):
        self.origin=event.globalPosition()
        self.setFocus()

    def _fraction(self,position):
        delta=position-self.origin
        vertical=self.edge in ('left','right')
        length=self.canvas.work_rect.width() if vertical else self.canvas.work_rect.height()
        return (delta.x() if vertical else delta.y())/max(1,length)

    def mouseMoveEvent(self,event):
        # Resize the tiles live while the divider is dragged.
        if self.origin is not None:
            self.canvas.preview_edge(self.index,self.edge,self._fraction(event.globalPosition()))

    def mouseReleaseEvent(self,event):
        if self.origin is not None:
            fraction=self._fraction(event.globalPosition())
            self.cancel()
            if abs(fraction)>1e-4:
                self.moved.emit(self.index,self.edge,fraction)
            return
        self.cancel()

    def cancel(self):
        self.origin=None
        self.canvas.end_preview()

    def keyPressEvent(self,event):
        if event.key()==Qt.Key.Key_Escape:
            self.cancel()
        elif event.key() in (Qt.Key.Key_Left,Qt.Key.Key_Up,Qt.Key.Key_Right,Qt.Key.Key_Down):
            self.moved.emit(self.index,self.edge, -.01 if event.key() in (Qt.Key.Key_Left,Qt.Key.Key_Up) else .01)
        else:
            super().keyPressEvent(event)


class TileCanvas(QWidget):
    selected = Signal(int)
    dropped = Signal(int, object)
    context_requested = Signal(int, object)
    keyboard_action = Signal(int, str)
    edge_moved = Signal(int, str, float)
    tiles_previewed = Signal(object)

    def __init__(self,settings,parent=None):
        super().__init__(parent)
        self.settings=settings
        self.cards=[]
        self.handles=[]
        self.profile=None
        self.ratio=16/9
        self.selected_index=0
        self.handle_preview=None
        self._drag_base=None
        self.active_hwnd=None
        self.work_rect=QRect(0,0,1,1)
        self.rectangles=[]
        self.setMinimumSize(260,180)
        self.setAccessibleName("Layout grid")

    def set_profile(self,profile,display,apps,windows,shared=None,pending=None,runtime=False):
        self.profile=profile
        self.native_area=display.work_area if display else Rect(0,0,1920,1080)
        if display:
            self.ratio=display.work_area.width/max(1,display.work_area.height)
        self.effective_preset,self.effective_tiles,count=effective_layout(profile,runtime=runtime)
        # An empty current space still exposes one drop target for an assignment.
        count=max(1,count)
        while len(self.cards)>count:
            card=self.cards.pop();card.hide();card.deleteLater()
        while len(self.cards)<count:
            card=TileCard(len(self.cards),self.settings,self)
            card.selected.connect(self.selected)
            card.dropped.connect(self.dropped)
            card.context_requested.connect(self.context_requested)
            card.keyboard_action.connect(self.keyboard_action)
            self.cards.append(card)
            card.show()
        apps_by_id={a.id:a for a in apps}
        windows_by_id={w.ref.hwnd:w for w in windows}
        self.selected_index=min(self.selected_index,count-1)
        used={a.window_id for a in profile.assignments if a and a.window_id in windows_by_id}
        for i,card in enumerate(self.cards):
            assignment=profile.assignments[i] if i<len(profile.assignments) else None
            card.settings=self.settings
            card.is_selected=i==self.selected_index
            window=windows_by_id.get(assignment.window_id) if assignment else None
            if assignment and not window and not runtime:
                candidates=[w for w in windows if w.app_id==assignment.app_id and w.eligible and w.ref.hwnd not in used]
                window=next(iter(sorted(candidates,key=lambda w:(w.display_id!=profile.display_id,w.ref.hwnd))),None)
                if window: used.add(window.ref.hwnd)
            card.describe(assignment,apps_by_id.get(assignment.app_id) if assignment else None,
                          window,
                          (shared or {}).get(i,()),i in (pending or ()))
        self._layout_cards()

    def select(self,index,focus=False):
        self.selected_index=index
        for i,card in enumerate(self.cards):
            card.is_selected=i==index
            card.update()
        if focus and 0<=index<len(self.cards):
            self.cards[index].setFocus()

    def _layout_cards(self):
        if not self.profile:
            return
        width,height=max(1,self.width()-24),max(1,self.height()-24)
        if width/height>self.ratio:
            width=round(height*self.ratio)
        else:
            height=round(width/self.ratio)
        self.work_rect=QRect((self.width()-width)//2,(self.height()-height)//2,width,height)
        padding=self.settings.margins if self.settings.independent_padding else self.settings.padding
        area=self.native_area
        try:
            native_rects=resolve_layout(area,len(self.cards),self.effective_preset,gap=self.settings.gap,
                padding=padding,ratio=self.profile.ratio,tiles=self.effective_tiles)
            self.rectangles=[Rect(self.work_rect.x()+round((r.x-area.x)*width/area.width),
                self.work_rect.y()+round((r.y-area.y)*height/area.height),
                max(1,round(r.width*width/area.width)),max(1,round(r.height*height/area.height))) for r in native_rects]
        except (ValueError,TypeError) as error:
            self.rectangles=[]
            self.setToolTip(str(error))
        for card in self.cards:
            card.setVisible(bool(self.rectangles))
        for card,rect in zip(self.cards,self.rectangles):
            card.setGeometry(rect.x,rect.y,max(1,rect.width),max(1,rect.height))
        if self._drag_base is not None and len(self.handles):
            self._place_handles()
            self.update()
            return
        for handle in self.handles:
            handle.deleteLater()
        self.handles=[]
        if self.effective_preset=='custom':
            for i,tile in enumerate(self.effective_tiles):
                if i>=len(self.rectangles):
                    continue
                rect=self.rectangles[i]
                for edge,bound in [('right',tile.x+tile.width),('bottom',tile.y+tile.height)]:
                    if bound>=.999:
                        continue
                    handle=EdgeHandle(i,edge,self)
                    if edge=='right':
                        handle.setGeometry(rect.right-5,rect.y+8,10,max(12,rect.height-16))
                    else:
                        handle.setGeometry(rect.x+8,rect.bottom-5,max(12,rect.width-16),10)
                    handle.moved.connect(self.edge_moved)
                    self.handles.append(handle)
                    handle.show()
                    handle.raise_()
        self.update()

    def _handle_geometry(self,index,edge):
        rect=self.rectangles[index]
        if edge=='right':
            return rect.right-5,rect.y+8,10,max(12,rect.height-16)
        return rect.x+8,rect.bottom-5,max(12,rect.width-16),10

    def _place_handles(self):
        for handle in self.handles:
            if handle.index<len(self.rectangles):
                handle.setGeometry(*self._handle_geometry(handle.index,handle.edge))

    def preview_edge(self,index,edge,fraction):
        """Live divider drag: move the edge on a copy of the tiles and relayout."""
        if self._drag_base is None:
            self._drag_base=list(self.effective_tiles)
        try:
            tiles=move_edge(self._drag_base,index,edge,fraction)
        except (ValueError,IndexError):
            return
        if any(t.width<.03 or t.height<.03 for t in tiles):
            return
        self.effective_tiles=tiles
        self._layout_cards()
        self.tiles_previewed.emit(tiles)

    def end_preview(self):
        if self._drag_base is not None:
            self.effective_tiles=self._drag_base
            self._drag_base=None
            self._layout_cards()

    def neighbor(self,index,direction):
        return directional_neighbor(self.rectangles,index,direction)

    def resizeEvent(self,event):
        self._layout_cards()
        super().resizeEvent(event)

    def paintEvent(self,event):
        painter=QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        t=tokens(self.settings)
        painter.setPen(QPen(QColor('#2d3b44' if self.settings.theme=='dark' else t['line']),1))
        painter.setBrush(QColor(t['canvas']))
        painter.drawRoundedRect(QRectF(self.work_rect.adjusted(-8,-8,8,8)),17,17)
        if self.handle_preview:
            edge,global_pos=self.handle_preview
            pos=self.mapFromGlobal(global_pos.toPoint())
            painter.setPen(QPen(QColor(t['accent']),2,Qt.PenStyle.DashLine))
            if edge in ('left','right'):
                painter.drawLine(pos.x(),self.work_rect.top(),pos.x(),self.work_rect.bottom())
            else:
                painter.drawLine(self.work_rect.left(),pos.y(),self.work_rect.right(),pos.y())
