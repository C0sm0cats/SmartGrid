"""Small Adwaita-like controls for the system-themed Preferences window."""
from PySide6.QtCore import Qt, QRectF, QSize, QPropertyAnimation, Property
from PySide6.QtGui import QColor, QPainter
from PySide6.QtWidgets import QCheckBox, QWidget
from smartgrid.core.geometry import effective_layout, resolve_layout
from smartgrid.core.models import Rect


class Switch(QCheckBox):
    """Adw.SwitchRow toggle: rounded track, sliding knob, accent when on."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self._offset = 0.0
        self.accent = '#3584e4'
        self.slide = QPropertyAnimation(self, b'offset', self); self.slide.setDuration(120)
        self.toggled.connect(self._animate)

    def sizeHint(self):
        return QSize(48, 26)

    def hitButton(self, pos):
        return self.rect().contains(pos)

    def _animate(self, checked):
        self.slide.stop(); self.slide.setStartValue(self._offset); self.slide.setEndValue(1.0 if checked else 0.0); self.slide.start()

    def get_offset(self): return self._offset
    def set_offset(self, value): self._offset = value; self.update()
    offset = Property(float, get_offset, set_offset)

    def setChecked(self, checked):
        super().setChecked(checked); self.slide.stop(); self._offset = 1.0 if checked else 0.0; self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        dark = self.palette().window().color().lightness() < 128
        off = QColor('#545454' if dark else '#d0d0d0')
        track = QColor(self.accent) if self.isChecked() else off
        if not self.isEnabled(): track.setAlphaF(.4)
        painter.setPen(Qt.PenStyle.NoPen); painter.setBrush(track)
        painter.drawRoundedRect(QRectF(2, 2, 44, 22), 11, 11)
        painter.setBrush(QColor('#ffffff'))
        painter.drawEllipse(QRectF(4 + self._offset * 22, 4, 18, 18))


class SpacePreview(QWidget):
    """prefs.js Current space preview: plain tiles, focused and occupied shades."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedSize(220, 116)
        self.rects, self.focused, self.occupied = [], -1, set()
        self.setToolTip('Live layout of the active desktop, display and SmartGrid space')

    def set_layout(self, profile, display, settings, focused_hwnd):
        if profile is None or display is None:
            self.rects = []; self.update(); return
        area = display.work_area
        factor = min(self.width() / max(1, area.width), self.height() / max(1, area.height))
        width, height = area.width * factor, area.height * factor
        x, y = (self.width() - width) / 2, (self.height() - height) / 2
        preset, tiles, total = effective_layout(profile, runtime=True)
        padding = settings.margins if settings.independent_padding else settings.padding
        if isinstance(padding, dict): padding = {k: round(v * factor) for k, v in padding.items()}
        else: padding = round(padding * factor)
        scaled = Rect(round(x), round(y), max(1, round(width)), max(1, round(height)))
        try:
            self.rects = resolve_layout(scaled, total, preset=preset, gap=max(0, round(settings.gap * factor)),
                                        padding=padding, ratio=settings.master_ratio, tiles=tiles) if total else []
        except ValueError:
            self.rects = []
        self.occupied = {i for i, a in enumerate(profile.assignments) if a and a.window_id}
        self.focused = next((i for i, a in enumerate(profile.assignments) if a and a.window_id == focused_hwnd), -1)
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        dark = self.palette().window().color().lightness() < 128
        painter.fillRect(self.rect(), QColor.fromRgbF(*((.12, .16, .19) if dark else (.90, .93, .95))))
        for index, rect in enumerate(self.rects):
            if index == self.focused: rgb = (.29, .51, .75) if dark else (.26, .48, .70)
            elif index in self.occupied: rgb = (.37, .43, .48) if dark else (.68, .75, .80)
            else: rgb = (.23, .27, .30) if dark else (.79, .83, .86)
            painter.fillRect(rect.x, rect.y, rect.width, rect.height, QColor.fromRgbF(*rgb))


class KeyCaps(QWidget):
    """Gtk.ShortcutLabel: each key in a rounded cap, or a dimmed "Off"."""
    def __init__(self, parent=None):
        super().__init__(parent)
        self.keys = []

    def set_keys(self, keys):
        self.keys = keys
        from PySide6.QtGui import QFontMetrics
        metrics = QFontMetrics(self.font())
        width = sum(metrics.horizontalAdvance(k) + 14 for k in keys) + 5 * max(0, len(keys) - 1) if keys else metrics.horizontalAdvance('Off')
        self.setFixedSize(width + 2, 26); self.update()

    def paintEvent(self, event):
        painter = QPainter(self); painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        palette = self.palette()
        if not self.keys:
            painter.setPen(palette.placeholderText().color())
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight, 'Off'); return
        x = 1
        for key in self.keys:
            width = painter.fontMetrics().horizontalAdvance(key) + 14
            box = QRectF(x, 2, width, 22)
            painter.setPen(palette.mid().color()); painter.setBrush(palette.button())
            painter.drawRoundedRect(box, 5, 5)
            painter.setPen(palette.buttonText().color()); painter.drawText(box, Qt.AlignmentFlag.AlignCenter, key)
            x += width + 5


class ShortcutEditor(QWidget):
    """Shortcut row: key caps (or "Off"), a pencil to type the
    combination and a record button. Esc cancels recording, Backspace disables.

    Mirrors the QKeySequenceEdit API used by Preferences: keySequence(),
    setKeySequence() and keySequenceChanged.
    """
    from PySide6.QtCore import Signal as _Signal
    keySequenceChanged = _Signal(object)

    def __init__(self, sequence, title, parent=None):
        from PySide6.QtGui import QKeySequence
        from PySide6.QtWidgets import QHBoxLayout, QToolButton, QLabel, QMenu, QWidgetAction, QLineEdit
        super().__init__(parent)
        self.title, self._sequence, self.recording = title, QKeySequence(sequence), False
        layout = QHBoxLayout(self); layout.setContentsMargins(0, 0, 0, 0); layout.setSpacing(6)
        self.caps = KeyCaps()
        layout.addStretch(); layout.addWidget(self.caps)
        self.edit_button = QToolButton(); self.edit_button.setText('✎'); self.edit_button.setToolTip(f'Edit {title.lower()}')
        self.edit_button.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        menu = QMenu(self.edit_button); holder = QWidget(); row = QHBoxLayout(holder); row.setContentsMargins(12, 12, 12, 12)
        self.entry = QLineEdit(); self.entry.setPlaceholderText('Clear to disable'); self.entry.setMinimumWidth(220)
        apply = QToolButton(); apply.setText('✓'); apply.setToolTip('Apply shortcut')
        row.addWidget(self.entry); row.addWidget(apply)
        action = QWidgetAction(menu); action.setDefaultWidget(holder); menu.addAction(action)
        menu.aboutToShow.connect(lambda: (self.entry.setText(self._portable()), self.entry.setFocus()))
        commit = lambda: (self.setKeySequence(QKeySequence(self.entry.text().strip().replace('Win+', 'Meta+')), True), menu.close())
        self.entry.returnPressed.connect(commit); apply.clicked.connect(commit)
        self.edit_button.setMenu(menu)
        self.record_button = QToolButton(); self.record_button.setText('●'); self.record_button.setToolTip(f'Record {title.lower()}')
        self.record_button.clicked.connect(self._toggle_record)
        for button in (self.edit_button, self.record_button):
            button.setFixedSize(34, 30); layout.addWidget(button)
        self.edit_button.setStyleSheet('QToolButton::menu-indicator { image: none; width: 0; }')
        self.setAccessibleName(title)
        self._render()

    def _portable(self):
        from PySide6.QtGui import QKeySequence
        return self._sequence.toString(QKeySequence.SequenceFormat.PortableText).replace('Meta+', 'Win+')

    def _render(self):
        text = self._portable()
        order = {'Ctrl': 0, 'Alt': 1, 'Shift': 2, 'Win': 3}
        keys = text.split('+') if text else []
        # Modifier order: Ctrl, Alt, Shift, Win, then the key.
        self.caps.set_keys(sorted(keys[:-1], key=lambda k: order.get(k, 9)) + keys[-1:])
        self.record_button.setText('■' if self.recording else '●')

    def keySequence(self):
        return self._sequence

    def setKeySequence(self, sequence, emit=True):
        self._sequence = sequence; self._render()
        if emit: self.keySequenceChanged.emit(sequence)

    def setClearButtonEnabled(self, enabled):
        pass

    def _toggle_record(self):
        self.recording = not self.recording; self._render()
        if self.recording:
            self.record_button.setFocus(); self.record_button.installEventFilter(self)
            self.setToolTip('Press a shortcut · Esc cancels · Backspace disables')
        else:
            self.record_button.removeEventFilter(self)

    def eventFilter(self, watched, event):
        from PySide6.QtCore import QEvent
        from PySide6.QtGui import QKeySequence, QKeyCombination
        if not self.recording or event.type() != QEvent.Type.KeyPress:
            return False
        key = event.key()
        if key == Qt.Key.Key_Escape:
            self._toggle_record(); return True
        if key in (Qt.Key.Key_Backspace, Qt.Key.Key_Delete):
            self._toggle_record(); self.setKeySequence(QKeySequence()); return True
        if key in (Qt.Key.Key_Control, Qt.Key.Key_Alt, Qt.Key.Key_Shift, Qt.Key.Key_Meta):
            return True
        modifiers = event.modifiers() & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.AltModifier |
                                         Qt.KeyboardModifier.ShiftModifier | Qt.KeyboardModifier.MetaModifier)
        if not modifiers & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.AltModifier | Qt.KeyboardModifier.MetaModifier):
            return True
        self._toggle_record()
        self.setKeySequence(QKeySequence(QKeyCombination(modifiers, Qt.Key(key))))
        return True
