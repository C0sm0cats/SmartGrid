"""SmartGrid Studio: independent current drafts and durable template editing."""
from copy import deepcopy
from uuid import uuid4
from PySide6.QtCore import Qt, Signal, QSignalBlocker, QSize, QTimer, QPoint
from PySide6.QtGui import QAction, QKeySequence
from PySide6.QtWidgets import (QMainWindow,QWidget,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QToolButton,
    QComboBox,QSplitter,QScrollArea,QStackedWidget,QGroupBox,QFormLayout,QDoubleSpinBox,QCheckBox,QMenu,QInputDialog,
    QMessageBox,QFileDialog,QDialog,QDialogButtonBox,QTreeWidget,QTreeWidgetItem,QLineEdit,QGridLayout,QListWidget,QListWidgetItem)
from smartgrid.core.models import Assignment,SpaceProfile,Tile,Rect,app_display_name
from smartgrid.core.geometry import presets,preset_name,capacity,auto_preset,resolve_layout,split_tile,merge_tiles,move_edge,set_tile_geometry
from smartgrid.core.history import History
from .bridge import bridge_for
from .tile_view import TileCanvas, layout_icon
from .library import ApplicationLibrary
from .theme import apply_theme,studio_settings
from .backdrop import show_backdrops,hide_backdrops
from .dialog_surface import paint_surface,animate_open


def save_icon():
    """Monochrome save icon."""
    from PySide6.QtGui import QPixmap,QPainter,QPen,QColor,QIcon
    from PySide6.QtCore import QRectF,QPointF
    pixmap=QPixmap(32,32);pixmap.fill(Qt.GlobalColor.transparent)
    painter=QPainter(pixmap);painter.setRenderHint(QPainter.RenderHint.Antialiasing)
    painter.setPen(QPen(QColor('#c1cdd1'),2.4));painter.setBrush(Qt.BrushStyle.NoBrush)
    painter.drawRoundedRect(QRectF(5,5,22,22),3,3)
    painter.drawRect(QRectF(10,5,12,7));painter.drawRoundedRect(QRectF(10,17,12,10),1,1)
    painter.end()
    return QIcon(pixmap)


class Studio(QMainWindow):
    preferences_requested=Signal()

    def __init__(self,controller,display_id=None,parent=None):
        super().__init__(parent)
        self.controller=controller
        self.bridge=bridge_for(controller)
        self.display_id=display_id or controller.current_display_id()
        self.space=controller.active_spaces.get(self.display_id,0)
        self.mode='current'
        self.template_id=None
        self.editing_saved=False
        self.saved_picker_open=True
        self.local_profiles={}
        self.local_histories={}
        self.local_dirty=set()
        self._refreshing=False
        self._thumbnails=None
        self.selected_index=-1
        self.merge_candidates=[]
        self._merge_pending=None
        self.setWindowTitle("SmartGrid · Layout Studio")
        self.setWindowFlags(Qt.WindowType.Tool|Qt.WindowType.FramelessWindowHint|Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setMinimumSize(560,480)
        self.resize(1040,760)
        shell=QWidget()
        self.setCentralWidget(shell)
        layout=QVBoxLayout(shell)
        layout.setContentsMargins(20,16,20,16)
        layout.setSpacing(10)
        header=QHBoxLayout()
        brand=QLabel('<span style="font-size:11px;letter-spacing:3px;color:#8ce8c3;font-weight:800">SMARTGRID</span><br><span style="font-size:27px;font-weight:800">A place for every window.</span>')
        self.brand=brand
        brand.setStyleSheet("font-size: 22px; font-weight: bold;")
        header.addWidget(brand)
        header.addStretch()
        self.toggle_button=QPushButton("Arrange windows")
        self.toggle_button.setAccessibleName("Arrange windows")
        self.toggle_button.clicked.connect(lambda:self.bridge.submit('handle_action','toggle'))
        self.toggle_button.setParent(shell)
        self.toggle_button.hide()
        self.undo_button=QPushButton("Undo")
        self.undo_button.setToolTip("Undo · Ctrl+Z")
        self.undo_button.clicked.connect(self.undo)
        self.redo_button=QPushButton("Redo")
        self.redo_button.setToolTip("Redo · Ctrl+Shift+Z")
        self.redo_button.clicked.connect(self.redo)
        preferences=QPushButton("Preferences")
        preferences.clicked.connect(self._preferences)
        preferences.setParent(shell);preferences.hide()
        # Brand icon; Cancel/Close and Escape close the Studio.
        from .app import DesktopUI
        mark=QLabel();mark.setPixmap(DesktopUI.tray_icon('#8ce8c3').pixmap(36,36));mark.setAccessibleName('SmartGrid')
        header.addWidget(mark,0,Qt.AlignmentFlag.AlignTop)
        layout.addLayout(header)
        # Context chips: Display 1 … · Space 1 Space 2 Space 3.
        self.context_row=QWidget();context=QHBoxLayout(self.context_row);context.setContentsMargins(0,0,0,0)
        self.display_combo=QComboBox(self.context_row)
        self.display_combo.setAccessibleName("Display")
        self.display_combo.currentIndexChanged.connect(self._display_changed)
        self.display_combo.hide()
        self.display_chips=QHBoxLayout();self.display_chips.setSpacing(6);self.display_buttons=[]
        context.addLayout(self.display_chips)
        self.context_separator=QLabel("·");self.context_separator.setProperty('muted',True);context.addWidget(self.context_separator)
        self.space_buttons=[]
        for index in range(3):
            button=QPushButton(f"Space {index+1}")
            button.setCheckable(True)
            button.clicked.connect(lambda checked=False,s=index:self._space_changed(s))
            self.space_buttons.append(button)
            context.addWidget(button)
        self.runtime_status=QLabel(self.context_row)
        self.runtime_status.hide()
        context.addStretch()
        context.addWidget(self.undo_button);context.addWidget(self.redo_button)
        layout.addWidget(self.context_row)
        modes=QHBoxLayout()
        self.mode_combo=QComboBox()
        self.mode_combo.addItem("Current space",'current')
        self.mode_combo.addItem("Saved layouts",'edit')
        self.mode_combo.addItem("New layout",'new')
        self.mode_combo.setAccessibleName("Studio mode")
        self.mode_combo.currentIndexChanged.connect(self._mode_changed)
        self.mode_combo.setParent(shell)
        self.mode_combo.hide()
        self.mode_buttons={}
        for key,label in [('current','Current space'),('edit','Saved layouts ▾'),('new','New layout')]:
            button=QPushButton(label);button.setCheckable(True)
            button.clicked.connect(lambda checked=False,k=key:self._activate_mode(k))
            self.mode_buttons[key]=button;modes.addWidget(button)
        self.template_combo=QComboBox()
        self.template_combo.setMinimumWidth(180)
        self.template_combo.setAccessibleName("Saved layouts")
        self.template_combo.currentIndexChanged.connect(self._template_changed)
        self.template_search=QLineEdit()
        self.template_search.setPlaceholderText("Search saved layouts")
        self.template_search.setAccessibleName("Search saved layouts")
        self.template_search.textChanged.connect(self.refresh_from_controller)
        self.saved_controls=QWidget()
        saved_container=QVBoxLayout(self.saved_controls);saved_container.setContentsMargins(0,0,0,0)
        saved_selection=QHBoxLayout();saved_container.addLayout(saved_selection)
        saved_selection.addWidget(self.template_search)
        self.template_combo.setParent(self.saved_controls);self.template_combo.hide()
        self.saved_list=QListWidget();self.saved_list.setMaximumHeight(180);self.saved_list.setMinimumHeight(100)
        self.saved_list.setIconSize(QSize(60,36));self.saved_list.currentItemChanged.connect(self._saved_selected)
        self._saved_signature=None;saved_container.addWidget(self.saved_list)
        self.saved_list.itemClicked.connect(self._close_saved_picker)
        self.saved_list.itemActivated.connect(self._close_saved_picker)
        self.saved_list.setProperty('chooser',True)
        saved_layout=QHBoxLayout();saved_container.addLayout(saved_layout)
        self.saved_summary=QLabel();self.saved_summary.setProperty('muted',True)
        saved_layout.addWidget(self.saved_summary,1)
        self.rename_button=QPushButton("Rename layout")
        self.rename_button.clicked.connect(self.rename_template)
        self.edit_button=QPushButton('Edit layout')
        self.edit_button.clicked.connect(self.edit_saved)
        self.delete_button=QPushButton("Delete layout")
        self.delete_button.setProperty('danger',True)
        self.delete_button.clicked.connect(self.delete_template)
        self.undo_delete_button=QPushButton("Undo delete")
        self.undo_delete_button.clicked.connect(lambda:self.bridge.submit('undo_delete_template'))
        for button in (self.undo_delete_button,self.rename_button,self.edit_button,self.delete_button):
            saved_layout.addWidget(button)
        layout.addLayout(modes)
        layout.addWidget(self.saved_controls)
        toolbar=QHBoxLayout()
        self.preset_combo=QComboBox()
        self.preset_combo.setIconSize(QSize(64,40))
        for key,name in presets():
            self.preset_combo.addItem(layout_icon(key,controller.settings),name,key)
        self.preset_combo.setAccessibleName("Layout")
        self.preset_combo.currentIndexChanged.connect(self._preset_changed)
        self.preset_combo.setParent(shell);self.preset_combo.hide()
        self.preset_strip=QWidget()
        self.preset_grid=QGridLayout(self.preset_strip);self.preset_grid.setContentsMargins(0,0,0,0);self.preset_grid.setSpacing(5)
        self.preset_buttons={}
        for index in range(self.preset_combo.count()):
            key=self.preset_combo.itemData(index)
            button=QPushButton(preset_name(key))
            # Every preset except Auto is offered when composing a layout.
            if key=='auto': button.setParent(self.preset_strip);button.hide()
            button.setCheckable(True)
            button.setMinimumWidth(button.sizeHint().width())
            button.clicked.connect(lambda checked=False,k=key:self.preset_combo.setCurrentIndex(self.preset_combo.findData(k)))
            self.preset_buttons[key]=button
            if key!='auto': self.preset_grid.addWidget(button,(index-1)//9,(index-1)%9)
        self.current_layout=QLabel()
        self.current_layout.setAccessibleName("Current layout")
        self.current_layout.setStyleSheet('color:#8ce3c1;font-size:13px;font-weight:bold;')
        toolbar.addWidget(self.current_layout)
        self.ratio=QDoubleSpinBox()
        self.ratio.setRange(25,75)
        self.ratio.setSuffix(" %")
        self.ratio.setAccessibleName("Focus layout: large tile width")
        self.ratio.editingFinished.connect(self._ratio_changed)
        # The Focus ratio is a preference, not a Studio control.
        self.ratio.setParent(shell);self.ratio.hide()
        toolbar.addStretch()
        self.preview_button=QToolButton()
        self.preview_button.setText("Window previews")
        self.preview_button.setCheckable(True)
        self.preview_button.setEnabled(hasattr(controller.backend,'api'))
        self.preview_button.toggled.connect(self._refresh_thumbnails)
        self.preview_button.toggled.connect(lambda shown:self.preview_button.setText("Hide previews" if shown else "Window previews"))
        modes.addStretch()
        modes.addWidget(self.preview_button)
        self.overview_button=QToolButton(shell)
        self.overview_button.setCheckable(True)
        self.overview_button.toggled.connect(self._show_overview)
        self.overview_button.hide()
        self.library_button=QPushButton("Choose an app")
        self.library_button.setCheckable(True)
        self.library_button.toggled.connect(self._show_library)
        layout.addLayout(toolbar)
        layout.addWidget(self.preset_strip)
        # Below 820 px a single menu button replaces the preset chips.
        self.preset_menu_button=QPushButton();self.preset_menu=QMenu(self.preset_menu_button)
        self.preset_menu_button.setMenu(self.preset_menu);self.preset_menu_button.hide()
        self.preset_menu_button.setStyleSheet('QPushButton::menu-indicator { image: none; width: 0; }')
        for key,button in self.preset_buttons.items():
            if key!='auto':
                self.preset_menu.addAction(button.text(),lambda k=key:self.preset_combo.setCurrentIndex(self.preset_combo.findData(k)))
        compact_row=QHBoxLayout();compact_row.addWidget(self.preset_menu_button);compact_row.addStretch()
        layout.addLayout(compact_row)
        self.preset_status=QLabel();self.preset_status.setWordWrap(True);self.preset_status.setStyleSheet('color:#9bafb3;font-size:10px;')
        layout.addWidget(self.preset_status)
        self.legend=QWidget();legend=QHBoxLayout(self.legend);legend.setContentsMargins(0,0,0,0);legend.setSpacing(14)
        # Legend colours.
        for text,color in [('Active','#68d9ff'),('Selected','#ff9b91'),('Floating','#f3bd72'),('Pinned','#57e389'),('Empty','#53646d')]:
            item=QLabel(f'<span style="color:{color};font-size:11px">●</span>&nbsp;{text}');item.setProperty('legend',True)
            legend.addWidget(item)
        legend.addStretch()
        layout.addWidget(self.legend)
        self.overview=QTreeWidget()
        self.overview.setColumnCount(4)
        self.overview.setHeaderLabels(["Display / space","Layout","Assigned","State"])
        self.overview.setMaximumHeight(150)
        self.overview.hide()
        self.overview.itemActivated.connect(self._overview_activated)
        layout.addWidget(self.overview)
        # Custom controls sit above the grid while composing a custom layout.
        self.custom_controls=QWidget();custom=QVBoxLayout(self.custom_controls);custom.setContentsMargins(0,0,0,0)
        actions=QHBoxLayout();custom.addLayout(actions)
        self.custom_title=QLabel();self.custom_title.setStyleSheet('color:#b5c9c8;font-size:11px;');actions.addWidget(self.custom_title)
        self.split_buttons={}
        for axis,label in [('vertical','Split vertically'),('horizontal','Split horizontally')]:
            button=QPushButton(label)
            button.clicked.connect(lambda checked=False,a=axis:self.split_selected(a))
            self.split_buttons[axis]=button;actions.addWidget(button)
        self.merge_button=QPushButton("Merge tiles")
        self.merge_button.clicked.connect(self.merge_selected)
        actions.addWidget(self.merge_button);actions.addStretch()
        self.merge_choice=QWidget();choice=QHBoxLayout(self.merge_choice);choice.setContentsMargins(0,0,0,0)
        choice.addWidget(QLabel('Merged tile: keep which app?'))
        self.keep_buttons=[QPushButton(),QPushButton()]
        for position,button in enumerate(self.keep_buttons):
            button.clicked.connect(lambda checked=False,p=position:self._keep_merged(p));choice.addWidget(button)
        choice.addStretch();custom.addWidget(self.merge_choice);self.merge_choice.hide()
        self.merge_hint=QLabel();self.merge_hint.setStyleSheet('color:#f3bd72;font-size:10px;');custom.addWidget(self.merge_hint);self.merge_hint.hide()
        self.geometry_row=QWidget();values=QHBoxLayout(self.geometry_row);values.setContentsMargins(0,0,0,0)
        self.geometry_fields={}
        for key,label in [('x','X'),('y','Y'),('width','Width'),('height','Height')]:
            field=QDoubleSpinBox()
            field.setRange(0 if key in ('x','y') else 3,100)
            field.setDecimals(2)
            field.setSingleStep(1)
            field.setButtonSymbols(QDoubleSpinBox.ButtonSymbols.NoButtons);field.setFixedWidth(72)
            field.setAccessibleName(f"{label} in percent")
            field.editingFinished.connect(lambda k=key:self._geometry_changed(k))
            self.geometry_fields[key]=field
            caption=QLabel(label);caption.setStyleSheet('color:#b5c9c8;font-size:11px;')
            unit=QLabel('%');unit.setStyleSheet('color:#b5c9c8;font-size:11px;')
            values.addWidget(caption);values.addWidget(field);values.addWidget(unit);values.addSpacing(8)
        values.addStretch();custom.addWidget(self.geometry_row)
        layout.addWidget(self.custom_controls);self.custom_controls.hide()
        # The application library replaces the grid.
        self.stack=QStackedWidget()
        self.splitter=self.stack
        self.canvas=TileCanvas(studio_settings(controller.settings))
        self.canvas.selected.connect(self.select_tile)
        self.canvas.dropped.connect(self.drop_assignment)
        self.canvas.keyboard_action.connect(self.tile_keyboard)
        self.canvas.edge_moved.connect(self.move_custom_edge)
        self.canvas.tiles_previewed.connect(self._preview_values)
        self.stack.addWidget(self.canvas)
        self.library=ApplicationLibrary(controller,self.bridge,location=self._window_location)
        self.library.assign_requested.connect(lambda payload:self.drop_assignment(self.selected_index,payload))
        self.stack.addWidget(self.library)
        layout.addWidget(self.stack,1)
        self.effect=QLabel()
        self.effect.setWordWrap(True)
        self.effect.setStyleSheet('color:#a8b8be;font-size:12px;')
        layout.addWidget(self.effect)
        self.tile_actions=QWidget();tile_actions=QHBoxLayout(self.tile_actions);tile_actions.setContentsMargins(0,0,0,0)
        self.pin_button=QPushButton("Pin this app to this tile");self.pin_button.setCheckable(True)
        self.pin_button.clicked.connect(lambda checked:self._pin_changed(checked))
        self.clear_button=QPushButton("Clear tile")
        self.clear_button.clicked.connect(lambda:self.clear_tile(self.selected_index))
        for button in (self.library_button,self.pin_button,self.clear_button): tile_actions.addWidget(button)
        tile_actions.addStretch()
        layout.addWidget(self.tile_actions)
        self.message=QLabel()
        self.message.setWordWrap(True)
        self.message.hide()
        layout.addWidget(self.message)
        self.name_row=QWidget();name_layout=QHBoxLayout(self.name_row);name_layout.setContentsMargins(0,0,0,0)
        self.name_entry=QLineEdit();self.name_entry.setPlaceholderText('Layout name')
        self.name_entry.returnPressed.connect(self.confirm_name)
        self.name_entry.textChanged.connect(lambda:self.replace_row.hide())
        name_layout.addWidget(self.name_entry,1)
        self.name_save=QPushButton('Save');self.name_save.setProperty('primary',True);self.name_save.clicked.connect(self.confirm_name);name_layout.addWidget(self.name_save)
        name_cancel=QPushButton('Cancel');name_cancel.clicked.connect(self._cancel_name);name_layout.addWidget(name_cancel)
        self.name_row.hide()
        self.replace_row=QWidget();replace_layout=QHBoxLayout(self.replace_row);replace_layout.setContentsMargins(0,0,0,0)
        self.replace_label=QLabel();replace_layout.addWidget(self.replace_label,1)
        cancel_replace=QPushButton('Cancel');cancel_replace.setAccessibleName('Cancel saving this layout')
        cancel_replace.clicked.connect(lambda:(self.replace_row.hide(),self.name_row.hide(),self.name_entry.clear()));replace_layout.addWidget(cancel_replace)
        change_name=QPushButton('Change name');change_name.clicked.connect(lambda:(self.replace_row.hide(),self.name_entry.setFocus()));replace_layout.addWidget(change_name)
        replace_button=QPushButton('Replace layout');replace_button.setProperty('danger',True);replace_button.clicked.connect(lambda:self.confirm_name(True));replace_layout.addWidget(replace_button)
        self.replace_row.setObjectName('replaceRow');self.replace_row.setAttribute(Qt.WidgetAttribute.WA_StyledBackground)
        self.replace_row.setStyleSheet('#replaceRow { background:#392d2f; border:1px solid #825157; border-radius:9px; } #replaceRow QLabel { color:#f5d8d8; font-size:11px; }')
        replace_layout.setContentsMargins(9,7,9,7)
        self.replace_row.hide()
        # Naming stays inline in the saved-layout section.
        position=layout.indexOf(self.saved_controls)+1
        layout.insertWidget(position,self.name_row);layout.insertWidget(position+1,self.replace_row)
        footer=QHBoxLayout()
        self.close_button=QPushButton("Cancel")
        self.close_button.clicked.connect(self._close_or_back)
        self.back_button=QPushButton("Back to current space")
        self.back_button.clicked.connect(lambda:self._activate_mode('current'))
        self.reset_button=QPushButton("Reset Space")
        self.reset_button.clicked.connect(self.reset_space)
        self.save_button=QPushButton("Save layout…")
        self.save_button.setIcon(save_icon())
        self.save_button.setAccessibleName("Save current draft as a named layout")
        self.save_button.clicked.connect(self.save_template)
        self.apply_button=QPushButton("Apply arrangement")
        self.apply_button.setProperty('primary',True)
        self.apply_button.clicked.connect(self.apply)
        for button in (self.close_button,self.back_button,self.reset_button,self.save_button,self.apply_button):
            footer.addWidget(button,1)
        layout.addLayout(footer)
        for shortcut,fn in [('Ctrl+Z',self.undo),('Ctrl+Shift+Z',self.redo),('Ctrl+S',self.save_template),
                            ('Ctrl+Return',self.apply),('Ctrl+F',self.focus_library),('Escape',self._escape)]:
            action=QAction(self)
            action.setShortcut(QKeySequence(shortcut))
            action.triggered.connect(fn)
            self.addAction(action)
        self.bridge.changed.connect(self.refresh_from_controller,Qt.ConnectionType.QueuedConnection)
        self.bridge.error.connect(self.show_error,Qt.ConnectionType.QueuedConnection)
        self.bridge.status.connect(self.show_status,Qt.ConnectionType.QueuedConnection)
        self.bridge.busy_changed.connect(self._busy)
        apply_theme(studio_settings(controller.settings),target=self)
        self.refresh_from_controller()
        self.preview_timer=QTimer(self)
        self.preview_timer.setInterval(200)
        self.preview_timer.timeout.connect(self._refresh_thumbnails)

    def _refresh_thumbnails(self):
        active=set()
        enabled=self.preview_button.isChecked() and self.isVisible() and not self.isMinimized()
        if enabled:
            if self._thumbnails is None:
                from smartgrid.platform.windows.thumbnails import ThumbnailManager
                self._thumbnails=ThumbnailManager(self.controller.backend.api)
            windows={w.ref.hwnd:w for w in self.controller.windows}
            targets={}
            scale=self.devicePixelRatioF()
            for card in self.canvas.cards:
                assignment=card.assignment
                window=windows.get(card.window.ref.hwnd) if getattr(card,'window',None) else None
                if window and window.ref!=card.window.ref: window=None
                if not window or window.state not in ('normal','maximized') or not window.eligible or not card.isVisible():
                    continue
                box=card.rect().adjusted(10,30,-10,-64)
                if box.width()<40 or box.height()<30:
                    continue
                ratio=window.rect.width/max(1,window.rect.height)
                width=min(box.width(),round(box.height()*ratio))
                height=min(box.height(),round(width/ratio))
                origin=card.mapTo(self,QPoint(box.x()+(box.width()-width)//2,box.y()+(box.height()-height)//2))
                targets[card.index]=(window.ref,Rect(round(origin.x()*scale),round(origin.y()*scale),round(width*scale),round(height*scale)))
            active=self._thumbnails.sync(int(self.winId()),targets)
        elif self._thumbnails:
            self._thumbnails.clear()
        for card in self.canvas.cards:
            value=card.index in active
            if value!=card.preview_active:
                card.preview_active=value
                card.update()

    def paintEvent(self,event):
        paint_surface(self)

    def showEvent(self,event):
        super().showEvent(event)
        animate_open(self)
        show_backdrops(self)
        self.preview_timer.start()

    def hideEvent(self,event):
        hide_backdrops(self)
        self.preview_timer.stop()
        if self._thumbnails:
            self._thumbnails.clear()
        super().hideEvent(event)

    @property
    def local_key(self):
        return self.template_id if self.mode=='edit' else 'new'

    def current_profile(self):
        if self.mode=='current':
            return self.controller.draft(self.display_id,self.space).profile
        key=self.local_key
        if key not in self.local_profiles:
            if self.mode=='edit' and self.template_id:
                profile=self.controller.load_template_draft(self.template_id,self.display_id,self.space)
            else:
                profile=SpaceProfile(self.display_id,self.space,preset='2x2')
            self.local_profiles[key]=deepcopy(profile)
            self.local_histories[key]=History(30)
        profile=self.local_profiles[key]
        profile.display_id,self.space= self.display_id,self.space
        profile.space=self.space
        return profile

    def refresh_from_controller(self):
        if self._refreshing:
            return
        self._refreshing=True
        try:
            old_display=self.display_id
            with QSignalBlocker(self.display_combo):
                self.display_combo.clear()
                for display in self.controller.displays:
                    self.display_combo.addItem(f"{display.name} · {display.dpi/96:.0%}",display.id)
                index=self.display_combo.findData(old_display)
                if index<0 and self.display_combo.count():
                    index=0
                    self.display_id=self.display_combo.itemData(0)
                    self.space=self.controller.active_spaces.get(self.display_id,0)
                self.display_combo.setCurrentIndex(index)
            for index,button in enumerate(self.space_buttons):
                button.setChecked(index==self.space)
            self._refresh_display_chips()
            self.toggle_button.setEnabled(not self.bridge.pending)
            self.runtime_status.setText("PAUSED" if not self.controller.running or self.controller.paused else "ON")
            old_template=self.template_id or self.template_combo.currentData()
            with QSignalBlocker(self.template_combo):
                self.template_combo.clear()
                query=self.template_search.text().casefold()
                group=None
                for template in sorted(self.controller.templates,key=lambda t:(t.preset!='custom',capacity(t.preset,t.tiles),t.name.casefold())):
                    if query and query not in template.name.casefold():
                        continue
                    category='Custom layouts' if template.preset=='custom' else 'Preset layouts'
                    if category!=group:
                        self.template_combo.addItem(category,None)
                        self.template_combo.model().item(self.template_combo.count()-1).setEnabled(False)
                        group=category
                    count=max(sum(a is not None for a in template.assignments),capacity(template.preset,template.tiles))
                    meta='' if template.preset in ('custom','auto') else f"  ·  {preset_name(template.preset)}"
                    self.template_combo.addItem(f"{template.name}  ·  {count} TILES{meta}",template.id)
                index=self.template_combo.findData(old_template) if old_template else -1
                if index<0:
                    index=next((i for i in range(self.template_combo.count()) if self.template_combo.itemData(i)), -1)
                self.template_combo.setCurrentIndex(index)
            if self.mode=='edit' and self.template_id not in {t.id for t in self.controller.templates}:
                self.template_id=self.template_combo.currentData()
                if self.template_id is None:
                    self.mode='new'
                    with QSignalBlocker(self.mode_combo):
                        self.mode_combo.setCurrentIndex(self.mode_combo.findData('new'))
            visible=self.mode=='edit'
            for key,button in self.mode_buttons.items(): button.setChecked(key==self.mode)
            saved=next((t for t in self.controller.templates if t.id==self.template_id),None) if visible else None
            short=(saved.name[:19]+'…' if len(saved.name)>20 else saved.name) if saved else ''
            self.mode_buttons['edit'].setText(f'Saved · {short} ▾' if saved and not self.saved_picker_open else f'Saved layouts · {len(self.controller.templates)} ▾')
            self.mode_buttons['edit'].setEnabled(bool(self.controller.templates) or visible)
            deleted=getattr(self.controller,'_deleted_template',None)
            if saved:
                self._refresh_saved_summary(saved.id,repr(saved.to_dict()))
            elif deleted is not None:
                deleted=deleted[0] if isinstance(deleted,tuple) else deleted
                self.saved_summary.setText(f"Deleted “{getattr(deleted,'name','layout')}”")
            self.template_combo.hide()
            signature=tuple((self.template_combo.itemText(i),self.template_combo.itemData(i)) for i in range(self.template_combo.count()))
            if signature!=self._saved_signature:
                self._saved_signature=signature
                with QSignalBlocker(self.saved_list):
                    self.saved_list.clear()
                    for text,key in signature:
                        item=QListWidgetItem(text);item.setData(Qt.ItemDataRole.UserRole,key)
                        if not key: item.setFlags(Qt.ItemFlag.NoItemFlags)
                        else:
                            template=next(t for t in self.controller.templates if t.id==key)
                        self.saved_list.addItem(item)
            with QSignalBlocker(self.saved_list):
                for row in range(self.saved_list.count()):
                    if self.saved_list.item(row).data(Qt.ItemDataRole.UserRole)==self.template_combo.currentData():
                        self.saved_list.setCurrentRow(row);break
            self.template_search.setVisible(visible and self.saved_picker_open)
            self.saved_list.setVisible(visible and self.saved_picker_open)
            self.rename_button.setVisible(visible)
            self.delete_button.setVisible(visible)
            self.undo_delete_button.setVisible(visible)
            self.edit_button.setVisible(visible and not self.editing_saved)
            self.saved_summary.setVisible(visible and not self.editing_saved)
            self.undo_delete_button.setVisible(visible and bool(self.controller._deleted_template))
            editable=self.mode!='edit' or self.editing_saved
            self.canvas.setEnabled(editable)
            if not editable: self._show_library(False)
            self.save_button.setVisible(editable)
            if not self.display_id:
                self.effect.setText("No display available.")
                self.apply_button.setEnabled(False)
                return
            profile=self.current_profile()
            with QSignalBlocker(self.preset_combo):
                self.preset_combo.setCurrentIndex(self.preset_combo.findData(profile.preset))
            with QSignalBlocker(self.ratio):
                self.ratio.setValue(profile.ratio*100)
            show_presets=self.mode!='current' and editable
            compact=self.width()<820
            self.preset_strip.setVisible(show_presets and not compact)
            self.preset_menu_button.setVisible(show_presets and compact)
            self.preset_menu_button.setText(preset_name(profile.preset))
            for action,(key,button) in zip(self.preset_menu.actions(),[(k,b) for k,b in self.preset_buttons.items() if k!='auto']):
                action.setEnabled(not button.property('unavailable'))
            required=max((i+1 for i,a in enumerate(profile.assignments) if a),default=0)
            for key,button in self.preset_buttons.items():
                button.setChecked(key==profile.preset)
                insufficient=key not in ('auto','custom') and capacity(key)<required
                button.setProperty('unavailable',insufficient)
                button.setToolTip(f'Needs at least {required} tiles to preserve assignments' if insufficient else button.text())
                button.style().unpolish(button);button.style().polish(button)
            self.saved_controls.setVisible(visible)
            self.current_layout.setVisible(self.mode in ('current','edit'))
            self.ratio.hide()
            display=next((d for d in self.controller.displays if d.id==self.display_id),None)
            shared={}
            if self.mode=='current':
                for i,assignment in enumerate(profile.assignments):
                    if assignment and assignment.window_id:
                        spaces=[]
                        for s in range(3):
                            other=self.controller.draft(self.display_id,s).profile
                            if any(a and a.window_id==assignment.window_id for a in other.assignments):
                                spaces.append(s)
                        if len(spaces)>1:
                            shared[i]=spaces
            pending=getattr(self.controller,'pending',{})
            pending_slots=set()
            values=pending.values() if isinstance(pending,dict) else pending
            for p in values:
                if getattr(p,'display_id',None)==self.display_id and getattr(p,'space',None)==self.space:
                    pending_slots.add(getattr(p,'index',-1))
            self.canvas.settings=studio_settings(self.controller.settings)
            self.canvas.active_hwnd=getattr(self.controller,'_last_tiled_selection',None)
            self.canvas.set_profile(profile,display,self.controller.apps,self.controller.windows,shared,pending_slots,runtime=self.mode=='current')
            if self.mode=='edit':
                saved=next((t for t in self.controller.templates if t.id==self.template_id),None)
                self.current_layout.setText(f"{saved.name if saved else 'Saved layout'}  ·  {preset_name(self.canvas.effective_preset)}")
            else:
                self.current_layout.setText(f"Current layout: {preset_name(self.canvas.effective_preset) if any(profile.assignments) else 'Empty'}")
            self.selected_index=min(self.selected_index,len(self.canvas.cards)-1)
            self.canvas.select(self.selected_index)
            self._update_details()
            self._update_texts(profile,required)
            self._update_footer()
            history=self.local_histories.get(self.local_key)
            if self.mode=='current':
                histories=getattr(self.controller,'draft_histories',getattr(self.controller,'_draft_histories',{}))
                history=histories.get((self.display_id,self.space))
            # History chips: "Undo · N" / "Redo · N".
            undo,redo=(history.undo_count,history.redo_count) if history else (0,0)
            self.undo_button.setText(f'Undo · {undo}');self.redo_button.setText(f'Redo · {redo}')
            self.undo_button.setAccessibleName(f'Undo: {undo} actions available')
            self.redo_button.setAccessibleName(f'Redo: {redo} actions available')
            self.undo_button.setEnabled(bool(undo));self.redo_button.setEnabled(bool(redo))
            self.apply_button.setEnabled(not self.bridge.pending)
            if self.overview.isVisible():
                self._fill_overview()
            error=getattr(self.controller,'last_error','')
            if error:
                self.show_error(str(error))
        except Exception as exc:
            self.show_error(str(exc))
        finally:
            self._refreshing=False

    def _refresh_display_chips(self):
        signature=tuple(d.id for d in self.controller.displays)
        if getattr(self,'_display_signature',None)!=signature:
            self._display_signature=signature
            for button in self.display_buttons: button.deleteLater()
            self.display_buttons=[]
            for index,display in enumerate(self.controller.displays):
                button=QPushButton(f"Display {index+1}");button.setCheckable(True);button.setToolTip(display.name)
                button.clicked.connect(lambda checked=False,i=index:self.display_combo.setCurrentIndex(i))
                self.display_chips.addWidget(button);self.display_buttons.append(button)
        for display,button in zip(self.controller.displays,self.display_buttons):
            button.setChecked(display.id==self.display_id)
        # The Display/Space context is hidden while composing a layout.
        composing=self.mode=='new' or (self.mode=='edit' and self.editing_saved)
        for widget in self.display_buttons+self.space_buttons+[self.context_separator]: widget.setVisible(not composing)

    def _update_texts(self,profile,required=None):
        if required is None:
            required=max((i+1 for i,a in enumerate(profile.assignments) if a),default=0)
        windows=sum(1 for a in profile.assignments if a and a.window_id)
        pending=sum(1 for a in profile.assignments if a and not a.window_id and not a.pinned)
        tiles=len(self.canvas.cards)
        preset=profile.preset
        self.legend.setVisible(self.mode=='current')
        self.preset_status.setVisible(self.mode=='current' or self.mode=='new' or self.editing_saved)
        if self.mode=='current':
            effective=self.canvas.effective_preset
            if preset=='custom' and effective!='custom':
                behavior=f'This custom layout has {len(profile.tiles)} tiles; {preset_name(effective)} is shown and grows and shrinks with the windows.'
            elif effective=='custom':
                behavior='Custom keeps its geometry and tile positions.'
            else:
                behavior='Grows and shrinks with the number of windows. Pinned tiles stay reserved.'
            self.preset_status.setText(behavior+(f' · {required} tiles needed' if required>1 else ''))
            if pending:
                text=f"{windows} windows · {pending} {'app opens' if pending==1 else 'apps open'} on Apply · Select a tile to change or clear its app."
            elif windows:
                text=f"{windows} windows · {tiles} tiles · Select a tile to choose an app, clear or pin it. Drag to swap. Apply to confirm."
            else:
                text="Choose an app in an empty tile. Apply to confirm."
            if preset not in ('auto','custom') and capacity(preset)<len([a for a in profile.assignments if a]):
                text+=f" {preset_name(preset)} cannot fit {len([a for a in profile.assignments if a])} tiles; Auto is shown."
        elif self.mode=='edit' and not self.editing_saved:
            text='' if self.saved_picker_open else 'Saved layout preview only. Restore reuses matching windows, opens missing apps and minimizes extras. Unsaved Studio edits are not included.'
        else:
            self.preset_status.setText('Split or merge tiles, then assign apps.' if preset=='custom' else 'Choose a layout, then select a tile to assign an app.')
            saved=next((t for t in self.controller.templates if t.id==self.template_id),None)
            prefix=f"EDIT · {saved.name if saved else 'Layout'}" if self.mode=='edit' else 'NEW LAYOUT'
            apps=sum(1 for a in profile.assignments if a)
            if preset=='custom':
                tail=(f'Click an orange tile to merge it with tile {self.selected_index+1}.' if self.merge_candidates else
                      'Merge: 1. Select a tile → 2. Merge tiles → 3. Click its neighbor. Drag green dividers to resize.')
            else:
                tail='Choose a tile to assign an installed app. Saving does not change your desktop.'
            text=f"{prefix} · {apps} apps · {tiles} tiles · {tail}"
        if self.stack.currentWidget() is self.library and self.selected_index>=0:
            tile=self.selected_index+1
            text=(f'Choose an app for tile {tile}. Existing windows are reused; missing apps open only on Apply.' if self.mode=='current'
                  else f'Choose an open window or installed app for tile {tile}. Saving does not move windows or launch apps.')
        self.effect.setText(text)
        self.effect.setVisible(bool(text) and self.height()>=600)

    def _refresh_saved_summary(self,template_id,content=''):
        key=(template_id,content,self.display_id,self.space,tuple((w.ref.hwnd,w.state) for w in self.controller.windows))
        if getattr(self,'_summary_key',None)==key: return
        self._summary_key=key
        self.saved_summary.setText('PREVIEW')
        def ready(plan):
            if getattr(self,'_summary_key',None)!=key: return
            count=lambda v:len(v) if isinstance(v,(list,tuple,set,dict)) else int(v or 0)
            self.saved_summary.setText(f"PREVIEW · {count(plan.get('reuse'))} reuse · {count(plan.get('open'))} open · {count(plan.get('hide'))} minimize")
        self.bridge.submit('preview_template',template_id,self.display_id,self.space,on_success=ready)

    def _update_footer(self):
        current=self.mode=='current'
        preview=self.mode=='edit' and not self.editing_saved
        chooser=preview and self.saved_picker_open
        composing=self.mode=='new' or (self.mode=='edit' and self.editing_saved)
        self.close_button.setText('Cancel' if current else 'Close')
        self.close_button.setVisible(not composing)
        self.back_button.setText('Back to layout' if chooser else 'Back to current space')
        self.back_button.setVisible(not current)
        self.reset_button.setVisible(current)
        self.save_button.setText('Save changes…' if self.mode=='edit' else 'Save layout…')
        self.save_button.setVisible(current or composing)
        self.apply_button.setText('Restore in this space' if preview else 'Apply arrangement')
        self.apply_button.setVisible((current or preview) and not chooser)
        self.undo_button.setVisible(not chooser);self.redo_button.setVisible(not chooser)
        QTimer.singleShot(0,self._fit_canvas)
        # The saved-layout chooser replaces the grid.
        self.splitter.setVisible(not chooser)
        self.current_layout.setVisible(current)
        self.preview_button.setVisible(current)
        if not current and self.preview_button.isChecked(): self.preview_button.setChecked(False)

    def _fit_canvas(self):
        """Give the grid the same size in every view (current space, saved
        preview, new or edited layout): up to 480 px high at the display's
        aspect ratio. The Studio grows to make room, within the screen."""
        display=next((d for d in self.controller.displays if d.id==self.display_id),None)
        ratio=display.work_area.height/max(1,display.work_area.width) if display else 9/16
        wanted=min(480,round((self.width()-40)*ratio))
        screen=self.screen().availableGeometry() if self.screen() else None
        if screen is not None:
            # Keep room for the header, controls and footer on small screens.
            wanted=min(wanted,max(160,screen.height()-420))
        self.canvas.setMinimumHeight(wanted)
        need=self.sizeHint().height()
        if screen is not None and need>self.height():
            height=min(need,screen.height()-24)
            self.resize(self.width(),height)
            if self.geometry().bottom()>screen.bottom():
                self.move(self.x(),max(screen.top()+12,screen.bottom()-height-12))

    def _close_or_back(self):
        self.close()

    def _display_changed(self,index):
        if self._refreshing or index<0:
            return
        self.display_id=self.display_combo.itemData(index)
        self.space=self.controller.active_spaces.get(self.display_id,self.space)
        self.selected_index=-1
        self.refresh_from_controller()

    def _space_changed(self,space):
        self.space=space
        self.selected_index=-1
        self.refresh_from_controller()

    def _mode_changed(self,index):
        if self._refreshing:
            return
        self.mode=self.mode_combo.itemData(index)
        self.editing_saved=False
        self.name_row.hide();self.replace_row.hide();self._cancel_merge()
        if self.mode=='edit':
            self.template_id=self.template_combo.currentData()
        self.selected_index=-1
        self.refresh_from_controller()

    def _activate_mode(self,key):
        if key=='edit' and self.mode=='edit':
            self.saved_picker_open=not self.saved_picker_open;self.refresh_from_controller()
        else:
            self.saved_picker_open=True
            self.mode_combo.setCurrentIndex(self.mode_combo.findData(key))

    def _close_saved_picker(self,item):
        if item and item.data(Qt.ItemDataRole.UserRole):
            self.saved_picker_open=False;self.refresh_from_controller()

    def _saved_selected(self,item,previous):
        key=item.data(Qt.ItemDataRole.UserRole) if item else None
        if key: self.template_combo.setCurrentIndex(self.template_combo.findData(key))

    def _template_changed(self,index):
        if self._refreshing or self.mode!='edit':
            return
        self.template_id=self.template_combo.itemData(index)
        self.editing_saved=False
        self.selected_index=-1
        self.refresh_from_controller()

    def edit_saved(self):
        if self.template_id:
            self.editing_saved=True
            self.refresh_from_controller()

    def _mutate_profile(self,fn):
        if self.mode=='edit' and not self.editing_saved: return
        profile=deepcopy(self.current_profile())
        try:
            fn(profile)
        except (ValueError,IndexError,TypeError) as exc:
            self.show_error(str(exc))
            self._update_details()
            return
        if self.mode=='current':
            self.bridge.submit('set_draft_profile',self.display_id,self.space,profile)
        else:
            self.local_histories[self.local_key].push(self.current_profile())
            self.local_profiles[self.local_key]=profile
            self.local_dirty.add(self.local_key)
            self.refresh_from_controller()

    def _preset_changed(self,index):
        if self._refreshing or index<0:
            return
        preset=self.preset_combo.itemData(index)
        profile=self.current_profile()
        required=max((i+1 for i,a in enumerate(profile.assignments) if a),default=0)
        if self.mode!='current' and preset not in ('auto','custom') and capacity(preset)<required:
            self.show_error(f"{preset_name(preset)}: {capacity(preset)} tiles; this draft needs {required}")
            with QSignalBlocker(self.preset_combo):
                self.preset_combo.setCurrentIndex(self.preset_combo.findData(profile.preset))
            return
        if preset=='custom':
            def custom(profile):
                if self.mode=='new' and not required and not profile.tiles:
                    profile.preset='custom'
                    profile.tiles=[Tile(uuid4().hex,0,0,1,1)]
                    profile.assignments=[None]
                else:
                    self._ensure_custom(profile)
            self._mutate_profile(custom)
        elif self.mode=='current':
            self.bridge.submit('set_preset',self.display_id,self.space,preset)
        else:
            def mutate(profile):
                profile.preset=preset
                profile.tiles=[]
                profile.resize_tiles=[]
                profile.assignments=profile.assignments[:required]
            self._mutate_profile(mutate)

    def _ensure_custom(self,profile):
        if profile.resize_tiles:
            profile.preset='custom'
            profile.tiles=deepcopy(profile.resize_tiles)
            profile.resize_tiles=[]
        if profile.preset=='custom' and profile.tiles:
            return
        count=max(1,capacity(profile.preset,profile.tiles),len(profile.assignments))
        preset=auto_preset(count) if profile.preset=='auto' else profile.preset
        count=max(count,capacity(preset))
        rects=resolve_layout(Rect(0,0,10000,10000),count,preset,gap=0,padding=0,ratio=profile.ratio)
        profile.tiles=[Tile(uuid4().hex,r.x/10000,r.y/10000,r.width/10000,r.height/10000) for r in rects]
        profile.preset='custom'
        self._pad_assignments(profile,len(profile.tiles))

    @staticmethod
    def _pad_assignments(profile,count):
        profile.assignments.extend([None]*max(0,count-len(profile.assignments)))

    def _ratio_changed(self):
        if not self._refreshing:
            self._mutate_profile(lambda p:setattr(p,'ratio',self.ratio.value()/100))

    def select_tile(self,index):
        if self.merge_candidates:
            if index in self.merge_candidates: self._complete_merge(index)
            return
        profile=self.current_profile()
        custom=profile.preset=='custom'
        # Clicking a selected tile deselects it, except while composing a custom layout.
        if self.mode=='current' and index==self.selected_index:
            index=-1
        self.selected_index=index
        self.canvas.select(index)
        assignment=profile.assignments[index] if 0<=index<len(profile.assignments) else None
        self._show_library(index>=0 and assignment is None and not (self.mode!='current' and custom))
        self._update_details()

    def _show_library(self,visible):
        visible=bool(visible) and self.selected_index>=0
        with QSignalBlocker(self.library_button): self.library_button.setChecked(visible)
        self.library_button.setText("Back to layout" if visible else "Choose an app")
        self.stack.setCurrentWidget(self.library if visible else self.canvas)
        if hasattr(self,'pin_button'):
            self._update_details()
            if self.display_id: self._update_texts(self.current_profile())

    def _window_location(self,window):
        """Library caption: spaces holding this window, or the display it moves from."""
        displays=self.controller.displays
        spaces=[s for s in range(3) if any(a and a.window_id==window.ref.hwnd
                for a in self.controller.draft(window.display_id,s).profile.assignments)]
        names=', '.join(f'Space {s+1}' for s in spaces)
        if window.display_id==self.display_id:
            return names
        index=next((i for i,d in enumerate(displays) if d.id==window.display_id),None)
        source=f'Move from Display {index+1}' if index is not None else 'Move from another display'
        return f'{source} · {names}' if names else source

    def _update_details(self):
        profile=self.current_profile()
        index=self.selected_index
        assignment=profile.assignments[index] if 0<=index<len(profile.assignments) else None
        preview=self.mode=='edit' and not self.editing_saved
        chooser=preview and self.saved_picker_open
        library=self.stack.currentWidget() is self.library
        self.tile_actions.setVisible(index>=0 and not preview and not chooser)
        self.pin_button.setVisible(assignment is not None and not library)
        self.clear_button.setVisible(assignment is not None and not library)
        pinned=bool(assignment and assignment.pinned)
        with QSignalBlocker(self.pin_button): self.pin_button.setChecked(pinned)
        if self.mode=='current':
            self.pin_button.setText('Unpin this app from this tile' if pinned else 'Pin this app to this tile')
        else:
            self.pin_button.setText('Unpin this app' if pinned else 'Pin this app to this tile')
        self._update_custom_controls(profile)

    def _update_custom_controls(self,profile):
        # Custom geometry is edited while composing a layout, and also in the
        # current space when it uses a custom layout that fits its windows.
        composing=self.mode=='new' or (self.mode=='edit' and self.editing_saved)
        custom=(composing or self.mode=='current') and profile.preset=='custom' and self.canvas.effective_preset=='custom'
        self.custom_controls.setVisible(custom)
        if not custom:
            return
        index=self.selected_index
        tile=profile.tiles[index] if 0<=index<len(profile.tiles) else None
        self.custom_title.setText(f'Tile {index+1}' if tile else 'Select a tile')
        for axis,button in self.split_buttons.items():
            try: possible=bool(tile) and len(profile.tiles)<30 and all(t.width>=.03 and t.height>=.03 for t in split_tile(profile.tiles,index,axis))
            except (ValueError,IndexError): possible=False
            button.setEnabled(possible)
        neighbors=self._merge_neighbors(profile) if tile else []
        self.merge_button.setText('Cancel merge' if self.merge_candidates else 'Merge tiles')
        self.merge_button.setEnabled(bool(neighbors) or bool(self.merge_candidates))
        pending=getattr(self,'_merge_pending',None)
        self.merge_choice.setVisible(bool(pending))
        self.merge_hint.setVisible(bool(self.merge_candidates) and not pending)
        self.merge_hint.setText(f'Click an orange tile to merge it with tile {index+1}.')
        self.geometry_row.setVisible(bool(tile) and not self.merge_candidates and not pending)
        if tile:
            fixed={'x':tile.x<=1e-9,'y':tile.y<=1e-9,'width':tile.width>=1-1e-9,'height':tile.height>=1-1e-9}
            for key,field in self.geometry_fields.items():
                with QSignalBlocker(field): field.setValue(getattr(tile,key)*100)
                field.setEnabled(not fixed[key])
                field.setAccessibleName(f"{key.title() if key not in ('x','y') else key.upper()} in percent for tile {index+1}"+(', fixed outer edge' if fixed[key] else ''))

    def _preview_values(self,tiles):
        """Keep X/Y/Width/Height in step with a divider being dragged."""
        index=self.selected_index
        if 0<=index<len(tiles) and self.geometry_row.isVisible():
            for key,field in self.geometry_fields.items():
                with QSignalBlocker(field): field.setValue(getattr(tiles[index],key)*100)

    def _merge_neighbors(self,profile):
        result=[]
        for other in range(len(profile.tiles)):
            if other==self.selected_index: continue
            try: merge_tiles(profile.tiles,self.selected_index,other)
            except (ValueError,IndexError): continue
            result.append(other)
        return result

    def drop_assignment(self,index,payload):
        if not isinstance(payload,dict) or index<0:
            return
        if 'tile' in payload:
            source=int(payload['tile'])
            if source==index:
                return
            self.selected_index=index;self.canvas.select(index)
            if self.mode=='current':
                self.bridge.submit('swap_draft',self.display_id,self.space,source,index)
            else:
                def swap(profile):
                    self._pad_assignments(profile,max(source,index)+1)
                    profile.assignments[source],profile.assignments[index]=profile.assignments[index],profile.assignments[source]
                self._mutate_profile(swap)
        elif payload.get('app_id'):
            app_id,window_id=payload['app_id'],payload.get('window_id')
            if self.mode=='current':
                display_id,space=self.display_id,self.space
                def assign():
                    controller=self.controller
                    wid=window_id
                    if wid is None:
                        # Reuse an open window of the app on this display.
                        live=next((w for w in controller.windows if w.app_id==app_id and w.display_id==display_id
                                   and w.eligible and not w.floating),None)
                        wid=live.ref.hwnd if live else None
                    window=controller._window(wid) if wid else None
                    profile=controller.draft(display_id,space).profile
                    previous=profile.assignments[index] if index<len(profile.assignments) else None
                    pinned=bool(previous and previous.pinned and previous.app_id==app_id)
                    if window and window.display_id!=display_id:
                        # A window from another display moves; same-display spaces share it.
                        for other in range(3):
                            draft=controller.draft(window.display_id,other).profile
                            for i,a in enumerate(draft.assignments):
                                if a and a.window_id==wid: controller.set_assignment(window.display_id,other,i,None)
                    controller.set_assignment(display_id,space,index,Assignment(app_id,wid,pinned))
                self.bridge.submit(assign)
            else:
                def assign(profile):
                    self._pad_assignments(profile,index+1)
                    for i,a in enumerate(profile.assignments):
                        if i!=index and a and a.app_id==app_id: profile.assignments[i]=None
                    profile.assignments[index]=Assignment(app_id,window_id,False)
                self._mutate_profile(assign)
            self._show_library(False)

    def clear_tile(self,index):
        if index<0: return
        if self.mode=='current':
            self.bridge.submit('set_assignment',self.display_id,self.space,index,None)
        else:
            def clear(profile):
                self._pad_assignments(profile,index+1)
                profile.assignments[index]=None
            self._mutate_profile(clear)

    def _pin_changed(self,pinned):
        if self._refreshing or self.selected_index<0:
            return
        index=self.selected_index
        def pin(profile):
            if index<len(profile.assignments) and profile.assignments[index]:
                app_id=profile.assignments[index].app_id
                # One pin per application in a space.
                for a in profile.assignments:
                    if pinned and a and a.app_id==app_id: a.pinned=False
                profile.assignments[index].pinned=pinned
        self._mutate_profile(pin)

    def tile_keyboard(self,index,action):
        if action=='assign':
            self.selected_index=index;self.canvas.select(index);self.focus_library()
        elif action=='clear':
            self.clear_tile(index)
        elif action.startswith(('select_','swap_')):
            kind,direction=action.split('_',1)
            other=self.canvas.neighbor(index,direction)
            if other is not None:
                if kind=='swap':
                    self.drop_assignment(other,{'tile':index})
                self.selected_index=other
                self.canvas.select(other,True)
                self._update_details()

    def split_selected(self,axis):
        index=self.selected_index
        if index<0: return
        def split(profile):
            self._ensure_custom(profile)
            if len(profile.tiles)>=30:
                raise ValueError("A custom layout accepts at most 30 tiles.")
            profile.tiles=split_tile(profile.tiles,index,axis)
            if any(t.width<.03 or t.height<.03 for t in profile.tiles):
                raise ValueError("Tiles must be at least 3 % wide and tall.")
            self._pad_assignments(profile,len(profile.tiles)-1)
            profile.assignments.insert(index+1,None)
        self._cancel_merge()
        self._mutate_profile(split)

    def merge_selected(self):
        if self.merge_candidates or getattr(self,'_merge_pending',None):
            self._cancel_merge();self._update_details();return
        profile=deepcopy(self.current_profile())
        if profile.preset!='custom' and not profile.resize_tiles:
            return
        self._ensure_custom(profile)
        candidates=self._merge_neighbors(profile)
        if not candidates:
            return
        self.merge_candidates=candidates
        for i,card in enumerate(self.canvas.cards): card.merge_target=i in candidates;card.update()
        self._update_details();self._update_texts(self.current_profile())

    def _complete_merge(self,other):
        profile=deepcopy(self.current_profile());self._ensure_custom(profile)
        index=self.selected_index
        self._pad_assignments(profile,len(profile.tiles))
        first,second=profile.assignments[index],profile.assignments[other]
        if first and second and first.app_id!=second.app_id:
            # Ask inline which app the merged tile keeps.
            self._merge_pending=(index,other)
            for button,assignment in zip(self.keep_buttons,(first,second)):
                app=next((a for a in self.controller.apps if a.id==assignment.app_id),None)
                name=app.name if app else app_display_name(assignment.app_id)
                button.setText(f"Keep {name[:23]+'…' if len(name)>24 else name}")
                button.setAccessibleName(f'Keep {name} in the merged tile')
            self.merge_candidates=[]
            for card in self.canvas.cards: card.merge_target=False;card.update()
            self._update_details();return
        self._finish_merge(index,other,index if first else other)

    def _keep_merged(self,position):
        pending=getattr(self,'_merge_pending',None)
        if pending: self._finish_merge(*pending,pending[position])

    def _finish_merge(self,index,other,keep_index):
        profile=deepcopy(self.current_profile());self._ensure_custom(profile)
        self._pad_assignments(profile,len(profile.tiles))
        keep=deepcopy(profile.assignments[keep_index])
        self._cancel_merge()
        def merge(p):
            self._ensure_custom(p)
            self._pad_assignments(p,len(p.tiles))
            p.tiles=merge_tiles(p.tiles,index,other)
            low,high=sorted([index,other])
            p.assignments.pop(high)
            p.assignments[low]=keep
        self.selected_index=min(index,other)
        self._mutate_profile(merge)

    def _cancel_merge(self):
        self.merge_candidates=[]
        self._merge_pending=None
        for card in self.canvas.cards: card.merge_target=False;card.update()

    def _escape(self):
        if self.merge_candidates or getattr(self,'_merge_pending',None): self._cancel_merge();self._update_details()
        elif self.stack.currentWidget() is self.library: self._show_library(False)
        elif self.mode=='new' or (self.mode=='edit' and self.editing_saved): self._activate_mode('current')
        else: self.close()

    def move_custom_edge(self,index,edge,delta):
        def move(profile):
            self._ensure_custom(profile)
            profile.tiles=move_edge(profile.tiles,index,edge,delta)
            if any(t.width<.03 or t.height<.03 for t in profile.tiles):
                raise ValueError("Tiles must be at least 3 % wide and tall.")
        self._mutate_profile(move)

    def _geometry_changed(self,key):
        if self._refreshing or self.selected_index<0:
            return
        value=self.geometry_fields[key].value()/100
        def geometry(profile):
            self._ensure_custom(profile)
            profile.tiles=set_tile_geometry(profile.tiles,self.selected_index,**{key:value})
            if any(t.width<.03 or t.height<.03 for t in profile.tiles):
                raise ValueError("Tiles must be at least 3 % wide and tall.")
        self._mutate_profile(geometry)

    def undo(self):
        if self.mode=='current':
            self.bridge.submit('undo_draft',self.display_id,self.space)
        else:
            restored=self.local_histories[self.local_key].undo(self.current_profile())
            if restored is not None:
                self.local_profiles[self.local_key]=restored
                self.local_dirty.add(self.local_key)
                self.refresh_from_controller()

    def redo(self):
        if self.mode=='current':
            self.bridge.submit('redo_draft',self.display_id,self.space)
        else:
            restored=self.local_histories[self.local_key].redo(self.current_profile())
            if restored is not None:
                self.local_profiles[self.local_key]=restored
                self.local_dirty.add(self.local_key)
                self.refresh_from_controller()

    def save_template(self):
        current=next((t for t in self.controller.templates if t.id==self.template_id),None) if self.mode=='edit' else None
        self._name_operation='save'
        self.name_save.setText('Save')
        if self.mode=='current' and self.name_row.isVisible():
            self.name_row.hide();self.replace_row.hide();return
        if current: self.name_entry.setText(current.name)
        self.name_row.show();self.name_entry.setFocus()

    def confirm_name(self,replace=False):
        name=self.name_entry.text().strip()[:60]
        if not name:
            self.name_entry.setPlaceholderText('Enter a layout name')
            if getattr(self,'_name_operation','save')=='rename': self.show_error('Enter a layout name.')
            return
        if getattr(self,'_name_operation','save')=='rename':
            if any(t.id!=self.template_id and t.name.casefold()==name.casefold() for t in self.controller.templates):
                self.show_error('A layout with this name already exists.');return
            self.bridge.submit('rename_template',self.template_id,name,on_success=lambda _:(self.name_row.hide(),self.message.hide()))
            return
        existing=next((t for t in self.controller.templates if t.name.casefold()==name.casefold()),None)
        template_id=self.template_id if self.mode=='edit' else None
        if existing and not replace:
            self.replace_label.setText(f'“{existing.name}” already exists. Replace it?')
            self.replace_row.show();return
        if existing: template_id=existing.id
        profile=deepcopy(self.current_profile())
        key=self.local_key
        def saved(template):
            self.local_dirty.discard(key)
            self.name_row.hide();self.replace_row.hide();self.name_entry.clear()
            # Show the saved layout's preview after saving.
            if template is not None and getattr(template,'id',None):
                self.local_profiles.pop(key,None);self.local_histories.pop(key,None)
                self.template_id=template.id;self.editing_saved=False;self.saved_picker_open=False
                with QSignalBlocker(self.mode_combo): self.mode_combo.setCurrentIndex(self.mode_combo.findData('edit'))
                self.mode='edit';self.selected_index=-1
                self.refresh_from_controller()
        self.bridge.submit('save_profile_template',name,profile,template_id,on_success=saved)

    def apply(self):
        if self.mode=='current':
            self.bridge.submit('apply_drafts',on_success=self._applied)
        else:
            profile=deepcopy(self.current_profile())
            self.bridge.submit('restore_profile',profile,on_success=self._applied)

    def _applied(self,result):
        errors=getattr(result,'failures',None) or getattr(result,'unavailable',None) or getattr(result,'errors',None)
        if isinstance(result,dict):
            errors=result.get('errors') or result.get('failed')
        if errors:
            self.show_error("Partially applied: "+str(errors))
        else:
            # Close the Studio once the arrangement is applied.
            self.close()

    def rename_template(self):
        template=next((t for t in self.controller.templates if t.id==self.template_id),None)
        if template:
            self._name_operation='rename';self.name_entry.setText(template.name)
            self.name_save.setText('Save name')
            self.name_row.show();self.name_entry.setFocus()

    def delete_template(self):
        template=next((t for t in self.controller.templates if t.id==self.template_id),None)
        if template:
            self.bridge.submit('delete_template',template.id)

    def reset_space(self):
        if self.mode!='current':
            return
        # Reset Space edits the draft only: Auto, no assignments.
        def reset():
            self.controller.clear_space(self.display_id,self.space)
            self.controller.set_preset(self.display_id,self.space,'auto')
        self.bridge.submit(reset,message='Space layout cleared. Other spaces are unchanged. Apply to confirm.')

    def import_archive(self):
        path,_=QFileDialog.getOpenFileName(self,"Import SmartGrid layouts","","JSON files (*.json)")
        if path:
            def imported(report):
                self.show_status(f"Imported {report['added']} layouts and {report['profiles_added']} profiles")
            self.bridge.submit('import_archive',path,on_success=imported)

    def export_archive(self):
        path,_=QFileDialog.getSaveFileName(self,"Export SmartGrid layouts","smartgrid-layouts.json","JSON files (*.json)")
        if path:
            self.bridge.submit('export_archive',path,message="Layouts and profiles exported")

    def _show_overview(self,visible):
        self.overview.setVisible(visible)
        if visible:
            self._fill_overview()

    def _fill_overview(self):
        self.overview.clear()
        for display in self.controller.displays:
            parent=QTreeWidgetItem([display.name,"","",""])
            self.overview.addTopLevelItem(parent)
            for space in range(3):
                draft=self.controller.draft(display.id,space)
                active=self.controller.active_spaces.get(display.id,0)==space
                item=QTreeWidgetItem([f"Space {space+1}",preset_name(draft.profile.preset),str(sum(a is not None for a in draft.profile.assignments)),
                                      "Edited" if draft.dirty else "Active" if active else ""])
                item.setData(0,Qt.ItemDataRole.UserRole,(display.id,space))
                parent.addChild(item)
            parent.setExpanded(True)

    def _overview_activated(self,item,column):
        context=item.data(0,Qt.ItemDataRole.UserRole)
        if context:
            self.display_id,self.space=context
            self.mode='current'
            with QSignalBlocker(self.mode_combo):
                self.mode_combo.setCurrentIndex(0)
            self.refresh_from_controller()

    def focus_library(self):
        if self.selected_index<0: return
        self._show_library(True)
        self.library.focus_search()

    def _cancel_name(self):
        current=next((t for t in self.controller.templates if t.id==self.template_id),None) if self.mode=='edit' else None
        self.name_entry.setText(current.name if current else '')
        self.name_row.hide();self.replace_row.hide()

    def _preferences(self):
        from .preferences import Preferences
        self._preferences_window=Preferences(self.controller,self)
        self._preferences_window.show()

    def show_error(self,text):
        self.message.setProperty('error',True)
        self.message.setText(str(text))
        self.message.style().unpolish(self.message)
        self.message.style().polish(self.message)
        self.message.show()

    def show_status(self,text):
        self.message.setProperty('error',False)
        self.message.setText(str(text))
        self.message.style().unpolish(self.message)
        self.message.style().polish(self.message)
        self.message.show()

    def _busy(self,busy):
        self.toggle_button.setEnabled(not busy)
        self.apply_button.setEnabled(not busy)
        self.save_button.setEnabled(not busy)

    def resizeEvent(self,event):
        if hasattr(self,'preset_buttons'):
            columns=max(3,(event.size().width()-40)//85)
            for i,button in enumerate(b for k,b in self.preset_buttons.items() if k!='auto'):
                self.preset_grid.addWidget(button,i//columns,i%columns)
        if hasattr(self,'preset_menu_button') and not self._refreshing:
            QTimer.singleShot(0,self.refresh_from_controller)
        if hasattr(self,'effect'):
            self.effect.setVisible(bool(self.effect.text()) and event.size().height()>=600)
        super().resizeEvent(event)

    def closeEvent(self,event):
        # Closing preserves every draft, as Escape does in the reference.
        event.accept()
