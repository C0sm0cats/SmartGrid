"""Explicit restoration picker showing the controller's actual reuse/open/hide plan."""
from PySide6.QtCore import Qt, QSignalBlocker,QSize
from PySide6.QtGui import QAction,QKeySequence
from PySide6.QtWidgets import QDialog,QVBoxLayout,QHBoxLayout,QLabel,QPushButton,QComboBox,QListWidget,QListWidgetItem,QSplitter,QWidget,QLineEdit
from smartgrid.core.geometry import capacity,preset_name
from .bridge import bridge_for
from .tile_view import TileCanvas,layout_icon
from .rows import CardDelegate,CardButton,quick_preview
from .theme import apply_theme,studio_settings
from .backdrop import show_backdrops,hide_backdrops
from .dialog_surface import paint_surface,animate_open


def count_effect(value):
    if isinstance(value,(list,tuple,set,dict)):
        return len(value)
    return int(value or 0)


def summary(parts):
    return '  ·  '.join(f'{count} {label}' for label,count in parts if count>0) or 'No windows to open or hide'


class QuickSwitcher(QDialog):
    def __init__(self,controller,display_id=None,parent=None):
        super().__init__(parent)
        self.controller,self.bridge=controller,bridge_for(controller)
        self.display_id=display_id or controller.current_display_id()
        self.space=controller.active_spaces.get(self.display_id,0)
        self.previews={}
        self._preview_generation=0
        self.setWindowTitle("SmartGrid · Change layout")
        self.setWindowFlags(Qt.WindowType.Tool|Qt.WindowType.FramelessWindowHint|Qt.WindowType.WindowStaysOnTopHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        apply_theme(studio_settings(controller.settings),target=self)
        self.resize(560,570)
        self.setMinimumSize(500,420)
        layout=QVBoxLayout(self)
        layout.setContentsMargins(20,20,20,20)
        top=QHBoxLayout();titles=QVBoxLayout();titles.setSpacing(3);top.addLayout(titles,1)
        from .app import DesktopUI
        mark=QLabel();mark.setPixmap(DesktopUI.tray_icon('#8ce8c3').pixmap(30,30));top.addWidget(mark,0,Qt.AlignmentFlag.AlignTop)
        eyebrow=QLabel("SMARTGRID");eyebrow.setStyleSheet('color:#8ce8c3;font-size:9px;font-weight:800;letter-spacing:2px;');titles.addWidget(eyebrow)
        heading=QLabel("Change layout")
        heading.setStyleSheet("font-size: 23px; font-weight: 800; color: #f0f6f4;")
        titles.addWidget(heading)
        self.context_label=QLabel();self.context_label.setStyleSheet('color:#a6b9bd;font-size:11px;')
        titles.addWidget(self.context_label);layout.addLayout(top);layout.addSpacing(8)
        context=QHBoxLayout()
        self.displays=QComboBox()
        for display in controller.displays:
            self.displays.addItem(display.name,display.id)
        self.displays.setCurrentIndex(max(0,self.displays.findData(self.display_id)))
        self.displays.setAccessibleName("Display")
        self.displays.hide()
        self.displays.currentIndexChanged.connect(self._context_changed)
        self.spaces=QComboBox()
        for i in range(3):
            self.spaces.addItem(f"Space {i+1}",i)
        self.spaces.setCurrentIndex(self.space)
        self.spaces.hide()
        self.spaces.currentIndexChanged.connect(self._context_changed)
        context.addWidget(self.displays,1)
        context.addWidget(self.spaces)
        layout.addLayout(context)
        self.arrange=CardButton();self.arrange.setFixedHeight(76)
        self.arrange.setIconSize(QSize(84,54))
        self.arrange.clicked.connect(self.arrange_open)

        layout.addWidget(self.arrange)
        self.search=QLineEdit()
        self.search.setPlaceholderText("Search saved layouts")
        self.search.setAccessibleName("Search saved layouts")
        self.search.textChanged.connect(self.filter_templates)
        layout.addWidget(self.search)
        splitter=QSplitter(Qt.Orientation.Horizontal)
        self.list=QListWidget()
        self.list.setAccessibleName("Saved layouts")
        self.list.currentItemChanged.connect(self._selection_changed)
        self.list.itemActivated.connect(lambda:self.restore_selected())
        self.list.itemClicked.connect(lambda item:self.restore_selected() if item.data(Qt.ItemDataRole.UserRole) else None)
        self.list.setIconSize(QSize(84,54))
        self.list.setItemDelegate(CardDelegate(self.list,switcher=True));self.list.setMouseTracking(True)
        self.list.setSpacing(0)
        self.list.setStyleSheet('QListWidget { background:#10161d; border:1px solid #304049; border-radius:13px; padding:9px; }'
            ' QListWidget::item { background:#222e36; border:1px solid transparent; padding:9px; border-radius:11px; margin:2px 0; color:#edf5f2; }'
            ' QListWidget::item:selected, QListWidget::item:hover { background:#30443f; border-color:#8ce8c3; }'
            ' QListWidget::item:disabled { background:transparent; border:none; color:#8ce8c3; font-size:9px; font-weight:800; padding:9px 7px 4px; }')
        splitter.addWidget(self.list)
        preview_container=QWidget()
        preview_layout=QVBoxLayout(preview_container)
        self.preview=TileCanvas(studio_settings(controller.settings))
        self.preview.setEnabled(False)
        preview_layout.addWidget(self.preview,1)
        self.effects=QLabel("Select a saved layout.")
        self.effects.setWordWrap(True)
        preview_layout.addWidget(self.effects)
        splitter.addWidget(preview_container)
        preview_container.hide()
        splitter.setSizes([340,360])
        layout.addWidget(splitter,1)
        self.empty=QLabel('SAVED LAYOUTS\nCreate a layout in Studio to see it here.');self.empty.setWordWrap(True)
        self.empty.setProperty('muted',True);layout.addWidget(self.empty)
        self.message=QLabel()
        self.message.setWordWrap(True)
        self.message.setProperty('error',True)
        layout.addWidget(self.message)
        footer=QHBoxLayout()
        footer.addStretch()
        self.restore=QPushButton("Restore in this space")
        self.restore.setProperty('primary',True)
        self.restore.setEnabled(False)
        self.restore.clicked.connect(self.restore_selected)
        footer.addWidget(self.restore);self.restore.hide()
        close=QPushButton("Cancel")
        close.clicked.connect(self.reject)
        footer.addWidget(close)
        layout.addLayout(footer)
        self.bridge.error.connect(self.message.setText,Qt.ConnectionType.QueuedConnection)
        self.bridge.changed.connect(self.refresh_from_controller,Qt.ConnectionType.QueuedConnection)
        self.bridge.busy_changed.connect(lambda busy:self.arrange.setEnabled(not busy))
        self._template_signature=None
        self.refresh_from_controller()
        # Catalogue refresh is lazy: opening the picker is an explicit need for it.
        self.bridge.submit('refresh_apps',on_success=lambda _:self.refresh_previews())

    def paintEvent(self,event):
        paint_surface(self)

    def showEvent(self,event):
        super().showEvent(event);show_backdrops(self);animate_open(self)

    def hideEvent(self,event):
        hide_backdrops(self);super().hideEvent(event)

    def _context_changed(self):
        self.display_id=self.displays.currentData()
        self.space=self.spaces.currentData()
        self._refresh_context()
        self.refresh_previews()

    def _refresh_context(self):
        index=next((i for i,d in enumerate(self.controller.displays) if d.id==self.display_id),0)
        self.context_label.setText(f'Display {index+1}  ·  Space {self.space+1}')
        windows=[w for w in self.controller.windows if w.display_id==self.display_id and w.eligible
                 and not w.floating and w.state in ('normal','maximized')]
        count=len(windows)
        self.arrange.setIcon(quick_preview('auto',max(1,count),set(range(max(1,count)))))
        self.arrange.setText(f"Arrange open windows\nAuto · {count} open window{'' if count==1 else 's'}\nRearrange open windows  ·  No apps opened")
        self.arrange.setAccessibleName(self.arrange.text().replace('\n','. '))

    def refresh_from_controller(self):
        choices=tuple((d.id,d.name) for d in self.controller.displays)
        if choices!=tuple((self.displays.itemData(i),self.displays.itemText(i)) for i in range(self.displays.count())):
            with QSignalBlocker(self.displays):
                self.displays.clear()
                for key,name in choices: self.displays.addItem(name,key)
                if self.display_id not in {key for key,name in choices}:
                    self.display_id=self.controller.current_display_id()
                    self.space=self.controller.active_spaces.get(self.display_id,0)
                self.displays.setCurrentIndex(self.displays.findData(self.display_id))
            with QSignalBlocker(self.spaces): self.spaces.setCurrentIndex(self.space)
        self._refresh_context()
        self.search.setVisible(bool(self.controller.templates))
        signature=(self.display_id,self.space,tuple(repr(t.to_dict()) for t in self.controller.templates),
            tuple((w.ref,w.app_id,w.state,w.display_id,w.eligible) for w in self.controller.windows),
            tuple(a.id for a in self.controller.apps))
        if signature!=self._template_signature:
            self._template_signature=signature
            current=self.list.currentItem().data(Qt.ItemDataRole.UserRole) if self.list.currentItem() else None
            with QSignalBlocker(self.list):
                self.list.clear()
                category=None
                for template in sorted(self.controller.templates,key=lambda t:(t.preset!='custom',capacity(t.preset,t.tiles),t.name.casefold())):
                    group='SAVED CUSTOM LAYOUTS' if template.preset=='custom' else 'SAVED PRESET LAYOUTS'
                    if group!=category:
                        header=QListWidgetItem(group);header.setFlags(Qt.ItemFlag.NoItemFlags)
                        header.setData(Qt.ItemDataRole.UserRole+1,group)
                        self.list.addItem(header);category=group
                    item=QListWidgetItem(f"{template.name}\n{self.metadata(template)}\nCalculating effects…")
                    item.setData(Qt.ItemDataRole.UserRole,template.id)
                    self.list.addItem(item)
                    if template.id==current:
                        self.list.setCurrentItem(item)
                if self.list.count() and self.list.currentRow()<0:
                    self.list.setCurrentRow(next((i for i in range(self.list.count()) if self.list.item(i).data(Qt.ItemDataRole.UserRole)),0))
            if not self.list.count():
                self.effects.setText("Create a layout in Studio to see it here.")
            self.refresh_previews()
            self.filter_templates()
        self._selection_changed()

    def filter_templates(self):
        query=self.search.text().casefold().strip()
        templates={t.id:t for t in self.controller.templates}
        first=None
        visible_groups=set()
        for row in range(self.list.count()):
            item=self.list.item(row)
            template=templates.get(item.data(Qt.ItemDataRole.UserRole))
            hidden=not template or bool(query and query not in template.name.casefold())
            item.setHidden(hidden)
            if not hidden: visible_groups.add('SAVED CUSTOM LAYOUTS' if template.preset=='custom' else 'SAVED PRESET LAYOUTS')
            if not hidden and first is None:
                first=item
        for row in range(self.list.count()):
            item=self.list.item(row);group=item.data(Qt.ItemDataRole.UserRole+1)
            if group: item.setHidden(group not in visible_groups)
        current=self.list.currentItem()
        if current is None or current.isHidden():
            self.list.setCurrentItem(first)
        self.empty.setVisible(first is None)
        self.empty.setText("No matching layouts." if query else "SAVED LAYOUTS\nCreate a layout in Studio to see it here.")
        if first is None:
            self.restore.setEnabled(False)
            self.preview.hide()
            self.effects.setText("No matching layouts." if query else "Create a layout in Studio to see it here.")
        else:
            self.preview.show()
        self._selection_changed()

    def refresh_previews(self):
        if not self.display_id:
            return
        self._preview_generation+=1
        generation=self._preview_generation
        display_id,space=self.display_id,self.space
        ids=[t.id for t in self.controller.templates]
        self.restore.setEnabled(False)
        self.effects.setText("Calculating effects…")
        def preview_all():
            return {key:self.controller.preview_template(key,display_id,space) for key in ids}
        def ready(previews):
            if generation!=self._preview_generation:
                return
            self.previews=previews
            for row in range(self.list.count()):
                item=self.list.item(row)
                key=item.data(Qt.ItemDataRole.UserRole)
                template=next((t for t in self.controller.templates if t.id==key),None)
                if template:
                    item.setText(f"{template.name}\n{self.metadata(template)}\n{self.effect_text(previews.get(key,{}))}")
                    item.setToolTip(item.text())
                    count=max(sum(a is not None for a in template.assignments),capacity(template.preset,template.tiles))
                    item.setIcon(quick_preview(template.preset,count,{i for i,a in enumerate(template.assignments) if a},
                                               {i for i,a in enumerate(template.assignments) if a and a.pinned},template.tiles,template.ratio))
                    profile=self.controller.profile(self.display_id,self.space)
                    item.setData(Qt.ItemDataRole.UserRole+3,self.controller.running and profile.preset==template.preset
                                 and [a.app_id if a else None for a in profile.assignments]==[a.app_id if a else None for a in template.assignments])
                    item.setSizeHint(QSize(0,82))
        self._selection_changed()
        self.bridge.submit(preview_all,on_success=ready)

    @staticmethod
    def metadata(template):
        count=max(sum(a is not None for a in template.assignments),capacity(template.preset,template.tiles))
        return f"{count} TILES" if template.preset in ('custom','auto') else f"{count} TILES  ·  {preset_name(template.preset)}"

    @staticmethod
    def effect_text(preview):
        return summary([('reused',count_effect(preview.get('reuse'))),('to open',count_effect(preview.get('open'))),
                        ('to hide',count_effect(preview.get('hide'))),('unavailable',count_effect(preview.get('unavailable')))])

    def _selection_changed(self,*args):
        item=self.list.currentItem()
        key=item.data(Qt.ItemDataRole.UserRole) if item and not item.isHidden() else None
        self.restore.setEnabled(bool(key and key in self.previews) and not self.bridge.pending)
        if key:
            template=next((t for t in self.controller.templates if t.id==key),None)
            if template:
                try:
                    profile=self.controller.load_template_draft(key,self.display_id,self.space)
                    display=next((d for d in self.controller.displays if d.id==self.display_id),None)
                    self.preview.set_profile(profile,display,self.controller.apps,[])
                    self.effects.setText(self.effect_text(self.previews[key]) if key in self.previews else "Calculating effects…")
                except ValueError as exc:
                    self.message.setText(str(exc))

    def restore_selected(self):
        item=self.list.currentItem()
        if not item or item.isHidden() or not self.restore.isEnabled():
            return
        key=item.data(Qt.ItemDataRole.UserRole)
        self.restore.setEnabled(False)
        def restored(result):
            if getattr(result,'success',True):
                self.accept()
            else:
                self.message.setText(getattr(result,'summary',"Partially restored."))
                self.refresh_previews()
        self.bridge.submit('restore_template',key,self.display_id,self.space,on_success=restored)

    def arrange_open(self):
        display_id,space=self.display_id,self.space
        def arrange():
            # Match the selected picker context, rather than the old active space.
            self.controller.switch_space(display_id,space)
            return self.controller.arrange(display_id)
        def arranged(result):
            if getattr(result,'success',True):
                self.accept()
            else:
                self.message.setText(getattr(result,'summary','Partially arranged.'))
        self.bridge.submit(arrange,on_success=arranged)
