"""Two collapsed, independently searchable application sections."""
import json
from PySide6.QtCore import Qt,Signal,QMimeData,QTimer,QSignalBlocker,QSize
from PySide6.QtGui import QDrag
from PySide6.QtWidgets import QWidget,QVBoxLayout,QHBoxLayout,QLabel,QLineEdit,QListWidget,QListWidgetItem,QTabWidget,QPushButton,QAbstractItemView,QToolButton
from smartgrid.core.models import app_display_name
from .tile_view import MIME,application_icon
from .rows import CardDelegate


class AssignmentList(QListWidget):
    def __init__(self,parent=None):
        super().__init__(parent);self.setDragEnabled(True)
        self.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
    def startDrag(self,supported_actions):
        item=self.currentItem()
        if not item or not item.data(Qt.ItemDataRole.UserRole) or item.data(Qt.ItemDataRole.UserRole+1): return
        mime=QMimeData();mime.setData(MIME,json.dumps(item.data(Qt.ItemDataRole.UserRole)).encode())
        drag=QDrag(self);drag.setMimeData(mime);drag.setPixmap(item.icon().pixmap(32,32));drag.exec(Qt.DropAction.CopyAction)


class ApplicationLibrary(QWidget):
    assign_requested=Signal(object)
    def __init__(self,controller,bridge,parent=None,location=None):
        super().__init__(parent)
        self.controller,self.bridge=controller,bridge
        self.location=location
        self.apps_loaded=False;self.apps_loading=False;self._content_signatures={}
        layout=QVBoxLayout(self);layout.setContentsMargins(0,0,0,0)
        # The Studio's Back to layout button labels this view.
        self.tabs=QTabWidget(self);self.tabs.addTab(QWidget(),'Windows');self.tabs.addTab(QWidget(),'Apps');self.tabs.hide()
        self.sections=[];self.section_buttons=[];self.search_fields=[]
        self.windows,self.apps=AssignmentList(),AssignmentList()
        for index,title,listing in [(0,'Open windows',self.windows),(1,'Installed apps',self.apps)]:
            button=QToolButton();button.setText(title);button.setCheckable(True)
            button.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon);button.setArrowType(Qt.ArrowType.RightArrow)
            button.setAccessibleName(title);layout.addWidget(button)
            section=QWidget();content=QVBoxLayout(section);content.setContentsMargins(0,0,0,0)
            search=QLineEdit();search.setPlaceholderText('Search '+title.lower());search.setAccessibleName('Search '+title.lower())
            listing.setItemDelegate(CardDelegate(listing));listing.setMouseTracking(True);listing.setIconSize(QSize(30,30))
            content.addWidget(search);content.addWidget(listing,1);listing.setMinimumHeight(130)
            section.hide();layout.addWidget(section,1)
            self.sections.append(section);self.section_buttons.append(button);self.search_fields.append(search)
            button.toggled.connect(lambda checked,i=index:self._expanded(i,checked))
            search.textChanged.connect(lambda text,i=index:self._search_changed(i))
            listing.itemActivated.connect(self._activate_item)
            # A single click assigns.
            listing.itemClicked.connect(self._activate_item)
            listing.currentItemChanged.connect(lambda current,previous,i=index:self._selected(i))
        self.stretch=QWidget();layout.addWidget(self.stretch,1)
        self.hint=QLabel();self.hint.setProperty('muted',True);self.hint.hide()
        self.hint.setWordWrap(True);layout.addWidget(self.hint)
        actions=QHBoxLayout();self.assign=QPushButton('Assign to selected tile');self.assign.clicked.connect(self.assign_current)
        actions.addWidget(self.assign);self.assign.hide();self.include=QPushButton('Include');self.include.clicked.connect(self.include_current)
        self.include.setToolTip('Explicitly include this excluded application');actions.addWidget(self.include);layout.addLayout(actions)
        self.timer=QTimer(self);self.timer.setSingleShot(True);self.timer.setInterval(120);self.timer.timeout.connect(self.refresh_from_controller)
        self.tabs.currentChanged.connect(self._tab_changed)
        bridge.changed.connect(self.refresh_from_controller,Qt.ConnectionType.QueuedConnection)
        bridge.error.connect(self._load_failed,Qt.ConnectionType.QueuedConnection)
        self._selection_changed()

    def _load_failed(self,message):
        if self.apps_loading:
            self.apps_loading=False
            self.hint.setText(message);self.hint.show()

    @property
    def search(self): return self.search_fields[self.tabs.currentIndex()]

    def _tab_changed(self,index):
        self.section_buttons[index].setChecked(True);self._selection_changed()

    def _expanded(self,index,expanded):
        self.sections[index].setVisible(expanded)
        self.stretch.setVisible(not any(section.isVisibleTo(self) for section in self.sections))
        self.section_buttons[index].setArrowType(Qt.ArrowType.DownArrow if expanded else Qt.ArrowType.RightArrow)
        if expanded:
            with QSignalBlocker(self.tabs): self.tabs.setCurrentIndex(index)
            if index==1 and not self.apps_loaded and not self.apps_loading:
                self.apps_loading=True
                self.bridge.submit('refresh_apps',on_success=self._apps_ready)
            self.refresh_from_controller()
        self._selection_changed()

    def _apps_ready(self,result):
        self.apps_loaded=True;self.apps_loading=False;self.refresh_from_controller()

    def _search_changed(self,index):
        with QSignalBlocker(self.tabs): self.tabs.setCurrentIndex(index)
        self.timer.start()

    def _selected(self,index):
        with QSignalBlocker(self.tabs): self.tabs.setCurrentIndex(index)
        self._selection_changed()

    def _activate_item(self,item):
        if item and not item.data(Qt.ItemDataRole.UserRole+1): self.assign_requested.emit(item.data(Qt.ItemDataRole.UserRole))

    def refresh_from_controller(self):
        ignored={'Not visible','Child window','Tool window','Does not activate','Windows desktop or shell surface','SmartGrid interface','Owned popup','Cloaked or on another Windows virtual desktop'}
        windows=[w for w in self.controller.windows if w.exclusion_reason not in ignored]
        apps={a.id:a for a in self.controller.apps};displays={d.id:d.name for d in self.controller.displays}
        for index,listing in enumerate((self.windows,self.apps)):
            if not self.section_buttons[index].isChecked(): continue
            query=self.search_fields[index].text().casefold()
            signature=(query,self.apps_loaded,self.apps_loading,tuple((a.id,a.name,a.icon_path) for a in apps.values()),
                tuple((w.ref,w.app_id,w.title,w.display_id,w.state,w.eligible,w.exclusion_reason,w.floating) for w in windows) if index==0 else (),
                tuple(displays.items()) if index==0 else (),
                tuple(self.location(w) for w in windows) if index==0 and self.location else ())
            if signature==self._content_signatures.get(index): continue
            self._content_signatures[index]=signature
            selected=listing.currentItem().data(Qt.ItemDataRole.UserRole) if listing.currentItem() else None
            with QSignalBlocker(listing):
                listing.clear()
                for window in sorted(windows,key=lambda w:(w.app_id,w.ref.hwnd)) if index==0 else ():
                    app=apps.get(window.app_id);name=app.name if app else app_display_name(window.app_id)
                    if query and query not in (name+' '+window.title+' '+displays.get(window.display_id,'')).casefold(): continue
                    source=displays.get(window.display_id,window.display_id)
                    where=self.location(window) if self.location else source
                    item=QListWidgetItem(application_icon(app),f'{window.title or name}\n{name} · {where}' if where else f'{window.title or name}\n{name}')
                    item.setData(Qt.ItemDataRole.AccessibleTextRole,f'Assign window {window.title or name} to the selected tile')
                    item.setData(Qt.ItemDataRole.UserRole,{'app_id':window.app_id,'window_id':window.ref.hwnd})
                    item.setData(Qt.ItemDataRole.UserRole+1,not window.eligible)
                    item.setToolTip(f'{window.title}\n{source}\n{window.exclusion_reason or window.state}');listing.addItem(item)
                    if selected==item.data(Qt.ItemDataRole.UserRole): listing.setCurrentItem(item)
                if index==1 and self.apps_loaded:
                    for app in sorted((a for a in apps.values() if a.installed),key=lambda a:a.name.casefold()):
                        if query and query not in (app.name+' '+app.id).casefold(): continue
                        item=QListWidgetItem(application_icon(app),f'{app.name}\n{app.id}');item.setData(Qt.ItemDataRole.UserRole,{'app_id':app.id,'window_id':None})
                        item.setToolTip(f'Assign {app.name} ({app.id}) to the selected tile');listing.addItem(item)
                        if selected==item.data(Qt.ItemDataRole.UserRole): listing.setCurrentItem(item)
        empty=any(b.isChecked() and l.count()==0 and f.text() for b,l,f in zip(self.section_buttons,(self.windows,self.apps),self.search_fields))
        self.hint.setText('Loading installed applications…' if self.apps_loading else 'No matches.' if empty else '')
        self.hint.setVisible(bool(self.hint.text()))
        counts=(len(windows),sum(a.installed for a in apps.values()) if self.apps_loaded else None)
        for button,title,count in zip(self.section_buttons,('Open windows','Installed apps'),counts):
            button.setText(title if count is None else f'{title} · {count}')
        self._selection_changed()

    def _selection_changed(self):
        index=self.tabs.currentIndex();listing=(self.windows,self.apps)[index]
        item=listing.currentItem() if self.section_buttons[index].isChecked() else None
        excluded=bool(item and item.data(Qt.ItemDataRole.UserRole+1))
        self.assign.setEnabled(bool(item) and not excluded);self.include.setVisible(excluded)

    def assign_current(self):
        listing=(self.windows,self.apps)[self.tabs.currentIndex()];item=listing.currentItem()
        if item and self.assign.isEnabled(): self.assign_requested.emit(item.data(Qt.ItemDataRole.UserRole))

    def include_current(self):
        item=self.windows.currentItem()
        if item: self.bridge.submit('include_application',item.data(Qt.ItemDataRole.UserRole)['app_id'],message='Application included.')

    def focus_search(self):
        self.section_buttons[self.tabs.currentIndex()].setChecked(True);self.search.setFocus();self.search.selectAll()
