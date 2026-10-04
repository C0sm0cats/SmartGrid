"""Preferences window with debounced, immediate persistence."""
from copy import deepcopy
from dataclasses import asdict
from PySide6.QtCore import Qt,QTimer,QSignalBlocker
from PySide6.QtGui import QColor,QKeySequence
from PySide6.QtWidgets import (QDialog,QWidget,QVBoxLayout,QHBoxLayout,QFormLayout,QLabel,QPushButton,
    QTabWidget,QScrollArea,QComboBox,QSpinBox,QDoubleSpinBox,QCheckBox,QColorDialog,QKeySequenceEdit,
    QPlainTextEdit,QGroupBox,QFileDialog,QLineEdit,QTreeWidget,QTreeWidgetItem)
from smartgrid.core.models import Settings,app_display_name
from smartgrid.core.controller import DEFAULT_HOTKEYS
from .bridge import bridge_for
from .tile_view import TileCanvas,application_icon
from .theme import apply_theme
from smartgrid import __version__
from smartgrid.core.geometry import preset_name,effective_layout
from .widgets import Switch,SpacePreview,ShortcutEditor

# Shortcut titles, in display order.
HOTKEY_LABELS={'toggle':'Toggle tiling','arrange':'Arrange again','studio':'Open Layout Studio',
    'quick_switcher':'Change layout','float':'Float focused window','swap':'Swap mode',
    'undo':'Undo','redo':'Redo','focus_left':'Focus window to the left',
    'focus_right':'Focus window to the right','focus_up':'Focus window above','focus_down':'Focus window below',
    'space1':'Monitor space 1','space2':'Monitor space 2','space3':'Monitor space 3','stop':'Stop and restore'}


class Preferences(QDialog):
    def __init__(self,controller,parent=None):
        super().__init__(parent)
        self.controller,self.bridge=controller,bridge_for(controller)
        self.settings=deepcopy(controller.settings)
        self.controls,self.resets,self.margin_controls,self.hotkeys={},{},{},{}
        self._building=True
        self._apps_loaded=False
        self._rules_signature=None
        self._last_submitted=None
        self.setWindowTitle('SmartGrid Preferences')
        self.resize(660,720)
        self.setMinimumSize(440,420)
        layout=QVBoxLayout(self)
        layout.setContentsMargins(16,12,16,12)
        self.tabs=QTabWidget()
        # Adw.ViewSwitcher: page tabs centred above the content.
        self.tabs.tabBar().setExpanding(False)
        layout.addWidget(self.tabs,1)
        main=self._page(f'SmartGrid v{__version__}')
        room=self._group(main,'Make room','Fine-tune the space around your windows.')
        self._spin(room,'gap','Window spacing',0,64)
        self._spin(room,'padding','Screen edge spacing',0,64)
        self._check(room,'independent_padding','Separate screen edges','Set a different margin on each side')
        self.edge_rows=[]
        for edge in ('top','right','bottom','left'):
            spin=QSpinBox();spin.setRange(0,max(64,self.settings.margins[edge]))
            spin.setValue(self.settings.margins[edge])
            spin.setAccessibleName(f'{edge.title()} edge spacing')
            self.margin_controls[edge]=spin
            container=QWidget();row=QHBoxLayout(container);row.setContentsMargins(0,0,0,0)
            reset=QPushButton('↶');reset.setFixedWidth(38);reset.setToolTip(f'Reset {edge} spacing to default')
            reset.clicked.connect(lambda checked=False,e=edge:self.margin_controls[e].setValue(Settings().margins[e]))
            spin.setFixedWidth(120);row.addStretch();row.addWidget(reset);row.addWidget(spin)
            room.addRow(f'{edge.title()} edge spacing',container)
            self.edge_rows.append(container)
            spin.valueChanged.connect(self._changed)
        self._spin(room,'master_ratio','Focus layout: large tile width',25,75,percent=True,
            subtitle='Percentage of the available width reserved for the left tile')
        self.preview_label=QLabel('Start arranging windows to see the current layout')
        self.preview_label.setToolTip('Live layout of the active desktop, display and SmartGrid space')
        self.preview_label.setWordWrap(True)
        self.preview_label.setProperty('muted',True)
        self.preview=SpacePreview()
        preview_row=QWidget();preview_layout=QHBoxLayout(preview_row);preview_layout.setContentsMargins(0,0,0,0)
        preview_layout.addStretch();preview_layout.addWidget(self.preview)
        self.preview_title=self._title('Current space preview','')
        self.preview_label.hide()
        room.addRow(self.preview_title,preview_row)
        flow=self._group(main,'Keep your flow')
        for key,title in [('active_border','Highlight the focused window'),('animations','Animate placement guides'),
                          ('compact_minimize','Close gaps when minimizing'),('compact_close','Close gaps when closing')]:
            self._check(flow,key,title)
        self.focus_group=self._group(main,'Focus outline','Customize the outline around the focused tiled window.')
        self._check(self.focus_group,'border_custom_color','Custom focus color','Off: follow the Windows accent color')
        self.color_button=QPushButton()
        self._paint_color_button(self.settings.border_color)
        self.color_button.setAccessibleName('Focus color')
        self.color_button.clicked.connect(self.choose_focus_color)
        self._row(self.focus_group,'border_color','Focus color',self.color_button)
        self._spin(self.focus_group,'border_width','Outline thickness',1,6,subtitle='Pixels')
        self._combo(self.focus_group,'border_style','Outline style',[('outline','Outline'),('glow','Subtle halo')])
        self.animation_group=self._group(main,'Animations','Adjust guides, focus effects and space transitions. Window placement stays immediate.')
        self._combo(self.animation_group,'animation_speed','Animation speed',[('fast','Fast'),('normal','Normal'),('slow','Slow'),('custom','Custom')])
        self._spin(self.animation_group,'animation_duration','Animation duration',40,500,
            subtitle='Base duration in milliseconds; longer transitions scale proportionally')
        self._combo(self.animation_group,'animation_curve','Animation curve',[('ease-out','Ease out'),('linear','Linear'),('ease-in-out','Ease in and out')])
        archive=self._group(main,'Back up and share','Export named layouts and per-space profiles to a JSON file. Import adds new items without replacing your existing ones.')
        self.export_button=QPushButton('Export…');self.export_button.clicked.connect(self.export_archive)
        self.import_button=QPushButton('Import…');self.import_button.clicked.connect(self.import_archive)
        archive.addRow(self._title('Export layouts and profiles','Save a portable SmartGrid JSON file'),self.export_button)
        archive.addRow(self._title('Import layouts and profiles','Add layouts; keep existing profiles when their space already exists'),self.import_button)
        shortcuts=self._group(main,'Keyboard shortcuts','Record a combination or use the pencil to edit it. Clear it to disable.')
        mapping=DEFAULT_HOTKEYS|self.settings.hotkeys
        for action in [a for a in HOTKEY_LABELS if a in DEFAULT_HOTKEYS]:
            # Rows: key caps, a pencil to type the shortcut, a record button.
            edit=ShortcutEditor(mapping.get(action,'').replace('Win+','Meta+'),HOTKEY_LABELS.get(action,action))
            self.hotkeys[action]=edit
            shortcuts.addRow(HOTKEY_LABELS.get(action,action),edit)
            edit.keySequenceChanged.connect(self._changed)
        self.conflicts=QLabel();self.conflicts.setWordWrap(True);self.conflicts.setProperty('error',True)
        shortcuts.addRow(self.conflicts)
        self._group(main,'Three spaces, per display','SmartGrid parks inactive-space windows by minimizing them. Stopping restores your windows. Windows virtual desktops remain separate. Configure other tilers to avoid competing shortcuts and placements.')
        apps_page=self._page('Applications')
        rules=self._group(apps_page,'Application rules','Search an application, then choose how SmartGrid handles its windows.')
        self._check(rules,'builtin_exclusions','Keep common overlays out of the grid',
                    'Media players, game launchers, streaming and monitoring tools, call windows and other window managers float by default')
        self.app_search=QLineEdit();self.app_search.setPlaceholderText('Search applications')
        self.app_search.setAccessibleName('Search applications');self.app_search.textChanged.connect(self.refresh_rules)
        rules.addRow(self.app_search)
        self.rules=QTreeWidget();self.rules.setHeaderHidden(True);self.rules.setColumnCount(1)
        self.rules.setMinimumHeight(300);self.rules.itemChanged.connect(self._rule_changed)
        rules.addRow(self.rules)
        note=QLabel('Scale-to-fit in the GNOME compositor has no equivalent in this native Win32 backend. Minimum client sizes remain enforced.')
        note.setWordWrap(True);note.setProperty('muted',True);rules.addRow(note)
        advanced=self._group(main,'Windows tools and advanced settings','Placement and native animation options for Windows.')
        # Collapsed by default: the checkbox expands the group.
        box=advanced.parentWidget();box.setCheckable(True);box.setChecked(False)
        box.setFlat(True)
        self._check(advanced,'force_resize','Force windows into their tiles',
                    'Resize windows below their minimum size and give fixed-size windows a resizable frame')
        self._check(advanced,'window_animations','Animate native window movement')
        self._spin(advanced,'animation_fps','Native animation rate',1,240)
        self._combo(advanced,'animation_effect','Native movement curve',[('crit_damped','Damped'),('spring','Spring'),('curved','Curved'),('linear','Linear')])
        self._spin(advanced,'debounce_ms','Reconciliation delay',0,10000)
        self._spin(advanced,'tile_retries','Placement retries',0,20)
        timeout=QDoubleSpinBox();timeout.setRange(.01,120);timeout.setValue(self.settings.tile_timeout)
        self._row(advanced,'tile_timeout','Placement time budget',timeout);timeout.valueChanged.connect(self._changed)
        self.included=QPlainTextEdit('\n'.join(self.settings.included_apps))
        self.excluded=QPlainTextEdit('\n'.join(self.settings.excluded_apps))
        # The rule lists are edited on the Applications page; these hold their values.
        for title,edit in [('Explicitly included app IDs',self.included),('Always floating app IDs',self.excluded)]:
            edit.setParent(box);edit.hide();edit.textChanged.connect(self._changed)
        def expand(on,form=advanced):
            for i in range(form.rowCount()): form.setRowVisible(i,on)
            box.setMaximumHeight(16777215 if on else 26)
            box.setProperty('collapsed',not on);box.style().unpolish(box);box.style().polish(box)
            self._refresh_visibility() if not self._building else None
        box.toggled.connect(expand);expand(False)
        self.message=QLabel();self.message.setWordWrap(True);self.message.setProperty('error',True)
        layout.addWidget(self.message)
        # Adwaita preference windows have no footer: the title bar closes them.
        self.save_timer=QTimer(self);self.save_timer.setSingleShot(True);self.save_timer.setInterval(150)
        self.save_timer.timeout.connect(self.save_settings)
        self.bridge.error.connect(self._failed,Qt.ConnectionType.QueuedConnection)
        self.bridge.changed.connect(self.refresh_from_controller,Qt.ConnectionType.QueuedConnection)
        self.tabs.currentChanged.connect(self._page_changed)
        apply_theme(self.settings,target=self)
        self._building=False
        self._refresh_visibility();self.refresh_preview();self.refresh_conflicts()

    def _failed(self,message):
        self._last_submitted=None
        self.message.setText(message)

    def _page(self,title):
        scroll=QScrollArea();scroll.setWidgetResizable(True)
        # Rows wrap their text instead of widening the page (no horizontal scroll).
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        page=QWidget();layout=QVBoxLayout(page);layout.setContentsMargins(12,12,12,12);layout.setSpacing(16)
        scroll.setWidget(page);self.tabs.addTab(scroll,title)
        return layout

    def _group(self,page,title,description=''):
        group=QGroupBox(title);group.setProperty('preferences',True)
        form=QFormLayout(group);form.setSpacing(10)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        if description:
            label=QLabel(description);label.setWordWrap(True);label.setProperty('muted',True);form.addRow(label)
        page.addWidget(group)
        return form

    @staticmethod
    def _title(title,subtitle=''):
        # Adwaita rows show a title with a dimmed subtitle beneath it.
        label=QLabel(f'{title}<br><span style="font-size:small;color:palette(placeholder-text)">{subtitle}</span>' if subtitle else title)
        label.setTextFormat(Qt.TextFormat.RichText if subtitle else Qt.TextFormat.PlainText)
        label.setWordWrap(True)
        # Wrap only when the window is narrow: keep the title on one line up to 300 px.
        label.setMinimumWidth(min(label.fontMetrics().horizontalAdvance(title)+8,300))
        return label

    def _row(self,form,key,title,control,subtitle=''):
        self.controls[key]=control
        control.setAccessibleName(title)
        if subtitle: control.setAccessibleDescription(subtitle)
        container=QWidget();row=QHBoxLayout(container);row.setContentsMargins(0,0,0,0)
        reset=QPushButton('↶');reset.setFixedWidth(38);reset.setToolTip(f'Reset {title} to default')
        reset.clicked.connect(lambda checked=False,k=key:self.reset_setting(k))
        self.resets[key]=reset
        # Adwaita rows keep their control compact at the right edge.
        if isinstance(control,QSpinBox) or isinstance(control,QDoubleSpinBox): control.setFixedWidth(120)
        elif isinstance(control,(QComboBox,QPushButton)): control.setFixedWidth(180)
        row.addStretch();row.addWidget(reset);row.addWidget(control);form.addRow(self._title(title,subtitle),container)
        control._preference_row=(form,container)

    def _spin(self,form,key,title,lower,upper,percent=False,subtitle=''):
        control=QSpinBox();value=getattr(self.settings,key)*100 if percent else getattr(self.settings,key)
        control.setRange(min(lower,round(value)),max(upper,round(value)));control.setValue(round(value))
        if percent: control.setSuffix(' %');control.setSingleStep(5)
        self._row(form,key,title,control,subtitle);control.valueChanged.connect(self._changed)

    def _check(self,form,key,title,subtitle=''):
        control=Switch();control.setChecked(getattr(self.settings,key))
        control.accent=self.controller.focus_color
        self._row(form,key,title,control,subtitle);control.toggled.connect(self._changed)

    def _combo(self,form,key,title,choices):
        control=QComboBox()
        for value,label in choices: control.addItem(label,value)
        index=control.findData(getattr(self.settings,key))
        if index<0:
            control.addItem(str(getattr(self.settings,key)),getattr(self.settings,key));index=control.count()-1
        control.setCurrentIndex(index);self._row(form,key,title,control)
        control.currentIndexChanged.connect(self._changed)

    def reset_setting(self,key):
        value=getattr(Settings(),key);control=self.controls[key]
        if key=='border_color':
            self.settings.border_color=value;self._paint_color_button(value);self._changed()
        elif isinstance(control,QComboBox): control.setCurrentIndex(control.findData(value))
        elif isinstance(control,QCheckBox): control.setChecked(value)
        else: control.setValue(round(value*100) if key=='master_ratio' else value)

    def _paint_color_button(self,color):
        # Gtk.ColorDialogButton: a colour swatch, not a hex label.
        self.color_button.setText('');self.color_button.setFixedSize(56,30)
        self.color_button.setToolTip(color)
        self.color_button.setStyleSheet(f'QPushButton {{ background:{color}; border:2px solid palette(mid); border-radius:6px; }}'
                                        ' QPushButton:disabled { border-style:dashed; }')

    def choose_focus_color(self):
        color=QColorDialog.getColor(QColor(self.settings.border_color),self,'Focus color')
        if color.isValid():
            self.settings.border_color=color.name();self._paint_color_button(color.name());self._changed()

    def collect_settings(self):
        settings=deepcopy(self.settings)
        for key,control in self.controls.items():
            if key=='border_color': continue
            value=control.currentData() if isinstance(control,QComboBox) else control.isChecked() if isinstance(control,QCheckBox) else control.value()
            setattr(settings,key,value/100 if key=='master_ratio' else value)
        settings.margins={edge:control.value() for edge,control in self.margin_controls.items()}
        settings.included_apps=list(dict.fromkeys(x.strip() for x in self.included.toPlainText().splitlines() if x.strip()))
        settings.excluded_apps=list(dict.fromkeys(x.strip() for x in self.excluded.toPlainText().splitlines() if x.strip()))
        mapping=deepcopy(self.settings.hotkeys)
        mapping.update({action:edit.keySequence().toString(QKeySequence.SequenceFormat.PortableText) for action,edit in self.hotkeys.items()})
        settings.hotkeys=mapping
        return Settings.from_dict(asdict(settings))

    def _changed(self,*args):
        if self._building: return
        self._refresh_visibility();self.refresh_preview();self.save_timer.start()

    def _refresh_visibility(self):
        independent=self.controls['independent_padding'].isChecked()
        self.controls['padding'].setEnabled(not independent)
        form=self.controls['padding']._preference_row[0]
        for row in self.edge_rows: form.setRowVisible(row,independent)
        self.focus_group.parentWidget().setEnabled(self.controls['active_border'].isChecked())
        self.color_button.setEnabled(self.controls['border_custom_color'].isChecked())
        self.animation_group.parentWidget().setEnabled(self.controls['animations'].isChecked())
        duration=self.controls['animation_duration']
        duration._preference_row[0].setRowVisible(duration._preference_row[1],self.controls['animation_speed'].currentData()=='custom')
        defaults=Settings()
        for key,reset in self.resets.items():
            control=self.controls[key]
            value=self.settings.border_color if key=='border_color' else control.currentData() if isinstance(control,QComboBox) else control.isChecked() if isinstance(control,QCheckBox) else control.value()/100 if key=='master_ratio' else control.value()
            reset.setEnabled(value!=getattr(defaults,key))

    def _set_preview_subtitle(self,text):
        self.preview_label.setText(text)
        self.preview_title.setText(f'Current space preview<br><span style="font-size:small;color:palette(placeholder-text)">{text}</span>')
        self.preview_title.setTextFormat(Qt.TextFormat.RichText)

    def refresh_preview(self,*args):
        if not hasattr(self,'preview') or self._building: return
        settings=self.collect_settings()
        display_id=self.controller.current_display_id()
        space=self.controller.active_spaces.get(display_id,0)
        display=next((d for d in self.controller.displays if d.id==display_id),None)
        if display and self.controller.running:
            profile=deepcopy(self.controller.profile(display_id,space));profile.ratio=settings.master_ratio
            preset=effective_layout(profile,runtime=True)[0]
            self.preview.set_layout(profile,display,settings,getattr(self.controller,'_last_border',None))
            assigned=any(a and a.window_id for a in profile.assignments)
            name=preset_name(preset) if assigned else 'No tiled windows'
            self._set_preview_subtitle(f'Display {self.controller.displays.index(display)+1} · Space {space+1} · {name}')
            visible=assigned and preset=='master_stack'
        else:
            # The preview area stays, empty.
            self.preview.set_layout(None,None,settings,None)
            self._set_preview_subtitle('Start arranging windows to see the current layout');visible=False
        ratio=self.controls['master_ratio'];ratio._preference_row[0].setRowVisible(ratio._preference_row[1],visible)

    def save_settings(self):
        if self._building: return
        try:
            settings=self.collect_settings();used={}
            from smartgrid.platform.windows.hotkeys import parse_shortcut
            for action,shortcut in settings.hotkeys.items():
                if shortcut:
                    normalized=parse_shortcut(shortcut)
                    # Only shortcuts with Ctrl, Alt or Win are accepted.
                    if not normalized[0]&(1|2|8):
                        raise ValueError('Use a valid shortcut with Ctrl, Alt or Win.')
                    if normalized in used: raise ValueError(f'Duplicate shortcut: {shortcut}')
                    used[normalized]=action
        except ValueError as exc:
            self.message.setText(str(exc));return
        data=asdict(settings)
        if data==self._last_submitted: return
        self._last_submitted=data
        def saved(_):
            self.settings=self.collect_settings()
            apply_theme(settings,target=self)
            self.message.clear();self.refresh_conflicts()
        self.bridge.submit('update_settings',settings,on_success=saved)

    def refresh_conflicts(self):
        self.conflicts.setText('\n'.join(f'{HOTKEY_LABELS.get(action,action)}: {error}' for action,error in self.controller.hotkey_conflicts.items()))

    def refresh_from_controller(self):
        self.refresh_conflicts();self.refresh_preview()
        if self._apps_loaded: self.refresh_rules()

    def _page_changed(self,index):
        if index==1 and not self._apps_loaded:
            self._apps_loaded=True
            self.bridge.submit('refresh_apps',on_success=lambda _:self.refresh_rules())

    def refresh_rules(self,*args):
        if not hasattr(self,'rules') or not self._apps_loaded: return
        query=self.app_search.text().casefold()
        included=set(self.included.toPlainText().splitlines());excluded=set(self.excluded.toPlainText().splitlines())
        known={a.id:a for a in self.controller.apps}
        ids=set(known)|included|excluded|{w.app_id for w in self.controller.windows if w.eligible}
        signature=(query,tuple(sorted(included)),tuple(sorted(excluded)),tuple((i,known[i].name,known[i].icon_path) for i in sorted(known)),tuple(sorted(ids)))
        if signature==self._rules_signature: return
        self._rules_signature=signature
        expanded={self.rules.topLevelItem(i).child(j).data(0,Qt.ItemDataRole.UserRole)
                  for i in range(self.rules.topLevelItemCount())
                  for j in range(self.rules.topLevelItem(i).childCount())
                  if self.rules.topLevelItem(i).child(j).isExpanded()}
        with QSignalBlocker(self.rules):
            self.rules.clear();configured=QTreeWidgetItem(['Configured applications']);available=QTreeWidgetItem(['Other applications'])
            configured.setToolTip(0,'Applications with at least one SmartGrid rule enabled.')
            self.rules.addTopLevelItems([configured,available])
            for app_id in sorted(ids,key=lambda i:known[i].name.casefold() if i in known else app_display_name(i).casefold()):
                app=known.get(app_id);name=app.name if app else app_display_name(app_id)
                if query and query not in (name+' '+app_id).casefold(): continue
                item=QTreeWidgetItem([name]);item.setIcon(0,application_icon(app));item.setToolTip(0,app_id if app else f'{app_id} · Saved app ID')
                item.setData(0,Qt.ItemDataRole.UserRole,app_id)
                item.setExpanded(app_id in expanded)
                (configured if app_id in included|excluded else available).addChild(item)
                item.setExpanded(app_id in expanded)
                for rule,label,subtitle,state in [('excluded','Always floating','Keep its windows outside the tiled grid',app_id in excluded),
                                                  ('included','Explicitly include (Windows)','Tile windows that Windows reports as excluded',app_id in included)]:
                    toggle=QTreeWidgetItem([label]);toggle.setToolTip(0,subtitle);toggle.setFlags(toggle.flags()|Qt.ItemFlag.ItemIsUserCheckable)
                    toggle.setCheckState(0,Qt.CheckState.Checked if state else Qt.CheckState.Unchecked)
                    toggle.setData(0,Qt.ItemDataRole.UserRole,(app_id,rule));item.addChild(toggle)
            for group in (configured,available):
                group.setText(0,group.text(0)+f' · {group.childCount()}');group.setExpanded(True);group.setHidden(not group.childCount())
            if not configured.childCount() and not available.childCount(): self.rules.addTopLevelItem(QTreeWidgetItem(['No matching applications']))

    def _rule_changed(self,item,column):
        data=item.data(0,Qt.ItemDataRole.UserRole)
        if not isinstance(data,tuple): return
        app_id,rule=data;editor=self.excluded if rule=='excluded' else self.included
        values=editor.toPlainText().splitlines();values=[x for x in values if x!=app_id]
        if item.checkState(0)==Qt.CheckState.Checked:
            values.append(app_id)
            other=self.included if rule=='excluded' else self.excluded
            with QSignalBlocker(other): other.setPlainText('\n'.join(x for x in other.toPlainText().splitlines() if x!=app_id))
        with QSignalBlocker(editor): editor.setPlainText('\n'.join(values))
        self._changed()
        QTimer.singleShot(0,self.refresh_rules)

    def export_archive(self):
        path,_=QFileDialog.getSaveFileName(self,'Export SmartGrid layouts','smartgrid-layouts.json','JSON files (*.json)')
        if path: self.bridge.submit('export_archive',path,message='Layouts and profiles exported')

    def import_archive(self):
        path,_=QFileDialog.getOpenFileName(self,'Import SmartGrid layouts','','JSON files (*.json)')
        if path: self.bridge.submit('import_archive',path)

    def closeEvent(self,event):
        if self.save_timer.isActive():
            self.save_timer.stop();self.save_settings()
        super().closeEvent(event)

    def reject(self):
        if self.save_timer.isActive():
            self.save_timer.stop();self.save_settings()
        super().reject()
