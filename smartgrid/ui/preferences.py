"""Preferences window with debounced, immediate persistence."""
from copy import deepcopy
from dataclasses import asdict
from PySide6.QtCore import Qt,QTimer,QSignalBlocker,QSize
from PySide6.QtGui import QColor,QKeySequence,QFont,QFontMetrics
from PySide6.QtWidgets import (QDialog,QWidget,QVBoxLayout,QHBoxLayout,QFormLayout,QLabel,QPushButton,
    QTabWidget,QScrollArea,QComboBox,QSpinBox,QDoubleSpinBox,QCheckBox,QColorDialog,QKeySequenceEdit,
    QPlainTextEdit,QGroupBox,QFileDialog,QLineEdit,QTreeWidget,QTreeWidgetItem,QListWidget,QListWidgetItem,QApplication)
from smartgrid.core.models import Settings,app_display_name
from smartgrid.core.exclusions import TITLE_WORDS, excluded_app, normalize_word, overlay_words
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


class RowTitle(QLabel):
    """A row title with a dimmed subtitle beneath it.

    Its minimum width follows the current font (the theme is applied after
    construction): the text stays on one line up to cap pixels, then wraps.
    """
    def __init__(self,title,subtitle='',cap=300):
        super().__init__();self.setWordWrap(True);self.cap=cap;self.set_text(title,subtitle)

    def set_text(self,title,subtitle=''):
        self.title,self.subtitle=title,subtitle
        self.setTextFormat(Qt.TextFormat.RichText if subtitle else Qt.TextFormat.PlainText)
        self.setText(f'{title}<br><span style="font-size:small;color:palette(placeholder-text)">{subtitle}</span>' if subtitle else title)
        self.updateGeometry()

    def minimumSizeHint(self):
        hint=super().minimumSizeHint()
        if not self.cap: return hint
        small=QFont(self.font());small.setPointSizeF(max(6,small.pointSizeF()*.83))
        widest=max(self.fontMetrics().horizontalAdvance(self.title),QFontMetrics(small).horizontalAdvance(self.subtitle))
        width=min(widest+8,self.cap)
        return hint.expandedTo(QSize(width,self.heightForWidth(width)))


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
        # Page tabs centred above the content.
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
            row.addStretch();row.addWidget(reset);row.addWidget(self._stepper(spin))
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
        self._combo(self.animation_group,'animation_curve','Animation curve',[('ease-out','Ease out'),('linear','Linear'),('ease-in-out','Ease in and out'),('spring','Spring')])
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
        overlays=self._group(apps_page,'Common overlays','A list built into SmartGrid of windows that usually float.')
        self._check(overlays,'builtin_exclusions','Keep common overlays out of the grid',
                    'Media players, game launchers, streaming and monitoring tools, call windows and other window managers float')
        self.controls['builtin_exclusions'].toggled.connect(lambda _:QTimer.singleShot(0,self.refresh_rules))
        # The built-in list itself is hidden until asked for.
        self.customize_overlays=QPushButton('Customize the list  ▸');self.customize_overlays.setCheckable(True)
        self.customize_overlays.setFlat(True);self.customize_overlays.setProperty('disclosure',True)
        self.customize_overlays.setAccessibleName('Customize the common overlays list')
        overlays.addRow(self.customize_overlays)
        self._keyword_editor(overlays)
        def disclose(on):
            self.keyword_box.setVisible(on);self.customize_overlays.setText('Customize the list  ▾' if on else 'Customize the list  ▸')
        self.customize_overlays.toggled.connect(disclose);disclose(False)
        rules=self._group(apps_page,'Your applications','Every application is tiled unless you tick Always floating.')
        self.app_search=QLineEdit();self.app_search.setPlaceholderText('Search applications')
        self.app_search.setAccessibleName('Search applications');self.app_search.textChanged.connect(self.refresh_rules)
        rules.addRow(self.app_search)
        self.rules=QTreeWidget();self.rules.setHeaderHidden(True);self.rules.setColumnCount(1)
        self.rules.setMinimumHeight(300);self.rules.itemChanged.connect(self._rule_changed)
        rules.addRow(self.rules)
        note=QLabel('Windows that cannot shrink to their tile are handled by Force windows into their tiles, in Windows tools and advanced settings.')
        note.setWordWrap(True);note.setProperty('muted',True);rules.addRow(note)
        advanced=self._group(main,'Windows tools and advanced settings','Placement options for Windows.')
        # Collapsed by default: the checkbox expands the group.
        box=advanced.parentWidget();box.setCheckable(True);box.setChecked(False)
        box.setFlat(True)
        self._heading(advanced,'Window size')
        self._check(advanced,'force_resize','Force windows into their tiles',
                    'Resize windows below their minimum size and give fixed-size windows a resizable frame')
        self._heading(advanced,'Responsiveness and reliability')
        self._spin(advanced,'debounce_ms','Reconciliation delay',0,10000)
        self._spin(advanced,'tile_retries','Placement retries',0,20)
        timeout=QDoubleSpinBox();timeout.setRange(.01,120);timeout.setValue(self.settings.tile_timeout)
        self._row(advanced,'tile_timeout','Placement time budget',timeout);timeout.valueChanged.connect(self._changed)
        self.included=QPlainTextEdit('\n'.join(self.settings.included_apps))
        self.excluded=QPlainTextEdit('\n'.join(self.settings.excluded_apps))
        # The rule lists are edited on the Applications page; these hold their values.
        for title,edit in [('Apps tiled despite the common overlays list',self.included),('Always floating app IDs',self.excluded)]:
            edit.setParent(box);edit.hide();edit.textChanged.connect(self._changed)
        def expand(on,form=advanced):
            for i in range(form.rowCount()): form.setRowVisible(i,on)
            box.setMaximumHeight(16777215 if on else 26)
            box.setProperty('collapsed',not on);box.style().unpolish(box);box.style().polish(box)
            self._refresh_visibility() if not self._building else None
        box.toggled.connect(expand);expand(False)
        self.message=QLabel();self.message.setWordWrap(True);self.message.setProperty('error',True)
        layout.addWidget(self.message)
        # No footer: the title bar closes the window.
        self.save_timer=QTimer(self);self.save_timer.setSingleShot(True);self.save_timer.setInterval(150)
        self.save_timer.timeout.connect(self.save_settings)
        self.bridge.error.connect(self._failed,Qt.ConnectionType.QueuedConnection)
        self.bridge.changed.connect(self.refresh_from_controller,Qt.ConnectionType.QueuedConnection)
        self.tabs.currentChanged.connect(self._page_changed)
        apply_theme(self.settings,target=self)
        self._fit_to_content()
        self._building=False
        self._refresh_visibility();self.refresh_preview();self.refresh_conflicts()

    def _fit_to_content(self):
        """Open wide enough for every page: no row, list or button is clipped
        (pages scroll vertically only)."""
        self.ensurePolished()
        for i in range(self.tabs.count()): self.tabs.widget(i).widget().ensurePolished()
        content=max(self.tabs.widget(i).widget().minimumSizeHint().width() for i in range(self.tabs.count()))
        scroll=self.style().pixelMetric(self.style().PixelMetric.PM_ScrollBarExtent)
        margins=self.layout().contentsMargins()
        width=content+scroll+margins.left()+margins.right()+2*self.tabs.style().pixelMetric(self.style().PixelMetric.PM_DefaultFrameWidth)+8
        screen=(self.screen() or QApplication.primaryScreen()).availableGeometry()
        width=min(width,screen.width()-40)
        self.setMinimumWidth(max(440,width))
        self.resize(max(660,width),min(720,screen.height()-60))

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
        form=QFormLayout(group);form.setSpacing(10);form.setVerticalSpacing(12)
        form.setFieldGrowthPolicy(QFormLayout.FieldGrowthPolicy.AllNonFixedFieldsGrow)
        if description:
            label=QLabel(description);label.setWordWrap(True);label.setProperty('muted',True);form.addRow(label)
        page.addWidget(group)
        return form

    @staticmethod
    def _heading(form,title):
        """Subsection title inside a group."""
        label=QLabel(title.upper());label.setProperty('pref_heading',True);form.addRow(label)

    @staticmethod
    def _title(title,subtitle='',cap=300):
        return RowTitle(title,subtitle,cap)

    def _row(self,form,key,title,control,subtitle=''):
        self.controls[key]=control
        control.setAccessibleName(title)
        if subtitle: control.setAccessibleDescription(subtitle)
        container=QWidget();row=QHBoxLayout(container);row.setContentsMargins(0,0,0,0)
        reset=QPushButton('↶');reset.setFixedWidth(38);reset.setToolTip(f'Reset {title} to default')
        reset.clicked.connect(lambda checked=False,k=key:self.reset_setting(k))
        self.resets[key]=reset
        # Rows keep their control compact at the right edge.
        widget=control
        if isinstance(control,(QSpinBox,QDoubleSpinBox)): widget=self._stepper(control)
        elif isinstance(control,(QComboBox,QPushButton)) and control is not getattr(self,'color_button',None): control.setFixedWidth(180)
        # Every row is tall enough for its control on Windows (Segoe UI metrics).
        container.setMinimumHeight(36);row.setContentsMargins(0,1,0,1)
        row.addStretch();row.addWidget(reset);row.addWidget(widget);form.addRow(self._title(title,subtitle),container)
        control._preference_row=(form,container)

    @staticmethod
    def _stepper(spin):
        """The value followed by large − and + buttons."""
        spin.setButtonSymbols(QSpinBox.ButtonSymbols.NoButtons)
        spin.setAlignment(Qt.AlignmentFlag.AlignCenter);spin.setFixedWidth(72)
        box=QWidget();layout=QHBoxLayout(box);layout.setContentsMargins(0,1,0,1);layout.setSpacing(4)
        box.setMinimumHeight(36);spin.setFixedHeight(32)
        layout.addWidget(spin)
        for text,step,tip in (('−',-1,'Decrease'),('+',1,'Increase')):
            button=QPushButton(text);button.setFixedSize(34,32);button.setToolTip(tip);button.setAccessibleName(tip)
            button.setAutoRepeat(True);button.setAutoRepeatDelay(400);button.setAutoRepeatInterval(60)
            button.clicked.connect(lambda checked=False,s=step:spin.stepBy(s))
            layout.addWidget(button)
        def refresh(*args):
            buttons=box.findChildren(QPushButton)
            buttons[0].setEnabled(spin.value()>spin.minimum());buttons[1].setEnabled(spin.value()<spin.maximum())
        spin.valueChanged.connect(refresh);refresh()
        return box

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

    def _keyword_editor(self,form):
        """The two parts of the overlay list, side by side: keywords and the
        applications they catch. Each can be restored to its default.

        Only changes are saved (overlay_words_added / _removed, included_apps),
        so words added to the built-in list by a later version still apply.
        """
        self.words_added=list(self.settings.overlay_words_added)
        self.words_removed=list(self.settings.overlay_words_removed)
        self.apps_added=list(self.settings.overlay_apps_added)
        box=QWidget();columns=QHBoxLayout(box);columns.setContentsMargins(0,0,0,0);columns.setSpacing(12)
        left=QVBoxLayout();right=QVBoxLayout()
        self.words=QListWidget();self.words.setMinimumHeight(220);self.words.setAccessibleName('Overlay keywords')
        self.words.itemChanged.connect(self._word_changed)
        add=QHBoxLayout();self.word_input=QLineEdit();self.word_input.setPlaceholderText('Add a word or phrase')
        self.word_input.setAccessibleName('New overlay keyword');self.word_input.returnPressed.connect(self._add_word)
        add_button=QPushButton('Add');add_button.clicked.connect(self._add_word)
        add.addWidget(self.word_input,1);add.addWidget(add_button)
        self.words_reset=QPushButton('↶ Restore default keywords');self.words_reset.clicked.connect(self._reset_words)
        left.addWidget(self._title('Overlay keywords','Whole words in a window title or application name. Untick a word to stop using it.',0))
        left.addWidget(self.words);left.addLayout(add);left.addWidget(self.words_reset)
        self.overlay_apps=QListWidget();self.overlay_apps.setMinimumHeight(220);self.overlay_apps.setAccessibleName('Overlay applications');self.overlay_apps.setWordWrap(True)
        self.overlay_apps.itemChanged.connect(self._overlay_app_changed)
        add_app=QHBoxLayout();self.app_input=QComboBox();self.app_input.setEditable(True)
        self.app_input.setSizeAdjustPolicy(QComboBox.SizeAdjustPolicy.AdjustToMinimumContentsLengthWithIcon);self.app_input.setMinimumContentsLength(8)
        self.app_input.lineEdit().setPlaceholderText('Add an application');self.app_input.setAccessibleName('New overlay application')
        self.app_input.completer().setFilterMode(Qt.MatchFlag.MatchContains)
        self.app_input.completer().setCompletionMode(self.app_input.completer().CompletionMode.PopupCompletion)
        self.app_input.lineEdit().returnPressed.connect(self._add_overlay_app)
        add_app_button=QPushButton('Add');add_app_button.clicked.connect(self._add_overlay_app)
        add_app.addWidget(self.app_input,1);add_app.addWidget(add_app_button)
        self.apps_reset=QPushButton('↶ Restore default applications');self.apps_reset.clicked.connect(self._reset_overlay_apps)
        right.addWidget(self._title('Overlay applications','Applications caught by a keyword, and the ones you add. Ticked: floating. Untick to tile it.',0))
        right.addWidget(self.overlay_apps,1);right.addLayout(add_app);right.addWidget(self.apps_reset)
        columns.addLayout(left,1);columns.addLayout(right,1)
        form.addRow(box);self.keyword_box=box
        self._fill_words()

    def _fill_words(self):
        removed={normalize_word(w) for w in self.words_removed}
        with QSignalBlocker(self.words):
            self.words.clear()
            for word in sorted(set(TITLE_WORDS)|{normalize_word(w) for w in self.words_added}):
                own=word not in TITLE_WORDS
                item=QListWidgetItem(word+(' (added)' if own else ''));item.setData(Qt.ItemDataRole.UserRole,word)
                item.setFlags(item.flags()|Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Unchecked if word in removed else Qt.CheckState.Checked)
                item.setToolTip('Your word: untick to delete it' if own else 'Built-in word')
                self.words.addItem(item)
        self.words_reset.setEnabled(bool(self.words_added or self.words_removed))

    def _word_changed(self,item):
        word=item.data(Qt.ItemDataRole.UserRole);on=item.checkState()==Qt.CheckState.Checked
        if word in TITLE_WORDS:
            self.words_removed=[w for w in self.words_removed if normalize_word(w)!=word]+([] if on else [word])
        elif not on:
            self.words_added=[w for w in self.words_added if normalize_word(w)!=word]
        QTimer.singleShot(0,self._fill_words);self._words_changed()

    def _add_word(self):
        word=normalize_word(self.word_input.text());self.word_input.clear()
        if not word or len(word)>100: return
        if word in TITLE_WORDS:
            self.words_removed=[w for w in self.words_removed if normalize_word(w)!=word]
        elif word not in {normalize_word(w) for w in self.words_added}:
            self.words_added.append(word)
        self._fill_words();self._words_changed()

    def _reset_words(self):
        self.words_added,self.words_removed=[],[]
        self._fill_words();self._words_changed()

    def _overlay_app_changed(self,item):
        app_id=item.data(Qt.ItemDataRole.UserRole);on=item.checkState()==Qt.CheckState.Checked
        if app_id in self.apps_added:
            # An added application is removed by unticking it.
            if not on:
                self.apps_added=[a for a in self.apps_added if a!=app_id]
                self._rules_signature=None;QTimer.singleShot(0,self.refresh_rules);self._changed()
            return
        self._set_floating(app_id,True,on)

    def _add_overlay_app(self):
        index=self.app_input.findText(self.app_input.currentText().strip(),Qt.MatchFlag.MatchFixedString)
        app_id=self.app_input.itemData(index) if index>=0 else None
        self.app_input.setEditText('')
        if not app_id or app_id in self.apps_added: return
        self.apps_added.append(app_id)
        # It now floats through the overlay list, not through its own rule.
        for name in ('included','excluded'):
            editor=getattr(self,name)
            with QSignalBlocker(editor): editor.setPlainText('\n'.join(x for x in editor.toPlainText().splitlines() if x and x!=app_id))
        self._rules_signature=None;self.refresh_rules();self._changed()

    def _reset_overlay_apps(self):
        """Every overlay application floats again; added applications are removed."""
        self.apps_added=[]
        values=[x for x in self.included.toPlainText().splitlines() if x and x not in getattr(self,'overlay_ids',set())]
        with QSignalBlocker(self.included): self.included.setPlainText('\n'.join(values))
        self._rules_signature=None;QTimer.singleShot(0,self.refresh_rules);self._changed()

    def _words_changed(self):
        self._rules_signature=None;QTimer.singleShot(0,self.refresh_rules);self._changed()

    def _paint_color_button(self,color):
        # A colour swatch, not a hex label.
        self.color_button.setText('');self.color_button.setFixedSize(56,32)
        self.color_button.setToolTip(color)
        self.color_button.setStyleSheet(f'QPushButton {{ background:{color}; border:2px solid palette(mid); border-radius:6px; padding:0; min-height:0; }}'
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
        settings.overlay_words_added=list(self.words_added);settings.overlay_words_removed=list(self.words_removed)
        settings.overlay_apps_added=list(self.apps_added)
        mapping=deepcopy(self.settings.hotkeys)
        mapping.update({action:edit.keySequence().toString(QKeySequence.SequenceFormat.PortableText) for action,edit in self.hotkeys.items()})
        settings.hotkeys=mapping
        return Settings.from_dict(asdict(settings))

    def _changed(self,*args):
        if self._building: return
        self._refresh_visibility();self.refresh_preview();self.save_timer.start()

    def _refresh_visibility(self):
        independent=self.controls['independent_padding'].isChecked()
        self.controls['padding'].parentWidget().setEnabled(not independent)
        form=self.controls['padding']._preference_row[0]
        for row in self.edge_rows: form.setRowVisible(row,independent)
        self.focus_group.parentWidget().setEnabled(self.controls['active_border'].isChecked())
        self.keyword_box.setEnabled(self.controls['builtin_exclusions'].isChecked())
        self.customize_overlays.setEnabled(self.controls['builtin_exclusions'].isChecked())
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
        self.preview_title.set_text('Current space preview',text)

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
        builtin=self.controls['builtin_exclusions'].isChecked()
        words=overlay_words(self.words_added,self.words_removed)
        signature=(query,builtin,words,tuple(self.apps_added),tuple(sorted(included)),tuple(sorted(excluded)),tuple((i,known[i].name,known[i].icon_path) for i in sorted(known)),tuple(sorted(ids)))
        if signature==self._rules_signature: return
        self._rules_signature=signature
        expanded={self.rules.topLevelItem(i).child(j).data(0,Qt.ItemDataRole.UserRole)
                  for i in range(self.rules.topLevelItemCount())
                  for j in range(self.rules.topLevelItem(i).childCount())
                  if self.rules.topLevelItem(i).child(j).isExpanded()}
        overlays=[]
        with QSignalBlocker(self.rules):
            self.rules.clear();configured=QTreeWidgetItem(['Configured applications']);available=QTreeWidgetItem(['Other applications'])
            configured.setToolTip(0,'Applications you keep floating.')
            self.rules.addTopLevelItems([configured,available])
            for app_id in sorted(ids,key=lambda i:known[i].name.casefold() if i in known else app_display_name(i).casefold()):
                app=known.get(app_id);name=app.name if app else app_display_name(app_id)
                # Common overlays are listed under Overlay applications instead.
                if builtin and (excluded_app(name,words) or app_id in self.apps_added):
                    overlays.append((name,app,app_id));continue
                floating=app_id in excluded
                if query and query not in (name+' '+app_id).casefold(): continue
                item=QTreeWidgetItem([name]);item.setIcon(0,application_icon(app));item.setToolTip(0,app_id if app else f'{app_id} · Saved app ID')
                item.setData(0,Qt.ItemDataRole.UserRole,app_id)
                item.setExpanded(app_id in expanded)
                (configured if floating else available).addChild(item)
                item.setExpanded(app_id in expanded)
                toggle=QTreeWidgetItem(['Always floating']);toggle.setToolTip(0,'Keep its windows outside the tiled grid');toggle.setFlags(toggle.flags()|Qt.ItemFlag.ItemIsUserCheckable)
                toggle.setCheckState(0,Qt.CheckState.Checked if floating else Qt.CheckState.Unchecked)
                toggle.setData(0,Qt.ItemDataRole.UserRole,(app_id,False));item.addChild(toggle)
            for group in (configured,available):
                group.setText(0,group.text(0)+f' · {group.childCount()}');group.setExpanded(True);group.setHidden(not group.childCount())
            if not configured.childCount() and not available.childCount(): self.rules.addTopLevelItem(QTreeWidgetItem(['No matching applications']))
        with QSignalBlocker(self.overlay_apps):
            self.overlay_apps.clear()
            for name,app,app_id in overlays:
                keyword=excluded_app(name,words)
                item=QListWidgetItem(application_icon(app),name if keyword else f'{name} (added)');item.setData(Qt.ItemDataRole.UserRole,app_id)
                item.setToolTip(f'{app_id} · keyword: {keyword}' if keyword else f'{app_id} · added by you: untick to remove it')
                item.setFlags(item.flags()|Qt.ItemFlag.ItemIsUserCheckable)
                item.setCheckState(Qt.CheckState.Unchecked if keyword and app_id in included else Qt.CheckState.Checked)
                self.overlay_apps.addItem(item)
            if not overlays:
                empty=QListWidgetItem('No installed or open application matches a keyword');empty.setFlags(Qt.ItemFlag.NoItemFlags)
                self.overlay_apps.addItem(empty)
        self.overlay_ids={app_id for _,_,app_id in overlays}
        self.apps_reset.setEnabled(bool(self.overlay_ids&included or self.apps_added))
        # Applications that can still be added: everything not already listed.
        with QSignalBlocker(self.app_input):
            text=self.app_input.currentText();self.app_input.clear()
            for app_id in sorted(ids-self.overlay_ids,key=lambda i:(known[i].name if i in known else app_display_name(i)).casefold()):
                app=known.get(app_id);self.app_input.addItem(application_icon(app),app.name if app else app_display_name(app_id),app_id)
            self.app_input.setCurrentIndex(-1);self.app_input.setEditText(text)

    def _rule_changed(self,item,column):
        data=item.data(0,Qt.ItemDataRole.UserRole)
        if not isinstance(data,tuple): return
        self._set_floating(data[0],data[1],item.checkState(0)==Qt.CheckState.Checked)

    def _set_floating(self,app_id,default,floating):
        # A common overlay floats by default: unticking it records an exception
        # (included), ticking it again simply removes the exception.
        lists={'excluded':floating and not default,'included':not floating and default}
        for name,keep in lists.items():
            editor=getattr(self,name)
            values=[x for x in editor.toPlainText().splitlines() if x and x!=app_id]+([app_id] if keep else [])
            with QSignalBlocker(editor): editor.setPlainText('\n'.join(values))
        self._rules_signature=None;QTimer.singleShot(0,self.refresh_rules)
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
