"""Qt application lifetime, tray commands and DPI-aware reserved-slot cards."""
import logging
import time
from PySide6.QtCore import Qt, QTimer, QRect
from PySide6.QtGui import QAction, QActionGroup, QColor, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import QApplication, QMenu, QSystemTrayIcon, QPushButton, QMessageBox, QWidgetAction, QLabel
from smartgrid.core.models import app_display_name, Rect
from .bridge import bridge_for
from .studio import Studio
from .quick_switcher import QuickSwitcher
from .preferences import Preferences
from .theme import apply_theme
from .focus_frame import FocusFrame
from .window_actions import WindowActions
from .guides import TargetGuide,SpaceOsd,PinPlaceholder,PlacementGhost,SourceGuide
from .tile_view import application_icon
from smartgrid.core.geometry import effective_layout,preset_name

log = logging.getLogger(__name__)


class DesktopUI:
    def __init__(self, controller, app):
        self.controller, self.app = controller, app
        self.bridge = bridge_for(controller)
        self.windows, self.cards = {}, {}
        self.focus_frame = FocusFrame()
        self.target_guide=TargetGuide();self.space_osd=SpaceOsd()
        self._last_hints=None
        self._last_space_event=controller.space_event
        self.swap_from=SourceGuide();self.swap_to=TargetGuide()
        self._seen_motion=len(controller.motion_events) and controller.motion_events[-1][4]
        self._last_swap_event=None
        self.swap_timer=QTimer(app);self.swap_timer.setSingleShot(True)
        self.swap_timer.timeout.connect(lambda:(self.swap_from.hide(),self.swap_to.hide()))
        controller.native_border=False
        # Native window movement runs at the display refresh rate.
        def refresh_rate(*args):
            controller.refresh_rate=max(30,min(240,round(max((s.refreshRate() for s in app.screens()),default=60))))
        refresh_rate();app.screenAdded.connect(refresh_rate);app.screenRemoved.connect(refresh_rate)
        self._frame_rect=None
        self.window_actions=WindowActions()
        self.action_reference=None
        self.window_actions.action.connect(self.perform_window_action)
        self.resize_frames = []
        self.quitting = False
        self.last_error = ''
        apply_theme(controller.settings)
        self._icon_color=None
        self.tray = QSystemTrayIcon(app)
        self.menu = QMenu()
        from .theme import studio_settings
        apply_theme(studio_settings(controller.settings),target=self.menu)
        # Text rows are labels, not greyed actions, so they never read as unavailable commands.
        self.status_item=self._label_row('SmartGrid  ·  PAUSED','status')
        self.toggle = self.add('Arrange windows', 'toggle');self.toggle.setCheckable(True)
        self.menu.addSeparator()
        self.display_anchor=self.menu.addSeparator()
        self.display_menus={};self._display_signature=None
        self.space_actions=[]
        self._label_row('FOCUSED WINDOW','heading')
        self.float_action=self.add('Float / tile focused window','float')
        self.swap_action=self.add('Swap focused window','swap')
        self.swap_action.setStatusTip('Swap focused window. Arrows move, Enter accepts, Escape cancels.')
        self.undo_action=self.add('Undo last arrangement','undo')
        self.redo_action=self.add('Redo last arrangement','redo')
        self.menu.addSeparator()
        self.arrange_action=self.add('Arrange again','arrange')
        self.add('Change layout…','quick_switcher')
        self.add('Layout Studio…','studio')
        self.add('Preferences','preferences')
        self.menu.addSeparator()
        self.stop_action=self.add('Stop and restore windows','stop')
        advanced=self.menu.addMenu('Windows tools')
        self.add('Open Windows Recycle Bin','recycle_bin',advanced)
        self.add('Quit and restore windows','quit',advanced)
        self.menu.aboutToShow.connect(self.prepare_menu)
        self.tray.setContextMenu(self.menu)
        self.tray.activated.connect(lambda reason: self.open('studio') if reason == QSystemTrayIcon.ActivationReason.DoubleClick else None)
        self.bridge.ui_action.connect(self.open, Qt.ConnectionType.QueuedConnection)
        self.bridge.changed.connect(self.refresh, Qt.ConnectionType.QueuedConnection)
        self.bridge.guides.connect(self.refresh_drag, Qt.ConnectionType.QueuedConnection)
        self.bridge.follow.connect(self.follow_window, Qt.ConnectionType.QueuedConnection)
        self.bridge.error.connect(self.error)
        self.tray.show()
        self.refresh()
        self.timer=QTimer(app)
        self.timer.setInterval(100)
        self.timer.timeout.connect(self.refresh_focus)
        self.timer.start()
        self.bridge.submit('start_runtime')
        # Load the Start menu catalogue once in the background: running windows
        # then show their real application names (File Explorer, not explorer).
        self.bridge.submit('refresh_apps')

    def _label_row(self, text, kind):
        label=QLabel(text);label.setProperty('menu_'+kind,True)
        row=QWidgetAction(self.menu);row.setDefaultWidget(label);row.setEnabled(False)
        self.menu.addAction(row)
        return label

    @staticmethod
    def tray_icon(color):
        pixmap = QPixmap(64,64)
        pixmap.fill(Qt.GlobalColor.transparent)
        painter = QPainter(pixmap)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setBrush(QColor(color))
        painter.setPen(Qt.PenStyle.NoPen)
        for x,y,w,h in ((4,4,32,56),(44,4,16,24),(44,36,16,24)):
            painter.drawRoundedRect(x,y,w,h,5,5)
        painter.end()
        return QIcon(pixmap)

    def add(self, text, action, menu=None):
        item = QAction(text, self.menu)
        item.triggered.connect(lambda checked=False,name=action:self.dispatch(name))
        (menu or self.menu).addAction(item)
        return item

    def dispatch(self,action):
        reference=getattr(self,'_menu_reference',None)
        window=next((w for w in self.controller.windows if reference and w.ref==reference),None)
        if action in ('studio','quick_switcher','preferences','quit'):
            return self.open(action,window.display_id if window else None)
        if action=='float' and window:
            return self.bridge.submit('window_action','float',reference)
        if action=='arrange' and window:
            return self.bridge.submit('arrange',window.display_id)
        if action=='swap' and window:
            def swap():
                self.controller.refresh(auto=False)
                live=self.controller._window(reference.hwnd)
                if live and live.ref==reference and self.controller.backend.focus(reference.hwnd):
                    self.controller.begin_swap()
            return self.bridge.submit(swap)
        self.bridge.submit('handle_action',action)

    def prepare_menu(self):
        # Opening the tray menu activates the taskbar: act on the last application
        # window that had the focus.
        hwnd=self.controller.backend.foreground()
        current=next((w.ref for w in self.controller.windows if w.ref.hwnd==hwnd and w.eligible),None)
        last=getattr(self,'_last_app_window',None)
        if current is None and last and any(w.ref==last for w in self.controller.windows):
            current=last
        self._menu_reference=current
        self.refresh()

    def _rebuild_displays(self):
        signature=tuple(d.id for d in self.controller.displays)
        if signature==self._display_signature: return
        self._display_signature=signature
        for menu,actions in self.display_menus.values():
            self.menu.removeAction(menu.menuAction());menu.deleteLater()
        self.display_menus={};self.space_actions=[]
        for display in self.controller.displays:
            menu=QMenu(self.menu);self.menu.insertMenu(self.display_anchor,menu)
            actions=[];group=QActionGroup(menu);group.setExclusive(True)
            for space in range(3):
                action=menu.addAction(f'Space {space+1}');action.setCheckable(True);group.addAction(action)
                action.triggered.connect(lambda checked=False,d=display.id,s=space:self.bridge.submit('switch_space',d,s))
                actions.append(action);self.space_actions.append(action)
            self.display_menus[display.id]=(menu,actions)

    def perform_window_action(self,action):
        reference=self.action_reference
        self.window_actions.dismiss()
        if reference: self.bridge.submit('window_action',action,reference)

    def open(self, action, display_id=None):
        if action == 'quit':
            return self.request_quit()
        if self.quitting:
            return
        widget = self.windows.get(action)
        if widget is None:
            cls = {'studio':Studio, 'quick_switcher':QuickSwitcher, 'preferences':Preferences}.get(action)
            if cls is None:
                return
            widget = cls(self.controller,display_id=display_id) if action in ('studio','quick_switcher') else cls(self.controller)
            self.windows[action] = widget
            if action == 'studio':
                widget.preferences_requested.connect(lambda: self.open('preferences'))
        if display_id and action=='studio':
            widget.display_id=display_id
            widget.space=self.controller.active_spaces.get(display_id,0)
            widget.refresh_from_controller()
        elif display_id and action=='quick_switcher':
            widget.displays.setCurrentIndex(widget.displays.findData(display_id))
            widget.spaces.setCurrentIndex(self.controller.active_spaces.get(display_id,0))
            widget.refresh_from_controller()
        elif action == 'quick_switcher':
            widget.refresh_from_controller()
        widget.show()
        widget.raise_()
        widget.activateWindow()

    def error(self, message):
        self.last_error = message
        self.tray.showMessage('SmartGrid',message,QSystemTrayIcon.MessageIcon.Warning,6000)
        if 'studio' in self.windows:
            self.windows['studio'].show_error(message)

    def refresh(self):
        controller = self.controller
        active=controller.running and not controller.paused
        self.status_item.setText('SmartGrid  ·  '+('ON' if active else 'PAUSED'))
        # The tray icon takes the accent colour while running.
        color=controller.focus_color if active else '#9baeb6'
        if color!=self._icon_color:
            self._icon_color=color;self.tray.setIcon(self.tray_icon(color))
        self.toggle.setChecked(active)
        self.tray.setToolTip('SmartGrid · '+controller.status)
        self._rebuild_displays()
        for index,display in enumerate(controller.displays):
            menu,actions=self.display_menus[display.id]
            space=controller.active_spaces.get(display.id,0)
            preset,_,_=effective_layout(controller.profile(display.id,space),runtime=True)
            menu.setTitle(f'Display {index+1}  ·  {preset_name(preset)}  ·  Space {space+1}')
            menu.menuAction().setEnabled(controller.running)
            for s,action in enumerate(actions): action.setChecked(s==space)
        foreground=getattr(self,'_menu_reference',None)
        hwnd=foreground.hwnd if foreground else controller.backend.foreground()
        window=next((w for w in controller.windows if w.ref.hwnd==hwnd),None)
        usable=bool(controller.running and window and window.eligible and window.state=='normal')
        self.float_action.setEnabled(usable)
        occupants=[a for a in controller.profile(window.display_id,controller.active_spaces.get(window.display_id,0)).assignments if a and a.window_id] if usable else []
        self.swap_action.setEnabled(usable and not window.floating and len(occupants)>1)
        self.undo_action.setEnabled(controller.running and controller.history.can_undo)
        self.redo_action.setEnabled(controller.running and controller.history.can_redo)
        self.arrange_action.setEnabled(controller.running)
        self.stop_action.setEnabled(controller.running)
        if controller.last_error and controller.last_error != self.last_error:
            self.error(controller.last_error)
        slots = {(s['display_id'],s['space'],s['index']):s for s in controller.reserved_slots()}
        # Hidden while the Studio, the switcher or a window drag is active.
        hide_cards=bool(controller._interacting) or any(w.isVisible() for k,w in self.windows.items() if k in ('studio','quick_switcher'))
        for key in self.cards.keys() - slots.keys():
            self.cards.pop(key).deleteLater()
        for key,slot in slots.items():
            card = self.cards.get(key)
            if card is None:
                card = PinPlaceholder()
                card.clicked.connect(lambda k=key: self.bridge.submit('activate_reserved',*k))
                self.cards[key] = card
            app = next((a for a in controller.apps if a.id == slot['app_id']),None)
            name = app.name if app else app_display_name(slot['app_id'])
            state = slot['state']
            unavailable = state=='CLOSED' and not app
            # Minimized and floating windows return, a closed app opens.
            actionable = state in ('MINIMIZED','FLOATING') or (state=='CLOSED' and not unavailable)
            card.describe(application_icon(app),name,f"PINNED · {'UNAVAILABLE' if unavailable else state}",actionable)
            rect = self._logical(key[0],slot['rect'])
            if rect is None or hide_cards:
                card.hide();continue
            # The card fills the reserved tile with a 7 px inset.
            card.setGeometry(rect.adjusted(7,7,-7,-7))
            if not card.isVisible():
                card.show()

    def _screen(self, display):
        """Qt screen of a native display: device name first, then physical size."""
        screens = self.app.screens()
        names = (display.id.casefold(),display.name.casefold(),getattr(self.controller.backend,'monitor_devices',{}).get(display.id,'').casefold())
        screen = next((s for s in screens if s.name().casefold() in names),None)
        if screen is None:
            # Qt names screens after the monitor model ("DELL U2720Q"), Windows
            # after the device: match the physical size of the work area instead.
            screen = next((s for s in screens if abs(s.availableGeometry().width()*s.devicePixelRatio()-display.work_area.width)<4
                           and abs(s.availableGeometry().height()*s.devicePixelRatio()-display.work_area.height)<4),None)
        if screen is None:
            screen = next((s for s in screens if abs(s.geometry().width()*s.devicePixelRatio()-display.work_area.width)<64
                           and abs(s.geometry().height()*s.devicePixelRatio()-display.work_area.height)<160),None)
        if screen is None and len(screens) == 1:
            screen = screens[0]
        if screen is None and display.primary:
            screen = self.app.primaryScreen()
        return screen

    def _logical(self, display_id, rect):
        """Native physical rectangle to Qt logical coordinates, or None if unmapped."""
        display = next((d for d in self.controller.displays if d.id == display_id),None)
        screen = self._screen(display) if display else None
        if screen is None:
            return None
        scale = screen.devicePixelRatio() or 1
        native, logical = display.work_area, screen.availableGeometry()
        return QRect(logical.x()+round((rect.x-native.x)/scale),logical.y()+round((rect.y-native.y)/scale),
                     round(rect.width/scale),round(rect.height/scale))

    def refresh_guides(self):
        controller=self.controller
        color=controller.focus_color
        # Drop target while a window is dragged.
        guide=controller.drag_guide if not getattr(controller.backend,'is_fake',False) else None
        target=self._logical(guide[0],guide[1]) if guide else None
        if target is None:
            self.target_guide.hide()
        else:
            self.target_guide.color,self.target_guide.label=color,guide[2]
            app=next((a for a in controller.apps if a.id==guide[3]),None)
            self.target_guide.content=(application_icon(app),app.name if app else app_display_name(guide[3]))
            if self.target_guide.geometry()!=target: self.target_guide.setGeometry(target)
            self.target_guide.update()
            if not self.target_guide.isVisible(): self.target_guide.show()
            self.target_guide.raise_()
        # Space OSD after a space switch.
        event=controller.space_event
        if event and event!=self._last_space_event:
            self._last_space_event=event
            display=next((d for d in controller.displays if d.id==event[0]),None)
            area=self._logical(event[0],display.work_area) if display else None
            if area is not None and not getattr(controller.backend,'is_fake',False):
                index=controller.displays.index(display)
                self.space_osd.color=color
                self.space_osd.show_space(area.center().x(),area.y(),event[1],f'Space {event[1]+1} · Display {index+1}',
                                          controller.settings.visual_duration(90))
        fake=getattr(controller.backend,'is_fake',False)
        # Placement ghosts from the previous tile to the new one.
        from PySide6.QtCore import QEasingCurve
        curve=QEasingCurve({'ease-out':QEasingCurve.Type.OutCubic,'linear':QEasingCurve.Type.Linear,
                            'ease-in-out':QEasingCurve.Type.InOutCubic,'spring':QEasingCurve.Type.OutBack}[controller.settings.animation_curve])
        for display_id,start,end,app_id,stamp in list(controller.motion_events):
            if stamp<=self._seen_motion: continue
            self._seen_motion=stamp
            a,b=self._logical(display_id,start),self._logical(display_id,end)
            duration=controller.settings.visual_duration(210)
            if fake or a is None or b is None or not duration or time.monotonic()-stamp>.5: continue
            app=next((x for x in controller.apps if x.id==app_id),None)
            PlacementGhost(application_icon(app),app.name if app else app_display_name(app_id),a,b,duration,curve).start()
        # Keyboard swap: source and target guides for 360 ms.
        event=controller.swap_event
        if event and event!=self._last_swap_event:
            self._last_swap_event=event
            a,b=self._logical(event[0],event[1]),self._logical(event[0],event[2])
            if a is not None and b is not None and not fake:
                self.swap_from.setGeometry(a);self.swap_to.setGeometry(b)
                self.swap_to.color,self.swap_to.label=color,event[3]
                self.swap_from.show();self.swap_to.show();self.swap_to.update()
                self.swap_timer.start(360)

    def _swap_arrows(self,rect):
        """Swap mode: one badge per neighbouring window, on the edge they share.

        rect is the window's live native frame; badges are returned in Qt
        logical coordinates and painted by the focus outline, which is
        reliably shown above the window.
        """
        controller=self.controller
        hints=controller.swap_hints()
        arrows=[]
        if hints:
            for direction,target,x,y,primary in hints[2]:
                if direction in ('left','right'):
                    x=rect.x if direction=='left' else rect.right
                    y=min(max(y,rect.y+13),rect.bottom-13)
                else:
                    y=rect.y if direction=='up' else rect.bottom
                    x=min(max(x,rect.x+13),rect.right-13)
                point=self._logical(hints[0],Rect(round(x),round(y),1,1))
                if point is not None: arrows.append((direction,primary,point.x(),point.y()))
        key=(hints[0],hints[1],[a[:2] for a in hints[2]]) if hints else bool(controller._swap_snapshot)
        if key!=self._last_hints:
            self._last_hints=key
            if hints: log.info('Swap arrows for tile %s: %s',hints[1],arrows)
            elif key: log.info('Swap mode without arrows: window %s is not a normal tiled window',controller._swap_hwnd)
        return arrows

    def _refresh_resize_frames(self):
        """Linked-resize preview outlines while a window edge is dragged."""
        controller=self.controller
        previews=[] if getattr(controller.backend,'is_fake',False) else list(controller.preview_rectangles)
        while len(self.resize_frames)>len(previews):
            self.resize_frames.pop().deleteLater()
        while len(self.resize_frames)<len(previews):
            self.resize_frames.append(FocusFrame(topmost=True))
        for frame,(display_id,rect) in zip(self.resize_frames,previews):
            display=next((d for d in controller.displays if d.id==display_id),None)
            device=getattr(controller.backend,'monitor_devices',{}).get(display_id,'')
            screen=self._screen(display) if display else None
            if not display or not screen:
                frame.hide()
                continue
            native,logical=display.work_area,screen.availableGeometry()
            scale=screen.devicePixelRatio() or 1
            frame.move_preview(QRect(logical.x()+round((rect.x-native.x)/scale),logical.y()+round((rect.y-native.y)/scale),round(rect.width/scale),round(rect.height/scale)),controller.settings)
            frame.color=controller.settings.accent
            frame.width_px=2
            frame.style='outline'
            frame.update()
            if not frame.isVisible():
                frame.show()

    def refresh_drag(self):
        """Drag guides only: called as soon as they change, without waiting
        for the 100 ms timer."""
        try:
            self._refresh_resize_frames()
            self.refresh_guides()
        except Exception:
            log.exception('Drag guides could not be drawn')

    def refresh_focus(self):
        try:
            self._refresh_focus()
        finally:
            # Guides are drawn after the focus outline so they stay above it:
            # swap arrows sit on the outline.
            try:
                self.refresh_guides()
            except Exception:
                # Logged once per message: a timer slot error is otherwise invisible under pythonw.
                import traceback
                text=traceback.format_exc()
                if text!=getattr(self,'_guide_error',None):
                    self._guide_error=text
                    log.error('Guides could not be drawn:\n%s',text)

    def _place_window_actions(self, rect, area):
        def position(size):
            # windowActionPosition: 4 px inside the left edge, vertically centred, clamped.
            x=min(max(rect.x()+4,area.x()+6),max(area.x()+6,area.right()-size.width()-6))
            y=min(max(rect.y()+(rect.height()-size.height())//2,area.y()+8),max(area.y()+8,area.bottom()-size.height()-8))
            return x,y
        self.window_actions.move(*position(self.window_actions.size()))
        self.window_actions.handle.move(*position(self.window_actions.handle.size()))

    def follow_window(self, hwnd):
        """Native move of a window: only the palette of that window follows
        its edge; everything else waits for the regular refresh."""
        self.bridge.follow_queued=False
        controller=self.controller
        reference=self.action_reference
        if (not reference or reference.hwnd!=hwnd or controller._interacting
                or not (self.window_actions.isVisible() or self.window_actions.handle.isVisible())):
            return
        try:
            window=next((w for w in controller.windows if w.ref.hwnd==hwnd),None)
            live=getattr(controller.backend,'window_rect',lambda h: None)(hwnd)
            display=next((d for d in controller.displays if window and d.id==window.display_id),None)
            screen=self._screen(display) if display else None
            if not live or not screen:
                return
            centre_x,centre_y=live.x+live.width/2,live.y+live.height/2
            area=display.work_area
            if not (area.x<=centre_x<area.x+area.width and area.y<=centre_y<area.y+area.height):
                # Moved to another display: the full refresh maps it to its screen.
                return self.refresh_focus()
            self._place_window_actions(self._logical(display.id,live),screen.availableGeometry())
        except Exception:
            log.exception('Window actions could not follow the window')

    def _refresh_focus(self):
        controller=self.controller
        hwnd=controller.backend.foreground()
        selected=next((w for w in controller.windows if w.ref.hwnd==hwnd),None)
        if selected and selected.eligible and not self.menu.isVisible():
            self._last_app_window=selected.ref
        display=next((d for d in controller.displays if selected and d.id==selected.display_id),None)
        device=getattr(controller.backend,'monitor_devices',{}).get(display.id,'') if display else ''
        screen=self._screen(display) if display else None
        # The palette is hidden while the Studio or the switcher is open.
        dialogs=any(w.isVisible() for k,w in self.windows.items() if k in ('studio','quick_switcher'))
        allowed=bool(controller.running and not controller.paused and selected and selected.eligible and selected.state in ('normal','maximized') and not controller._interacting and not controller._swap_snapshot and not self.app.activeModalWidget() and not dialogs)
        if allowed and display and screen:
            if self.action_reference!=selected.ref:
                self.window_actions.dismiss()
                self._handle_due=time.monotonic()+.14
            live=getattr(controller.backend,'window_rect',None)
            rect=self._logical(display.id,(live(selected.ref.hwnd) if live else None) or selected.rect)
            self._place_window_actions(rect,screen.availableGeometry())
            self.window_actions.set_state(selected.floating,selected.state=='maximized')
            self.window_actions.set_accent(controller.focus_color)
            self.window_actions.duration=controller.settings.visual_duration(140)
            self.action_reference=selected.ref
            self.window_actions.handle.setProperty('available',True)
            if (not self.window_actions.isVisible() and not self.window_actions.handle.isVisible()
                    and time.monotonic()>=getattr(self,'_handle_due',0)):
                self.window_actions.show_handle(controller.settings.visual_duration(120))
        else:
            self.window_actions.dismiss();self.action_reference=None
        self._refresh_resize_frames()
        hwnd=controller._last_border
        window=next((w for w in controller.windows if w.ref.hwnd==hwnd),None)
        if getattr(controller.backend,'is_fake',False):
            window=None
        dialogs=any(w.isVisible() for k,w in self.windows.items() if k in ('studio','quick_switcher'))
        swapping=bool(controller._swap_snapshot)
        if not window or window.state!='normal' or not controller.running or controller.paused or dialogs or (
                not swapping and (not controller.settings.border_width or not controller.settings.active_border)):
            self.focus_frame.hide()
            return
        display=next((d for d in controller.displays if d.id==window.display_id),None)
        screens=self.app.screens()
        screen=self._screen(display) if display else None
        if screen is None:
            # Do not put a frame onto an unrelated monitor when mapping is unknown.
            self.focus_frame.hide()
            return
        # Live geometry, not the last reconciled record: the outline must follow
        # the window. While it moves (arrangement, drag, animation) the outline
        # is hidden and reappears once two successive reads agree.
        live=getattr(controller.backend,'window_rect',None)
        rect=(live(window.ref.hwnd) if live else None) or window.rect
        stable=rect==self._frame_rect
        self._frame_rect=rect
        if not stable or self.bridge.pending or controller._interacting:
            self.focus_frame.hide()
            return
        logical=self._logical(display.id,rect)
        if logical is None:
            self.focus_frame.hide();return
        # Accent outline at 80 % opacity; in swap mode the same accent,
        # fully opaque, thicker and with a soft band.
        self.focus_frame.width_px=controller.settings.border_width
        self.focus_frame.style=controller.settings.border_style
        self.focus_frame.emphasized=swapping
        self.focus_frame.arrows=self._swap_arrows(rect) if swapping else []
        self.focus_frame.color=controller.focus_color
        stack=getattr(controller.backend,'stack_above',None)
        restack=(lambda:stack(int(self.focus_frame.winId()),window.ref.hwnd)) if stack else None
        self.focus_frame.show_for(window.ref,logical,controller.settings.visual_duration(140),restack)

    def request_quit(self):
        """Restore the windows, then always quit: a window that could not be put
        back exactly is reported, but never keeps SmartGrid running."""
        if self.quitting:
            return
        self.quitting = True
        def shutdown():
            restored = True
            try:
                restored = self.controller.stop()
            except Exception:
                log.exception('Restoring windows before quitting failed')
                restored = False
            try:
                self.controller.quit()
            except Exception:
                log.exception('Shutdown did not complete cleanly')
            return restored
        def done(restored):
            if restored is False:
                self.tray.showMessage('SmartGrid', 'Some windows could not be restored exactly.',
                                      QSystemTrayIcon.MessageIcon.Warning, 3000)
            self.app.quit()
        future = self.bridge.submit(shutdown, on_success=done)
        # Never stay running if the shutdown hangs or fails.
        QTimer.singleShot(8000, self.app.quit)
        if future is None:
            self.app.quit()

    def cleanup(self):
        self.tray.hide()
        # Do not wait for a command stuck on a hung application.
        self.bridge.shutdown(wait=False)
        try:
            self.controller.quit()
        except Exception:
            log.exception('Shutdown did not complete cleanly')
        self.timer.stop()
        self.focus_frame.close()
        for widget in [self.target_guide,self.space_osd,self.swap_from,self.swap_to]:
            widget.close()
        self.window_actions.close()
        for frame in self.resize_frames:
            frame.close()
        self.tray.hide()
        for card in self.cards.values():
            card.close()


def run_ui(controller, argv=None):
    app = QApplication.instance() or QApplication(argv or [])
    app.setApplicationName('SmartGrid')
    app.setOrganizationName('SmartGrid')
    app.setQuitOnLastWindowClosed(False)
    ui = DesktopUI(controller,app)
    app.aboutToQuit.connect(ui.cleanup)
    return app.exec()
