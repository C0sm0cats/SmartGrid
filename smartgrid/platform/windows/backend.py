from __future__ import annotations
import ctypes as C
from dataclasses import dataclass, replace
import logging
import math
import os
import queue
import threading
import time
from functools import wraps
from smartgrid.core.models import Rect, Display, WindowRef, WindowRecord
from smartgrid.core.exclusions import excluded_class
from .api import Win32, RECT, POINT, DWORD, UINT, BOOL, MONITORINFOEX, DISPLAY_DEVICE, MONITORPROC, ENUMPROC, WINDOWPLACEMENT, MINMAXINFO, ULONG_PTR, WINEVENTPROC
from .apps import AppCatalogue
from .messages import MessageThread
from .virtual_desktops import VirtualDesktops,scoped_display

log = logging.getLogger(__name__)

@dataclass(frozen=True)
class PlacementResult:
    success: bool
    requested: Rect
    actual: Rect | None
    attempts: int
    reason: str = ''


def physical(method):
    @wraps(method)
    def wrapped(self, *args, **kwargs):
        with self.api.physical_coordinates():
            return method(self, *args, **kwargs)
    return wrapped


def rect_from_native(rect):
    return Rect(rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top)


def eligibility(*, visible, cloaked, owner, style, exstyle, class_name, own_process=False):
    """Structural policy: titles and legitimate application classes are not filters."""
    if own_process:
        return False, 'SmartGrid interface'
    if not visible:
        return False, 'Not visible'
    if cloaked:
        return False, 'Cloaked or on another Windows virtual desktop'
    if style & 0x40000000:
        return False, 'Child window'
    if exstyle & 0x08000000:
        return False, 'Does not activate'
    if exstyle & 0x80 and not exstyle & 0x40000:
        return False, 'Tool window'
    if owner and not exstyle & 0x40000:
        return False, 'Owned popup'
    if excluded_class(class_name):
        return False, 'Windows desktop or shell surface'
    return True, ''

def animation_rect(source, target, progress, effect='crit_damped', minimum=(1, 1)):
    """Physical-coordinate animation frame; its final frame equals the requested tile."""
    t = max(0.0, min(1.0, float(progress)))
    if t in (0.0, 1.0):
        return source if t == 0.0 else target
    if effect == 'spring':
        response = lambda p: 1 - math.exp(-5 * p) * (math.cos(10 * p) + .5 * math.sin(10 * p))
        ease = response(t) / response(1)
    elif effect == 'crit_damped':
        response = lambda p: 1 - math.exp(-10 * p) * (1 + 10 * p)
        ease = response(t) / response(1)
    elif effect in ('curved', 'curve'):
        ease = .5 * (1 - math.cos(math.pi * t))
    else:
        ease = t
    values = [round(a + (b - a) * ease) for a, b in zip((source.x, source.y, source.width, source.height), (target.x, target.y, target.width, target.height))]
    if effect in ('curved', 'curve'):
        dx, dy = target.x - source.x, target.y - source.y
        distance = math.hypot(dx, dy)
        if distance:
            arc = math.sin(math.pi * t) * min(64, distance * .12)
            values[0] += round(-dy / distance * arc)
            values[1] += round(dx / distance * arc)
    values[2], values[3] = max(minimum[0], values[2]), max(minimum[1], values[3])
    return Rect(*values)


def abgr_to_hex(value):
    """Registry accent DWORD (0xAABBGGRR) to #rrggbb."""
    return f'#{value & 0xff:02x}{(value >> 8) & 0xff:02x}{(value >> 16) & 0xff:02x}'


def accent_from_registry():
    try:
        import winreg
    except ImportError:
        return None
    for path, name in ((r'Software\Microsoft\Windows\DWM', 'AccentColor'),
                       (r'Software\Microsoft\Windows\CurrentVersion\Explorer\Accent', 'AccentColorMenu')):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                value, kind = winreg.QueryValueEx(key, name)
            if kind == winreg.REG_DWORD:
                return abgr_to_hex(value)
        except OSError:
            continue
    return None


class WindowsBackend:
    capabilities = {'physical_coordinates': True, 'win_events': True, 'virtual_desktops': 'IVirtualDesktopManager; independent groups per desktop and display', 'native_border': 'Windows 11 DWM border color', 'window_constraints': True}

    def __init__(self, api=None):
        self.api = api or Win32()
        if self.api.SetProcessDpiAwarenessContext:
            # ERROR_ACCESS_DENIED means the manifest/Qt already fixed awareness.
            self.api.SetProcessDpiAwarenessContext(C.c_void_p(-4))
        self.catalogue = AppCatalogue(self.api)
        self.last_result = None
        self._monitor_ids = {}
        self._monitor_areas = {}
        self._generations = {}
        self._discovery_records = {}
        self._generation_counter = 0
        self._events = None
        self._event_queue = queue.Queue(maxsize=4096)
        self._event_dispatcher = None
        self._event_stop = threading.Event()
        self._callback = None
        self._hooks = []
        self._event_lock = threading.Lock()
        self.virtual_desktops=VirtualDesktops()
        self.desktop_id=None
        self.desktop_known=not self.virtual_desktops.available
        self.discover_displays()

    @physical
    def discover_displays(self):
        current,self.desktop_known=self.virtual_desktops.resolve(lambda:[self.foreground()]+self._visible_top_level())
        if current: self.desktop_id=current
        displays, monitor_ids, monitor_areas = [], {}, {}
        monitor_devices = {}
        def enum(monitor, dc, rect, data):
            info = MONITORINFOEX()
            info.cbSize = C.sizeof(info)
            if not self.api.GetMonitorInfoW(monitor, C.byref(info)):
                return True
            device = DISPLAY_DEVICE()
            device.cb = C.sizeof(device)
            stable, name = info.szDevice, info.szDevice
            if self.api.EnumDisplayDevicesW(info.szDevice, 0, C.byref(device), 1):
                stable = device.DeviceID or device.DeviceKey or stable
                name = device.DeviceString or name
            x, y = UINT(96), UINT(96)
            self.api.GetDpiForMonitor(monitor, 0, C.byref(x), C.byref(y))
            identifier = 'monitor:' + stable.casefold()
            monitor_devices[identifier] = info.szDevice
            monitor_ids[monitor] = identifier
            monitor_areas[monitor] = rect_from_native(info.rcMonitor)
            displays.append(Display(identifier, name, rect_from_native(info.rcWork), int(x.value), bool(info.dwFlags & 1)))
            return True
        self.api.check(self.api.EnumDisplayMonitors(None, None, MONITORPROC(enum), 0), 'EnumDisplayMonitors')
        self._monitor_ids, self._monitor_areas = monitor_ids, monitor_areas
        self.monitor_devices = {scoped_display(key,self.desktop_id):value for key,value in monitor_devices.items()}
        displays=[replace(display,id=scoped_display(display.id,self.desktop_id)) for display in displays]
        return sorted(displays, key=lambda display: (not display.primary, display.id))

    def _visible_top_level(self, limit=64):
        """Visible, uncloaked top-level windows of other processes, to identify the desktop."""
        found = []
        def enum(hwnd, data):
            pid = DWORD()
            self.api.GetWindowThreadProcessId(hwnd, C.byref(pid))
            if pid.value != os.getpid() and self.api.IsWindowVisible(hwnd) and not self._attribute(hwnd, 14):
                found.append(hwnd)
            return len(found) < limit
        self.api.EnumWindows(ENUMPROC(enum), 0)
        return found

    def _attribute(self, hwnd, number, initial=0):
        result = DWORD(initial)
        if self.api.DwmGetWindowAttribute(hwnd, number, C.byref(result), C.sizeof(result)) == 0:
            return result.value
        return initial

    def _raw_rect(self, hwnd):
        rect = RECT()
        if not self.api.GetWindowRect(hwnd, C.byref(rect)):
            return None
        return rect_from_native(rect)

    @physical
    def visible_frame(self, hwnd):
        """The visible frame inside the window rectangle, without the invisible
        resize borders: left, top, right, bottom relative to the window."""
        raw, frame = self._raw_rect(hwnd), self.window_rect(hwnd)
        if not raw or not frame:
            return None
        return (frame.x - raw.x, frame.y - raw.y, frame.right - raw.x, frame.bottom - raw.y)

    @physical
    def window_rect(self, hwnd):
        rect = RECT()
        if not self.api.IsIconic(hwnd) and self.api.DwmGetWindowAttribute(hwnd, 9, C.byref(rect), C.sizeof(rect)) == 0:
            if rect.right > rect.left and rect.bottom > rect.top:
                return rect_from_native(rect)
        return self._raw_rect(hwnd)

    @physical
    def gesture_kind(self, hwnd):
        """'move' or 'resize' for a gesture that has just started on hwnd.

        Windows reports the start of a move or resize loop without saying which.
        The cursor tells: on a visible edge or outside the frame (the invisible
        resize border) it is a resize, anywhere else a move. A window resizing
        itself during a move (applications enforcing a minimum size) is then not
        mistaken for a resize."""
        point = POINT()
        rect = self.window_rect(hwnd)
        if not rect or not self.api.GetCursorPos(C.byref(point)):
            return None
        margin = 4
        inside = (rect.x + margin <= point.x < rect.right - margin and
                  rect.y + margin <= point.y < rect.bottom - margin)
        return 'move' if inside else 'resize'

    def _identity(self, hwnd, *, pid=None, scan_identities=None):
        if pid is None:
            native_pid=DWORD()
            self.api.GetWindowThreadProcessId(hwnd,C.byref(native_pid))
            pid=native_pid.value
        if scan_identities is not None and pid in scan_identities:
            app,creation=scan_identities[pid]
        else:
            app,creation=self.catalogue.process_identity(pid)
            if scan_identities is not None:
                scan_identities[pid]=(app,creation)
        key = (pid, creation)
        old = self._generations.get(hwnd)
        if old is None or old[:2] != key:
            self._generation_counter += 1
            self._generations[hwnd] = (*key, f'{pid}:{creation}:{self._generation_counter}')
        return WindowRef(int(hwnd), pid, self._generations[hwnd][2]), app

    @physical
    def discover_windows(self):
        windows = []
        found = set()
        scan_identities={}
        previous=getattr(self,"_discovery_records",{})
        def enum(hwnd, data):
            try:
                found.add(int(hwnd))
                pid=DWORD()
                self.api.GetWindowThreadProcessId(hwnd,C.byref(pid))
                if pid.value==os.getpid():
                    return True
                visible=bool(self.api.IsWindowVisible(hwnd))
                # Never inspect thousands of never-managed hidden helper windows.
                # Previously discovered app windows remain tracked while hidden.
                if not visible and int(hwnd) not in previous:
                    return True
                class_buf = C.create_unicode_buffer(256)
                self.api.GetClassNameW(hwnd, class_buf, len(class_buf))
                style,exstyle=self.api.GetWindowLong(hwnd,-16),self.api.GetWindowLong(hwnd,-20)
                allowed,reason=eligibility(visible=visible,cloaked=self._attribute(hwnd,14),owner=self.api.GetWindow(hwnd,4),style=style,exstyle=exstyle,class_name=class_buf.value)
                if not allowed and reason in {'Child window','Tool window','Does not activate','Windows desktop or shell surface','Owned popup'} and int(hwnd) not in previous:
                    return True
                ref,app=self._identity(hwnd,pid=pid.value,scan_identities=scan_identities)
                if class_buf.value == 'ApplicationFrameWindow':
                    def child_identity(child, data):
                        nonlocal app
                        child_class = C.create_unicode_buffer(256)
                        self.api.GetClassNameW(child, child_class, len(child_class))
                        if child_class.value == 'Windows.UI.Core.CoreWindow':
                            child_pid = DWORD()
                            self.api.GetWindowThreadProcessId(child, C.byref(child_pid))
                            candidate, _ = self.catalogue.process_identity(child_pid.value)
                            if candidate.aumid:
                                app = candidate
                                return False
                        return True
                    self.api.EnumChildWindows(hwnd, ENUMPROC(child_identity), 0)
                length = min(self.api.GetWindowTextLengthW(hwnd) + 1, 32768)
                title = C.create_unicode_buffer(max(1, length))
                self.api.GetWindowTextW(hwnd, title, len(title))
                rect = self.window_rect(hwnd)
                if not rect:
                    return True
                monitor = self.api.MonitorFromWindow(hwnd, 2)
                state = 'minimized' if self.api.IsIconic(hwnd) else ('maximized' if self.api.IsZoomed(hwnd) else 'normal')
                area = self._monitor_areas.get(monitor)
                if state == 'normal' and area and not style & 0x00c00000 and self._close_rect(rect, area):
                    state = 'fullscreen'
                    allowed, reason = False, 'Fullscreen; preserved until explicitly restored'
                desktop=self.desktop_id
                if desktop_query:
                    on_current=desktop_query.on_current(hwnd)
                    if on_current is False:
                        allowed,reason=False,'Cloaked or on another Windows virtual desktop'
                        desktop=desktop_query.desktop_id(hwnd) or desktop
                identifier=scoped_display(self._monitor_ids.get(monitor,''),desktop)
                windows.append(WindowRecord(ref, app.id, title.value, rect, identifier, state, allowed, reason))
            except Exception:
                log.exception('Failed to inspect window %s', hwnd)
            return True
        with self.virtual_desktops.query() as desktop_query:
            self.api.check(self.api.EnumWindows(ENUMPROC(enum), 0), 'EnumWindows')
        self._generations = {hwnd: value for hwnd, value in self._generations.items() if hwnd in found}
        self._discovery_records={w.ref.hwnd:w for w in windows}
        return sorted(windows, key=lambda window: (window.ref.pid, window.ref.generation, window.ref.hwnd))

    def discover_apps(self):
        self.discover_windows()
        return self.catalogue.discover()

    @physical
    def snapshot(self, hwnd):
        if not self.alive(hwnd):
            return {}
        placement = WINDOWPLACEMENT()
        placement.length = C.sizeof(placement)
        if not self.api.GetWindowPlacement(hwnd, C.byref(placement)):
            return {}
        ref, _ = self._identity(hwnd)
        rect = self.window_rect(hwnd)
        return {'hwnd': hwnd, 'pid': ref.pid, 'generation': ref.generation, 'rect': [rect.x, rect.y, rect.width, rect.height] if rect else None,
                'placement': {'flags': placement.flags, 'showCmd': placement.showCmd, 'min': [placement.ptMinPosition.x, placement.ptMinPosition.y], 'max': [placement.ptMaxPosition.x, placement.ptMaxPosition.y], 'normal': [placement.rcNormalPosition.left, placement.rcNormalPosition.top, placement.rcNormalPosition.right, placement.rcNormalPosition.bottom]},
                'style': self.api.GetWindowLong(hwnd, -16), 'exstyle': self.api.GetWindowLong(hwnd, -20), 'border_color': self._attribute(hwnd, 34, 0xffffffff)}

    def restore_centered(self, hwnd, snapshot, area):
        """Restore a window as it was before tiling (styles, normal size), but
        centred on area and not maximized. Used when a tiled window floats."""
        placement = snapshot.get('placement') or {}
        if not placement:
            return self.restore(hwnd, snapshot)
        l, t, r, b = placement['normal']
        # The visible frame (DWM bounds) when the window was normal; otherwise
        # its normal rectangle, which also includes the invisible resize border.
        x, y, w, h = snapshot['rect'] if snapshot.get('rect') and placement.get('showCmd') == 1 else (l, t, r - l, b - t)
        nw, nh = min(w, area.width * 9 // 10), min(h, area.height * 9 // 10)
        nx, ny = area.x + (area.width - nw) // 2, area.y + (area.height - nh) // 2
        # rcNormalPosition uses workspace coordinates: move and shrink it by the
        # same amounts as the visible frame, whatever the coordinate origin.
        normal = [l + nx - x, t + ny - y, r + (nx + nw) - (x + w), b + (ny + nh) - (y + h)]
        return self.restore(hwnd, {**snapshot, 'rect': [nx, ny, nw, nh],
                                   'placement': {**placement, 'normal': normal, 'showCmd': 1}})

    @physical
    def restore(self, hwnd, snapshot):
        if not self.alive(hwnd) or not snapshot:
            return False
        ref, _ = self._identity(hwnd)
        if snapshot.get('generation') != ref.generation or snapshot.get('pid') != ref.pid:
            return False
        try:
            saved = snapshot['placement']
            placement = WINDOWPLACEMENT()
            placement.length = C.sizeof(placement)
            placement.flags, placement.showCmd = saved['flags'], saved['showCmd']
            placement.ptMinPosition.x, placement.ptMinPosition.y = saved['min']
            placement.ptMaxPosition.x, placement.ptMaxPosition.y = saved['max']
            placement.rcNormalPosition = RECT(*saved['normal'])
            # No resize styles are added during management; restore only differences.
            for key, index in [('style', -16), ('exstyle', -20)]:
                if self.api.GetWindowLong(hwnd, index) != snapshot[key]:
                    C.set_last_error(0)
                    result = self.api.SetWindowLong(hwnd, index, snapshot[key])
                    if result == 0 and C.get_last_error():
                        return False
                    self.api.SetWindowPos(hwnd, None, 0, 0, 0, 0, 0x4000 | 0x27)
            color = DWORD(snapshot.get('border_color', 0xffffffff))
            self.api.DwmSetWindowAttribute(hwnd, 34, C.byref(color), C.sizeof(color))
            # WPF_ASYNCWINDOWPLACEMENT keeps a hung third-party process bounded.
            placement.flags |= 4
            if not self.api.SetWindowPlacement(hwnd, C.byref(placement)):
                return False
            deadline = time.monotonic() + .75
            while time.monotonic() < deadline and self.alive(hwnd):
                if saved['showCmd'] in (2, 6, 7, 11):
                    restored = bool(self.api.IsIconic(hwnd))
                elif saved['showCmd'] == 3:
                    restored = bool(self.api.IsZoomed(hwnd))
                else:
                    restored = not self.api.IsIconic(hwnd) and self._close_rect(self.window_rect(hwnd), Rect(*snapshot['rect']))
                if restored:
                    return True
                time.sleep(.015)
            return False
        except (KeyError, TypeError, ValueError):
            log.exception('Invalid window restoration snapshot')
            return False

    @physical
    def min_size(self, hwnd):
        info, result = MINMAXINFO(), ULONG_PTR()
        # SMTO_ABORTIFHUNG|SMTO_BLOCK, maximum 100 ms.
        if not self.api.SendMessageTimeoutW(hwnd, 0x24, 0, C.addressof(info), 3, 100, C.byref(result)):
            return 1, 1
        raw, visible = self._raw_rect(hwnd), self.window_rect(hwnd)
        dx = raw.width - visible.width if raw and visible else 0
        dy = raw.height - visible.height if raw and visible else 0
        return max(1, info.ptMinTrackSize.x - dx), max(1, info.ptMinTrackSize.y - dy)

    @staticmethod
    def _close_rect(actual, expected, tolerance=2):
        return actual is not None and all(abs(a - b) <= tolerance for a, b in zip((actual.x, actual.y, actual.width, actual.height), (expected.x, expected.y, expected.width, expected.height)))

    def _place_once(self, hwnd, rect, force=False):
        raw, visible = self._raw_rect(hwnd), self.window_rect(hwnd)
        if not raw or not visible:
            return False
        # Requested bounds refer to visible DWM frame, Win32 receives outer bounds.
        return bool(self.api.SetWindowPos(hwnd, None, rect.x + raw.x - visible.x, rect.y + raw.y - visible.y,
                                         rect.width + raw.width - visible.width, rect.height + raw.height - visible.height,
                                         self._placement_flags(hwnd, force)))

    def _placement_flags(self, hwnd, force):
        """SWP_NOZORDER|SWP_NOACTIVATE, asynchronous by default.

        Forced placement adds SWP_NOSENDCHANGING, which skips WM_WINDOWPOSCHANGING
        where Windows enforces the application's minimum track size, and
        SWP_FRAMECHANGED. It is synchronous like a direct SetWindowPos (an
        asynchronous request is replayed by the application's own thread, which
        can enforce its minimum again); a hung application stays asynchronous.
        """
        if not force:
            return 0x4000 | 0x14
        hung = getattr(self.api, 'IsHungAppWindow', None)
        asynchronous = 0x4000 if hung and hung(hwnd) else 0
        return asynchronous | 0x14 | 0x0400 | 0x0020

    def force_resizable(self, hwnd):
        """Give a fixed-size window a resizable frame (WS_THICKFRAME) so it can be
        sized into its tile. The original styles are part of the restoration
        snapshot and come back when tiling stops."""
        style = self.api.GetWindowLong(hwnd, -16)
        if style & 0x00040000 and not style & 0x01000000:  # already WS_THICKFRAME, not WS_MAXIMIZE
            return False
        self.api.SetWindowLong(hwnd, -16, (style | 0x00040000) & ~0x01000000)
        # SWP_NOMOVE|SWP_NOSIZE|SWP_NOZORDER|SWP_NOACTIVATE|SWP_FRAMECHANGED
        self.api.SetWindowPos(hwnd, None, 0, 0, 0, 0, 0x0002 | 0x0001 | 0x0004 | 0x0010 | 0x0020)
        return True

    @physical
    def place_result(self, hwnd, rect, animate=False, timeout=2.0, retries=3, cancel=None, force=False, **kwargs):
        deadline = time.monotonic() + max(0.05, min(float(timeout), 30.0))
        actual, attempts = self.window_rect(hwnd), 0
        if not self.alive(hwnd):
            return PlacementResult(False, rect, actual, 0, 'Window no longer exists')
        if rect.width <= 0 or rect.height <= 0:
            return PlacementResult(False, rect, actual, 0, 'Invalid geometry')
        if force:
            # Force the window into its tile even when it reports a larger minimum
            # size or no resizable frame; Windows applies whatever it accepts.
            self.force_resizable(hwnd)
            minimum = (1, 1)
        else:
            minimum = self.min_size(hwnd)
            if rect.width < minimum[0] or rect.height < minimum[1]:
                return PlacementResult(False, rect, actual, 0, f'Application minimum size is {minimum[0]}×{minimum[1]}')
        if self.api.IsIconic(hwnd) or self.api.IsZoomed(hwnd):
            self.show(hwnd)
        if animate and actual:
            duration = min(float(kwargs.get('duration', .14)), max(0, deadline - time.monotonic()) / 2)
            started = time.monotonic()
            fps = max(1, min(240, int(kwargs.get('fps', 60))))
            effect = kwargs.get('effect', 'crit_damped')
            while duration > 0 and time.monotonic() - started < duration:
                if cancel and cancel():
                    return PlacementResult(False, rect, self.window_rect(hwnd), attempts, 'Cancelled')
                t = min(1, (time.monotonic() - started) / duration)
                self._place_once(hwnd, animation_rect(actual, rect, t, effect, minimum), force)
                time.sleep(min(1 / fps, max(0, started + duration - time.monotonic()), max(0, deadline - time.monotonic())))
        for attempt in range(max(1, min(int(retries), 20))):
            if cancel and cancel():
                return PlacementResult(False, rect, self.window_rect(hwnd), attempts, 'Cancelled')
            if time.monotonic() >= deadline:
                break
            attempts += 1
            if not self._place_once(hwnd, rect, force):
                return PlacementResult(False, rect, self.window_rect(hwnd), attempts, f'Windows denied placement ({C.get_last_error()})')
            if force:
                # Correct a window that grows back at once, as soon as it does,
                # instead of letting it overlap its neighbours while waiting.
                time.sleep(.015)
                if not self._close_rect(self.window_rect(hwnd), rect):
                    self._force_fit(hwnd, rect)
            settle = min(deadline, time.monotonic() + .15)
            stable_since = None
            while time.monotonic() < settle:
                if cancel and cancel():
                    return PlacementResult(False, rect, self.window_rect(hwnd), attempts, 'Cancelled')
                actual = self.window_rect(hwnd)
                if self._close_rect(actual, rect):
                    if stable_since is None:
                        stable_since = time.monotonic()
                    if time.monotonic() - stable_since >= .045:
                        return PlacementResult(True, rect, actual, attempts)
                else:
                    stable_since = None
                time.sleep(min(.015, max(0, settle - time.monotonic())))
        actual = self.window_rect(hwnd)
        if force and not self._close_rect(actual, rect) and self.alive(hwnd):
            actual = self._force_fit(hwnd, rect)
        return PlacementResult(self._close_rect(actual, rect), rect, actual, attempts, '' if self._close_rect(actual, rect) else 'Application refused requested geometry or placement timed out')

    def _force_fit(self, hwnd, rect):
        """Last resort for windows that resize themselves back (GTK, Qt, Electron):
        correct by the measured difference, then MoveWindow to the tile."""
        hung = getattr(self.api, 'IsHungAppWindow', None)
        if hung and hung(hwnd):
            return self.window_rect(hwnd)
        for _ in range(2):
            raw, visible = self._raw_rect(hwnd), self.window_rect(hwnd)
            if not raw or not visible or self._close_rect(visible, rect):
                break
            # Shrink the outer rectangle by exactly what the visible frame overshoots.
            x = raw.x + rect.x - visible.x
            y = raw.y + rect.y - visible.y
            width = raw.width + rect.width - visible.width
            height = raw.height + rect.height - visible.height
            self.api.SetWindowPos(hwnd, None, x, y, max(1, width), max(1, height), 0x14 | 0x0400 | 0x0020)
            time.sleep(.015)
        visible = self.window_rect(hwnd)
        if not self._close_rect(visible, rect, 6):
            raw = self._raw_rect(hwnd)
            if raw and visible:
                self.api.MoveWindow(hwnd, raw.x + rect.x - visible.x, raw.y + rect.y - visible.y,
                                    max(1, raw.width + rect.width - visible.width),
                                    max(1, raw.height + rect.height - visible.height), True)
                time.sleep(.015)
        visible = self.window_rect(hwnd)
        if visible and not self._close_rect(visible, rect):
            log.info('Forced fit of %s ended at %sx%s for a %sx%s tile', hwnd, visible.width, visible.height, rect.width, rect.height)
        return visible

    def is_visible(self, hwnd):
        return bool(self.api.IsWindowVisible(hwnd))

    def mouse_down(self):
        """Left (primary) mouse button held, read live."""
        state = getattr(self.api, 'GetAsyncKeyState', None)
        return bool(state and state(0x01) & 0x8000)

    @physical
    def window_monitor(self, hwnd):
        """Live monitor and work area of the monitor holding the window."""
        info = MONITORINFOEX()
        info.cbSize = C.sizeof(info)
        monitor = self.api.MonitorFromWindow(hwnd, 2)
        if not monitor or not self.api.GetMonitorInfoW(monitor, C.byref(info)):
            return None
        return rect_from_native(info.rcMonitor), rect_from_native(info.rcWork)

    @physical
    def clamp(self, hwnd, rect):
        """Put a tiled window back in its tile at once, below its minimum size
        if need be; no waiting for it to settle (the slot guard)."""
        self._place_once(hwnd, rect, force=True)
        if not self._close_rect(self.window_rect(hwnd), rect):
            self._force_fit(hwnd, rect)

    def place(self, hwnd, rect, animate=False, **kwargs):
        self.last_result = self.place_result(hwnd, rect, animate, **kwargs)
        if not self.last_result.success:
            actual = self.last_result.actual
            log.warning('Placement %s: %s (requested %sx%s, got %s)', hwnd, self.last_result.reason, rect.width, rect.height,
                        f'{actual.width}x{actual.height}' if actual else 'unknown')
        return self.last_result.success

    def alive(self, hwnd):
        return bool(self.api.IsWindow(hwnd))

    def is_maximized(self, hwnd):
        return bool(self.api.IsZoomed(hwnd))

    def is_minimized(self, hwnd):
        return bool(self.api.IsIconic(hwnd))

    def minimize(self, hwnd):
        if not self.alive(hwnd):
            return False
        if not self.api.ShowWindowAsync(hwnd, 6):
            return bool(self.api.IsIconic(hwnd))
        deadline=time.monotonic()+.5
        while self.alive(hwnd) and time.monotonic()<deadline:
            if self.api.IsIconic(hwnd):
                return True
            time.sleep(.01)
        return False

    def show(self, hwnd):
        if not self.alive(hwnd):
            return False
        # SW_SHOWNOACTIVATE restores without making each tiled client steal focus.
        # Focus changes are confined to focus(), an explicit user command.
        self.api.ShowWindowAsync(hwnd, 4)
        deadline=time.monotonic()+.5
        while self.alive(hwnd) and time.monotonic()<deadline:
            if not self.api.IsIconic(hwnd) and not self.api.IsZoomed(hwnd) and self.api.IsWindowVisible(hwnd):
                return True
            time.sleep(.01)
        return False

    SHELL_SURFACES = {'Shell_TrayWnd', 'Shell_SecondaryTrayWnd', 'NotifyIconOverflowWindow',
                      'TopLevelWindowForOverflowXamlIsland', 'XamlExplorerHostIslandWindow'}

    def is_shell_surface(self, hwnd):
        """Taskbar, notification area or a SmartGrid window (its tray menu, Studio)."""
        pid = DWORD()
        self.api.GetWindowThreadProcessId(hwnd, C.byref(pid))
        if pid.value == os.getpid():
            return True
        name = C.create_unicode_buffer(256)
        self.api.GetClassNameW(hwnd, name, len(name))
        return name.value in self.SHELL_SURFACES

    def stack_above(self, overlay, target):
        """Place an overlay window directly above target in the Z order, not topmost.

        The focus outline stays with the window content, so Studio,
        Preferences or any window above the target also covers the outline.
        """
        above = self.api.GetWindow(target, 3)  # GW_HWNDPREV
        if above and int(above) == int(overlay):
            return True  # already directly above the target
        # Insert below the window that is above the target; with nothing above,
        # the target is the topmost normal window and the overlay goes to the top.
        insert = above if above else C.c_void_p(0)  # HWND_TOP
        # SWP_NOMOVE|SWP_NOSIZE|SWP_NOACTIVATE|SWP_NOOWNERZORDER
        return bool(self.api.SetWindowPos(overlay, insert, 0, 0, 0, 0, 0x0002 | 0x0001 | 0x0010 | 0x0200))

    def foreground(self):
        return int(self.api.GetForegroundWindow() or 0)

    def maximize(self, hwnd):
        return bool(self.alive(hwnd) and self.api.ShowWindowAsync(hwnd,3))

    def close(self, hwnd):
        # WM_CLOSE leaves unsaved-document confirmation to the application.
        return bool(self.alive(hwnd) and self.api.PostMessageW(hwnd,0x10,0,0))

    def system_accent_color(self):
        """Windows accent colour (Settings > Personalisation > Colours).

        DwmGetColorizationColor is the blended title-bar colorization, not the
        accent: on Windows 11 it can be pink or grey. The accent is stored as
        ABGR in the DWM and Explorer keys; colorization is only a last resort.
        """
        now=time.monotonic()
        if now-getattr(self,'_accent_time',-10)>2:
            self._accent_color=accent_from_registry()
            if not self._accent_color:
                color,opaque=DWORD(),BOOL()
                self._accent_color=f'#{color.value&0xffffff:06x}' if self.api.DwmGetColorizationColor(C.byref(color),C.byref(opaque))==0 else '#3584e4'
            self._accent_time=now
        return self._accent_color

    def focus(self, hwnd):
        if not self.show(hwnd):
            return False
        return bool(self.api.SetForegroundWindow(hwnd))

    def launch(self, app):
        target = 'shell:AppsFolder\\' + app.aumid if app.aumid else app.executable
        if not target:
            return False
        return self.api.ShellExecuteW(None, 'open', target, None, None, 1) > 32

    def border(self, hwnd, color):
        try:
            value = 0xffffffff if color is None else int(color.lstrip('#'), 16)
            if color is not None:
                value = ((value & 0xff) << 16) | (value & 0xff00) | ((value >> 16) & 0xff)
            native = DWORD(value)
            return self.api.DwmSetWindowAttribute(hwnd, 34, C.byref(native), C.sizeof(native)) == 0
        except (ValueError, AttributeError):
            return False

    def _enqueue(self, event):
        if self._event_stop.is_set():
            return
        try:
            self._event_queue.put_nowait(event)
        except queue.Full:
            # Drop noisy locations, force a reconciliation for lifecycle overflow.
            if event['type'] != 'location':
                log.warning('Native event queue overflow; reconciliation required')

    def start_events(self, callback):
        with self._event_lock:
            if self._events:
                self._callback = callback
                return
            self._callback = callback
            self._event_stop.clear()
            def window_message(hwnd, message, wparam, lparam):
                if message in (0x7e, 0x2e0, 0x1a):
                    self._enqueue({'type':'display_change', 'hwnd':0})
                return None
            self._events = MessageThread(self.api, 'SmartGrid WinEvents', window_message)
            def install():
                event_types = {0x8000:'created', 0x8001:'destroyed', 0x8002:'shown', 0x8003:'hidden', 0x800b:'location', 0x800c:'title', 3:'foreground', 0xa:'move_start', 0xb:'move_end', 0x16:'minimized', 0x17:'restored'}
                def event_proc(hook, event, hwnd, object_id, child_id, thread_id, event_time):
                    if not hwnd or (event >= 0x8000 and (object_id != 0 or child_id != 0)):
                        return
                    self._enqueue({'type':event_types.get(event, 'changed'), 'hwnd':int(hwnd), 'native_time':event_time})
                self._winevent_proc = WINEVENTPROC(event_proc)
                for event in event_types:
                    hook = self.api.SetWinEventHook(event, event, None, self._winevent_proc, 0, 0, 2)
                    self.api.check(hook, 'SetWinEventHook')
                    self._hooks.append(hook)
                return True
            def uninstall():
                for hook in self._hooks:
                    self.api.UnhookWinEvent(hook)
                self._hooks.clear()
            self._events._cleanup.append(uninstall)
            try:
                self._events.call(install)
            except Exception:
                self._events.stop()
                self._events = None
                raise
            def dispatch():
                while not self._event_stop.is_set():
                    try:
                        event = self._event_queue.get(timeout=.2)
                    except queue.Empty:
                        continue
                    try:
                        if event['type'] == 'destroyed':
                            self._generations.pop(event['hwnd'], None)
                        if self._callback:
                            self._callback(event)
                    except Exception:
                        log.exception('Backend event subscriber failed')
            self._event_dispatcher = threading.Thread(target=dispatch, name='SmartGrid event dispatch', daemon=False)
            self._event_dispatcher.start()


    def instance_guard(self, name='SmartGrid'):
        from .instance import InstanceGuard
        return InstanceGuard(name, self.api)

    def stop(self):
        with self._event_lock:
            self._event_stop.set()
            if self._events:
                self._events.stop()
                self._events = None
            dispatcher, self._event_dispatcher = self._event_dispatcher, None
        if dispatcher and dispatcher is not threading.current_thread():
            dispatcher.join(5)
            if dispatcher.is_alive():
                raise TimeoutError('Event subscriber did not finish')
