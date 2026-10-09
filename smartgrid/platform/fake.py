"""Deterministic desktop simulator: constraints, lifecycle events and failures are real inputs."""
from __future__ import annotations
from copy import deepcopy
from dataclasses import replace
from smartgrid.core.models import Rect, Display, ApplicationRef, WindowRef, WindowRecord
from .windows.backend import PlacementResult

class FakeBackend:
    is_fake = True
    capabilities = {'physical_coordinates': True, 'win_events': True, 'virtual_desktops': False, 'native_border': True, 'window_constraints': True}

    def __init__(self, displays=None, windows=None, apps=None):
        self.displays = list(displays) if displays is not None else [Display('fake-primary', 'Test display', Rect(0, 0, 1920, 1080), 96, True)]
        self.windows = {window.ref.hwnd: deepcopy(window) for window in (windows or [])}
        # Explicit catalogue entries behave like Start menu applications.
        self.apps = {app.id: replace(deepcopy(app), installed=True) for app in (apps or [])}
        for window in self.windows.values():
            self.apps.setdefault(window.app_id, ApplicationRef(window.app_id, window.app_id))
        self.callback = None
        self.minimums = {}
        self.failures = set()
        self.borders = {}
        self.styles = {}
        self.operations = []
        self.launches = []
        self.auto_launch = True
        self.stopped = False
        self.foreground_hwnd = next(iter(self.windows), 0)
        self.last_result = None
        self._generation = 0

    def discover_displays(self):
        return deepcopy(self.displays)

    def discover_windows(self):
        return deepcopy(list(self.windows.values()))

    def discover_apps(self):
        return deepcopy(list(self.apps.values()))

    def snapshot(self, hwnd):
        window = self.windows.get(hwnd)
        if not window:
            return {}
        return {'hwnd':hwnd, 'pid':window.ref.pid, 'generation':window.ref.generation, 'rect':[window.rect.x, window.rect.y, window.rect.width, window.rect.height], 'state':window.state, 'display_id':window.display_id, 'style':self.styles.get(hwnd, 0), 'border_color':self.borders.get(hwnd)}

    def restore_centered(self, hwnd, snapshot, area):
        _, _, w, h = snapshot['rect']
        w, h = min(w, area.width * 9 // 10), min(h, area.height * 9 // 10)
        rect = [area.x + (area.width - w) // 2, area.y + (area.height - h) // 2, w, h]
        return self.restore(hwnd, {**snapshot, 'rect': rect, 'state': 'normal'})

    def restore(self, hwnd, snapshot):
        window = self.windows.get(hwnd)
        if not window or hwnd in self.failures or window.ref.generation != snapshot.get('generation'):
            return False
        self.windows[hwnd] = replace(window, rect=Rect(*snapshot['rect']), state=snapshot['state'], display_id=snapshot['display_id'])
        self.styles[hwnd] = snapshot.get('style', 0)
        self.borders[hwnd] = snapshot.get('border_color')
        self.operations.append(('restore', hwnd))
        self.emit('restored', hwnd)
        return True

    def min_size(self, hwnd):
        return self.minimums.get(hwnd, (1, 1))

    def window_rect(self, hwnd):
        return self.windows[hwnd].rect if hwnd in self.windows else None

    def place_result(self, hwnd, rect, animate=False, timeout=2.0, retries=3, cancel=None, force=False, **kwargs):
        window = self.windows.get(hwnd)
        reason = ''
        if cancel and cancel():
            reason = 'Cancelled'
        elif not window:
            reason = 'Window no longer exists'
        elif hwnd in self.failures:
            reason = 'Simulated placement failure'
        elif not force and (rect.width < self.min_size(hwnd)[0] or rect.height < self.min_size(hwnd)[1]):
            reason = 'Application minimum size'
        if reason:
            return PlacementResult(False, rect, window.rect if window else None, 0, reason)
        display = next((display for display in self.displays if display.work_area.contains(rect.x + rect.width // 2, rect.y + rect.height // 2)), None)
        self.windows[hwnd] = replace(window, rect=rect, state='normal', display_id=display.id if display else window.display_id)
        self.operations.append(('place', hwnd, rect))
        self.emit('location', hwnd)
        return PlacementResult(True, rect, rect, 1)

    mouse_pressed = False

    def mouse_down(self):
        return self.mouse_pressed

    def window_monitor(self, hwnd):
        window = self.windows.get(hwnd)
        display = next((d for d in self.displays if window and d.id == window.display_id), None)
        return (display.work_area, display.work_area) if display else None

    def clamp(self, hwnd, rect):
        self.place(hwnd, rect, force=True)

    def place(self, hwnd, rect, animate=False, **kwargs):
        self.last_result = self.place_result(hwnd, rect, animate, **kwargs)
        return self.last_result.success

    def alive(self, hwnd):
        return hwnd in self.windows

    def is_maximized(self, hwnd):
        return hwnd in self.windows and self.windows[hwnd].state == 'maximized'

    def is_minimized(self, hwnd):
        return hwnd in self.windows and self.windows[hwnd].state == 'minimized'

    def minimize(self, hwnd):
        if not self.alive(hwnd) or hwnd in self.failures:
            return False
        self.windows[hwnd] = replace(self.windows[hwnd], state='minimized')
        self.operations.append(('minimize', hwnd))
        self.emit('minimized', hwnd)
        return True

    def show(self, hwnd):
        if not self.alive(hwnd) or hwnd in self.failures:
            return False
        self.windows[hwnd] = replace(self.windows[hwnd], state='normal')
        self.operations.append(('show', hwnd))
        self.emit('restored', hwnd)
        return True

    def foreground(self):
        return self.foreground_hwnd

    def maximize(self, hwnd):
        if not self.alive(hwnd) or hwnd in self.failures:
            return False
        self.windows[hwnd]=replace(self.windows[hwnd],state='maximized')
        self.operations.append(('maximize',hwnd))
        return True

    def close(self, hwnd):
        if not self.alive(hwnd) or hwnd in self.failures:
            return False
        self.operations.append(('close',hwnd))
        self.remove_window(hwnd)
        return True

    def focus(self, hwnd):
        if not self.show(hwnd):
            return False
        self.foreground_hwnd = hwnd
        self.operations.append(('focus', hwnd))
        self.emit('foreground', hwnd)
        return True

    def launch(self, app):
        if app.id not in self.apps or app.id in self.failures:
            return False
        self.launches.append(app.id)
        if self.auto_launch:
            self.add_window(app.id)
        return True

    def add_window(self, app_id, hwnd=None, rect=None, title=None, state='normal'):
        hwnd = hwnd or max(self.windows, default=100) + 1
        self._generation += 1
        app = self.apps.setdefault(app_id, ApplicationRef(app_id, app_id))
        self.windows[hwnd] = WindowRecord(WindowRef(hwnd, hwnd, f'fake:{self._generation}'), app_id, title or app.name, rect or Rect(40, 40, 600, 400), self.displays[0].id if self.displays else '', state)
        self.emit('created', hwnd)
        return self.windows[hwnd]

    def remove_window(self, hwnd):
        if hwnd in self.windows:
            del self.windows[hwnd]
            self.emit('destroyed', hwnd)

    def set_displays(self, displays):
        self.displays = deepcopy(list(displays))
        self.emit('display_change', 0)

    def border(self, hwnd, color):
        if not self.alive(hwnd):
            return False
        self.borders[hwnd] = color
        return True

    def start_events(self, callback):
        self.callback = callback
        self.stopped = False

    def emit(self, event_type, hwnd=0, **kwargs):
        if self.callback:
            self.callback({'type':event_type, 'hwnd':hwnd, **kwargs})

    def stop(self):
        self.callback = None
        self.stopped = True
