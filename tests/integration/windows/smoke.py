"""Executable native smoke test; only moves windows spawned by this harness.

Run from checkout: python tests/integration/windows/smoke.py --allow-desktop
This is a native interactive Windows test, never a Linux/simulated validation.
"""
from __future__ import annotations
import argparse
import ctypes as C
import json
from pathlib import Path
import subprocess
import sys
import threading
import time

sys.path.insert(0, str(Path(__file__).resolve().parents[3]))

def wait_for(predicate, timeout=3):
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if predicate():
            return True
        time.sleep(.05)
    return False

def fixture():
    from smartgrid.platform.windows.api import Win32, MINMAXINFO
    from smartgrid.platform.windows.messages import MessageThread
    api = Win32()
    def callback(hwnd, message, wp, lp):
        if message == 0x24:
            api.DefWindowProcW(hwnd, message, wp, lp)
            info = C.cast(lp, C.POINTER(MINMAXINFO)).contents
            info.ptMinTrackSize.x, info.ptMinTrackSize.y = 160, 120
            return 0
        return None
    owner = MessageThread(api, 'SmartGrid test fixture', callback)
    owner._class_name = 'MozillaWindowClass'
    owner.start()
    def show():
        api.SetWindowLong(owner.hwnd, -20, 0)
        api.SetWindowLong(owner.hwnd, -16, 0x00cf0000)
        api.SetWindowTextW(owner.hwnd, 'Getting started — Recall notes — Firefox fixture')
        api.SetWindowPos(owner.hwnd, None, 60, 60, 400, 300, 0x74)
        api.ShowWindowAsync(owner.hwnd, 1)
    owner.call(show)
    print(json.dumps({'hwnd':int(owner.hwnd)}), flush=True)
    try:
        sys.stdin.readline()
    finally:
        owner.stop()

def smoke():
    from smartgrid.core.models import Rect
    from smartgrid.platform.windows import WindowsBackend, HotkeyManager, InstanceGuard
    child = subprocess.Popen([sys.executable, __file__, '--fixture'], stdin=subprocess.PIPE, stdout=subprocess.PIPE, text=True)
    backend = None
    hotkeys = None
    try:
        hwnd = json.loads(child.stdout.readline())['hwnd']
        backend = WindowsBackend()
        displays = backend.discover_displays()
        assert displays and len({display.id for display in displays}) == len(displays)
        windows = backend.discover_windows()
        record = next(window for window in windows if window.ref.hwnd == hwnd)
        assert record.eligible, record.exclusion_reason
        snapshot = backend.snapshot(hwnd)
        events = []
        backend.start_events(events.append)
        area = displays[0].work_area
        target = Rect(area.x + 40, area.y + 40, min(500, area.width - 80), min(360, area.height - 80))
        result = backend.place_result(hwnd, target, timeout=2)
        assert result.success, result
        assert wait_for(lambda: any(event['hwnd'] == hwnd for event in events)), 'No native events received'
        with backend.virtual_desktops.query() as desktops:
            assert desktops is not None, 'IVirtualDesktopManager unavailable'
            assert desktops.on_current(hwnd), 'Fixture is not on the current desktop'
            assert desktops.desktop_id(hwnd), 'Missing native desktop GUID'
        assert backend.minimize(hwnd)
        assert wait_for(lambda: bool(backend.api.IsIconic(hwnd)))
        assert backend.restore(hwnd, snapshot)
        assert wait_for(lambda: backend._close_rect(backend.window_rect(hwnd), Rect(*snapshot['rect']))), 'Original geometry was not restored'
        hotkeys = HotkeyManager(lambda action: None)
        conflicts = hotkeys.register({'smoke':'Ctrl+Alt+Shift+F24'})
        assert not conflicts, conflicts
        hotkeys.unregister()
        with InstanceGuard('SmartGridSmoke') as first:
            assert first.acquired
            with InstanceGuard('SmartGridSmoke') as second:
                assert not second.acquired
        print(json.dumps({'status':'passed', 'displays':[{'id':d.id, 'dpi':d.dpi, 'work_area':d.work_area.to_dict()} for d in displays], 'events':len(events), 'limitations':['Physical hotkey keypress, mixed DPI/monitor unplug, packaged app identities, blocked/elevated third-party apps require the manual matrix.']}))
    finally:
        if hotkeys:
            hotkeys.stop()
        if backend:
            backend.stop()
        if child.stdin:
            child.stdin.write('quit\n')
            child.stdin.flush()
        try:
            child.wait(timeout=5)
        except subprocess.TimeoutExpired:
            child.terminate()
            child.wait(timeout=5)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--fixture', action='store_true')
    parser.add_argument('--allow-desktop', action='store_true')
    args = parser.parse_args()
    if sys.platform != 'win32':
        parser.exit(2, 'Native smoke requires an interactive Windows session.\n')
    if args.fixture:
        fixture()
    elif args.allow_desktop:
        smoke()
    else:
        parser.error('Pass --allow-desktop to create and move temporary test windows')
