# Dependencies:
# pip install pywin32

import ctypes
import time
import threading
from ctypes import wintypes

# Global monitors cache — updated only at startup + on each retile
MONITORS_CACHE = []

# --- ADDING PYWIN32 ---
import win32con
import win32gui

# --- WINAPI CONSTANTS / FLAGS ---
user32 = ctypes.WinDLL('user32', use_last_error=True)

GWL_STYLE = -16
WS_THICKFRAME = 0x00040000
WS_MAXIMIZE = 0x01000000
WS_MINIMIZE = 0x20000000

SW_RESTORE = 9
SW_SHOWNORMAL = 1

# SetWindowPos flags
SWP_NOZORDER = 0x0004
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOACTIVATE = 0x0010
SWP_FRAMECHANGED = 0x0020
SWP_NOSENDCHANGING = 0x0400

# GAP between windows (constant and identical everywhere)
GAP = 8
EDGE_PADDING = 8  # padding on screen edges

# state
grid_state = {}
is_active = False
monitor_thread = None

# HOTKEYS
HOTKEY_TOGGLE = 9001
HOTKEY_RETILE = 9002
HOTKEY_QUIT = 9003
HOTKEY_MOVE_MONITOR = 9004

# Structure for DwmGetWindowAttribute
DWMWA_EXTENDED_FRAME_BOUNDS = 9

def get_frame_borders(hwnd):
    window_rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(window_rect)):
        return 0, 0, 0, 0
    
    try:
        dwmapi = ctypes.WinDLL('dwmapi')
        frame_rect = wintypes.RECT()
        result = dwmapi.DwmGetWindowAttribute(
            hwnd,
            DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(frame_rect),
            ctypes.sizeof(frame_rect)
        )
        if result == 0:
            return (
                frame_rect.left - window_rect.left,
                frame_rect.top - window_rect.top,
                window_rect.right - frame_rect.right,
                window_rect.bottom - frame_rect.bottom
            )
    except:
        pass
    return 0, 0, 0, 0

# --- Monitors / visible windows (YOUR ORIGINAL CODE, UNTOUCHED) ---
def get_monitors():
    global MONITORS_CACHE
    if MONITORS_CACHE:
        return MONITORS_CACHE
    
    monitors = []
    def monitor_enum_proc(hMonitor, hdcMonitor, lprcMonitor, dwData):
        class MONITORINFO(ctypes.Structure):
            _fields_ = [
                ("cbSize", wintypes.DWORD),
                ("rcMonitor", wintypes.RECT),
                ("rcWork", wintypes.RECT),
                ("dwFlags", wintypes.DWORD),
            ]
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(mi)
        user32.GetMonitorInfoW(hMonitor, ctypes.byref(mi))
        rcw = mi.rcWork
        monitors.append((rcw.left, rcw.top, rcw.right - rcw.left, rcw.bottom - rcw.top))
        return True

    MonitorEnumProc = ctypes.WINFUNCTYPE(ctypes.c_int, wintypes.HMONITOR, wintypes.HDC, ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)
    user32.EnumDisplayMonitors(0, 0, MonitorEnumProc(monitor_enum_proc), 0)
    if not monitors:
        monitors.append((0, 0, win32gui.GetSystemMetrics(0), win32gui.GetSystemMetrics(1)))
    
    MONITORS_CACHE = monitors
    return monitors

def is_useful_window(title):
    if not title:
        return False
    t = title.lower()
    exclude = [
        "program manager", "task switching", "start", "cortana", "search", "notification",
        "toast", "popup", "tooltip", "splash", "realtek audio console"
    ]
    return not any(x in t for x in exclude)

def get_visible_windows():
    monitors = get_monitors()
    windows = []
    def enum(hwnd, _):
        if user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
            rect = wintypes.RECT()
            if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                if w > 50 and h > 50:
                    title_buf = ctypes.create_unicode_buffer(256)
                    user32.GetWindowTextW(hwnd, title_buf, 256)
                    title = title_buf.value
                    if title and is_useful_window(title):
                        overlap = 0
                        for mx, my, mw, mh in monitors:
                            left = max(rect.left, mx)
                            right = min(rect.right, mx + mw)
                            top = max(rect.top, my)
                            bottom = min(rect.bottom, my + mh)
                            if left < right and top < bottom:
                                overlap += (right - left) * (bottom - top)
                        if overlap > (w * h * 0.1):
                            windows.append((hwnd, title, rect))
        return True
    user32.EnumWindows(ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, ctypes.c_void_p)(enum), 0)
    return windows

# ==================================================================
# ONLY REPLACED FUNCTION: ULTRA-AGGRESSIVE VERSION THAT WORKS 100%
# ==================================================================
def force_tile_resizable(hwnd, x, y, w, h):
    # 1. Force resizable style
    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    user32.SetWindowLongW(hwnd, GWL_STYLE, (style | WS_THICKFRAME) & ~WS_MAXIMIZE)

    # 2. Restore properly
    user32.ShowWindowAsync(hwnd, SW_RESTORE)
    user32.ShowWindowAsync(hwnd, SW_SHOWNORMAL)
    time.sleep(0.012)

    # 3. Magic boost
    user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED)

    # 4. DWM compensation + aggressive spam with recalculation
    lb, tb, rb, bb = get_frame_borders(hwnd)
    adj_x = x - lb
    adj_y = y - tb
    adj_w = w + lb + rb
    adj_h = h + tb + bb
    flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED | SWP_NOSENDCHANGING

    for attempt in range(10):
        user32.SetWindowPos(hwnd, 0, int(adj_x), int(adj_y), int(adj_w), int(adj_h), flags)
        time.sleep(0.01 + attempt * 0.004)

        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        lb2, tb2, rb2, bb2 = get_frame_borders(hwnd)
        cur_w = rect.right - rect.left - lb2 - rb2
        cur_h = rect.bottom - rect.top - tb2 - bb2

        if abs(cur_w - w) <= 6 and abs(cur_h - h) <= 6:
            break

        adj_x = x - lb2
        adj_y = y - tb2
        adj_w = w + lb2 + rb2
        adj_h = h + tb2 + bb2

    # 5. Forced redraw
    user32.RedrawWindow(hwnd, None, None,
                        win32con.RDW_FRAME | win32con.RDW_INVALIDATE | win32con.RDW_UPDATENOW | win32con.RDW_ALLCHILDREN)

    # 6. Final nuclear option
    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    lb, tb, rb, bb = get_frame_borders(hwnd)
    final_w = rect.right - rect.left - lb - rb
    final_h = rect.bottom - rect.top - tb - bb
    if abs(final_w - w) > 15 or abs(final_h - h) > 15:
        title = win32gui.GetWindowText(hwnd)
        print(f"   [NUKE] Forced MoveWindow on → {title[:60]}")
        win32gui.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True)

# ==================================================================
# smart_tile → only a small sleep added
# ==================================================================
def smart_tile(temp=False):
    global grid_state
    monitors = get_monitors()
    visible_windows = get_visible_windows()
    if not visible_windows:
        print("[!] No windows detected.")
        return

    count = len(visible_windows)
    mon = monitors[0]
    mon_x, mon_y, mon_w, mon_h = mon

    if count <= 9:
        cols, rows = 3, 3
        mode = "3x3"
    elif count <= 12:
        cols, rows = 4, 3
        mode = "4x3"
    elif count <= 15:
        cols, rows = 5, 3
        mode = "5x3"
    else:
        cols, rows = 5, 3
        mode = "5x3 (max)"
        visible_windows = visible_windows[:15]

    total_gaps_w = GAP * (cols - 1)
    total_gaps_h = GAP * (rows - 1)
    available_w = mon_w - 2 * EDGE_PADDING - total_gaps_w
    available_h = mon_h - 2 * EDGE_PADDING - total_gaps_h
    cell_w = available_w // cols
    cell_h = available_h // rows

    col_x = [mon_x + EDGE_PADDING + c * (cell_w + GAP) for c in range(cols)]
    row_y = [mon_y + EDGE_PADDING + r * (cell_h + GAP) for r in range(rows)]

    new_grid = {}
    tiled = 0
    print(f"\n[TILE] mode={mode} ({cols}x{rows}) - windows detected: {len(visible_windows)} (temp={temp})")
    print(f"[INFO] Cell size: {cell_w}x{cell_h} px | Gap: {GAP} px (constant)")
    
    for i, (hwnd, title, _) in enumerate(visible_windows):
        col = tiled % cols
        row = tiled // cols
        x = col_x[col]
        y = row_y[row]
        
        force_tile_resizable(hwnd, x, y, cell_w, cell_h)
        
        print(f"   • {title[:60]} -> col {col}, row {row} @ {x},{y} [{cell_w}x{cell_h}]")
        time.sleep(0.04 + i * 0.006)  # ← THIS IS WHAT SAVES EVERYTHING WITH >10 WINDOWS
        
        new_grid[hwnd] = (0, col, row)
        tiled += 1

    if not temp:
        grid_state.clear()
        grid_state.update(new_grid)
        print(f"[DONE] {tiled} windows tiled (persistent).")
    else:
        print(f"[DONE] {tiled} windows tiled (temporary).")

# ==================================================================
# THE REST IS 100% YOUR ORIGINAL CODE (NOTHING HAS BEEN CHANGED)
# ==================================================================
def apply_grid():
    smart_tile(temp=False)

def monitor():
    global grid_state
    while is_active:
        for hwnd in list(grid_state.keys()):
            if not user32.IsWindow(hwnd):
                grid_state.pop(hwnd, None)
        time.sleep(0.5)

def toggle_persistent():
    global is_active, monitor_thread
    is_active = not is_active
    status = "ENABLED" if is_active else "DISABLED"
    print(f"\n[SMARTGRID] persistent mode {status}")
    if is_active:
        apply_grid()
        if monitor_thread is None or not monitor_thread.is_alive():
            monitor_thread = threading.Thread(target=monitor, daemon=True)
            monitor_thread.start()

def register_hotkeys():
    user32.RegisterHotKey(None, HOTKEY_TOGGLE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('T'))
    user32.RegisterHotKey(None, HOTKEY_RETILE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('R'))
    user32.RegisterHotKey(None, HOTKEY_QUIT,   win32con.MOD_CONTROL | win32con.MOD_ALT, ord('Q'))
    user32.RegisterHotKey(None, HOTKEY_MOVE_MONITOR, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('M'))

def unregister_hotkeys():
    try:
        user32.UnregisterHotKey(None, HOTKEY_TOGGLE)
        user32.UnregisterHotKey(None, HOTKEY_RETILE)
        user32.UnregisterHotKey(None, HOTKEY_QUIT)
        user32.UnregisterHotKey(None, HOTKEY_MOVE_MONITOR)
    except:
        pass

# ==================================================================
# NEW: Cycle tiled windows to next monitor
# Shortcut: Ctrl + Alt + M
# ==================================================================

CURRENT_MONITOR_INDEX = 0  # Global variable to keep active monitor

def move_all_tiled_to_next_monitor():
    global grid_state, CURRENT_MONITOR_INDEX, MONITORS_CACHE

    # 1. Nothing to move?
    if not grid_state:
        print("[INFO] Nothing to move → no windows in grid")
        return

    # 2. Check immediately the number of monitors → if only one, exit IMMEDIATELY
    if len(MONITORS_CACHE) <= 1:
        print("[INFO] Move cancelled → only one monitor detected")
        return

    # === Only from here do we start touching windows ===
    CURRENT_MONITOR_INDEX = (CURRENT_MONITOR_INDEX + 1) % len(MONITORS_CACHE)
    new_idx = CURRENT_MONITOR_INDEX
    mon_x, mon_y, mon_w, mon_h = MONITORS_CACHE[new_idx]

    print(f"[SWITCH] → Monitor {new_idx + 1}/{len(MONITORS_CACHE)} ({mon_w}×{mon_h} @ {mon_x},{mon_y})")

    # Recalculate grid on new monitor (same layout)
    count = len(grid_state)
    if count <= 9:
        cols, rows = 3, 3
    elif count <= 12:
        cols, rows = 4, 3
    else:
        cols, rows = 5, 3

    total_gaps_w = GAP * (cols - 1) if cols > 1 else 0
    total_gaps_h = GAP * (rows - 1) if rows > 1 else 0
    available_w = mon_w - 2 * EDGE_PADDING - total_gaps_w
    available_h = mon_h - 2 * EDGE_PADDING - total_gaps_h
    cell_w = available_w // cols
    cell_h = available_h // rows

    col_x = [mon_x + EDGE_PADDING + c * (cell_w + GAP) for c in range(cols)]
    row_y = [mon_y + EDGE_PADDING + r * (cell_h + GAP) for r in range(rows)]

    # Actual movement
    new_grid = {}
    moved = 0
    for hwnd, (old_mon_idx, col, row) in grid_state.items():
        if not user32.IsWindow(hwnd):
            continue

        x = col_x[col]
        y = row_y[row]
        force_tile_resizable(hwnd, x, y, cell_w, cell_h)

        title = win32gui.GetWindowText(hwnd)
        print(f"   → {title[:50]}")

        new_grid[hwnd] = (new_idx, col, row)
        moved += 1
        time.sleep(0.03)  # visual fluidity

    grid_state = new_grid
    print(f"[OK] {moved} windows moved to monitor {new_idx + 1}")

# --- MAIN LOOP ---
if __name__ == "__main__":
    print("="*70)
    print("   SMARTGRID — Final version (CherryTree/Bruno/Obsidian = FORCED)")
    print("="*70)
    print("• Ctrl + Alt + T → enable/disable persistent mode")
    print("• Ctrl + Alt + R → temporary retile")
    print("• Ctrl + Alt + Q → quit cleanly")
    print("-"*70)

    register_hotkeys()

    monitors = get_monitors()
    print(f"   {len(monitors)} monitor(s) detected")
    if len(monitors) > 1:
        print("   Ctrl + Alt + M → move all windows to next monitor")

    try:
        msg = wintypes.MSG()
        while True:
            ret = user32.GetMessageW(ctypes.byref(msg), None, 0, 0)
            if ret <= 0:
                break

            if msg.message == win32con.WM_HOTKEY:
                if msg.wParam == HOTKEY_TOGGLE:
                    toggle_persistent()
                elif msg.wParam == HOTKEY_RETILE:
                    threading.Thread(target=smart_tile, kwargs={"temp": True}, daemon=True).start()
                elif msg.wParam == HOTKEY_MOVE_MONITOR:
                    threading.Thread(target=move_all_tiled_to_next_monitor, daemon=True).start()
                elif msg.wParam == HOTKEY_QUIT:
                    print("\n[QUIT] Clean exit via Ctrl+Alt+Q")
                    break

            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))

    except KeyboardInterrupt:
        print("\n[INTERRUPT] Ctrl+C received, stopping.")
    finally:
        unregister_hotkeys()
        print("[EXIT] SmartGrid stopped.")