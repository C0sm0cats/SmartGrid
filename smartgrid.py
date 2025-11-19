# pip install pywin32

import ctypes
import time
import threading
from ctypes import wintypes
import win32con
import win32gui

user32 = ctypes.WinDLL('user32', use_last_error=True)
dwmapi = ctypes.WinDLL('dwmapi')

DWMWA_BORDER_COLOR = 34
DWMWA_COLOR_NONE   = 0xFFFFFFFF

current_hwnd = None

def remove_border(hwnd):
    if hwnd:
        try:
            dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_BORDER_COLOR,
                ctypes.byref(ctypes.c_uint(DWMWA_COLOR_NONE)),
                ctypes.sizeof(ctypes.c_uint)
            )
        except:
            pass

def apply_border(hwnd):
    global current_hwnd
    if current_hwnd and current_hwnd != hwnd:
        remove_border(current_hwnd)
    if hwnd:
        color = ctypes.c_uint(0x0000FF00)  # pure green
        dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_BORDER_COLOR,
            ctypes.byref(color), ctypes.sizeof(ctypes.c_uint)
        )
        current_hwnd = hwnd

# ==================================================================
# Constants & config
# ==================================================================
MONITORS_CACHE = []

GWL_STYLE = -16
WS_THICKFRAME = 0x00040000
WS_MAXIMIZE = 0x01000000

SW_RESTORE = 9
SW_SHOWNORMAL = 1
SWP_NOZORDER = 0x0004
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOACTIVATE = 0x0010
SWP_FRAMECHANGED = 0x0020
SWP_NOSENDCHANGING = 0x0400

GAP = 8
EDGE_PADDING = 8

grid_state = {}
last_visible_count = 0
is_active = False
monitor_thread = None

HOTKEY_TOGGLE = 9001
HOTKEY_RETILE = 9002
HOTKEY_QUIT = 9003
HOTKEY_MOVE_MONITOR = 9004

DWMWA_EXTENDED_FRAME_BOUNDS = 9

def get_frame_borders(hwnd):
    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return 0, 0, 0, 0
    try:
        ext = wintypes.RECT()
        if dwmapi.DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS,
                                        ctypes.byref(ext), ctypes.sizeof(ext)) == 0:
            return (ext.left - rect.left, ext.top - rect.top,
                    rect.right - ext.right, rect.bottom - ext.bottom)
    except:
        pass
    return 0, 0, 0, 0

def get_monitors():
    global MONITORS_CACHE
    if MONITORS_CACHE:
        return MONITORS_CACHE
    monitors = []
    def enum_proc(hMonitor, hdc, lprc, data):
        class MONITORINFO(ctypes.Structure):
            _fields_ = [("cbSize", wintypes.DWORD), ("rcMonitor", wintypes.RECT),
                        ("rcWork", wintypes.RECT), ("dwFlags", wintypes.DWORD)]
        mi = MONITORINFO()
        mi.cbSize = ctypes.sizeof(mi)
        user32.GetMonitorInfoW(hMonitor, ctypes.byref(mi))
        r = mi.rcWork
        monitors.append((r.left, r.top, r.right - r.left, r.bottom - r.top))
        return True
    user32.EnumDisplayMonitors(None, None,
        ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HMONITOR, wintypes.HDC,
                           ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)(enum_proc), 0)
    if not monitors:
        monitors = [(0, 0, win32gui.GetSystemMetrics(0), win32gui.GetSystemMetrics(1))]
    MONITORS_CACHE = monitors
    return monitors

def is_useful_window(title, class_name=""):
    if not title:
        return False

    title_lower = title.lower()
    class_lower = class_name.lower() if class_name else ""

    # === HARD EXCLUDE BY TITLE ===
    bad_titles = [
        "spotify", "discord", "steam", "call", "meeting", "join", "incoming call",
        "obs", "streamlabs", "twitch studio", "nvidia overlay", "geforce experience",
        "shadowplay", "radeon software", "amd relive", "rainmeter", "wallpaper engine",
        "lively wallpaper", "msi afterburner", "rtss", "rivatuner", "hwinfo", "hwmonitor",
        "displayfusion", "actual window", "aquasnap", "powertoys", "fancyzones",
        "picture in picture", "pip", "miniplayer", "mini player", "youtube music",
        "vlc media player", "media player classic", "battle.net", "origin", "epic games",
        "gog galaxy", "uplay", "ubisoft connect", "ea app", "game bar", "xbox",
        "notification", "toast", "popup", "tooltip", "splash", "alert", "flyout", "volume control", "brightness",
        "program manager", "start", "cortana", "search", "realtek audio console",
        "operationstatuswindow", "shell_secondarytraywnd"
    ]

    if any(bad in title_lower for bad in bad_titles):
        return False

    # === HARD EXCLUDE BY CLASS NAME (even if title is empty or sneaky) ===
    bad_classes = [
        "chrome_renderwidgethosthwnd",   # Chrome PIP
        "mozillawindowclass",            # Firefox PIP
        "operationstatuswindow",         # Win11 flyouts
        "windows.ui.core.corewindow",    # UWP popups
        "foregroundstaging",             # Teams call window
        "workerw",                       # Desktop wallpaper tricks
        "progman",                       # Program Manager
        "shell_traywnd",                 # Taskbar
        "realtimedisplay",               # Some overlays
    ]

    if class_lower in bad_classes:
        return False

    # === SIZE FILTER (tiny windows = overlays) ===
    # We'll do this in get_visible_windows() instead — cleaner

    return True

def get_visible_windows():
    monitors = get_monitors()
    windows = []
    def enum(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            rect = wintypes.RECT()
            if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                if w > 180 and h > 180:  # tighter than 50x50
                    title_buf = ctypes.create_unicode_buffer(256)
                    user32.GetWindowTextW(hwnd, title_buf, 256)
                    title = title_buf.value or ""
                    class_name = win32gui.GetClassName(hwnd)

                    if is_useful_window(title, class_name):
                        overlap = sum(
                            max(0, min(rect.right, mx + mw) - max(rect.left, mx)) *
                            max(0, min(rect.bottom, my + mh) - max(rect.top, my))
                            for mx, my, mw, mh in monitors
                        )
                        if overlap > (w * h * 0.15):  # slightly stricter
                            windows.append((hwnd, title, rect))
        return True

    user32.EnumWindows(
        ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(enum), 0
    )
    return windows

def force_tile_resizable(hwnd, x, y, w, h):
    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    user32.SetWindowLongW(hwnd, GWL_STYLE, (style | WS_THICKFRAME) & ~WS_MAXIMIZE)
    user32.ShowWindowAsync(hwnd, SW_RESTORE)
    time.sleep(0.012)
    user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED)

    lb, tb, rb, bb = get_frame_borders(hwnd)
    ax, ay = x - lb, y - tb
    aw, ah = w + lb + rb, h + tb + bb
    flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED | SWP_NOSENDCHANGING

    for _ in range(10):
        user32.SetWindowPos(hwnd, 0, int(ax), int(ay), int(aw), int(ah), flags)
        time.sleep(0.014)
        lb2, tb2, rb2, bb2 = get_frame_borders(hwnd)
        rect = wintypes.RECT()
        user32.GetWindowRect(hwnd, ctypes.byref(rect))
        cur_w = rect.right - rect.left - lb2 - rb2
        cur_h = rect.bottom - rect.top - tb2 - bb2
        if abs(cur_w - w) <= 6 and abs(cur_h - h) <= 6:
            break
        ax, ay = x - lb2, y - tb2
        aw, ah = w + lb2 + rb2, h + tb2 + bb2

    user32.RedrawWindow(hwnd, None, None,
                        win32con.RDW_FRAME | win32con.RDW_INVALIDATE | win32con.RDW_UPDATENOW | win32con.RDW_ALLCHILDREN)

    rect = wintypes.RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    lb, tb, rb, bb = get_frame_borders(hwnd)
    final_w = rect.right - rect.left - lb - rb
    final_h = rect.bottom - rect.top - tb - bb
    if abs(final_w - w) > 15 or abs(final_h - h) > 15:
        title = win32gui.GetWindowText(hwnd)
        print(f"   [NUKE] MoveWindow forced on -> {title[:60]}")
        win32gui.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True)

# ==================================================================
# Smart layout chooser
# ==================================================================
def choose_layout(count):
    if count == 1: return "full", None
    if count == 2: return "side_by_side", None
    if count == 3: return "master_stack", None
    if count == 4: return "grid", (2, 2)
    if count <= 6: return "grid", (3, 2)
    if count <= 9: return "grid", (3, 3)
    if count <= 12: return "grid", (4, 3)
    return "grid", (5, 3)

def clear_all_borders():
    def enum_callback(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            try:
                none = ctypes.c_uint(DWMWA_COLOR_NONE)
                dwmapi.DwmSetWindowAttribute(hwnd, DWMWA_BORDER_COLOR,
                                             ctypes.byref(none), ctypes.sizeof(ctypes.c_uint))
            except:
                pass
        return True
    
    enum_proc = ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(enum_callback)
    user32.EnumWindows(enum_proc, 0)

# ==================================================================
# smart_tile with intelligent layouts + your grid fallback
# ==================================================================
def smart_tile(temp=False):
    global grid_state
    
    if temp:
        clear_all_borders()
        time.sleep(0.05)
    
    monitors = get_monitors()
    visible_windows = get_visible_windows()
    if not visible_windows:
        print("[TILE] No windows detected.")
        return

    count = len(visible_windows)
    mon_x, mon_y, mon_w, mon_h = monitors[0]
    layout, info = choose_layout(count)

    positions = []

    if layout == "full":
        x = mon_x + EDGE_PADDING
        y = mon_y + EDGE_PADDING
        w = mon_w - 2 * EDGE_PADDING
        h = mon_h - 2 * EDGE_PADDING
        positions = [(x, y, w, h)]

    elif layout == "side_by_side":
        cw = (mon_w - 2*EDGE_PADDING - GAP) // 2
        ch = mon_h - 2*EDGE_PADDING
        positions = [
            (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, cw, ch),
            (mon_x + EDGE_PADDING + cw + GAP, mon_y + EDGE_PADDING, cw, ch)
        ]

    elif layout == "master_stack":
        mw = (mon_w - 2*EDGE_PADDING - GAP) * 3 // 5
        sw = mon_w - 2*EDGE_PADDING - mw - GAP
        sh = (mon_h - 2*EDGE_PADDING - GAP) // 2
        positions = [
            (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, mw, mon_h - 2*EDGE_PADDING),
            (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING, sw, sh),
            (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING + sh + GAP, sw, sh)
        ]

    else:
        cols, rows = info
        total_gaps_w = GAP * (cols - 1) if cols > 1 else 0
        total_gaps_h = GAP * (rows - 1) if rows > 1 else 0
        cell_w = (mon_w - 2*EDGE_PADDING - total_gaps_w) // cols
        cell_h = (mon_h - 2*EDGE_PADDING - total_gaps_h) // rows
        for r in range(rows):
            for c in range(cols):
                if len(positions) >= count:
                    break
                x = mon_x + EDGE_PADDING + c * (cell_w + GAP)
                y = mon_y + EDGE_PADDING + r * (cell_h + GAP)
                positions.append((x, y, cell_w, cell_h))

    print(f"\n[TILE] {count} windows -> {layout} layout (temp={temp})")
    new_grid = {}
    for i, (hwnd, title, _) in enumerate(visible_windows[:len(positions)]):
        x, y, w, h = positions[i]
        force_tile_resizable(hwnd, x, y, w, h)
        print(f"   -> {title[:60]}")
        time.sleep(0.04 + i * 0.006)

        if layout == "grid":
            cols = info[0]
            new_grid[hwnd] = (0, i % cols, i // cols)
        else:
            new_grid[hwnd] = (0, i, 0)   # ← for side-by-side / master-stack

    grid_state = new_grid

    time.sleep(0.15)
    active = user32.GetForegroundWindow()
    if active and user32.IsWindowVisible(active):
        apply_border(active)

# ==================================================================
# Multi-monitor
# ==================================================================
CURRENT_MONITOR_INDEX = 0

def move_all_tiled_to_next_monitor():
    global grid_state, CURRENT_MONITOR_INDEX, MONITORS_CACHE
    if not grid_state:
        print("[SWITCH] Nothing to move - no windows in grid")
        return
    if len(MONITORS_CACHE) <= 1:
        print("[SWITCH] Only one monitor detected")
        return

    clear_all_borders()
    time.sleep(0.05)

    CURRENT_MONITOR_INDEX = (CURRENT_MONITOR_INDEX + 1) % len(MONITORS_CACHE)
    mon = MONITORS_CACHE[CURRENT_MONITOR_INDEX]
    mon_x, mon_y, mon_w, mon_h = mon
    count = len(grid_state)
    layout, info = choose_layout(count)

    positions = []
    if layout == "full":
        positions = [(mon_x + EDGE_PADDING, mon_y + EDGE_PADDING,
                      mon_w - 2*EDGE_PADDING, mon_h - 2*EDGE_PADDING)]
    elif layout == "side_by_side":
        cw = (mon_w - 2*EDGE_PADDING - GAP) // 2
        positions = [(mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, cw, mon_h - 2*EDGE_PADDING),
                     (mon_x + EDGE_PADDING + cw + GAP, mon_y + EDGE_PADDING, cw, mon_h - 2*EDGE_PADDING)]
    elif layout == "master_stack":
        mw = (mon_w - 2*EDGE_PADDING - GAP) * 3 // 5
        sw = mon_w - 2*EDGE_PADDING - mw - GAP
        sh = (mon_h - 2*EDGE_PADDING - GAP) // 2
        positions = [(mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, mw, mon_h - 2*EDGE_PADDING),
                     (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING, sw, sh),
                     (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING + sh + GAP, sw, sh)]
    else:
        cols, rows = info
        total_gaps_w = GAP * (cols - 1) if cols > 1 else 0
        total_gaps_h = GAP * (rows - 1) if rows > 1 else 0
        cell_w = (mon_w - 2*EDGE_PADDING - total_gaps_w) // cols
        cell_h = (mon_h - 2*EDGE_PADDING - total_gaps_h) // rows
        for r in range(rows):
            for c in range(cols):
                if len(positions) >= count: break
                positions.append((mon_x + EDGE_PADDING + c*(cell_w+GAP),
                                  mon_y + EDGE_PADDING + r*(cell_h+GAP), cell_w, cell_h))

    new_grid = {}
    for i, (hwnd, (old_idx, col, row)) in enumerate(grid_state.items()):
        if not user32.IsWindow(hwnd):
            continue
        x, y, w, h = positions[i]
        force_tile_resizable(hwnd, x, y, w, h)
        title = win32gui.GetWindowText(hwnd)
        print(f"   -> {title[:50]}")
        new_grid[hwnd] = (CURRENT_MONITOR_INDEX, col, row)
        time.sleep(0.03)

    grid_state = new_grid
    time.sleep(0.15)
    active = user32.GetForegroundWindow()
    if active and user32.IsWindowVisible(active):
        apply_border(active)
    print(f"[SWITCH] {len(new_grid)} windows moved to monitor {CURRENT_MONITOR_INDEX + 1}")

# ==============================================================================
# Polling monitor - Auto-retile when windows are shown/hidden/minimized/restored
# ==============================================================================
def monitor():
    global current_hwnd, grid_state, last_visible_count

    while True:
        if is_active:
            # Clean dead windows
            for hwnd in list(grid_state.keys()):
                if not user32.IsWindow(hwnd):
                    grid_state.pop(hwnd, None)

            # Get currently visible windows
            visible_windows = get_visible_windows()
            current_count = len(visible_windows)
            current_hwnds = {hwnd for hwnd, _, _ in visible_windows}

            # Add newly visible windows to grid_state (so they get tiled)
            updated = False
            for hwnd, title, _ in visible_windows:
                if hwnd not in grid_state:
                    grid_state[hwnd] = (0, 0, 0)   # temporary position
                    updated = True

            # Auto-retile only if number of visible windows changed
            if current_count != last_visible_count or updated:
                print(f"[AUTO-RETILE] {last_visible_count} → {current_count} visible windows")
                smart_tile(temp=True)
                last_visible_count = current_count
                time.sleep(0.2)  # small debounce

        # Update green border on active window
        active = user32.GetForegroundWindow()
        if active and user32.IsWindowVisible(active):
            if active != current_hwnd:
                apply_border(active)
        else:
            if current_hwnd:
                remove_border(current_hwnd)
                current_hwnd = None

        time.sleep(0.35)  # responsive but no CPU spam

# ==================================================================
# Hotkeys & main loop
# ==================================================================
def register_hotkeys():
    user32.RegisterHotKey(None, HOTKEY_TOGGLE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('T'))
    user32.RegisterHotKey(None, HOTKEY_RETILE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('R'))
    user32.RegisterHotKey(None, HOTKEY_QUIT,   win32con.MOD_CONTROL | win32con.MOD_ALT, ord('Q'))
    user32.RegisterHotKey(None, HOTKEY_MOVE_MONITOR, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('M'))

def unregister_hotkeys():
    for hk in (HOTKEY_TOGGLE, HOTKEY_RETILE, HOTKEY_QUIT, HOTKEY_MOVE_MONITOR):
        try: user32.UnregisterHotKey(None, hk)
        except: pass

def toggle_persistent():
    global is_active, monitor_thread
    is_active = not is_active
    print(f"\n[SMARTGRID] Persistent mode: {'ON' if is_active else 'OFF'}")
    if is_active:
        smart_tile(temp=False)

if __name__ == "__main__":
    print("="*70)
    print("   SMARTGRID - Intelligent layouts + green border")
    print("="*70)
    print("Ctrl+Alt+T  -> Toggle persistent tiling mode")
    print("Ctrl+Alt+R  -> One-shot re-tile of all visible windows")
    print("Ctrl+Alt+M  -> Cycle all tiled windows to the next monitor")
    print("Ctrl+Alt+Q  -> Quit SmartGrid")
    print("-"*70)

    clear_all_borders()
    time.sleep(0.05)

    # Start monitoring immediately (auto-retile + green border)
    threading.Thread(target=monitor, daemon=True).start()

    register_hotkeys()

    monitors = get_monitors()
    print(f"Detected {len(monitors)} monitor(s)")
    if len(monitors) > 1:
        print("Ctrl+Alt+M -> cycle tiled windows across monitors")

    msg = wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        if msg.message == win32con.WM_HOTKEY:
            if msg.wParam == HOTKEY_TOGGLE:
                toggle_persistent()
            elif msg.wParam == HOTKEY_RETILE:
                threading.Thread(target=smart_tile, kwargs={"temp": True}, daemon=True).start()
            elif msg.wParam == HOTKEY_MOVE_MONITOR:
                threading.Thread(target=move_all_tiled_to_next_monitor, daemon=True).start()
            elif msg.wParam == HOTKEY_QUIT:
                break
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

    if current_hwnd:
        remove_border(current_hwnd)
    clear_all_borders()    
    unregister_hotkeys()
    print("[EXIT] SmartGrid stopped.")