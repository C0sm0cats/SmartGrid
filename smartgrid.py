"""
SmartGrid – Advanced Windows Tiling Manager
Features: intelligent layouts, drag & drop snap, swap mode, workspaces per monitor,
          colored DWM borders, multi-monitor, systray, hotkeys
Author: C0sm0cats (2025)
"""

import os
import sys
import ctypes
import time
import threading
import winsound
from ctypes import wintypes
import win32gui
import win32api
import win32con
import win32process
from pystray import Icon, Menu, MenuItem
from PIL import Image, ImageDraw

# ==============================================================================
# CONFIGURATION & CONSTANTS
# ==============================================================================

DEBUG = 0  # Set to 1 for verbose logging during development

def log(*args, **kwargs):
    if DEBUG:
        print("[SmartGrid]", *args, **kwargs)

# Layout & appearance
GAP = 8           # Gap between tiled windows
EDGE_PADDING = 8  # Margin from screen edges

# DWM attributes
DWMWA_BORDER_COLOR = 34
DWMWA_COLOR_NONE = 0xFFFFFFFF
DWMWA_EXTENDED_FRAME_BOUNDS = 9

# Window styles
GWL_STYLE = -16
WS_THICKFRAME = 0x00040000
WS_MAXIMIZE = 0x01000000

# ShowWindow commands
SW_RESTORE = 9    # Restores a minimized or maximized window
SW_SHOWNORMAL = 1    # Activates and displays window (or restores if minimized)

# SetWindowPos flags (used in force_tile_resizable)
SWP_NOZORDER = 0x0004  # Ignores Z-order
SWP_NOMOVE = 0x0002  # Don't change position
SWP_NOSIZE = 0x0001  # Don't change size
SWP_NOACTIVATE = 0x0010  # Don't activate the window
SWP_FRAMECHANGED = 0x0020  # Force WM_NCCALCSIZE recalculation (border refresh)
SWP_NOSENDCHANGING = 0x0400 # Prevent WM_WINDOWPOSCHANGING (avoids conflicts)

# Hotkey IDs
HOTKEY_TOGGLE = 9001
HOTKEY_RETILE = 9002
HOTKEY_QUIT = 9003
HOTKEY_MOVE_MONITOR = 9004
HOTKEY_SWAP_MODE = 9005
HOTKEY_WS1 = 9101
HOTKEY_WS2 = 9102
HOTKEY_WS3 = 9103

# Swap mode arrow keys
HOTKEY_SWAP_LEFT = 9006
HOTKEY_SWAP_RIGHT = 9007
HOTKEY_SWAP_UP = 9008
HOTKEY_SWAP_DOWN = 9009
HOTKEY_SWAP_CONFIRM = 9010

# Custom toggle swap key - systray
CUSTOM_TOGGLE_SWAP = 0x9000

# Float toggle key
HOTKEY_FLOAT_TOGGLE = 9011

# Win32 API
user32 = ctypes.WinDLL('user32', use_last_error=True)   # Main Windows API
dwmapi = ctypes.WinDLL('dwmapi')  # Desktop Window Manager (for colored borders)

# ==============================================================================
# GLOBAL STATE
# ==============================================================================

# Window tracking
grid_state = {}     # hwnd → (monitor_idx, col, row)
minimized_windows = {}  # hwnd → (monitor, col, row) - saved positions
maximized_windows = {}  # hwnd → (monitor, col, row) - saved positions
current_hwnd = None   # Window with green border (active)
selected_hwnd = None   # Window with red border (swap mode)
last_active_hwnd = None   # Last known active/focused tiled window (preserved when focus is temporarily lost,

# Runtime flags
is_active = False            # Persistent tiling enabled?
swap_mode_lock = False       # Block auto-retile during swap
move_monitor_lock = False    # Block auto-retile during monitor move (Ctrl+Alt+M)
workspace_switching_lock = False  # Block auto-retile during workspace switch (Ctrl+Alt+1/2/3)
drag_drop_lock = False       # Block auto-retile during drag & drop
ignore_retile_until = 0.0    # Grace period after state changes

last_visible_count  = 0

# Background threads
monitor_thread        = None  # Thread running the main auto-retile + border monitor

# Multi-monitor & workspaces
MONITORS_CACHE = []
CURRENT_MONITOR_INDEX = 0
workspaces = {}
current_workspace = {}

# Overlay & systray
overlay_hwnd = None
preview_rect = None  # (x, y, w, h) of preview rectangle
tray_icon = None
main_thread_id = win32api.GetCurrentThreadId()

# Track the currently selected window in real-time
user_selected_hwnd = None  # Updated continuously to reflect user's current focus

# Float windows
override_windows = set()  # Set of HWNDs that should be excluded from tiling

# ==============================================================================
# UTILITY FUNCTIONS
# ==============================================================================

def get_monitors():
    """Return cached list of work area rectangles (x, y, w, h) for all monitors."""
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

def get_window_state(hwnd):
    """Return window state: 'normal', 'minimized', 'maximized', or 'hidden'"""
    if not hwnd or not user32.IsWindow(hwnd):
        return None
    
    if not user32.IsWindowVisible(hwnd):
        return 'hidden'
    
    if win32gui.IsIconic(hwnd):
        return 'minimized'
    
    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    if style & WS_MAXIMIZE:
        return 'maximized'
    
    return 'normal'

def get_frame_borders(hwnd):
    """Return (left, top, right, bottom) invisible border/shadow thickness"""
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

def set_window_border(hwnd, color):
    """Unique function to apply (or remove) a colored DWM border."""
    if not hwnd or not user32.IsWindow(hwnd):
        return
        
    try:
        if color is None:
            dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_BORDER_COLOR,
                ctypes.byref(ctypes.c_uint(DWMWA_COLOR_NONE)),
                ctypes.sizeof(ctypes.c_uint)
            )
        else:
            dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_BORDER_COLOR,
                ctypes.byref(ctypes.c_uint(color)),
                ctypes.sizeof(ctypes.c_uint)
            )
    except:
        pass

def get_process_name(hwnd):
    """Retrieves the name of the process associated with the window handle (e.g., 'ms-teams.exe').
    Returns an empty string if the handle is invalid or if an error occurs.
    """
    if not hwnd:
        return ""
    try:
        _, process_id = win32process.GetWindowThreadProcessId(hwnd)
        h_process = win32api.OpenProcess(0x0410, False, process_id)  # PROCESS_QUERY_INFORMATION | PROCESS_VM_READ
        path = win32process.GetModuleFileNameEx(h_process, 0)
        win32api.CloseHandle(h_process)
        return os.path.basename(path).lower()
    except:
        return ""

def get_window_size(hwnd):
    """Retrieves the width and height of a window using its Windows handle.
    Returns (0, 0) if the handle is invalid or if the window rectangle cannot be obtained.
    """
    if not hwnd or not user32.IsWindow(hwnd):
        return 0, 0
    rect = wintypes.RECT()
    if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        width = rect.right - rect.left
        height = rect.bottom - rect.top
        return width, height
    return 0, 0

def is_useful_window(title, class_name="", hwnd=None):
    """Filter out overlays, PIPs, taskbar, notifications, etc."""
    if not title:
        return False

    title_lower = title.lower()
    class_lower = class_name.lower() if class_name else ""

    # === SPECIAL CASE: Microsoft Teams
    if hwnd:
        process_name = get_process_name(hwnd)
        if process_name == "ms-teams.exe":
            w, h = get_window_size(hwnd)
            # Exclude only very small toast notifications (typical ~372x272)
            if w < 400 and h < 300:
                log(f"[FILTER] Excluded small Teams toast: {title[:50]} | size {w}x{h}")
                return False
            
            # Bonus for very short horizontal banners
            if h < 200:
                return False

    # === HARD EXCLUDE BY TITLE ===
    bad_titles = [
        "zscaler","spotify", "discord", "steam", "call", "meeting", "join", "incoming call",
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
        "credential dialog xaml host"    # Windows Security / zscaler
        # === Added to exclude Win+Tab / Alt+Tab task switcher ===
        "multitaskingviewframe",         # Win+Tab
        "taskswitcherwnd",               # Alt+Tab
        "xamlexplorerhostislandwindow",  # Alt+Tab / Win11 timeline
        "#32770",                        # MessageBox dialog class
        "windows.ui.popupwindowclass",   # WinUI menus
        "popuphostwindow",               # Menu dropdown host
        # === Added to exclude notepad menu popups (WinUI 3) ===
        "microsoft.ui.content.popupwindowsitebridge", # WinUI 3 menu popup
        "notepadshellexperiencehost"
    ]

    if class_lower in bad_classes:
        return False

    return True

def get_visible_windows():
    """Enumerate visible, non-minimized, non-maximized, tileable windows"""
    global overlay_hwnd

    monitors = get_monitors()
    windows = []
    
    def enum(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            state = get_window_state(hwnd)
            
            # Skip minimized/maximized
            if state in ('minimized', 'maximized'):
                return True
            elif state != 'normal':
                return True
        
            rect = wintypes.RECT()
            if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                if w > 180 and h > 180:
                    title_buf = ctypes.create_unicode_buffer(256)
                    user32.GetWindowTextW(hwnd, title_buf, 256)
                    title = title_buf.value or ""
                    class_name = win32gui.GetClassName(hwnd)

                    #proc = get_process_name(hwnd)
                    #log(f"[ENUM] hwnd={hwnd} title='{title[:60]}' class='{class_name}' state='{state}' size={w}x{h} proc='{proc}'")

                    if overlay_hwnd and hwnd == overlay_hwnd:
                        return True
                    
                    # override windows logic
                    useful = is_useful_window(title, class_name, hwnd=hwnd)
                    if hwnd in override_windows:
                        useful = not useful

                    if useful:
                        overlap = sum(
                            max(0, min(rect.right, mx + mw) - max(rect.left, mx)) *
                            max(0, min(rect.bottom, my + mh) - max(rect.top, my))
                            for mx, my, mw, mh in monitors
                        )
                        if overlap > (w * h * 0.15):
                            windows.append((hwnd, title, rect))
        return True

    user32.EnumWindows(
        ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(enum), 0
    )
    return windows

# ==============================================================================
# FLOATING WINDOW MANAGEMENT
# ==============================================================================

def toggle_floating_selected():
    """Ctrl+Alt+F → (tile → float, float → tile)   
    Uses user_selected_hwnd which is continuously updated by the monitor loop.
    This ensures we toggle the window the USER last clicked.
    """
    global user_selected_hwnd
    
    # Use the continuously-tracked user selection
    hwnd = user_selected_hwnd
    
    if not hwnd or not user32.IsWindow(hwnd) or not user32.IsWindowVisible(hwnd):
        winsound.MessageBeep(0xFFFFFFFF)
        log("[FLOAT] No valid window selected")
        return

    title = win32gui.GetWindowText(hwnd)
    class_name = win32gui.GetClassName(hwnd)
    default_useful = is_useful_window(title, class_name)

    if hwnd in override_windows:
        override_windows.remove(hwnd)
        log(f"[OVERRIDE] {title[:60]} → rollback to ({'tile' if default_useful else 'float'})")
        if default_useful:
            smart_tile_with_restore()
        winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)
    else:
        override_windows.add(hwnd)
        log(f"[OVERRIDE] {title[:60]} → override to ({'float' if default_useful else 'tile'})")
        if default_useful:
            if hwnd in grid_state:
                grid_state.pop(hwnd, None)
            set_window_border(hwnd, None)
        elif not default_useful:
            smart_tile_with_restore()
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS | winsound.SND_ASYNC)

    update_tray_menu()

# ==============================================================================
# LAYOUT ENGINE
# ==============================================================================

def choose_layout(count):
    """Choose optimal layout based on window count (full → side-by-side → grid)"""
    if count == 1: return "full", None
    if count == 2: return "side_by_side", None
    if count == 3: return "master_stack", None
    if count == 4: return "grid", (2, 2)
    if count <= 6: return "grid", (3, 2)
    if count <= 9: return "grid", (3, 3)
    if count <= 12: return "grid", (4, 3)
    return "grid", (5, 3)

def get_window_class(hwnd):
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value

def smart_tile_with_restore():
    """Smart tiling that respects saved grid positions"""
    global grid_state, ignore_retile_until

    if time.time() < ignore_retile_until:
        return
    
    ignore_retile_until = time.time() + 0.3

    monitors = get_monitors()

    # === STEP 1: AGGRESSIVE CLEANUP - Remove DEAD windows first ===
    dead_windows = []
    for hwnd in list(grid_state.keys()):
        # Check if window REALLY exists
        if not user32.IsWindow(hwnd):
            dead_windows.append(hwnd)
            grid_state.pop(hwnd, None)
            continue
        
        # Check if window is in min/max state (we don't tile those)
        state = get_window_state(hwnd)
        if state in ('minimized', 'maximized'):
            grid_state.pop(hwnd, None)
            continue
        
        # Check if window is actually visible
        if not user32.IsWindowVisible(hwnd):
            grid_state.pop(hwnd, None)
            continue
    
    if dead_windows:
        log(f"[CLEAN] Removed {len(dead_windows)} dead windows")

    # === STEP 2: SOFT CLEANUP - Remove GHOST windows (zombie-like but technically exist) ===
    ghost_windows = []
    for hwnd in list(grid_state.keys()):
        if not user32.IsWindow(hwnd):
            # Already removed in step 1, skip
            continue
        
        title = win32gui.GetWindowText(hwnd)
        rect = wintypes.RECT()
        
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            # Can't get rect = something's wrong
            grid_state.pop(hwnd, None)
            ghost_windows.append(hwnd)
            continue
        
        w = rect.right - rect.left
        h = rect.bottom - rect.top
        
        # Ghost = no title AND very small (likely leftover process window)
        if (not title or len(title.strip()) == 0) and (w < 50 or h < 50):
            log(f"[CLEAN] Ghost window: hwnd={hwnd} size={w}x{h}")
            grid_state.pop(hwnd, None)
            ghost_windows.append(hwnd)
    
    if ghost_windows:
        log(f"[CLEAN] Removed {len(ghost_windows)} ghost windows")

    visible_windows = get_visible_windows()
    
    if not visible_windows:
        log("[TILE] No windows detected.")
        return

    # Separate windows by monitor
    wins_by_monitor = {}
    for hwnd, title, rect in visible_windows:
        win_class = get_window_class(hwnd)
        # Check if window has a saved position (was minimized/maximized)
        if hwnd in minimized_windows:
            mon_idx, col, row = minimized_windows[hwnd]
            # Remove from minimized tracking
            del minimized_windows[hwnd]
            log(f"[RESTORE] Restoring minimized window to ({col},{row}): {title[:40]}")
        elif hwnd in maximized_windows:
            mon_idx, col, row = maximized_windows[hwnd]
            del maximized_windows[hwnd]
            log(f"[RESTORE] Restoring maximized window to ({col},{row}): {title[:40]}")
        elif hwnd in grid_state:
            # Keep existing position
            mon_idx, col, row = grid_state[hwnd]
        else:
            # New window - will be assigned below
            mon_idx = 0
            col, row = 0, 0
        
        wins_by_monitor.setdefault(mon_idx, []).append((hwnd, title, rect, col, row, win_class))
    
    # Process each monitor
    new_grid = {}
    
    for mon_idx, windows in wins_by_monitor.items():
        if mon_idx >= len(monitors):
            continue
            
        mon_x, mon_y, mon_w, mon_h = monitors[mon_idx]
        
        count = len(windows)
        layout, info = choose_layout(count)
        
        log(f"\n[TILE] Monitor {mon_idx+1}: {count} windows -> {layout} layout")
        
        # Build position map based on current layout
        if layout == "full":
            positions = [(mon_x + EDGE_PADDING, mon_y + EDGE_PADDING,
                         mon_w - 2*EDGE_PADDING, mon_h - 2*EDGE_PADDING)]
            grid_coords = [(0, 0)]
            
        elif layout == "side_by_side":
            cw = (mon_w - 2*EDGE_PADDING - GAP) // 2
            positions = [
                (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, cw, mon_h - 2*EDGE_PADDING),
                (mon_x + EDGE_PADDING + cw + GAP, mon_y + EDGE_PADDING, cw, mon_h - 2*EDGE_PADDING)
            ]
            grid_coords = [(0, 0), (1, 0)]
            
        elif layout == "master_stack":
            mw = (mon_w - 2*EDGE_PADDING - GAP) * 3 // 5
            sw = mon_w - 2*EDGE_PADDING - mw - GAP
            sh = (mon_h - 2*EDGE_PADDING - GAP) // 2
            positions = [
                (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, mw, mon_h - 2*EDGE_PADDING),
                (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING, sw, sh),
                (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING + sh + GAP, sw, sh)
            ]
            grid_coords = [(0, 0), (1, 0), (1, 1)]
            
        else:  # grid
            cols, rows = info
            total_gaps_w = GAP * (cols - 1) if cols > 1 else 0
            total_gaps_h = GAP * (rows - 1) if rows > 1 else 0
            cell_w = (mon_w - 2*EDGE_PADDING - total_gaps_w) // cols
            cell_h = (mon_h - 2*EDGE_PADDING - total_gaps_h) // rows
            
            positions = []
            grid_coords = []
            for r in range(rows):
                for c in range(cols):
                    if len(positions) >= count:
                        break
                    x = mon_x + EDGE_PADDING + c * (cell_w + GAP)
                    y = mon_y + EDGE_PADDING + r * (cell_h + GAP)
                    positions.append((x, y, cell_w, cell_h))
                    grid_coords.append((c, r))
        
        # Create position lookup
        pos_map = dict(zip(grid_coords, positions))
        
        # SMART ASSIGNMENT: Try to restore saved positions first
        assigned = set()
        unassigned_windows = []
        
        # Phase 1: Assign windows that have saved positions
        for hwnd, title, rect, saved_col, saved_row, win_class in windows:
            target_coords = (saved_col, saved_row)
            
            # Check if saved position is valid in current layout
            if target_coords in pos_map and target_coords not in assigned:
                # Perfect - restore to exact position
                x, y, w, h = pos_map[target_coords]
                force_tile_resizable(hwnd, x, y, w, h)
                new_grid[hwnd] = (mon_idx, saved_col, saved_row)
                assigned.add(target_coords)
                log(f"   ✓ RESTORED to ({saved_col},{saved_row}): {title[:50]} [{win_class}]")
                time.sleep(0.015)
            else:
                # Saved position doesn't exist in current layout - reassign later
                unassigned_windows.append((hwnd, title, rect, 0, 0, win_class))
        
        # Phase 2: Assign remaining windows to available positions
        available_positions = [coord for coord in grid_coords if coord not in assigned]
        
        for i, (hwnd, title, rect, saved_col, saved_row, win_class) in enumerate(unassigned_windows):
            if i < len(available_positions):
                col, row = available_positions[i]
                x, y, w, h = pos_map[(col, row)]
                force_tile_resizable(hwnd, x, y, w, h)
                new_grid[hwnd] = (mon_idx, col, row)
                log(f"   → NEW position ({col},{row}): {title[:50]} [{win_class}]")
                time.sleep(0.015)
    
    grid_state = new_grid
    
    time.sleep(0.06)

def force_tile_resizable(hwnd, x, y, w, h):
    """Move and resize window to exact coordinates, handling borders and restore from min/max states."""
    state = get_window_state(hwnd)
    if state in ('minimized', 'maximized'):
        return

    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    if not (style & WS_MAXIMIZE):
        user32.SetWindowLongW(hwnd, GWL_STYLE, (style | WS_THICKFRAME) & ~WS_MAXIMIZE)

    if state != 'normal':
        user32.ShowWindowAsync(hwnd, SW_RESTORE)
        for _ in range(10):
            if get_window_state(hwnd) == 'normal':
                break
            time.sleep(0.02)

    time.sleep(0.012)
    user32.SetWindowPos(hwnd, 0, 0, 0, 0, 0,
                        SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED)

    # Mandatory initialization step
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
        log(f"   [NUKE] MoveWindow forced on -> {title[:60]}")
        win32gui.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True)

def force_immediate_retile():
    """Ctrl+Alt+R → force immediate re-tile (bypass all grace delays)."""
    global ignore_retile_until, last_visible_count, swap_mode_lock
    
    if swap_mode_lock:  # swap mode enabled → we don't retile
        return
        
    log("\n[FORCE RETILE] Ctrl+Alt+R → Immediate full re-tile (bypasses all grace periods)")
    ignore_retile_until = 0          # Bypass the usual 0.6s grace period
    last_visible_count = 0           # Force change detection on next cycle
    try:
        winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS | winsound.SND_ASYNC)
    except:
        pass
    smart_tile_with_restore()

# ==============================================================================
# BORDER MANAGEMENT
# ==============================================================================

def apply_border(hwnd):
    """Central wrapper for green active window border."""
    if not hwnd or not user32.IsWindow(hwnd):
        return
    set_window_border(hwnd, 0x0000FF00)  # vert
    global current_hwnd
    current_hwnd = hwnd

def update_active_border():
    """
    Manage colored DWM borders for visual feedback.
    - Green border: marks the currently active tiled window
    - Red border: marks the selected window in swap mode
    Continuously reapplies borders as Windows DWM can remove them automatically.
    """
    global current_hwnd, last_active_hwnd

    # SWAP MODE: Continuously reapply red border
    if swap_mode_lock and selected_hwnd and user32.IsWindow(selected_hwnd):
        set_window_border(selected_hwnd, 0x000000FF)  # Red
        return

    active = user32.GetForegroundWindow()

    # If active window is tiled → apply/maintain green border
    if active in grid_state and user32.IsWindow(active):
        # Save as last active tiled window
        last_active_hwnd = active

        # Only change if active window is different
        if current_hwnd != active:
            # Remove border from previous tiled window
            if current_hwnd and user32.IsWindow(current_hwnd):
                set_window_border(current_hwnd, None)

            # Apply border to new window
            apply_border(active)
            current_hwnd = active
        else:
            # REAPPLY border even if same window
            # (Windows may have removed it automatically)
            set_window_border(current_hwnd, 0x0000FF00)

        return

    # If foreground is NON-tiled (taskbar, systray, menu, icon, popup)
    # → MAINTAIN green border on last known tiled window
    if last_active_hwnd and user32.IsWindow(last_active_hwnd) and last_active_hwnd in grid_state:
        # Continuously reapply border
        set_window_border(last_active_hwnd, 0x0000FF00)
        current_hwnd = last_active_hwnd

# ==============================================================================
# SWAP MODE
# ==============================================================================

def enter_swap_mode():
    """Enter swap mode: red border + arrow keys to swap window positions."""
    global selected_hwnd, last_active_hwnd, swap_mode_lock

    swap_mode_lock = True          # Prevent automatic re-tiling during swap mode
    time.sleep(0.25)
    
    # Force a quick update of grid_state in case it's empty or updating
    if not grid_state:
        visible_windows = get_visible_windows()
        for hwnd, title, _ in visible_windows:
            if hwnd not in grid_state:
                grid_state[hwnd] = (0, 0, 0)
        log(f"[SWAP] grid_state was empty → rebuilt with {len(grid_state)} windows")
        if not grid_state:
            log("[SWAP] No tiled windows detected. First press Ctrl+Alt+T or Ctrl+Alt+R")
            return

    # SMART SELECTION LOGIC FOR SWAP MODE (priority order)
    candidate = None
    
    # 1. Highest priority: last known active tiled window (preserves selection even if focus is temporarily lost)
    if last_active_hwnd and user32.IsWindow(last_active_hwnd) and last_active_hwnd in grid_state:
        candidate = last_active_hwnd
    # 2. Fallback: currently foreground window (if it's tiled)
    elif user32.GetForegroundWindow() in grid_state:
        candidate = user32.GetForegroundWindow()
    # 3. Last resort: first window in the current grid
    else:
        candidate = next(iter(grid_state.keys()), None)

    if not candidate:
        log("[SWAP] No valid window to select")
        return

    selected_hwnd = candidate

    time.sleep(0.05)
    set_window_border(selected_hwnd, 0x000000FF)  # Red

    log(f"[SWAP] Swap mode enabled — target window marked with solid red border")
    register_swap_hotkeys()
    update_tray_menu()
    
    title = win32gui.GetWindowText(selected_hwnd)[:50]
    log(f"\n[SWAP] ✓ Activated - Selected window: '{title}'")
    log(f"[SWAP] Selected hwnd: {selected_hwnd}")
    log("━" * 60)
    log("  DIRECT SWAP with arrow keys:")
    log("    ← → ↑ ↓  : Swap with adjacent window")
    log("    Enter     : Confirm selection")
    log("    Ctrl+Alt+S : Exit swap mode")
    log("  ")
    log("  The red window FOLLOWS your movements and swaps its position!")
    log("━" * 60)
    register_swap_hotkeys()
    # Refresh tray menu
    update_tray_menu()

def navigate_swap(direction):
    """Handle arrow key press in swap mode → find and swap with adjacent window"""
    global selected_hwnd
    
    if not swap_mode_lock or not selected_hwnd:
        log("[SWAP] Mode not active or no window selected")
        return
    
    log(f"[SWAP] Attempting to swap {direction}...")
    
    # Find the window in the specified direction
    target = find_window_in_direction(selected_hwnd, direction)
    
    if target:
        # DIRECT SWAP!
        if swap_windows(selected_hwnd, target):
            # The selected window (red) has moved, we follow it
            # selected_hwnd logically remains the same, but physically it has changed position
            
            # Clear and reapply the red border on the window that moved
            time.sleep(0.04)
            set_window_border(selected_hwnd, None)
            time.sleep(0.04)
            set_window_border(selected_hwnd, 0x000000FF)
            
            title = win32gui.GetWindowText(selected_hwnd)[:50]
            log(f"[SWAP] ✓ '{title}' swapped {direction}")
            
            # Focus on the window that moved
            user32.SetForegroundWindow(selected_hwnd)
        else:
            log(f"[SWAP] ✗ Swap failed")
    else:
        log(f"[SWAP] ✗ No window in the {direction} direction (grid limit)")

def confirm_swap():
    """Confirm current swap selection and exit swap mode (bound to Enter)"""
    global swap_mode_lock, selected_hwnd

    if not swap_mode_lock or not selected_hwnd:
        log("[SWAP] Swap mode not active or no window selected")
        return

    log(f"[SWAP] Confirming swap for hwnd {selected_hwnd}")

    exit_swap_mode()
    log("[SWAP] ✓ Swap confirmed and mode exited")

def exit_swap_mode():
    """Exit swap mode, clean borders, restore normal green border on active window."""
    global selected_hwnd, current_hwnd, swap_mode_lock
    
    if not swap_mode_lock:
        return
    
    # First, clear ALL borders
    log("[SWAP] Clearing borders...")
    if selected_hwnd and user32.IsWindow(selected_hwnd):
        set_window_border(selected_hwnd, None)
    selected_hwnd = None
    time.sleep(0.06)

    unregister_swap_hotkeys()
    selected_hwnd = None
    current_hwnd = None
    
    # Restore the green border on the active window
    active = user32.GetForegroundWindow()
    if active and user32.IsWindowVisible(active) and active in grid_state:
        apply_border(active)
        log(f"[SWAP] Green border restored on active window")

    swap_mode_lock = False    
    log("[SWAP] ✓ Deactivated\n")
    # Refresh tray menu
    update_tray_menu()

def find_window_in_direction(from_hwnd, direction):
    """Find the closest tiled window in the specified direction (left/right/up/down)"""
    if not grid_state or from_hwnd not in grid_state:
        return None

    mon_idx = grid_state[from_hwnd][0]

    # Get the actual position of the selected window
    from_rect = wintypes.RECT()
    if not user32.GetWindowRect(from_hwnd, ctypes.byref(from_rect)):
        return None

    fx1, fy1, fx2, fy2 = from_rect.left, from_rect.top, from_rect.right, from_rect.bottom
    fcx = (fx1 + fx2) // 2
    fcy = (fy1 + fy2) // 2

    from_width = fx2 - fx1
    from_height = fy2 - fy1

    best_hwnd = None
    best_distance = float('inf')  # We want the CLOSEST one

    for hwnd, (m, _, _) in grid_state.items():
        if hwnd == from_hwnd or m != mon_idx or not user32.IsWindow(hwnd):
            continue

        rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            continue

        x1, y1, x2, y2 = rect.left, rect.top, rect.right, rect.bottom
        cx = (x1 + x2) // 2
        cy = (y1 + y2) // 2

        dx = cx - fcx
        dy = cy - fcy

        # --- STRICT DIRECTION ---
        if direction == "right" and dx <= 30: continue
        if direction == "left"  and dx >= -30: continue
        if direction == "down"  and dy <= 30: continue
        if direction == "up"    and dy >= -30: continue

        # --- DYNAMIC OVERLAP THRESHOLD (NEW) ---
        # Calculate overlap in the perpendicular axis
        overlap_x = max(0, min(fx2, x2) - max(fx1, x1))
        overlap_y = max(0, min(fy2, y2) - max(fy1, y1))

        # Instead of hard-coded 50px, use percentage of source window size
        # This adapts to whether windows are huge or compressed
        if direction in ("left", "right"):
            # For horizontal movement, check vertical alignment
            # Require at least 20% of source height to overlap (was hard 50px)
            min_overlap_y = max(20, int(from_height * 0.20))
            
            if overlap_y < min_overlap_y:
                log(f"[SWAP] {hwnd} rejected: overlap_y={overlap_y} < min={min_overlap_y} (20% of {from_height})")
                continue
        else:
            # For vertical movement, check horizontal alignment
            # Require at least 20% of source width to overlap
            min_overlap_x = max(20, int(from_width * 0.20))
            
            if overlap_x < min_overlap_x:
                log(f"[SWAP] {hwnd} rejected: overlap_x={overlap_x} < min={min_overlap_x} (20% of {from_width})")
                continue

        # --- EUCLIDEAN DISTANCE (the key: we take the CLOSEST one) ---
        distance = (dx * dx) + (dy * dy)

        # Bonus: more overlap = better (in case of equal distance)
        alignment_bonus = (overlap_x if direction in ("up", "down") else overlap_y) * 10

        score = distance - alignment_bonus  # we MINIMIZE this score

        if score < best_distance:
            best_distance = score
            best_hwnd = hwnd

    return best_hwnd

def swap_windows(hwnd1, hwnd2):
    """Swap two windows' grid positions and physically move them (preserves sizes)"""
    if hwnd1 not in grid_state or hwnd2 not in grid_state:
        return False
    
    # First clear all borders
    set_window_border(hwnd1, None)
    set_window_border(hwnd2, None)
    time.sleep(0.05)
    
    # Retrieve the current positions
    rect1 = wintypes.RECT()
    rect2 = wintypes.RECT()
    
    if not user32.GetWindowRect(hwnd1, ctypes.byref(rect1)):
        return False
    if not user32.GetWindowRect(hwnd2, ctypes.byref(rect2)):
        return False
    
    lb1, tb1, rb1, bb1 = get_frame_borders(hwnd1)
    lb2, tb2, rb2, bb2 = get_frame_borders(hwnd2)
    
    # Calculate dimensions without borders
    x1, y1 = rect1.left + lb1, rect1.top + tb1
    w1 = rect1.right - rect1.left - lb1 - rb1
    h1 = rect1.bottom - rect1.top - tb1 - bb1
    
    x2, y2 = rect2.left + lb2, rect2.top + tb2
    w2 = rect2.right - rect2.left - lb2 - rb2
    h2 = rect2.bottom - rect2.top - tb2 - bb2
    
    title1 = win32gui.GetWindowText(hwnd1)[:40]
    title2 = win32gui.GetWindowText(hwnd2)[:40]
    log(f"[SWAP] '{title1}' ↔ '{title2}'")
    
    # Swap the positions in grid_state
    grid_state[hwnd1], grid_state[hwnd2] = grid_state[hwnd2], grid_state[hwnd1]
    
    # Physically move the windows (SWAPPING positions, not sizes)
    force_tile_resizable(hwnd1, x2, y2, w2, h2)
    time.sleep(0.04)
    force_tile_resizable(hwnd2, x1, y1, w1, h1)
    time.sleep(0.04)
    return True

# ==============================================================================
# DRAG & DROP SNAP + PREVIEW
# ==============================================================================

class PAINTSTRUCT(ctypes.Structure):
    _fields_ = [
        ("hdc", wintypes.HDC),
        ("fErase", wintypes.BOOL),
        ("rcPaint", wintypes.RECT),
        ("fRestore", wintypes.BOOL),
        ("fIncUpdate", wintypes.BOOL),
        ("rgbReserved", ctypes.c_char * 32)
    ]

def create_overlay_window():
    """Create transparent topmost overlay for drag-and-drop snap preview."""
    global overlay_hwnd

    if overlay_hwnd:
        return overlay_hwnd

    class_name = "SmartGridOverlay"
    
    # Properly define DefWindowProc with correct signature
    DefWindowProc = ctypes.windll.user32.DefWindowProcW
    DefWindowProc.argtypes = [wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM]
    DefWindowProc.restype = wintypes.LPARAM
    
    # Define WNDPROCTYPE properly
    WNDPROCTYPE = ctypes.WINFUNCTYPE(
        wintypes.LPARAM,
        wintypes.HWND,
        ctypes.c_uint,
        wintypes.WPARAM,
        wintypes.LPARAM
    )
    
    # Window procedure for overlay
    @WNDPROCTYPE
    def wnd_proc(hwnd, msg, wparam, lparam):
        try:
            if msg == win32con.WM_PAINT:
                ps = PAINTSTRUCT()
                hdc = user32.BeginPaint(hwnd, ctypes.byref(ps))

                if preview_rect:
                    _, _, w, h = preview_rect  # absolute coords kept, but drawing is local

                    brush = win32gui.CreateSolidBrush(win32api.RGB(100, 149, 237))
                    pen = win32gui.CreatePen(win32con.PS_SOLID, 4, win32api.RGB(65, 105, 225))

                    old_brush = win32gui.SelectObject(hdc, brush)
                    old_pen = win32gui.SelectObject(hdc, pen)

                    # Draw in local client coords (0..w, 0..h)
                    win32gui.Rectangle(hdc, 2, 2, w - 2, h - 2)

                    win32gui.SelectObject(hdc, old_brush)
                    win32gui.SelectObject(hdc, old_pen)
                    win32gui.DeleteObject(brush)
                    win32gui.DeleteObject(pen)

                user32.EndPaint(hwnd, ctypes.byref(ps))
                return 0

            elif msg == win32con.WM_DESTROY:
                return 0

            return DefWindowProc(hwnd, msg, wparam, lparam)

        except Exception as e:
            log(f"[OVERLAY] Error in wnd_proc: {e}")
            return 0

    wc = win32gui.WNDCLASS()
    wc.lpfnWndProc = wnd_proc
    wc.lpszClassName = class_name
    wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)

    try:
        win32gui.RegisterClass(wc)
    except:
        pass

    overlay_hwnd = win32gui.CreateWindowEx(
        win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT |
        win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
        class_name,
        "SmartGrid Preview",
        win32con.WS_POPUP,
        0, 0, 1, 1,
        0, 0, 0, None
    )

    win32gui.SetLayeredWindowAttributes(overlay_hwnd, 0, int(255 * 0.3), win32con.LWA_ALPHA)

    return overlay_hwnd


def show_snap_preview(x, y, w, h):
    """Show blue snap preview rectangle."""
    global preview_rect, overlay_hwnd

    if not overlay_hwnd:
        create_overlay_window()

    preview_rect = (x, y, w, h)
    
    # Position and show overlay window
    win32gui.SetWindowPos(
        overlay_hwnd, win32con.HWND_TOPMOST,
        int(x), int(y), int(w), int(h),
        win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE
    )
    
    # Force redraw
    win32gui.InvalidateRect(overlay_hwnd, None, True)
    win32gui.UpdateWindow(overlay_hwnd)

def hide_snap_preview():
    """Hide snap preview overlay."""
    global preview_rect, overlay_hwnd
    preview_rect = None
    if overlay_hwnd:
        win32gui.ShowWindow(overlay_hwnd, win32con.SW_HIDE)


def calculate_target_rect(source_hwnd, cursor_pos):
    """Compute target snap rectangle (absolute coords) with frame-borders included."""
    
    if source_hwnd not in grid_state:
        return None

    monitors = get_monitors()
    cx, cy = cursor_pos
    
    # Find target monitor
    target_mon_idx = 0
    for i, (mx, my, mw, mh) in enumerate(monitors):
        if mx <= cx < mx + mw and my <= cy < my + mh:
            target_mon_idx = i
            break

    mon_x, mon_y, mon_w, mon_h = monitors[target_mon_idx]
    
    # Get windows currently on this monitor
    wins_on_mon = [h for h, (m, _, _) in grid_state.items() 
                   if m == target_mon_idx and user32.IsWindow(h) and h != source_hwnd]

    count_including_source = len(wins_on_mon) + 1
    current_layout, current_info = choose_layout(count_including_source)

    # Default
    x = y = 0
    w = h = 0

    if current_layout == "master_stack":
        master_w = (mon_w - 2*EDGE_PADDING - GAP) * 3 // 5
        master_right = mon_x + EDGE_PADDING + master_w + GAP//2

        if cx < master_right:
            x = mon_x + EDGE_PADDING
            y = mon_y + EDGE_PADDING
            w = master_w
            h = mon_h - 2*EDGE_PADDING
        else:
            sw = mon_w - 2*EDGE_PADDING - master_w - GAP
            sh = (mon_h - 2*EDGE_PADDING - GAP) // 2
            mid = mon_y + mon_h // 2

            x = mon_x + EDGE_PADDING + master_w + GAP
            y = mon_y + EDGE_PADDING + (sh + GAP if cy >= mid else 0)
            w = sw
            h = sh

    elif current_layout == "side_by_side":
        cw = (mon_w - 2*EDGE_PADDING - GAP) // 2

        x = mon_x + EDGE_PADDING + (0 if cx < mon_x + mon_w//2 else cw + GAP)
        y = mon_y + EDGE_PADDING
        w = cw
        h = mon_h - 2*EDGE_PADDING

    elif current_layout == "full":
        x = mon_x + EDGE_PADDING
        y = mon_y + EDGE_PADDING
        w = mon_w - 2*EDGE_PADDING
        h = mon_h - 2*EDGE_PADDING

    else:  # Grid
        cols, rows = current_info if current_info else (2, 2)
        maxc = max((c for h,(m,c,r) in grid_state.items() if m == target_mon_idx), default=0)
        maxr = max((r for h,(m,c,r) in grid_state.items() if m == target_mon_idx), default=0)
        cols = max(cols, maxc + 1)
        rows = max(rows, maxr + 1)

        cw = (mon_w - 2*EDGE_PADDING - GAP*(cols-1)) // cols
        ch = (mon_h - 2*EDGE_PADDING - GAP*(rows-1)) // rows

        relx = cx - mon_x - EDGE_PADDING
        rely = cy - mon_y - EDGE_PADDING
        col = min(max(0, relx // (cw + GAP)), cols-1)
        row = min(max(0, rely // (ch + GAP)), rows-1)

        x = mon_x + EDGE_PADDING + col * (cw + GAP)
        y = mon_y + EDGE_PADDING + row * (ch + GAP)
        w = cw
        h = ch

    # ---- FRAME BORDER COMPENSATION (as tiling does) ----
    try:
        lb, tb, rb, bb = get_frame_borders(source_hwnd)
    except:
        lb = tb = rb = bb = 0

    return (
        int(x - lb),
        int(y - tb),
        int(w + lb + rb),
        int(h + tb + bb)
    )

def apply_grid_state():
    """Reapply all saved grid positions physically (used after snap, swap, or workspace change)."""
    global grid_state
    
    if not grid_state:
        return
    
    monitors = get_monitors()
    
    # Remove dead windows
    for hwnd in list(grid_state.keys()):
        if not user32.IsWindow(hwnd):
            grid_state.pop(hwnd, None)
    
    if not grid_state:
        return
    
    # Group windows by monitor
    wins_by_mon = {}
    for hwnd, (mon_idx, col, row) in grid_state.items():
        wins_by_mon.setdefault(mon_idx, []).append((hwnd, col, row))
    
    for mon_idx, windows in wins_by_mon.items():
        if mon_idx >= len(monitors):
            continue
        mon_x, mon_y, mon_w, mon_h = monitors[mon_idx]
        count = len(windows)
        layout, info = choose_layout(count)
        
        # === CALCULATE PHYSICAL POSITIONS FOR THIS MONITOR ===
        if layout == "full":
            positions = [(mon_x + EDGE_PADDING, mon_y + EDGE_PADDING,
                          mon_w - 2*EDGE_PADDING, mon_h - 2*EDGE_PADDING)]
            coord_list = [(0, 0)]
            
        elif layout == "side_by_side":
            cw = (mon_w - 2*EDGE_PADDING - GAP) // 2
            positions = [
                (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, cw, mon_h - 2*EDGE_PADDING),
                (mon_x + EDGE_PADDING + cw + GAP, mon_y + EDGE_PADDING, cw, mon_h - 2*EDGE_PADDING)
            ]
            coord_list = [(0, 0), (1, 0)]
            
        elif layout == "master_stack":
            mw = (mon_w - 2*EDGE_PADDING - GAP) * 3 // 5
            sw = mon_w - 2*EDGE_PADDING - mw - GAP
            sh = (mon_h - 2*EDGE_PADDING - GAP) // 2
            positions = [
                (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, mw, mon_h - 2*EDGE_PADDING),    # master
                (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING, sw, sh),             # Stack pane → top-right
                (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING + sh + GAP, sw, sh)   # Stack pane → bottom-right
            ]
            coord_list = [(0, 0), (1, 0), (1, 1)]
            
        else:  # grid → use the exact (col, row) positions stored in grid_state
            max_col = max((c for _, c, _ in windows), default=0)
            max_row = max((r for _, _, r in windows), default=0)
            cols = max(info[0], max_col + 1)
            rows = max(info[1], max_row + 1)
            
            total_gaps_w = GAP * (cols - 1) if cols > 1 else 0
            total_gaps_h = GAP * (rows - 1) if rows > 1 else 0
            cell_w = (mon_w - 2*EDGE_PADDING - total_gaps_w) // cols
            cell_h = (mon_h - 2*EDGE_PADDING - total_gaps_h) // rows
            
            positions = []
            coord_list = []
            for r in range(rows):
                for c in range(cols):
                    x = mon_x + EDGE_PADDING + c * (cell_w + GAP)
                    y = mon_y + EDGE_PADDING + r * (cell_h + GAP)
                    positions.append((x, y, cell_w, cell_h))
                    coord_list.append((c, r))
        
        # === APPLY POSITIONS TO WINDOWS ===
        pos_dict = dict(zip(coord_list, positions))
        
        for hwnd, col, row in windows:
            key = (col, row)
            if key in pos_dict:
                x, y, w, h = pos_dict[key]
                force_tile_resizable(hwnd, x, y, w, h)
                time.sleep(0.008)
    
    time.sleep(0.03)

def is_window_maximized(hwnd):
    """Return True if window is maximized. Used to block drag on maximized windows."""
    if not hwnd or not user32.IsWindow(hwnd):
        return False
    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    return bool(style & WS_MAXIMIZE)

def start_drag_snap_monitor():
    """Background thread: real-time drag detection with live snap preview."""
    global drag_drop_lock
    
    was_down = False
    drag_hwnd = None
    drag_start = None
    preview_active = False
    last_valid_rect = None

    while True:
        try:
            down = win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000
            
            # ----------- MOUSE DOWN -----------
            if down and not was_down:

                # fetch cursor position
                pt = win32api.GetCursorPos()

                # Try window under cursor
                hwnd = win32gui.WindowFromPoint(pt)

                # If our overlay is hit, fallback to foreground window
                if 'overlay_hwnd' in globals() and overlay_hwnd and hwnd == overlay_hwnd:
                    hwnd = user32.GetForegroundWindow()

                # Climb to top-level (WindowFromPoint returns child controls)
                try:
                    GA_ROOT = 2  # GA_ROOT
                    top = ctypes.windll.user32.GetAncestor(hwnd, GA_ROOT)
                    if top:
                        hwnd = top
                except Exception:
                    # fallback parent climb
                    try:
                        while True:
                            parent = win32gui.GetParent(hwnd)
                            if not parent:
                                break
                            hwnd = parent
                    except Exception:
                        pass

                # If unusable (invisible, tool window), fallback to foreground
                if not hwnd or not user32.IsWindowVisible(hwnd):
                    hwnd = user32.GetForegroundWindow()

                # Ignore maxed windows (original logic)
                if hwnd and is_window_maximized(hwnd):
                    was_down = True
                    continue

                # Only accept windows managed by SmartGrid
                if hwnd and hwnd in grid_state and user32.IsWindowVisible(hwnd):
                    drag_drop_lock = True
                    time.sleep(0.25) # Prevent automatic re-tiling during drag
                    drag_hwnd = hwnd
                    drag_start = pt
                    preview_active = False

            # ----------- DRAG IN PROGRESS -----------
            elif down and drag_hwnd:
                cursor_pos = win32api.GetCursorPos()

                if drag_start:
                    dx = abs(cursor_pos[0] - drag_start[0])
                    dy = abs(cursor_pos[1] - drag_start[1])
                    
                    # Show preview after 1px movement
                    if (dx > 1 or dy > 1) and not preview_active:    
                        preview_active = True
                    
                    # Update preview during drag
                    if preview_active:
                        target_rect = calculate_target_rect(drag_hwnd, cursor_pos)
                        if target_rect:
                            last_valid_rect = target_rect
                            show_snap_preview(*target_rect)
                        elif last_valid_rect:
                            show_snap_preview(*last_valid_rect)
                        else:
                            hide_snap_preview()
            
            # ----------- MOUSE UP (DROP) -----------
            elif not down and was_down and drag_hwnd:
                hide_snap_preview()
                cursor_pos = win32api.GetCursorPos()
                moved = False
                if drag_start:
                    dx = abs(cursor_pos[0] - drag_start[0])
                    dy = abs(cursor_pos[1] - drag_start[1])
                    moved = (dx > 10 or dy > 10)

                if moved:
                    # Delegate all snapping logic to the central function
                    # which already computes target monitor/col/row and updates grid_state.
                    handle_snap_drop(drag_hwnd, cursor_pos)
                    # handle_snap_drop must ensure re-application even when position unchanged
                else:
                    # released without movement → force re-apply positions so the window
                    # snaps perfectly back to its stored cell
                    apply_grid_state()

                drag_drop_lock = False
                drag_hwnd = None
                drag_start = None
                preview_active = False
            
            was_down = down
            time.sleep(0.005)

        except Exception as e:
            hide_snap_preview()
            drag_drop_lock = False
            drag_hwnd = None
            drag_start = None
            preview_active = False
            time.sleep(0.1)

def handle_snap_drop(source_hwnd, cursor_pos):
    """Handle window drop during drag → snap to target cell (swap or move) and reapply layout."""
    if source_hwnd not in grid_state:
        return

    monitors = get_monitors()
    cx, cy = cursor_pos

    # Find target monitor
    target_mon_idx = 0
    for i, (mx, my, mw, mh) in enumerate(monitors):
        if mx <= cx < mx + mw and my <= cy < my + mh:
            target_mon_idx = i
            break

    mon_x, mon_y, mon_w, mon_h = monitors[target_mon_idx]

    # Get windows currently on this monitor
    wins_on_mon = [h for h, (m, _, _) in grid_state.items() if m == target_mon_idx and user32.IsWindow(h) and h != source_hwnd]
    count_including_source = len(wins_on_mon) + 1  # +1 for the window being dropped
    current_layout, current_info = choose_layout(count_including_source)

    # Use the actual layout (critical during drag-and-drop operations)
    target_col = target_row = 0

    if current_layout == "master_stack":
        # Master pane = 3/5 of width
        master_width = (mon_w - 2*EDGE_PADDING - GAP) * 3 // 5
        master_right = mon_x + EDGE_PADDING + master_width + GAP//2

        if cx < master_right:
            target_col, target_row = 0, 0
        else:
            # Stack (right)
            mid_y = mon_y + mon_h // 2
            target_col = 1
            target_row = 0 if cy < mid_y else 1

    elif current_layout == "side_by_side":
        target_col = 0 if cx < mon_x + mon_w // 2 else 1
        target_row = 0

    elif current_layout == "full":
        target_col, target_row = 0, 0

    else:  # grid → extend dynamically to preserve existing window positions
        cols, rows = current_info if current_info else (2, 2)
        max_c = max((c for h,(m,c,r) in grid_state.items() if m == target_mon_idx), default=0)
        max_r = max((r for h,(m,c,r) in grid_state.items() if m == target_mon_idx), default=0)
        cols = max(cols, max_c + 1)
        rows = max(rows, max_r + 1)

        cell_w = (mon_w - 2*EDGE_PADDING - GAP*(cols-1)) // cols
        cell_h = (mon_h - 2*EDGE_PADDING - GAP*(rows-1)) // rows

        rel_x = cx - mon_x - EDGE_PADDING
        rel_y = cy - mon_y - EDGE_PADDING
        target_col = min(max(0, rel_x // (cell_w + GAP)), cols - 1)
        target_row = min(max(0, rel_y // (cell_h + GAP)), rows - 1)

    old_pos = grid_state[source_hwnd]
    new_pos = (target_mon_idx, target_col, target_row)

    if old_pos == new_pos:
        # Force re-apply so window returns perfectly to its cell (magnetic return)
        apply_grid_state()
        return

    # Find if the target cell is occupied
    target_hwnd = None
    for h, pos in grid_state.items():
        if pos == new_pos and h != source_hwnd and user32.IsWindow(h):
            target_hwnd = h
            break

    if target_hwnd:
        log(f"[SNAP] SWAP with '{win32gui.GetWindowText(target_hwnd)[:40]}'")
        grid_state[source_hwnd] = new_pos
        grid_state[target_hwnd] = old_pos
    else:
        log(f"[SNAP] MOVE to cell ({target_col},{target_row}) on monitor {target_mon_idx+1}")
        grid_state[source_hwnd] = new_pos

    apply_grid_state()

# ==============================================================================
# WORKSPACES
# ==============================================================================

def init_workspaces():
    """Initialize per-monitor workspace containers (3 workspaces each)."""
    monitors = get_monitors()
    ws = {}
    current = {}
    for i in range(len(monitors)):
        ws[i] = [{}, {}, {}]
        current[i] = 0
    return ws, current

def save_workspace(monitor_idx):
    """Save current workspace (positions + min/max states)."""
    global workspaces, grid_state

    if monitor_idx not in workspaces:
        return
    
    ws = current_workspace[monitor_idx]
    workspaces[monitor_idx][ws] = {}

    # Save normal windows
    for hwnd, (mon, col, row) in grid_state.items():
        if mon != monitor_idx or not user32.IsWindow(hwnd):
            continue

        rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            continue

        lb, tb, rb, bb = get_frame_borders(hwnd)
        x = rect.left + lb
        y = rect.top + tb
        w = rect.right - rect.left - lb - rb
        h = rect.bottom - rect.top - tb - bb

        workspaces[monitor_idx][ws][hwnd] = {
            'pos': (x, y, w, h),
            'grid': (col, row),
            'state': 'normal'
        }
    
    # Save minimized windows
    for hwnd, (mon, col, row) in minimized_windows.items():
        if mon == monitor_idx and user32.IsWindow(hwnd):
            workspaces[monitor_idx][ws][hwnd] = {
                'pos': (0, 0, 800, 600),  # Dummy pos
                'grid': (col, row),
                'state': 'minimized'
            }
    
    # Save maximized windows
    for hwnd, (mon, col, row) in maximized_windows.items():
        if mon == monitor_idx and user32.IsWindow(hwnd):
            workspaces[monitor_idx][ws][hwnd] = {
                'pos': (0, 0, 800, 600),  # Dummy pos
                'grid': (col, row),
                'state': 'maximized'
            }

    log(f"[WS] ✓ Workspace {ws+1} saved ({len(workspaces[monitor_idx][ws])} windows)")

def load_workspace(monitor_idx, ws_idx):
    """Load workspace and restore window positions intelligently."""
    global grid_state, workspaces
    global minimized_windows, maximized_windows

    if monitor_idx not in workspaces or ws_idx >= len(workspaces[monitor_idx]):
        return

    layout = workspaces[monitor_idx][ws_idx]
    
    if not layout:
        log(f"[WS] Workspace {ws_idx+1} is empty")
        return
    
    log(f"[WS] Loading workspace {ws_idx+1}...")
    
    ignore_retile_until = time.time() + 2.0

    for hwnd, data in layout.items():
        if not user32.IsWindow(hwnd):
            continue

        x, y, w, h = data['pos']
        col, row = data['grid']
        saved_state = data.get('state', 'normal')

        if saved_state == 'minimized':
            user32.ShowWindowAsync(hwnd, win32con.SW_MINIMIZE)
            minimized_windows[hwnd] = (monitor_idx, col, row)
        elif saved_state == 'maximized':
            user32.ShowWindowAsync(hwnd, win32con.SW_MAXIMIZE)
            maximized_windows[hwnd] = (monitor_idx, col, row)
        else:
            if win32gui.IsIconic(hwnd):
                user32.ShowWindowAsync(hwnd, SW_RESTORE)
                time.sleep(0.08)
            if not user32.IsWindowVisible(hwnd):
                user32.ShowWindowAsync(hwnd, SW_SHOWNORMAL)
                time.sleep(0.08)
            grid_state[hwnd] = (monitor_idx, col, row)
        
        time.sleep(0.015)

    time.sleep(0.15)
    smart_tile_with_restore()
    
    log(f"[WS] ✓ Workspace {ws_idx+1} restored")
    ignore_retile_until = time.time() + 2.0

def ws_switch(ws_idx):
    """Switch to specified workspace (0-2) on current monitor. Saves old, hides its windows, loads new."""
    global current_workspace, grid_state, is_active, last_visible_count, workspace_switching_lock
    
    workspace_switching_lock = True # Prevent automatic re-tiling during workspace switch
    time.sleep(0.25)
    try:
        mon = CURRENT_MONITOR_INDEX

        if mon not in workspaces:
            log(f"[WS] ✗ Monitor {mon} not initialized")
            return

        if ws_idx == current_workspace.get(mon, 0):
            log(f"[WS] Already on workspace {ws_idx+1}")
            return

        # FREEZE AUTO-RETILE during workspace switch
        was_active = is_active
        is_active = False
        
        # 1) Save current workspace
        save_workspace(mon)

        # 2) Hide windows from current workspace (smooth - no minimize animation)
        hidden = 0
        for hwnd, (mon_idx, col, row) in list(grid_state.items()):
            if mon_idx == mon and user32.IsWindow(hwnd):
                user32.ShowWindowAsync(hwnd, win32con.SW_HIDE)
                del grid_state[hwnd]
                hidden += 1
        
        log(f"[WS] Hidden {hidden} windows from workspace {current_workspace.get(mon, 0)+1}")

        # 3) Update current workspace index
        current_workspace[mon] = ws_idx

        # 4) Load new workspace
        time.sleep(0.1)
        load_workspace(mon, ws_idx)
        
        # UPDATE last_visible_count to prevent immediate re-tile
        visible_windows = get_visible_windows()
        last_visible_count = len(visible_windows)
        
        # RESTORE AUTO-RETILE state
        time.sleep(0.2)  # Let windows settle
        is_active = was_active
        
        log(f"[WS] ✓ Switched to workspace {ws_idx+1}")
    finally:
        workspace_switching_lock = False

def assign_grid_position(grid_dict, hwnd, monitor_idx, layout, info, index):
    """Assign true grid coordinates (col, row) based on layout (critical for correct snapping/swapping)."""
    if layout == "full":
        grid_dict[hwnd] = (monitor_idx, 0, 0)          # Full-screen → single cell
    elif layout == "side_by_side":
        col = 0 if index == 0 else 1
        grid_dict[hwnd] = (monitor_idx, col, 0)        # Two equal columns side by side
    elif layout == "master_stack":
        if index == 0:
            grid_dict[hwnd] = (monitor_idx, 0, 0)      # Master pane (left, full heigh)
        elif index == 1:
            grid_dict[hwnd] = (monitor_idx, 1, 0)      # Stack pane → top-right
        elif index == 2:
            grid_dict[hwnd] = (monitor_idx, 1, 1)      # Stack pane → bottom-right
    else:  # grid
        cols, rows = info
        grid_dict[hwnd] = (monitor_idx, index % cols, index // cols) # Classic grid layout

def move_current_workspace_to_next_monitor():
    """Move all windows from current workspace to the next monitor (circular). Keeps layout and workspace index."""
    global grid_state, CURRENT_MONITOR_INDEX, MONITORS_CACHE, current_workspace, move_monitor_lock
    
    move_monitor_lock = True # Prevent automatic re-tiling during move
    time.sleep(0.25)
    try:
        if not grid_state:
            log("[MOVE] Nothing to move - no windows in grid")
            return
        if len(MONITORS_CACHE) <= 1:
            log("[MOVE] Only one monitor detected")
            return

        # Save current workspace before moving
        old_mon = CURRENT_MONITOR_INDEX
        old_ws = current_workspace.get(old_mon, 0)
        save_workspace(old_mon)

        # Calculate target monitor
        new_mon = (CURRENT_MONITOR_INDEX + 1) % len(MONITORS_CACHE)
        CURRENT_MONITOR_INDEX = new_mon
        
        mon = MONITORS_CACHE[new_mon]
        mon_x, mon_y, mon_w, mon_h = mon
        
        # Get windows from current workspace on OLD monitor only
        windows_to_move = [(hwnd, pos) for hwnd, pos in grid_state.items() 
                        if pos[0] == old_mon and user32.IsWindow(hwnd)]
        
        if not windows_to_move:
            log(f"[MOVE] No windows to move from monitor {old_mon+1}")
            return
        
        count = len(windows_to_move)
        layout, info = choose_layout(count)

        time.sleep(0.05)

        # Calculate new positions on target monitor
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

        # Move windows to new monitor
        new_grid = {}
        log(f"\n[MOVE] Moving workspace {old_ws+1} from monitor {old_mon+1} → {new_mon+1}")
        for i, (hwnd, (old_idx, col, row)) in enumerate(windows_to_move):
            x, y, w, h = positions[i]
            force_tile_resizable(hwnd, x, y, w, h)
            assign_grid_position(new_grid, hwnd, new_mon, layout, info, i)
            title = win32gui.GetWindowText(hwnd)
            log(f"   -> {title[:50]}")
            time.sleep(0.015)

        # Update grid_state (remove old entries, add new ones)
        for hwnd, _ in windows_to_move:
            grid_state.pop(hwnd, None)
        grid_state.update(new_grid)

        time.sleep(0.15)
        active = user32.GetForegroundWindow()
        if active and user32.IsWindowVisible(active):
            apply_border(active)
        
        # Save to workspace on new monitor (same workspace number)
        current_workspace[new_mon] = old_ws
        save_workspace(new_mon)
        
        log(f"[MOVE] ✓ Workspace {old_ws+1} now on monitor {new_mon+1}")
    finally:
        move_monitor_lock = False 

# ==============================================================================
# HOTKEYS & SYSTRAY
# ==============================================================================

def register_hotkeys():
    """Register global hotkeys: Ctrl+Alt+T/R/Q/M/S for main features""" 
    user32.RegisterHotKey(None, HOTKEY_TOGGLE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('T'))
    user32.RegisterHotKey(None, HOTKEY_RETILE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('R'))
    user32.RegisterHotKey(None, HOTKEY_QUIT,   win32con.MOD_CONTROL | win32con.MOD_ALT, ord('Q'))
    user32.RegisterHotKey(None, HOTKEY_MOVE_MONITOR, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('M'))
    user32.RegisterHotKey(None, HOTKEY_SWAP_MODE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('S'))
    user32.RegisterHotKey(None, HOTKEY_WS1, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('1'))
    user32.RegisterHotKey(None, HOTKEY_WS2, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('2'))
    user32.RegisterHotKey(None, HOTKEY_WS3, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('3'))
    user32.RegisterHotKey(None, HOTKEY_FLOAT_TOGGLE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('F')) 

def unregister_hotkeys():
    """Unregister all main global hotkeys (cleanup on exit)"""
    for hk in (HOTKEY_TOGGLE, HOTKEY_RETILE, HOTKEY_QUIT, HOTKEY_MOVE_MONITOR,
               HOTKEY_SWAP_MODE, HOTKEY_WS1, HOTKEY_WS2, HOTKEY_WS3, HOTKEY_FLOAT_TOGGLE):
        try: user32.UnregisterHotKey(None, hk)
        except: pass

def register_swap_hotkeys():
    """Register arrow keys + Enter for navigation during swap mode"""   
    user32.RegisterHotKey(None, HOTKEY_SWAP_LEFT,    0, win32con.VK_LEFT)
    user32.RegisterHotKey(None, HOTKEY_SWAP_RIGHT,   0, win32con.VK_RIGHT)
    user32.RegisterHotKey(None, HOTKEY_SWAP_UP,      0, win32con.VK_UP)
    user32.RegisterHotKey(None, HOTKEY_SWAP_DOWN,     0, win32con.VK_DOWN)
    user32.RegisterHotKey(None, HOTKEY_SWAP_CONFIRM, 0, win32con.VK_RETURN)

def unregister_swap_hotkeys():
    """Unregister swap-mode hotkeys when leaving the mode"""
    for id in (HOTKEY_SWAP_LEFT, HOTKEY_SWAP_RIGHT, HOTKEY_SWAP_UP, HOTKEY_SWAP_DOWN, HOTKEY_SWAP_CONFIRM):
        try:
            user32.UnregisterHotKey(None, id)
        except:
            pass

def create_icon_image():
    """Create SmartGrid icon (green square with white grid)"""
    width = 64
    height = 64
    # Green background
    image = Image.new('RGB', (width, height), (0, 180, 0))
    dc = ImageDraw.Draw(image)
    
    # White border
    dc.rectangle((4, 4, width-4, height-4), outline=(255, 255, 255), width=4)
    
    # Grid pattern (2x2)
    mid_x = width // 2
    mid_y = height // 2
    dc.line((mid_x, 12, mid_x, height-12), fill=(255, 255, 255), width=3)
    dc.line((12, mid_y, width-12, mid_y), fill=(255, 255, 255), width=3)
    
    return image

def update_tray_menu():
    """Refresh system tray menu with correct checkmarks (tiling state, swap mode)."""
    global tray_icon
    if tray_icon:
        tray_icon.menu = create_tray_menu()
        tray_icon.update_menu()

def toggle_persistent():
    """Toggle persistent auto-tiling mode on/off"""
    global is_active, monitor_thread
    is_active = not is_active
    log(f"\n[SMARTGRID] Persistent mode: {'ON' if is_active else 'OFF'}")
    if is_active:
        grid_state.clear()
        smart_tile_with_restore()
    else:
        for hwnd in list(grid_state.keys()):
            if user32.IsWindow(hwnd):
                set_window_border(hwnd, None)
        grid_state.clear()
    # Refresh tray menu
    update_tray_menu()

def create_tray_menu():
    """Create the context menu structure"""
    return Menu(
        MenuItem(
            f"Tiling: {'ON' if is_active else 'OFF'}", 
            lambda: (toggle_persistent(), update_tray_menu()),
            checked=lambda item: is_active
        ),
        MenuItem('Force Re-tile All Windows (Ctrl+Alt+R)', 
         lambda: threading.Thread(target=force_immediate_retile, daemon=True).start()),
        MenuItem(
            f"Swap Mode: {'ON' if swap_mode_lock else 'OFF'} (Ctrl+Alt+S)",
            lambda: user32.PostThreadMessageW(main_thread_id, CUSTOM_TOGGLE_SWAP, 0, 0),
            checked=lambda item: swap_mode_lock
        ),
        MenuItem('Move Workspace to Next Monitor (Ctrl+Alt+M)', lambda: threading.Thread(target=move_current_workspace_to_next_monitor, daemon=True).start()),
        MenuItem('Toggle Floating Selected Window (Ctrl+Alt+F)',
            lambda: threading.Thread(target=toggle_floating_selected, daemon=True).start()
        ),
        Menu.SEPARATOR,
        MenuItem('Workspaces', Menu(
            MenuItem('Switch to Workspace 1 (Ctrl+Alt+1)', lambda: ws_switch(0)),
            MenuItem('Switch to Workspace 2 (Ctrl+Alt+2)', lambda: ws_switch(1)),
            MenuItem('Switch to Workspace 3 (Ctrl+Alt+3)', lambda: ws_switch(2)),
        )),
        Menu.SEPARATOR,
        MenuItem('Settings (Gap & Padding)', 
         lambda: threading.Thread(target=show_settings_dialog, daemon=True).start()),
        MenuItem('Hotkeys Cheatsheet', lambda: threading.Thread(target=show_hotkeys_tooltip, daemon=True).start()),
        MenuItem('Quit SmartGrid (Ctrl+Alt+Q)', on_quit_from_tray)
    )

def on_quit_from_tray(icon, item):
    """Quit from systray menu"""
    global tray_icon
    log("[TRAY] Quit requested from systray")
    
    # Stop the tray icon first
    if tray_icon:
        tray_icon.stop()
    
    # Clean up borders
    if current_hwnd:
        set_window_border(current_hwnd, None)
    
    if swap_mode_lock:
        exit_swap_mode()

    # Unregister hotkeys
    unregister_hotkeys()
    
    # Force exit the entire process
    log("[EXIT] SmartGrid stopped.")
    os._exit(0)  # Force exit (kills all threads)

def show_hotkeys_tooltip():
    """Show a notification with hotkeys"""
    ctypes.windll.user32.MessageBoxW(
        0,
        "---------- MAIN HOTKEYS\n\n"
        "Ctrl+Alt+T       →  Toggle tiling on/off\n"
        "Ctrl+Alt+R       →  Force re-tile all windows now\n"
        "Ctrl+Alt+S       →  Enter Swap Mode (red border + arrows)\n"
        "Ctrl+Alt+M      →  Move workspace to next monitor\n"
        "Ctrl+Alt+F       →  Toggle Floating Selected Window\n\n"
        "---------- WORKSPACES\n\n"
        "Ctrl+Alt+1/2/3   →  Switch workspace\n\n"
        "---------- EXIT\n\n"
        "Ctrl+Alt+Q       →  Quit",
        "SmartGrid Hotkeys",
        0x40  # MB_ICONINFORMATION
    )

def show_settings_dialog():
    """Show dialog to modify GAP and EDGE_PADDING with dropdown menus"""
    global GAP, EDGE_PADDING, last_visible_count
    
    import tkinter as tk
    from tkinter import ttk
    
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    
    dialog = tk.Toplevel(root)
    dialog.title("SmartGrid Settings")
    dialog.geometry("350x250")
    dialog.attributes('-topmost', True)
    dialog.resizable(False, False)
    
    # Title
    title = tk.Label(dialog, text="SmartGrid Settings", font=("Arial", 12, "bold"))
    title.pack(pady=10)
    
    # GAP Section
    gap_frame = tk.LabelFrame(dialog, text="Gap Between Windows", padx=10, pady=10)
    gap_frame.pack(fill=tk.X, padx=15, pady=5)
    
    gap_options = ["2px", "4px", "6px", "8px", "10px", "12px", "16px", "20px"]
    gap_values = [2, 4, 6, 8, 10, 12, 16, 20]
    current_gap_idx = gap_values.index(GAP) if GAP in gap_values else 3
    
    gap_var = tk.StringVar(value=gap_options[current_gap_idx])
    gap_dropdown = ttk.Combobox(gap_frame, textvariable=gap_var, values=gap_options, 
                                 state="readonly", width=15)
    gap_dropdown.pack(side=tk.LEFT, padx=5)
    
    gap_label = tk.Label(gap_frame, text=f"(current: {GAP}px)", font=("Arial", 9, "italic"))
    gap_label.pack(side=tk.LEFT, padx=10)
    
    # EDGE_PADDING Section
    padding_frame = tk.LabelFrame(dialog, text="Edge Padding (margin)", padx=10, pady=10)
    padding_frame.pack(fill=tk.X, padx=15, pady=5)
    
    padding_options = ["0px", "4px", "8px", "12px", "16px", "20px", "30px", "40px", "50px", "80px", "100px"]
    padding_values = [0, 4, 8, 12, 16, 20, 30, 40, 50, 80, 100]
    current_padding_idx = padding_values.index(EDGE_PADDING) if EDGE_PADDING in padding_values else 2
    
    padding_var = tk.StringVar(value=padding_options[current_padding_idx])
    padding_dropdown = ttk.Combobox(padding_frame, textvariable=padding_var, 
                                     values=padding_options, state="readonly", width=15)
    padding_dropdown.pack(side=tk.LEFT, padx=5)
    
    padding_label = tk.Label(padding_frame, text=f"(current: {EDGE_PADDING}px)", font=("Arial", 9, "italic"))
    padding_label.pack(side=tk.LEFT, padx=10)
    
    # Buttons
    button_frame = tk.Frame(dialog)
    button_frame.pack(pady=15)
    
    def apply_and_close():
        global GAP, EDGE_PADDING
        
        gap_str = gap_var.get().replace("px", "")
        padding_str = padding_var.get().replace("px", "")
        
        # stop preview if active
        hide_snap_preview()
        time.sleep(0.15)
        
        # apply new settings
        GAP = int(gap_str)
        EDGE_PADDING = int(padding_str)
        
        log(f"[SETTINGS] GAP={GAP}px, EDGE_PADDING={EDGE_PADDING}px")

        apply_new_settings()
        
        dialog.destroy()
        root.destroy()
        
    def cancel_and_close():
        dialog.destroy()
        root.destroy()
    
    def reset_defaults():
        nonlocal current_gap_idx, current_padding_idx
        gap_var.set(gap_options[3])  # 8px
        padding_var.set(padding_options[2])  # 8px
    
    tk.Button(button_frame, text="Apply", command=apply_and_close, width=12, bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Reset", command=reset_defaults, width=12).pack(side=tk.LEFT, padx=5)
    tk.Button(button_frame, text="Cancel", command=cancel_and_close, width=12).pack(side=tk.LEFT, padx=5)
    
    info = tk.Label(dialog, text="Changes apply immediately on next retile cycle", 
                    font=("Arial", 8), fg="gray")
    info.pack(pady=5)
    
    root.mainloop()

def apply_new_settings():
    global ignore_retile_until, last_visible_count, drag_drop_lock
    
    drag_drop_lock = True
    time.sleep(0.1)

    ignore_retile_until = 0
    last_visible_count = 0
    
    smart_tile_with_restore()
    apply_grid_state()

    time.sleep(0.1)
    drag_drop_lock = False


# ==============================================================================
# MAIN LOOP & THREADS
# ==============================================================================

def monitor():
    """Background loop: auto-retile + border tracking"""
    global current_hwnd, grid_state, last_visible_count, last_active_hwnd
    global ignore_retile_until, swap_mode_lock, drag_drop_lock
    global move_monitor_lock, workspace_switching_lock
    global user_selected_hwnd

    while True:

        # === TRACK USER'S CURRENT WINDOW SELECTION ===
        fg = user32.GetForegroundWindow()
        if fg and user32.IsWindow(fg) and user32.IsWindowVisible(fg):
            title = win32gui.GetWindowText(fg)
            class_name = win32gui.GetClassName(fg)
            if is_useful_window(title, class_name):
                user_selected_hwnd = fg

        update_active_border()
        
        if (swap_mode_lock or drag_drop_lock or move_monitor_lock or workspace_switching_lock):
            time.sleep(0.1)
            continue

        if is_active:
            # === LIGHTWEIGHT CLEANUP - Only remove DEAD windows (not ghosts) ===
            # Ghost detection is heavy, do it in smart_tile_with_restore() instead
            dead_count = 0
            for hwnd in list(grid_state.keys()):
                # Fast check: IsWindow() is cheap
                if not user32.IsWindow(hwnd):
                    grid_state.pop(hwnd, None)
                    dead_count += 1
                    continue

                # Fast check: window state
                state = get_window_state(hwnd)
                if state in ("minimized", "maximized"):
                    grid_state.pop(hwnd, None)
                    dead_count += 1
                    continue
                
                # Fast check: visible
                if not user32.IsWindowVisible(hwnd):
                    grid_state.pop(hwnd, None)
                    dead_count += 1
                    continue

            if dead_count > 0:
                log(f"[MONITOR] Cleaned {dead_count} dead windows")

            visible_windows = get_visible_windows()
            current_count = len(visible_windows)

            # Auto-retile with smart restore
            if time.time() >= ignore_retile_until and current_count > 0:
                if current_count != last_visible_count:
                    log(f"[AUTO-RETILE] {last_visible_count} → {current_count} windows")
                    smart_tile_with_restore()  # ← Ghost detection happens HERE
                    last_visible_count = current_count
                time.sleep(0.2)

        time.sleep(0.06)

def message_loop():
    """Main message loop for hotkeys and window messages"""
    msg = wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        if msg.message == win32con.WM_HOTKEY:
            if msg.wParam == HOTKEY_TOGGLE:
                toggle_persistent()
            elif msg.wParam == HOTKEY_RETILE:
                threading.Thread(target=force_immediate_retile, daemon=True).start()
            elif msg.wParam == HOTKEY_MOVE_MONITOR:
                threading.Thread(target=move_current_workspace_to_next_monitor, daemon=True).start()
            elif msg.wParam == HOTKEY_QUIT:
                break
            elif msg.wParam == HOTKEY_SWAP_MODE:
                if swap_mode_lock:
                    exit_swap_mode()
                else:
                    enter_swap_mode()
            elif swap_mode_lock:
                if msg.wParam == HOTKEY_SWAP_LEFT:
                    navigate_swap("left")
                elif msg.wParam == HOTKEY_SWAP_RIGHT:
                    navigate_swap("right")
                elif msg.wParam == HOTKEY_SWAP_UP:
                    navigate_swap("up")
                elif msg.wParam == HOTKEY_SWAP_DOWN:
                    navigate_swap("down")
                elif msg.wParam == HOTKEY_SWAP_CONFIRM:
                    exit_swap_mode()
            elif msg.wParam == HOTKEY_WS1:
                ws_switch(0)
            elif msg.wParam == HOTKEY_WS2:
                ws_switch(1)
            elif msg.wParam == HOTKEY_WS3:
                ws_switch(2)
            elif msg.wParam == HOTKEY_FLOAT_TOGGLE:
                toggle_floating_selected()
        elif msg.message == CUSTOM_TOGGLE_SWAP:
            if swap_mode_lock:
                exit_swap_mode()
            else:
                enter_swap_mode()

        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

if __name__ == "__main__":
    print("="*70)
    print(" SMARTGRID — Advanced Windows Tiling Manager")
    print("="*70)
    print("Ctrl+Alt+T     → Toggle tiling on/off")
    print("Ctrl+Alt+R     → Force re-tile all windows now")
    print("Ctrl+Alt+S     → Enter Swap Mode (red border + arrows)")
    print("Ctrl+Alt+M     → Move workspace to next monitor")
    print("Ctrl+Alt+F     → Toggle Floating Selected Window")
    print("Ctrl+Alt+1/2/3 → Switch workspace")
    print("Ctrl+Alt+Q     → Quit")
    print("-"*70)

    time.sleep(0.05)

    # Init
    monitors = get_monitors()
    log(f"Detected {len(monitors)} monitor(s)")
    workspaces, current_workspace = init_workspaces()
    current_workspace = {i: 0 for i in range(len(monitors))}
    log(f"Initialized {len(monitors)} × 3 workspaces")

    if len(monitors) > 1:
        log("Ctrl+Alt+M -> cycle tiled windows across monitors")

    # Background Threads
    threading.Thread(target=monitor, daemon=True).start()
    threading.Thread(target=start_drag_snap_monitor, daemon=True).start()

    # === REGISTER HOTKEYS ===
    register_hotkeys()
    log("[MAIN] All hotkeys registered (including arrows)")

    # === CREATE SYSTRAY ICON ===
    tray_icon = Icon(
        "SmartGrid",
        create_icon_image(),
        "SmartGrid - Tiling Window Manager",
        menu=create_tray_menu()
    )
    update_tray_menu()

    # === Give pystray a setup function that does nothing ===
    def setup(icon):
        icon.visible = True

    # === Launch pystray in the main thread, BUT without blocking ===
    tray_icon.run_detached(setup=setup)
    log("[MAIN] Systray launched with run_detached")

    # === OUR message_loop() takes over and receives ALL WM_HOTKEY ===
    try:
        message_loop()
    finally:
        unregister_hotkeys()
        if tray_icon:
            tray_icon.stop()
        if current_hwnd:
            set_window_border(current_hwnd, None)
        log("[EXIT] SmartGrid stopped.")
        os._exit(0)