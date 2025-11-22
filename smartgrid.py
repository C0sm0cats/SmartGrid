# pip install pywin32

import ctypes
import time
import threading
from ctypes import wintypes
import win32con
import win32gui
import win32api

# ==============================================================================
# Constants & config
# ==============================================================================
# Win32 API
user32 = ctypes.WinDLL('user32', use_last_error=True)   # Main Windows API
dwmapi = ctypes.WinDLL('dwmapi')  # Desktop Window Manager (for colored borders)

# Cached monitor work areas
MONITORS_CACHE = []

# Layout appearance
GAP          = 8      # Space between tiled windows
EDGE_PADDING = 8      # Margin from screen edges

# DWM (Desktop Window Manager) attributes
DWMWA_BORDER_COLOR          = 34
DWMWA_COLOR_NONE            = 0xFFFFFFFF
DWMWA_EXTENDED_FRAME_BOUNDS = 9

# Win32 window styles & flags
GWL_STYLE = -16
WS_THICKFRAME = 0x00040000
WS_MAXIMIZE = 0x01000000

# ShowWindow commands
SW_RESTORE    = 9    # Restores a minimized or maximized window
SW_SHOWNORMAL = 1    # Activates and displays window (or restores if minimized)

# SetWindowPos flags (used in force_tile_resizable)
SWP_NOZORDER      = 0x0004  # Ignores Z-order
SWP_NOMOVE        = 0x0002  # Don't change position
SWP_NOSIZE        = 0x0001  # Don't change size
SWP_NOACTIVATE    = 0x0010  # Don't activate the window
SWP_FRAMECHANGED  = 0x0020  # Force WM_NCCALCSIZE recalculation (border refresh)
SWP_NOSENDCHANGING = 0x0400 # Prevent WM_WINDOWPOSCHANGING (avoids conflicts)

# Hotkey identifiers
HOTKEY_TOGGLE = 9001
HOTKEY_RETILE = 9002
HOTKEY_QUIT = 9003
HOTKEY_MOVE_MONITOR = 9004
HOTKEY_SWAP_MODE = 9005
HOTKEY_SWAP_LEFT = 9006
HOTKEY_SWAP_RIGHT = 9007
HOTKEY_SWAP_UP = 9008
HOTKEY_SWAP_DOWN = 9009
HOTKEY_SWAP_CONFIRM = 9010
HOTKEY_WS1 = 9101
HOTKEY_WS2 = 9102
HOTKEY_WS3 = 9103

# ==============================================================================
# Application State
# ==============================================================================
# Window tracking
current_hwnd         = None   # Window with green border (active)
selected_hwnd        = None   # Window with red border (swap mode)
grid_state          = {}      # hwnd → (monitor_idx, col, row)

# Monitor state
CURRENT_MONITOR_INDEX = 0     # Target monitor when cycling with Ctrl+Alt+M

# Runtime flags
last_visible_count  = 0
is_active           = False   # Persistent tiling enabled?
swap_mode           = False   # Swap mode active?

# Background threads
monitor_thread        = None  # Thread running the main auto-retile + border monitor

# ==============================================================================
# Workspace System
# ==============================================================================
def init_workspaces():
    """Initialize workspace structure based on detected monitors"""
    monitors = get_monitors()
    ws = {}
    for i in range(len(monitors)):
        ws[i] = [{}, {}, {}]  # 3 workspaces per monitor
    return ws

workspaces = {}  # Will be initialized after monitor detection
current_workspace = {}  # monitor_index → current ws index
# ==============================================================================

def remove_border(hwnd):
    """Remove custom DWM border – restores default window frame"""
    if hwnd:
        try:
            dwmapi.DwmSetWindowAttribute(
                hwnd, DWMWA_BORDER_COLOR,
                ctypes.byref(ctypes.c_uint(DWMWA_COLOR_NONE)),
                ctypes.sizeof(ctypes.c_uint)
            )
        except:
            pass

def apply_border(hwnd, color=0x0000FF00):
    """Apply colored DWM border – green = active, red = swap mode"""
    global current_hwnd
    if current_hwnd and current_hwnd != hwnd:
        remove_border(current_hwnd)
    if hwnd:
        color_val = ctypes.c_uint(color)
        dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_BORDER_COLOR,
            ctypes.byref(color_val), ctypes.sizeof(ctypes.c_uint)
        )
        current_hwnd = hwnd

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

def get_monitors():
    """Return cached list of work-area rectangles for all monitors"""
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
    """Filter out overlays, PIPs, taskbar, notifications, etc."""
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
        # === Added to exclude Win+Tab / Alt+Tab task switcher ===
        "multitaskingviewframe",           # Win+Tab
        "taskswitcherwnd",                 # Alt+Tab
        "xamlexplorerhostislandwindow",    # Alt+Tab / Win11 timeline
    ]

    if class_lower in bad_classes:
        return False

    return True

def get_visible_windows():
    """Enumerate all visible, non-minimized, useful windows"""
    global overlay_hwnd  # ✅ Ajoute cette ligne en haut
    
    monitors = get_monitors()
    windows = []
    def enum(hwnd, _):
        if user32.IsWindowVisible(hwnd):
            rect = wintypes.RECT()
            if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                if w > 180 and h > 180:
                    title_buf = ctypes.create_unicode_buffer(256)
                    user32.GetWindowTextW(hwnd, title_buf, 256)
                    title = title_buf.value or ""
                    class_name = win32gui.GetClassName(hwnd)

                    # ✅ SKIP OVERLAY WINDOW
                    if overlay_hwnd and hwnd == overlay_hwnd:
                        return True
                    
                    if is_useful_window(title, class_name):
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

def force_tile_resizable(hwnd, x, y, w, h):
    """Move and resize window with pixel-perfect accuracy (border/shadow aware)"""
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

# ==============================================================================
# Smart layout chooser
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

def clear_all_borders():
    """Remove custom colored borders from all windows (cleanup utility)"""
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

# ==============================================================================
# smart_tile with intelligent layouts + your grid fallback
# ==============================================================================
def assign_grid_position(grid_dict, hwnd, monitor_idx, layout, info, index):
    """
    Assigns true grid coordinates (column, row) to a window.
    Critical for perfect Drag & Drop Snap behavior — especially in master/stack layouts.
    """
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

def smart_tile(temp=False):
    """Full retile: detect windows → choose layout → assign grid positions"""
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

        assign_grid_position(new_grid, hwnd, 0, layout, info, i)

    grid_state = new_grid

    time.sleep(0.15)
    active = user32.GetForegroundWindow()
    if active and user32.IsWindowVisible(active):
        apply_border(active)

# ==============================================================================
# Multi-monitor
# ==============================================================================
def move_current_workspace_to_next_monitor():
    """Move ONLY the current workspace to the next monitor (workspace-aware)"""
    global grid_state, CURRENT_MONITOR_INDEX, MONITORS_CACHE, current_workspace
    
    if not grid_state:
        print("[MOVE] Nothing to move - no windows in grid")
        return
    if len(MONITORS_CACHE) <= 1:
        print("[MOVE] Only one monitor detected")
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
        print(f"[MOVE] No windows to move from monitor {old_mon+1}")
        return
    
    count = len(windows_to_move)
    layout, info = choose_layout(count)

    # Clear borders before move
    clear_all_borders()
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
    print(f"\n[MOVE] Moving workspace {old_ws+1} from monitor {old_mon+1} → {new_mon+1}")
    for i, (hwnd, (old_idx, col, row)) in enumerate(windows_to_move):
        x, y, w, h = positions[i]
        force_tile_resizable(hwnd, x, y, w, h)
        assign_grid_position(new_grid, hwnd, new_mon, layout, info, i)
        title = win32gui.GetWindowText(hwnd)
        print(f"   -> {title[:50]}")
        time.sleep(0.03)

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
    
    print(f"[MOVE] ✓ Workspace {old_ws+1} now on monitor {new_mon+1}")

# ==============================================================================
# SWAP MODE
# ==============================================================================
def get_window_position_info(hwnd):
    """Return (center_x, center_y, monitor_index, grid_col, grid_row) for a tiled window"""
    if hwnd not in grid_state:
        return None
    
    rect = wintypes.RECT()
    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        return None
    
    x_center = (rect.left + rect.right) // 2
    y_center = (rect.top + rect.bottom) // 2
    mon_idx, col, row = grid_state[hwnd]
    
    return (x_center, y_center, mon_idx, col, row)

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

        # --- OVERLAP MINIMUM (pour éviter les diagonales) ---
        overlap_x = max(0, min(fx2, x2) - max(fx1, x1))
        overlap_y = max(0, min(fy2, y2) - max(fy1, y1))

        if direction in ("left", "right"):
            if overlap_y < 50: continue   # must be well aligned vertically
        else:
            if overlap_x < 50: continue   # must be well aligned horizontally

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
    remove_border(hwnd1)
    remove_border(hwnd2)
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
    print(f"[SWAP] '{title1}' ↔ '{title2}'")
    
    # Swap the positions in grid_state
    grid_state[hwnd1], grid_state[hwnd2] = grid_state[hwnd2], grid_state[hwnd1]
    
    # Physically move the windows (SWAPPING positions, not sizes)
    force_tile_resizable(hwnd1, x2, y2, w2, h2)
    time.sleep(0.08)
    force_tile_resizable(hwnd2, x1, y1, w1, h1)
    time.sleep(0.08)
    
    # Clear the borders again after the swap
    remove_border(hwnd1)
    remove_border(hwnd2)
    
    return True

def enter_swap_mode():
    """Activate swap mode with red border and arrow-key navigation"""
    global swap_mode, selected_hwnd
    
    # Wait a tiny bit for the tiling to stabilize
    time.sleep(0.15)
    
    # Force a quick update of grid_state in case it's empty or updating
    if not grid_state:
        visible_windows = get_visible_windows()
        for hwnd, title, _ in visible_windows:
            if hwnd not in grid_state:
                grid_state[hwnd] = (0, 0, 0)
        print(f"[SWAP] grid_state was empty → rebuilt with {len(grid_state)} windows")
    if not grid_state:
        print("[SWAP] No tiled windows detected. First press Ctrl+Alt+T or Ctrl+Alt+R")

        return
    
    # Get the active window
    active = user32.GetForegroundWindow()
    if not active or active not in grid_state:
        # Take the first window from grid_state
        active = next(iter(grid_state.keys()))
    
    swap_mode = True
    selected_hwnd = active
    
    remove_border(selected_hwnd)
    time.sleep(0.05)
    color_val = ctypes.c_uint(0x000000FF)  # Bright red (BGR format)
    dwmapi.DwmSetWindowAttribute(
        selected_hwnd, DWMWA_BORDER_COLOR,
        ctypes.byref(color_val), ctypes.sizeof(ctypes.c_uint)
    )
    
    title = win32gui.GetWindowText(selected_hwnd)[:50]
    print(f"\n[SWAP MODE] ✓ Activated - Selected window: '{title}'")
    print(f"[SWAP MODE] Selected hwnd: {selected_hwnd}")
    print("━" * 60)
    print("  DIRECT SWAP with arrow keys:")
    print("    ← → ↑ ↓  : Swap with adjacent window")
    print("    Ctrl+Alt+S : Exit swap mode")
    print("  ")
    print("  The red window FOLLOWS your movements and swaps its position!")
    print("━" * 60)
    register_swap_hotkeys()

def navigate_swap(direction):
    """Handle arrow key press in swap mode → find and swap with adjacent window"""
    global selected_hwnd
    
    if not swap_mode or not selected_hwnd:
        print("[SWAP] Mode not active or no window selected")
        return
    
    print(f"[SWAP] Attempting to swap {direction}...")
    
    # Find the window in the specified direction
    target = find_window_in_direction(selected_hwnd, direction)
    
    if target:
        # DIRECT SWAP!
        if swap_windows(selected_hwnd, target):
            # The selected window (red) has moved, we follow it
            # selected_hwnd logically remains the same, but physically it has changed position
            
            # Clear and reapply the red border on the window that moved
            time.sleep(0.1)
            remove_border(selected_hwnd)
            time.sleep(0.05)
            color_val = ctypes.c_uint(0x000000FF)
            dwmapi.DwmSetWindowAttribute(
                selected_hwnd, DWMWA_BORDER_COLOR,
                ctypes.byref(color_val), ctypes.sizeof(ctypes.c_uint)
            )
            
            title = win32gui.GetWindowText(selected_hwnd)[:50]
            print(f"[SWAP] ✓ '{title}' swapped {direction}")
            
            # Focus on the window that moved
            user32.SetForegroundWindow(selected_hwnd)
        else:
            print(f"[SWAP] ✗ Swap failed")
    else:
        print(f"[SWAP] ✗ No window in the {direction} direction (grid limit)")

def confirm_swap():
    """Confirm current swap selection and exit swap mode (bound to Enter)"""
    global swap_mode, selected_hwnd

    if not swap_mode or not selected_hwnd:
        print("[SWAP] Swap mode not active or no window selected")
        return

    print(f"[SWAP] Confirming swap for hwnd {selected_hwnd}")

    exit_swap_mode()
    print("[SWAP] ✓ Swap confirmed and mode exited")

def exit_swap_mode():
    """Exit swap mode and restore normal green border on active window"""
    global swap_mode, selected_hwnd, current_hwnd
    
    if not swap_mode:
        return
    
    # First, clear ALL borders
    print("[SWAP MODE] Clearing borders...")
    clear_all_borders()
    time.sleep(0.15)
    
    swap_mode = False
    unregister_swap_hotkeys()
    old_selected = selected_hwnd
    selected_hwnd = None
    current_hwnd = None
    
    # Restore the green border on the active window
    active = user32.GetForegroundWindow()
    if active and user32.IsWindowVisible(active) and active in grid_state:
        color_val = ctypes.c_uint(0x0000FF00)  # Vert
        dwmapi.DwmSetWindowAttribute(
            active, DWMWA_BORDER_COLOR,
            ctypes.byref(color_val), ctypes.sizeof(ctypes.c_uint)
        )
        current_hwnd = active
        print(f"[SWAP MODE] Green border restored on active window")
    
    print("[SWAP MODE] ✓ Deactivated\n")

# ==============================================================================
# Polling monitor - Auto-retile when windows are shown/hidden/minimized/restored
# ==============================================================================
def monitor():
    """Background loop: auto-retile on window events + track active window border"""
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
                    grid_state[hwnd] = (0, 0, 0)
                    updated = True

            # Auto-retile only if number of visible windows changed
            if current_count != last_visible_count or updated:
                print(f"[AUTO-RETILE] {last_visible_count} → {current_count} visible windows")
                smart_tile(temp=True)
                last_visible_count = current_count
                time.sleep(0.2)

        # Update border on active window (except in swap mode)
        if not swap_mode:
            active = user32.GetForegroundWindow()
            if active and user32.IsWindowVisible(active):
                if active != current_hwnd:
                    apply_border(active, color=0x0000FF00)  # Green
            else:
                if current_hwnd:
                    remove_border(current_hwnd)
                    current_hwnd = None

        time.sleep(0.35)

# ==============================================================================
# Hotkeys & main loop
# ==============================================================================
def toggle_persistent():
    """Toggle persistent auto-tiling mode on/off"""
    global is_active, monitor_thread
    is_active = not is_active
    print(f"\n[SMARTGRID] Persistent mode: {'ON' if is_active else 'OFF'}")
    if is_active:
        smart_tile(temp=False)

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

def unregister_hotkeys():
    """Unregister all main global hotkeys (cleanup on exit)"""
    for hk in (HOTKEY_TOGGLE, HOTKEY_RETILE, HOTKEY_QUIT, HOTKEY_MOVE_MONITOR,
               HOTKEY_SWAP_MODE, HOTKEY_WS1, HOTKEY_WS2, HOTKEY_WS3):
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

# ==============================================================================
# DRAG & DROP SNAP
# ==============================================================================
def apply_grid_state():
    """Re-apply grid positions — FOR ALL LAYOUTS (full, side_by_side, master_stack, grid)"""
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
                (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, mw, mon_h - 2*EDGE_PADDING),           # master
                (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING, sw, sh),                   # Stack pane → top-right
                (mon_x + EDGE_PADDING + mw + GAP, mon_y + EDGE_PADDING + sh + GAP, sw, sh)        # Stack pane → bottom-right
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
                time.sleep(0.015)
    
    # Apply green border to the active window
    time.sleep(0.08)
    active = user32.GetForegroundWindow()
    if active and user32.IsWindowVisible(active) and active in grid_state:
        apply_border(active)

def is_window_maximized(hwnd):
    """Return True if window is maximized – prevents false drags on internal splits"""
    if not hwnd or not user32.IsWindow(hwnd):
        return False
    style = user32.GetWindowLongW(hwnd, GWL_STYLE)
    return bool(style & WS_MAXIMIZE)

def start_drag_snap_monitor():
    """Background thread: detects title-bar drag → shows preview → snaps on drop"""
    was_down = False
    drag_hwnd = None
    drag_start = None
    preview_active = False

    while True:
        try:
            down = win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000
            
            if down and not was_down:
                hwnd = user32.GetForegroundWindow()
                # Ignore drag if window is maximized (prevents interference with internal splits)
                if hwnd and is_window_maximized(hwnd):
                    was_down = True
                    continue
                if hwnd and hwnd in grid_state and user32.IsWindowVisible(hwnd):
                    drag_hwnd = hwnd
                    drag_start = win32api.GetCursorPos()
                    preview_active = False
            
            elif down and drag_hwnd:
                # Dragging in progress
                cursor_pos = win32api.GetCursorPos()
                if drag_start:
                    dx = abs(cursor_pos[0] - drag_start[0])
                    dy = abs(cursor_pos[1] - drag_start[1])
                    
                    # Show preview after 20px movement
                    if (dx > 20 or dy > 20) and not preview_active:
                        preview_active = True
                    
                    # Update preview during drag
                    if preview_active:
                        target_rect = calculate_target_rect(drag_hwnd, cursor_pos)
                        if target_rect:
                            show_snap_preview(*target_rect)
            
            elif not down and was_down and drag_hwnd:
                # Drop detected
                hide_snap_preview()
                
                if drag_start:
                    dx = abs(win32api.GetCursorPos()[0] - drag_start[0])
                    dy = abs(win32api.GetCursorPos()[1] - drag_start[1])
                    if dx > 20 or dy > 20:
                        handle_snap_drop(drag_hwnd, win32api.GetCursorPos())
                
                drag_hwnd = None
                drag_start = None
                preview_active = False
            
            was_down = down
            time.sleep(0.015)
        except Exception as e:
            hide_snap_preview()
            time.sleep(0.1)

def handle_snap_drop(source_hwnd, cursor_pos):
    """Handle drop: swap or move window to target cell"""
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
        return

    # Find if the target cell is occupied
    target_hwnd = None
    for h, pos in grid_state.items():
        if pos == new_pos and h != source_hwnd and user32.IsWindow(h):
            target_hwnd = h
            break

    if target_hwnd:
        print(f"[SNAP] SWAP with '{win32gui.GetWindowText(target_hwnd)[:40]}'")
        grid_state[source_hwnd] = new_pos
        grid_state[target_hwnd] = old_pos
    else:
        print(f"[SNAP] MOVE to cell ({target_col},{target_row}) on monitor {target_mon_idx+1}")
        grid_state[source_hwnd] = new_pos

    apply_grid_state()

# ==============================================================================
# DRAG & DROP SNAP WITH PREVIEW OVERLAY
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

# Global overlay window
overlay_hwnd = None
preview_rect = None  # (x, y, w, h) of preview rectangle

def create_overlay_window():
    """Create transparent overlay window for snap preview"""
    global overlay_hwnd
    
    if overlay_hwnd:
        return overlay_hwnd
    
    class_name = "SmartGridOverlay"
    
    # ✅ Properly define DefWindowProc with correct signature
    DefWindowProc = ctypes.windll.user32.DefWindowProcW
    DefWindowProc.argtypes = [wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM]
    DefWindowProc.restype = wintypes.LPARAM
    
    # ✅ Define WNDPROCTYPE properly
    WNDPROCTYPE = ctypes.WINFUNCTYPE(
        wintypes.LPARAM,    # Return type (changed from c_long)
        wintypes.HWND,      # hwnd
        ctypes.c_uint,      # msg
        wintypes.WPARAM,    # wparam
        wintypes.LPARAM     # lparam
    )
    
    # Window procedure for overlay
    @WNDPROCTYPE
    def wnd_proc(hwnd, msg, wparam, lparam):
        try:
            if msg == win32con.WM_PAINT:
                ps = PAINTSTRUCT()
                hdc = user32.BeginPaint(hwnd, ctypes.byref(ps))
                
                if preview_rect:
                    x, y, w, h = preview_rect
                    
                    # Create semi-transparent brush (blue with 30% opacity)
                    brush = win32gui.CreateSolidBrush(win32api.RGB(100, 149, 237))  # Cornflower blue
                    pen = win32gui.CreatePen(win32con.PS_SOLID, 4, win32api.RGB(65, 105, 225))  # Royal blue border
                    
                    old_brush = win32gui.SelectObject(hdc, brush)
                    old_pen = win32gui.SelectObject(hdc, pen)
                    
                    # Draw rectangle
                    win32gui.Rectangle(hdc, x, y, x + w, y + h)
                    
                    win32gui.SelectObject(hdc, old_brush)
                    win32gui.SelectObject(hdc, old_pen)
                    win32gui.DeleteObject(brush)
                    win32gui.DeleteObject(pen)
                
                user32.EndPaint(hwnd, ctypes.byref(ps))
                return 0
            
            elif msg == win32con.WM_DESTROY:
                return 0
            
            # ✅ Use properly typed DefWindowProc
            return DefWindowProc(hwnd, msg, wparam, lparam)
        
        except Exception as e:
            print(f"[OVERLAY] Error in wnd_proc: {e}")
            return 0
    
    # Register window class
    wc = win32gui.WNDCLASS()
    wc.lpfnWndProc = wnd_proc
    wc.lpszClassName = class_name
    wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
    
    try:
        win32gui.RegisterClass(wc)
    except:
        pass  # Class already registered
    
    # Create layered window (for transparency)
    overlay_hwnd = win32gui.CreateWindowEx(
        win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT | win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
        class_name,
        "SmartGrid Preview",
        win32con.WS_POPUP,
        0, 0, 1, 1,  # Will be resized
        0, 0, 0, None
    )
    
    # Set 30% opacity
    win32gui.SetLayeredWindowAttributes(overlay_hwnd, 0, int(255 * 0.3), win32con.LWA_ALPHA)
    
    return overlay_hwnd

def show_snap_preview(x, y, w, h):
    """Show preview rectangle at specified position"""
    global preview_rect, overlay_hwnd
    
    if not overlay_hwnd:
        create_overlay_window()
    
    preview_rect = (x, y, w, h)
    
    # Position and show overlay window
    win32gui.SetWindowPos(
        overlay_hwnd, win32con.HWND_TOPMOST,
        x, y, w, h,
        win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE
    )
    
    # Force redraw
    win32gui.InvalidateRect(overlay_hwnd, None, True)
    win32gui.UpdateWindow(overlay_hwnd)

def hide_snap_preview():
    """Hide preview overlay"""
    global preview_rect, overlay_hwnd
    
    preview_rect = None
    
    if overlay_hwnd:
        win32gui.ShowWindow(overlay_hwnd, win32con.SW_HIDE)

def calculate_target_rect(source_hwnd, cursor_pos):
    """Calculate where window will snap (returns rect for preview)"""
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
    
    # Calculate target cell
    target_col = target_row = 0
    
    if current_layout == "master_stack":
        master_width = (mon_w - 2*EDGE_PADDING - GAP) * 3 // 5
        master_right = mon_x + EDGE_PADDING + master_width + GAP//2
        
        if cx < master_right:
            target_col, target_row = 0, 0
            # Master pane dimensions
            return (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING,
                   master_width, mon_h - 2*EDGE_PADDING)
        else:
            # Stack pane
            sw = mon_w - 2*EDGE_PADDING - master_width - GAP
            sh = (mon_h - 2*EDGE_PADDING - GAP) // 2
            mid_y = mon_y + mon_h // 2
            
            if cy < mid_y:
                # Top stack
                return (mon_x + EDGE_PADDING + master_width + GAP, mon_y + EDGE_PADDING,
                       sw, sh)
            else:
                # Bottom stack
                return (mon_x + EDGE_PADDING + master_width + GAP, 
                       mon_y + EDGE_PADDING + sh + GAP, sw, sh)
    
    elif current_layout == "side_by_side":
        cw = (mon_w - 2*EDGE_PADDING - GAP) // 2
        if cx < mon_x + mon_w // 2:
            # Left side
            return (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING, 
                   cw, mon_h - 2*EDGE_PADDING)
        else:
            # Right side
            return (mon_x + EDGE_PADDING + cw + GAP, mon_y + EDGE_PADDING,
                   cw, mon_h - 2*EDGE_PADDING)
    
    elif current_layout == "full":
        return (mon_x + EDGE_PADDING, mon_y + EDGE_PADDING,
               mon_w - 2*EDGE_PADDING, mon_h - 2*EDGE_PADDING)
    
    else:  # grid
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
        
        x = mon_x + EDGE_PADDING + target_col * (cell_w + GAP)
        y = mon_y + EDGE_PADDING + target_row * (cell_h + GAP)
        
        return (x, y, cell_w, cell_h)
    
    return None

# ==============================================================================
# WORKSPACE MANAGEMENT
# ==============================================================================
def save_workspace(monitor_idx):
    """Save current window layout to workspace (with grid positions!)"""
    global workspaces, grid_state

    if monitor_idx not in workspaces:
        return
    
    ws = current_workspace[monitor_idx]
    workspaces[monitor_idx][ws] = {}

    for hwnd, (mon, col, row) in grid_state.items():
        if mon != monitor_idx or not user32.IsWindow(hwnd):
            continue

        rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            continue

        # ✅ Get frame borders to calculate TRUE client area
        lb, tb, rb, bb = get_frame_borders(hwnd)
        
        # ✅ Save position WITHOUT invisible borders (pure client area)
        x = rect.left + lb
        y = rect.top + tb
        w = rect.right - rect.left - lb - rb
        h = rect.bottom - rect.top - tb - bb

        # Save both physical position AND grid coordinates
        workspaces[monitor_idx][ws][hwnd] = {
            'pos': (x, y, w, h),  # ✅ Clean dimensions
            'grid': (col, row)
        }

    print(f"[WS] ✓ Workspace {ws+1} saved on monitor {monitor_idx+1} ({len(workspaces[monitor_idx][ws])} windows)")

def load_workspace(monitor_idx, ws_idx):
    """Restore workspace layout - handles dead/minimized windows gracefully"""
    global grid_state, workspaces

    if monitor_idx not in workspaces or ws_idx >= len(workspaces[monitor_idx]):
        return

    layout = workspaces[monitor_idx][ws_idx]
    
    if not layout:
        print(f"[WS] Workspace {ws_idx+1} is empty")
        return
    
    print(f"[WS] Loading workspace {ws_idx+1} for monitor {monitor_idx+1}...")
    restored = 0

    for hwnd, data in layout.items():
        # Skip dead windows
        if not user32.IsWindow(hwnd):
            continue

        x, y, w, h = data['pos']
        col, row = data['grid']

        # Restore minimized windows
        if win32gui.IsIconic(hwnd):
            user32.ShowWindowAsync(hwnd, SW_RESTORE)
            time.sleep(0.08)

        # Restore hidden windows
        if not user32.IsWindowVisible(hwnd):
            user32.ShowWindowAsync(hwnd, SW_SHOWNORMAL)
            time.sleep(0.08)

        # Reposition window
        force_tile_resizable(hwnd, x, y, w, h)
        
        # Restore grid position
        grid_state[hwnd] = (monitor_idx, col, row)
        
        restored += 1
        time.sleep(0.04)

    time.sleep(0.15)
    
    # Apply green border to active window
    active = user32.GetForegroundWindow()
    if active and user32.IsWindowVisible(active) and active in grid_state:
        apply_border(active)
    
    print(f"[WS] ✓ Restored {restored}/{len(layout)} windows")

def ws_switch(ws_idx):
    """Switch to another workspace (smooth transition)"""
    global current_workspace, grid_state, is_active, last_visible_count
    
    mon = CURRENT_MONITOR_INDEX

    if mon not in workspaces:
        print(f"[WS] ✗ Monitor {mon} not initialized")
        return

    if ws_idx == current_workspace.get(mon, 0):
        print(f"[WS] Already on workspace {ws_idx+1}")
        return

    # ✅ FREEZE AUTO-RETILE during workspace switch
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
    
    print(f"[WS] Hidden {hidden} windows from workspace {current_workspace.get(mon, 0)+1}")

    # 3) Update current workspace index
    current_workspace[mon] = ws_idx

    # 4) Load new workspace
    time.sleep(0.1)
    load_workspace(mon, ws_idx)
    
    # ✅ UPDATE last_visible_count to prevent immediate re-tile
    visible_windows = get_visible_windows()
    last_visible_count = len(visible_windows)
    
    # ✅ RESTORE AUTO-RETILE state
    time.sleep(0.2)  # Let windows settle
    is_active = was_active
    
    print(f"[WS] ✓ Switched to workspace {ws_idx+1}")

# ==============================================================================
# MAIN
# ==============================================================================
if __name__ == "__main__":
    print("="*70)
    print("   SMARTGRID - Intelligent layouts + green border + SWAP MODE")
    print("="*70)
    print("Ctrl + Alt + T     → Toggle persistent tiling mode (on/off)")
    print("Ctrl + Alt + R     → Force re-tile all visible windows now")
    print("Ctrl + Alt + M     → Move current workspace to next monitor (merges if target workspace has windows)")
    print("Ctrl + Alt + S     → Enter SWAP MODE (red border + arrow keys) to exchange window positions")
    print("                     ↳ Use ← → ↑ ↓ to navigate, Enter or Ctrl+Alt+S to exit")
    print("")
    print("WORKSPACES         → Ctrl + Alt + 1/2/3 to switch workspace")
    print("                     ↳ Each monitor has 3 independent workspaces")
    print("")
    print("DRAG & DROP SNAP   → Grab any tiled window by title bar")
    print("                     ↳ Drop it anywhere → it snaps perfectly (swap or move)")
    print("                     ↳ Works even across monitors!")
    print("")
    print("Ctrl + Alt + Q     → Quit SmartGrid")
    print("-"*70)

    clear_all_borders()
    time.sleep(0.05)

    # Initialize monitors and workspaces
    monitors = get_monitors()
    print(f"Detected {len(monitors)} monitor(s)")
    
    # Initialize workspace structure dynamically
    workspaces = init_workspaces()
    current_workspace = {i: 0 for i in range(len(monitors))}
    print(f"Initialized {len(monitors)} × 3 workspaces")
    
    if len(monitors) > 1:
        print("Ctrl+Alt+M -> cycle tiled windows across monitors")

    # Background services ----------------------------------------------------
    threading.Thread(target=monitor, daemon=True).start()
    # → Watches for new/minimized/restored windows, auto-retiles, manages green border

    threading.Thread(target=start_drag_snap_monitor, daemon=True).start()
    # → Real-time drag detection: enables snap on drop (with swap/move)
    # ------------------------------------------------------------------------
    register_hotkeys()

    msg = wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
        if msg.message == win32con.WM_HOTKEY:
            if msg.wParam == HOTKEY_TOGGLE:
                toggle_persistent()
            elif msg.wParam == HOTKEY_RETILE:
                threading.Thread(target=smart_tile, kwargs={"temp": True}, daemon=True).start()
            elif msg.wParam == HOTKEY_MOVE_MONITOR:
                threading.Thread(target=move_current_workspace_to_next_monitor, daemon=True).start()
            elif msg.wParam == HOTKEY_QUIT:
                break
            elif msg.wParam == HOTKEY_SWAP_MODE:
                if swap_mode:
                    exit_swap_mode()
                else:
                    enter_swap_mode()
            elif swap_mode:
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

        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

    if current_hwnd:
        remove_border(current_hwnd)
    clear_all_borders()    
    unregister_hotkeys()
    print("[EXIT] SmartGrid stopped.")