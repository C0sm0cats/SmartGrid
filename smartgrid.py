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

# Layout & appearance (defaults - can be changed in settings)
DEFAULT_GAP = 8
DEFAULT_EDGE_PADDING = 8

# Window size constraints
MIN_WINDOW_WIDTH = 180
MIN_WINDOW_HEIGHT = 180
TEAMS_TOAST_MAX_WIDTH = 400
TEAMS_TOAST_MAX_HEIGHT = 300

# Timing constants
ANIMATION_DURATION = 0.20
ANIMATION_FPS = 60
DRAG_MONITOR_FPS = 60  # Reduced from 200
RETILE_DEBOUNCE = 0.5  # 500ms between auto-retiles
CACHE_TTL = 5.0  # Cache validity duration
TILE_TIMEOUT = 2.0  # Max time for tiling operation
MAX_TILE_RETRIES = 10
DRAG_THRESHOLD = 10

# DWM attributes
DWMWA_BORDER_COLOR = 34
DWMWA_COLOR_NONE = 0xFFFFFFFF
DWMWA_EXTENDED_FRAME_BOUNDS = 9

# Window styles
GWL_STYLE = -16
WS_THICKFRAME = 0x00040000
WS_MAXIMIZE = 0x01000000

# ShowWindow commands
SW_RESTORE = 9
SW_SHOWNORMAL = 1

# SetWindowPos flags
SWP_NOZORDER = 0x0004
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOACTIVATE = 0x0010
SWP_FRAMECHANGED = 0x0020
SWP_NOSENDCHANGING = 0x0400

# Hotkey IDs
HOTKEY_TOGGLE = 9001
HOTKEY_RETILE = 9002
HOTKEY_QUIT = 9003
HOTKEY_MOVE_MONITOR = 9004
HOTKEY_SWAP_MODE = 9005
HOTKEY_WS1 = 9101
HOTKEY_WS2 = 9102
HOTKEY_WS3 = 9103
HOTKEY_SWAP_LEFT = 9006
HOTKEY_SWAP_RIGHT = 9007
HOTKEY_SWAP_UP = 9008
HOTKEY_SWAP_DOWN = 9009
HOTKEY_SWAP_CONFIRM = 9010
HOTKEY_FLOAT_TOGGLE = 9011

# Custom messages
CUSTOM_TOGGLE_SWAP = 0x9000

# Win32 API
user32 = ctypes.WinDLL('user32', use_last_error=True)
dwmapi = ctypes.WinDLL('dwmapi')

# ==============================================================================
# UTILITY FUNCTIONS (Global helpers - stateless)
# ==============================================================================

def get_monitors():
    """Return list of work area rectangles (x, y, w, h) for all monitors."""
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
    
    try:
        user32.EnumDisplayMonitors(None, None,
            ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HMONITOR, wintypes.HDC,
                               ctypes.POINTER(wintypes.RECT), wintypes.LPARAM)(enum_proc), 0)
    except Exception as e:
        log(f"[ERROR] EnumDisplayMonitors failed: {e}")
    
    if not monitors:
        monitors = [(0, 0, win32gui.GetSystemMetrics(0), win32gui.GetSystemMetrics(1))]
    return monitors

def get_window_state(hwnd):
    """Return window state: 'normal', 'minimized', 'maximized', or 'hidden'"""
    if not hwnd or not user32.IsWindow(hwnd):
        return None
    
    try:
        if not user32.IsWindowVisible(hwnd):
            return 'hidden'
        
        if win32gui.IsIconic(hwnd):
            return 'minimized'

        try:
            if win32gui.IsZoomed(hwnd):
                return 'maximized'
        except Exception:
            style = user32.GetWindowLongW(hwnd, GWL_STYLE)
            if style & WS_MAXIMIZE:
                return 'maximized'
        
        return 'normal'
    except Exception as e:
        log(f"[ERROR] get_window_state failed for hwnd={hwnd}: {e}")
        return None

def get_frame_borders(hwnd):
    """Return (left, top, right, bottom) invisible border/shadow thickness"""
    try:
        rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return 0, 0, 0, 0
        
        ext = wintypes.RECT()
        if dwmapi.DwmGetWindowAttribute(hwnd, DWMWA_EXTENDED_FRAME_BOUNDS,
                                        ctypes.byref(ext), ctypes.sizeof(ext)) == 0:
            return (ext.left - rect.left, ext.top - rect.top,
                    rect.right - ext.right, rect.bottom - ext.bottom)
    except Exception as e:
        log(f"[ERROR] get_frame_borders failed: {e}")
    
    return 0, 0, 0, 0

def set_window_border(hwnd, color):
    """Apply (or remove) a colored DWM border."""
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
    except Exception as e:
        log(f"[ERROR] set_window_border failed: {e}")

def get_process_name(hwnd):
    """Return process name (e.g., 'ms-teams.exe') or empty string on error."""
    if not hwnd:
        return ""
    try:
        _, process_id = win32process.GetWindowThreadProcessId(hwnd)
        h_process = win32api.OpenProcess(0x0410, False, process_id)
        path = win32process.GetModuleFileNameEx(h_process, 0)
        win32api.CloseHandle(h_process)
        return os.path.basename(path).lower()
    except Exception:
        return ""

def get_window_size(hwnd):
    """Return (width, height) or (0, 0) on error."""
    if not hwnd or not user32.IsWindow(hwnd):
        return 0, 0
    try:
        rect = wintypes.RECT()
        if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return rect.right - rect.left, rect.bottom - rect.top
    except Exception:
        pass
    return 0, 0

def is_useful_window(title, class_name="", hwnd=None):
    """Filter out overlays, PIPs, taskbar, notifications, etc."""
    if not title:
        return False

    title_lower = title.lower()
    class_lower = class_name.lower() if class_name else ""

    # Special case: Microsoft Teams - exclude tiny toasts only
    if hwnd:
        process_name = get_process_name(hwnd)
        if process_name == "ms-teams.exe":
            w, h = get_window_size(hwnd)
            if w < TEAMS_TOAST_MAX_WIDTH and h < TEAMS_TOAST_MAX_HEIGHT:
                log(f"[FILTER] Excluded small Teams toast: {title[:50]} | size {w}x{h}")
                return False
            if h < 200:
                return False

    # Hard exclude by title
    bad_titles = [
        "zscaler", "spotify", "discord", "steam", "call", "meeting", "join", "incoming call",
        "obs", "streamlabs", "twitch studio", "nvidia overlay", "geforce experience",
        "shadowplay", "radeon software", "amd relive", "rainmeter", "wallpaper engine",
        "lively wallpaper", "msi afterburner", "rtss", "rivatuner", "hwinfo", "hwmonitor",
        "displayfusion", "actual window", "aquasnap", "powertoys", "fancyzones",
        "picture in picture", "pip", "miniplayer", "mini player", "youtube music",
        "vlc media player", "media player classic", "battle.net", "origin", "epic games",
        "gog galaxy", "uplay", "ubisoft connect", "ea app", "game bar", "xbox",
        "notification", "toast", "popup", "tooltip", "splash", "alert", "flyout",
        "volume control", "brightness", "program manager", "start", "cortana", "search",
        "realtek audio console", "operationstatuswindow", "shell_secondarytraywnd",
        "smartgrid settings", "tk"
    ]

    if any(bad in title_lower for bad in bad_titles):
        return False

    # Hard exclude by class name
    bad_classes = [
        "chrome_renderwidgethosthwnd", "mozillawindowclass", "operationstatuswindow",
        "windows.ui.core.corewindow", "foregroundstaging", "workerw", "progman",
        "shell_traywnd", "realtimedisplay", "credential dialog xaml host",
        "multitaskingviewframe", "taskswitcherwnd", "xamlexplorerhostislandwindow",
        "#32770", "windows.ui.popupwindowclass", "popuphostwindow",
        "microsoft.ui.content.popupwindowsitebridge", "notepadshellexperiencehost",
        "trectanglecapture", "tk", "toplevel" 
    ]

    if class_lower in bad_classes:
        return False

    return True

def get_window_class(hwnd):
    """Return window class name."""
    try:
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(hwnd, buf, 256)
        return buf.value
    except Exception:
        return ""

def create_icon_image():
    """Create SmartGrid icon (green square with white grid)"""
    width, height = 64, 64
    image = Image.new('RGB', (width, height), (0, 180, 0))
    dc = ImageDraw.Draw(image)
    
    # White border
    dc.rectangle((4, 4, width-4, height-4), outline=(255, 255, 255), width=4)
    
    # Grid pattern (2x2)
    mid_x, mid_y = width // 2, height // 2
    dc.line((mid_x, 12, mid_x, height-12), fill=(255, 255, 255), width=3)
    dc.line((12, mid_y, width-12, mid_y), fill=(255, 255, 255), width=3)
    
    return image

def animate_window_move(hwnd, target_x, target_y, target_w, target_h):
    """Animate window movement/resizing with easing"""
    try:
        # Current position
        rect = wintypes.RECT()
        if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return False
        
        lb, tb, rb, bb = get_frame_borders(hwnd)
        start_x = rect.left + lb
        start_y = rect.top + tb
        start_w = rect.right - rect.left - lb - rb
        start_h = rect.bottom - rect.top - tb - bb
        
        # If already at target position, skip
        if (abs(start_x - target_x) < 5 and abs(start_y - target_y) < 5 and
            abs(start_w - target_w) < 5 and abs(start_h - target_h) < 5):
            return False
        
        # Number of frames
        frames = int(ANIMATION_DURATION * ANIMATION_FPS)
        
        # Interpolation with easing
        for i in range(1, frames + 1):
            t = i / frames
            ease = 1 - (1 - t) * (1 - t) * (1 - 2 * t)  # Bounce effect

            x = start_x + (target_x - start_x) * ease
            y = start_y + (target_y - start_y) * ease
            w = start_w + (target_w - start_w) * ease
            h = start_h + (target_h - start_h) * ease
            
            # Apply position (including borders)
            ax = int(x - lb)
            ay = int(y - tb)
            aw = int(w + lb + rb)
            ah = int(h + tb + bb)
            
            user32.SetWindowPos(
                hwnd, 0, ax, ay, aw, ah,
                SWP_NOZORDER | SWP_NOACTIVATE | SWP_NOSENDCHANGING
            )
            
            time.sleep(1.0 / ANIMATION_FPS)
        
        return True
    
    except Exception as e:
        log(f"[ANIM] Error: {e}")
        return False

# ==============================================================================
# LAYOUT ENGINE (Centralized layout calculations)
# ==============================================================================

class LayoutEngine:
    """Calculates window positions for different layout types."""
    
    @staticmethod
    def choose_layout(count):
        """Choose optimal layout based on window count."""
        if count == 1: return "full", None
        if count == 2: return "side_by_side", None
        if count == 3: return "master_stack", None
        if count == 4: return "grid", (2, 2)
        if count <= 6: return "grid", (3, 2)
        if count <= 9: return "grid", (3, 3)
        if count <= 12: return "grid", (4, 3)
        return "grid", (5, 3)
    
    @staticmethod
    def calculate_positions(monitor_rect, count, gap, edge_padding, layout=None, info=None):
        """
        Calculate all window positions for a given layout.
        Returns: (positions_list, grid_coords_list)
            positions_list: [(x, y, w, h), ...]
            grid_coords_list: [(col, row), ...]
        """
        mon_x, mon_y, mon_w, mon_h = monitor_rect
        
        if layout is None:
            layout, info = LayoutEngine.choose_layout(count)
        
        positions = []
        grid_coords = []
        
        if layout == "full":
            positions = [(mon_x + edge_padding, mon_y + edge_padding,
                         mon_w - 2*edge_padding, mon_h - 2*edge_padding)]
            grid_coords = [(0, 0)]
        
        elif layout == "side_by_side":
            cw = (mon_w - 2*edge_padding - gap) // 2
            positions = [
                (mon_x + edge_padding, mon_y + edge_padding, cw, mon_h - 2*edge_padding),
                (mon_x + edge_padding + cw + gap, mon_y + edge_padding, cw, mon_h - 2*edge_padding)
            ]
            grid_coords = [(0, 0), (1, 0)]
        
        elif layout == "master_stack":
            mw = (mon_w - 2*edge_padding - gap) * 3 // 5
            sw = mon_w - 2*edge_padding - mw - gap
            sh = (mon_h - 2*edge_padding - gap) // 2
            positions = [
                (mon_x + edge_padding, mon_y + edge_padding, mw, mon_h - 2*edge_padding),
                (mon_x + edge_padding + mw + gap, mon_y + edge_padding, sw, sh),
                (mon_x + edge_padding + mw + gap, mon_y + edge_padding + sh + gap, sw, sh)
            ]
            grid_coords = [(0, 0), (1, 0), (1, 1)]
        
        else:  # grid
            cols, rows = info if info else (2, 2)
            total_gaps_w = gap * (cols - 1) if cols > 1 else 0
            total_gaps_h = gap * (rows - 1) if rows > 1 else 0
            cell_w = (mon_w - 2*edge_padding - total_gaps_w) // cols
            cell_h = (mon_h - 2*edge_padding - total_gaps_h) // rows
            
            for r in range(rows):
                for c in range(cols):
                    if len(positions) >= count:
                        break
                    x = mon_x + edge_padding + c * (cell_w + gap)
                    y = mon_y + edge_padding + r * (cell_h + gap)
                    positions.append((x, y, cell_w, cell_h))
                    grid_coords.append((c, r))
        
        return positions, grid_coords

# ==============================================================================
# WINDOW MANAGER (Handles grid_state, borders, tiling)
# ==============================================================================

class WindowManager:
    """Manages window grid state, borders, and physical tiling."""
    
    def __init__(self, gap=DEFAULT_GAP, edge_padding=DEFAULT_EDGE_PADDING):
        self.gap = gap
        self.edge_padding = edge_padding
        
        # Window tracking
        self.grid_state = {}  # hwnd → (monitor_idx, col, row)
        self.minimized_windows = {}
        self.maximized_windows = {}
        self.override_windows = set()  # Floating windows
        self.float_restore_slots = {}  # hwnd → (monitor_idx, col, row)
        
        # Active/selected windows
        self.current_hwnd = None  # Green border
        self.selected_hwnd = None  # Red border (swap mode)
        self.last_active_hwnd = None
        self.user_selected_hwnd = None
        
        # Cache for is_useful_window
        self.useful_cache = {}  # hwnd → (timestamp, is_useful)
        self.cache_ttl = CACHE_TTL
        
        # Thread safety
        self.lock = threading.Lock()
    
    def is_window_useful_cached(self, hwnd, title, class_name):
        """Cached version of is_useful_window to reduce overhead."""
        now = time.time()
        if hwnd in self.useful_cache:
            timestamp, result = self.useful_cache[hwnd]
            if now - timestamp < self.cache_ttl:
                return result
        
        result = is_useful_window(title, class_name, hwnd)
        self.useful_cache[hwnd] = (now, result)
        return result
    
    def get_visible_windows(self, monitors, overlay_hwnd=None):
        """Enumerate visible, tileable windows."""
        windows = []
        
        def enum(hwnd, _):
            try:
                if not user32.IsWindowVisible(hwnd):
                    return True
                
                state = get_window_state(hwnd)
                if state in ('minimized', 'maximized', 'hidden'):
                    return True
                
                rect = wintypes.RECT()
                if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                    return True
                
                w = rect.right - rect.left
                h = rect.bottom - rect.top
                
                if w <= MIN_WINDOW_WIDTH or h <= MIN_WINDOW_HEIGHT:
                    return True
                
                title_buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, title_buf, 256)
                title = title_buf.value or ""
                class_name = win32gui.GetClassName(hwnd)
                
                if overlay_hwnd and hwnd == overlay_hwnd:
                    return True
                
                # Check override (float toggle)
                useful = self.is_window_useful_cached(hwnd, title, class_name)
                if hwnd in self.override_windows:
                    useful = not useful
                
                if useful:
                    overlap = sum(
                        max(0, min(rect.right, mx + mw) - max(rect.left, mx)) *
                        max(0, min(rect.bottom, my + mh) - max(rect.top, my))
                        for mx, my, mw, mh in monitors
                    )
                    if overlap > (w * h * 0.15):
                        windows.append((hwnd, title, rect))
            
            except Exception as e:
                log(f"[ERROR] enum callback failed: {e}")
            
            return True
        
        try:
            user32.EnumWindows(
                ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(enum), 0
            )
        except Exception as e:
            log(f"[ERROR] EnumWindows failed: {e}")
        
        return windows
    
    def cleanup_dead_windows(self):
        """Remove dead windows from grid_state and override_windows."""
        dead_windows = []
        
        with self.lock:
            for hwnd in list(self.grid_state.keys()):
                if not user32.IsWindow(hwnd):
                    dead_windows.append(hwnd)
                    self.grid_state.pop(hwnd, None)
                    continue
                
                state = get_window_state(hwnd)
                if state == 'minimized':
                    self.minimized_windows[hwnd] = self.grid_state[hwnd]
                    self.grid_state.pop(hwnd, None)
                    continue
                if state == 'maximized':
                    self.maximized_windows[hwnd] = self.grid_state[hwnd]
                    self.grid_state.pop(hwnd, None)
                    continue
                
                if not user32.IsWindowVisible(hwnd):
                    self.grid_state.pop(hwnd, None)
                    continue

            # Cleanup minimized/maximized caches for dead windows
            for hwnd in list(self.minimized_windows.keys()):
                if not user32.IsWindow(hwnd):
                    self.minimized_windows.pop(hwnd, None)
            for hwnd in list(self.maximized_windows.keys()):
                if not user32.IsWindow(hwnd):
                    self.maximized_windows.pop(hwnd, None)
            
            # Cleanup override_windows
            dead_overrides = [
                hwnd for hwnd in self.override_windows 
                if not user32.IsWindow(hwnd)
            ]
            for hwnd in dead_overrides:
                self.override_windows.remove(hwnd)
            
            # Cleanup float_restore_slots
            for hwnd in list(self.float_restore_slots.keys()):
                if not user32.IsWindow(hwnd):
                    self.float_restore_slots.pop(hwnd, None)
            
            # Cleanup cache
            for hwnd in list(self.useful_cache.keys()):
                if not user32.IsWindow(hwnd):
                    del self.useful_cache[hwnd]
        
        if dead_windows:
            log(f"[CLEAN] Removed {len(dead_windows)} dead windows")
        
        return len(dead_windows)
    
    def cleanup_ghost_windows(self):
        """Remove ghost windows (zombie-like windows)."""
        ghost_windows = []
        
        with self.lock:
            for hwnd in list(self.grid_state.keys()):
                if not user32.IsWindow(hwnd):
                    continue
                
                try:
                    title = win32gui.GetWindowText(hwnd)
                    rect = wintypes.RECT()
                    
                    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                        self.grid_state.pop(hwnd, None)
                        ghost_windows.append(hwnd)
                        continue
                    
                    w = rect.right - rect.left
                    h = rect.bottom - rect.top
                    
                    if (not title or len(title.strip()) == 0) and (w < 50 or h < 50):
                        log(f"[CLEAN] Ghost window: hwnd={hwnd} size={w}x{h}")
                        self.grid_state.pop(hwnd, None)
                        ghost_windows.append(hwnd)
                
                except Exception as e:
                    log(f"[ERROR] cleanup_ghost_windows: {e}")
                    self.grid_state.pop(hwnd, None)
                    ghost_windows.append(hwnd)
        
        if ghost_windows:
            log(f"[CLEAN] Removed {len(ghost_windows)} ghost windows")
        
        return len(ghost_windows)
    
    def apply_border(self, hwnd, color):
        """Apply colored border and update tracking."""
        if not hwnd or not user32.IsWindow(hwnd):
            return
        
        set_window_border(hwnd, color)
        
        if color == 0x0000FF00:  # Green = active
            self.current_hwnd = hwnd
        elif color == 0x000000FF:  # Red = swap mode
            self.selected_hwnd = hwnd
    
    def force_tile_resizable(self, hwnd, x, y, w, h, animate=True):
        """Move and resize window to exact coordinates, handling borders."""
        start_time = time.time()
        
        try:
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
            
            if animate:
                animated = animate_window_move(hwnd, x, y, w, h)
                if animated:
                    # Exact final position after animation
                    lb, tb, rb, bb = get_frame_borders(hwnd)
                    ax, ay = x - lb, y - tb
                    aw, ah = w + lb + rb, h + tb + bb
                    user32.SetWindowPos(hwnd, 0, int(ax), int(ay), int(aw), int(ah),
                                        SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED | SWP_NOSENDCHANGING)
                    return
            
            # Fallback: classic method without animation
            lb, tb, rb, bb = get_frame_borders(hwnd)
            ax, ay = x - lb, y - tb
            aw, ah = w + lb + rb, h + tb + bb
            
            flags = SWP_NOZORDER | SWP_NOACTIVATE | SWP_FRAMECHANGED | SWP_NOSENDCHANGING
            
            for attempt in range(MAX_TILE_RETRIES):
                if time.time() - start_time > TILE_TIMEOUT:
                    log(f"[WARN] Tile timeout for hwnd={hwnd}")
                    break
                
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
                                win32con.RDW_FRAME | win32con.RDW_INVALIDATE | 
                                win32con.RDW_UPDATENOW | win32con.RDW_ALLCHILDREN)
            
            # Final size check
            rect = wintypes.RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            lb, tb, rb, bb = get_frame_borders(hwnd)
            final_w = rect.right - rect.left - lb - rb
            final_h = rect.bottom - rect.top - tb - bb
            
            if abs(final_w - w) > 15 or abs(final_h - h) > 15:
                title = win32gui.GetWindowText(hwnd)
                log(f"[NUKE] MoveWindow forced on -> {title[:60]}")
                win32gui.MoveWindow(hwnd, int(x), int(y), int(w), int(h), True)
        
        except Exception as e:
            log(f"[ERROR] force_tile_resizable failed for hwnd={hwnd}: {e}")

# ==============================================================================
# MAIN APPLICATION CLASS
# ==============================================================================

class SmartGrid:
    """Main application orchestrator."""
    
    def __init__(self):
        # Core components
        self.window_mgr = WindowManager()
        self.layout_engine = LayoutEngine()
        
        # Settings
        self.gap = DEFAULT_GAP
        self.edge_padding = DEFAULT_EDGE_PADDING
        
        # Runtime flags
        self.is_active = False
        self.swap_mode_lock = False
        self.move_monitor_lock = False
        self.workspace_switching_lock = False
        self.drag_drop_lock = False
        self.ignore_retile_until = 0.0
        self.last_visible_count = 0
        self.last_retile_time = 0.0
        self.last_known_count = 0
        self.layout_signature = {}
        self.layout_capacity = {}
        self._maximize_freeze_active = False
        
        # Multi-monitor & workspaces
        self.monitors_cache = []
        self.current_monitor_index = 0
        self.workspaces = {}
        self.current_workspace = {}
        
        # Overlay & UI
        self.overlay_hwnd = None
        self.preview_rect = None
        self.tray_icon = None
        self._wnd_proc_ref = None  # Keep reference to prevent GC
        self.overlay_brush = None  # Reusable GDI objects
        self.overlay_pen = None
        
        # Threading
        self.lock = threading.Lock()
        self._tiling_lock = threading.RLock()
        self._stop_event = threading.Event()
        self.main_thread_id = win32api.GetCurrentThreadId()
        
        # Initialize
        self._init_monitors()
        self._init_workspaces()
    
    def _init_monitors(self):
        """Initialize monitor cache."""
        self.monitors_cache = get_monitors()
        log(f"Detected {len(self.monitors_cache)} monitor(s)")
    
    def _init_workspaces(self):
        """Initialize 3 workspaces per monitor."""
        for i in range(len(self.monitors_cache)):
            self.workspaces[i] = [{}, {}, {}]
            self.current_workspace[i] = 0
        log(f"Initialized {len(self.monitors_cache)} × 3 workspaces")
    
    # ==========================================================================
    # TILING LOGIC
    # ==========================================================================
    
    def smart_tile_with_restore(self):
        """Smart tiling that respects saved grid positions."""
        with self._tiling_lock:
            if time.time() < self.ignore_retile_until:
                return
            
            # GLOBAL LOCK AT THE START
            with self.lock:
                self.ignore_retile_until = time.time() + 0.3
                
                # Cleanup
                self.window_mgr.cleanup_dead_windows()
                self.window_mgr.cleanup_ghost_windows()
            
            # Get visible windows (no lock needed)
            visible_windows = self.window_mgr.get_visible_windows(
                self.monitors_cache, self.overlay_hwnd
            )
            
            if not visible_windows:
                log("[TILE] No windows detected.")
                return
            
            # Separate windows by monitor
            wins_by_monitor = self._group_windows_by_monitor(visible_windows)

            # Snapshot maximized windows AFTER grouping, so restored windows are not
            # mistakenly treated as reserved slots.
            with self.lock:
                maximized_snapshot = dict(self.window_mgr.maximized_windows)
                grid_snapshot = dict(self.window_mgr.grid_state)

            reserved_slots_by_monitor = {}
            for _hwnd, (m_idx, col, row) in maximized_snapshot.items():
                reserved_slots_by_monitor.setdefault(m_idx, set()).add((col, row))

            # Process each monitor
            new_grid = {}
            for mon_idx, windows in wins_by_monitor.items():
                if mon_idx >= len(self.monitors_cache):
                    continue
                visible_count = len(windows)
                if visible_count <= 0:
                    continue

                # Hyprland-like behavior: when a window is maximized, keep its grid slot
                # reserved so background retiles can't steal it.
                reserved_slots = reserved_slots_by_monitor.get(mon_idx, set())
                if reserved_slots:
                    log(f"[TILE] Monitor {mon_idx+1}: reserved slots {sorted(reserved_slots)}")
                    # Hard freeze: while a window is maximized on this monitor, do NOT move any
                    # other tiled windows (Hyprland-like). Keep the previous grid_state for this
                    # monitor, and skip tiling it entirely.
                    kept = 0
                    for hwnd, (m, c, r) in grid_snapshot.items():
                        if m != mon_idx or not user32.IsWindow(hwnd):
                            continue
                        if get_window_state(hwnd) != 'normal':
                            continue
                        new_grid[hwnd] = (m, c, r)
                        kept += 1
                    log(f"[TILE] Monitor {mon_idx+1}: maximize freeze (kept {kept} windows)")
                    continue

                # While there are maximized windows on this monitor, avoid shrinking the
                # layout based solely on visible windows (prevents slot reassignment).
                effective_count = visible_count
                if reserved_slots:
                    effective_count = visible_count + len(reserved_slots)
                    prev_capacity = self.layout_capacity.get(mon_idx, 0)
                    if prev_capacity:
                        effective_count = max(effective_count, prev_capacity)

                layout, info = self.layout_engine.choose_layout(effective_count)
                capacity = self._layout_capacity(layout, info)
                self.layout_signature[mon_idx] = (layout, info)
                self.layout_capacity[mon_idx] = capacity
                self._tile_monitor(
                    mon_idx,
                    windows,
                    new_grid,
                    layout,
                    info,
                    capacity,
                    reserved_slots=reserved_slots,
                )
            
            # Update grid_state with lock - ONLY ONCE
            with self.lock:
                self.window_mgr.grid_state.clear()  # ← CLEAR BEFORE UPDATE
                self.window_mgr.grid_state.update(new_grid)
            
            time.sleep(0.06)
    
    def _group_windows_by_monitor(self, visible_windows):
        """Group windows by their assigned monitor, preserving saved positions."""
        wins_by_monitor = {}
        
        # Atomic copy of necessary dictionaries
        with self.lock:
            minimized_snapshot = dict(self.window_mgr.minimized_windows)
            maximized_snapshot = dict(self.window_mgr.maximized_windows)
            grid_snapshot = dict(self.window_mgr.grid_state)
        
        for hwnd, title, rect in visible_windows:
            win_class = get_window_class(hwnd)
            
            # DETERMINE THE PHYSICAL MONITOR FIRST
            w_center_x = (rect.left + rect.right) // 2
            w_center_y = (rect.top + rect.bottom) // 2
            
            physical_mon_idx = 0
            for i, (mx, my, mw, mh) in enumerate(self.monitors_cache):
                if mx <= w_center_x < mx + mw and my <= w_center_y < my + mh:
                    physical_mon_idx = i
                    break
            
            # Check for saved position (use snapshots)
            if hwnd in minimized_snapshot:
                mon_idx, col, row = minimized_snapshot[hwnd]
                # Atomic removal
                with self.lock:
                    if hwnd in self.window_mgr.minimized_windows:
                        del self.window_mgr.minimized_windows[hwnd]
                log(f"[RESTORE] Restoring minimized window to ({col},{row}): {title[:40]}")
            elif hwnd in maximized_snapshot:
                mon_idx, col, row = maximized_snapshot[hwnd]
                # Atomic removal
                with self.lock:
                    if hwnd in self.window_mgr.maximized_windows:
                        del self.window_mgr.maximized_windows[hwnd]
                log(f"[RESTORE] Restoring maximized window to ({col},{row}): {title[:40]}")
            elif hwnd in grid_snapshot:
                mon_idx, col, row = grid_snapshot[hwnd]
            else:
                # NEW WINDOW: use physical monitor
                mon_idx = physical_mon_idx
                col, row = 0, 0
            
            wins_by_monitor.setdefault(mon_idx, []).append(
                (hwnd, title, rect, col, row, win_class)
            )
        
        return wins_by_monitor
    
    def _tile_monitor(self, mon_idx, windows, new_grid, layout=None, info=None, capacity=None, reserved_slots=None):
        """Tile windows on a specific monitor."""
        monitor_rect = self.monitors_cache[mon_idx]
        visible_count = len(windows)
        if layout is None or capacity is None:
            layout, info = self.layout_engine.choose_layout(visible_count)
            capacity = self._layout_capacity(layout, info)
        reserved_slots = reserved_slots or set()
        log(f"\n[TILE] Monitor {mon_idx+1}: {visible_count}/{capacity} windows -> {layout} layout")
        
        # Calculate positions
        positions, grid_coords = self.layout_engine.calculate_positions(
            monitor_rect, capacity, self.gap, self.edge_padding, layout, info
        )
        
        pos_map = dict(zip(grid_coords, positions))
        
        # Phase 1: Restore saved positions
        reserved_in_layout = {slot for slot in reserved_slots if slot in pos_map}
        assigned = set(reserved_in_layout)
        unassigned_windows = []
        
        for hwnd, title, rect, saved_col, saved_row, win_class in windows:
            target_coords = (saved_col, saved_row)
            
            # CHECK IF COORDINATES ARE VALID
            if (target_coords in pos_map and 
                target_coords not in assigned and
                saved_col < 10 and saved_row < 10):  # ← ADD VALIDATION
                x, y, w, h = pos_map[target_coords]
                self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                new_grid[hwnd] = (mon_idx, saved_col, saved_row)
                assigned.add(target_coords)
                log(f"   ✓ RESTORED to ({saved_col},{saved_row}): {title[:50]} [{win_class}]")
                time.sleep(0.015)
            else:
                # Invalid or already occupied position
                desired = target_coords if (target_coords in pos_map and saved_col < 10 and saved_row < 10) else None
                unassigned_windows.append((hwnd, title, rect, desired, win_class))
        
        # Phase 2: Assign remaining windows
        available_positions = [coord for coord in grid_coords if coord not in assigned]
        coord_order = {coord: i for i, coord in enumerate(grid_coords)}
        
        for hwnd, title, rect, desired, win_class in unassigned_windows:
            if not available_positions:
                break
            
            if desired and desired in pos_map:
                col, row = min(
                    available_positions,
                    key=lambda coord: (
                        abs(coord[0] - desired[0]) + abs(coord[1] - desired[1]),
                        coord_order[coord],
                    ),
                )
            else:
                col, row = available_positions[0]
            
            available_positions.remove((col, row))
            x, y, w, h = pos_map[(col, row)]
            self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
            new_grid[hwnd] = (mon_idx, col, row)
            log(f"   → NEW position ({col},{row}): {title[:50]} [{win_class}]")
            time.sleep(0.015)

    def _sync_window_state_changes(self):
        """Track min/max/restore state transitions for stable retile decisions."""
        restored = []
        minimized_moved = 0
        maximized_moved = 0
        with self.lock:
            # Move minimized/maximized windows out of grid_state (keep their slot)
            for hwnd, (mon, col, row) in list(self.window_mgr.grid_state.items()):
                if not user32.IsWindow(hwnd):
                    continue
                state = get_window_state(hwnd)
                if state == 'minimized':
                    self.window_mgr.minimized_windows[hwnd] = (mon, col, row)
                    self.window_mgr.grid_state.pop(hwnd, None)
                    minimized_moved += 1
                elif state == 'maximized':
                    self.window_mgr.maximized_windows[hwnd] = (mon, col, row)
                    self.window_mgr.grid_state.pop(hwnd, None)
                    maximized_moved += 1

            # Restore minimized windows that returned to normal
            for hwnd, (mon, col, row) in list(self.window_mgr.minimized_windows.items()):
                if not user32.IsWindow(hwnd):
                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    continue
                state = get_window_state(hwnd)
                if state == 'normal' and user32.IsWindowVisible(hwnd):
                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    self.window_mgr.grid_state[hwnd] = (mon, col, row)
                    restored.append((hwnd, mon, col, row))
                elif state == 'maximized':
                    self.window_mgr.maximized_windows[hwnd] = (mon, col, row)
                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    maximized_moved += 1

            # Restore maximized windows that returned to normal
            for hwnd, (mon, col, row) in list(self.window_mgr.maximized_windows.items()):
                if not user32.IsWindow(hwnd):
                    self.window_mgr.maximized_windows.pop(hwnd, None)
                    continue
                state = get_window_state(hwnd)
                if state == 'normal' and user32.IsWindowVisible(hwnd):
                    self.window_mgr.maximized_windows.pop(hwnd, None)
                    self.window_mgr.grid_state[hwnd] = (mon, col, row)
                    restored.append((hwnd, mon, col, row))
                elif state == 'minimized':
                    self.window_mgr.minimized_windows[hwnd] = (mon, col, row)
                    self.window_mgr.maximized_windows.pop(hwnd, None)
                    minimized_moved += 1

        return restored, minimized_moved, maximized_moved

    def _layout_capacity(self, layout, info):
        if layout == "full":
            return 1
        if layout == "side_by_side":
            return 2
        if layout == "master_stack":
            return 3
        if layout == "grid":
            cols, rows = info if info else (2, 2)
            return cols * rows
        return 1

    def _count_visible_by_monitor(self, visible_windows):
        counts = {}
        for _, _, rect in visible_windows:
            w_center_x = (rect.left + rect.right) // 2
            w_center_y = (rect.top + rect.bottom) // 2
            mon_idx = 0
            for i, (mx, my, mw, mh) in enumerate(self.monitors_cache):
                if mx <= w_center_x < mx + mw and my <= w_center_y < my + mh:
                    mon_idx = i
                    break
            counts[mon_idx] = counts.get(mon_idx, 0) + 1
        return counts

    def _get_layout_count_for_monitor(self, mon_idx):
        with self.lock:
            base = self.layout_capacity.get(mon_idx, 0)
        if base > 0:
            return base
        count = 0
        with self.lock:
            for hwnd, (m, _, _) in self.window_mgr.grid_state.items():
                if m == mon_idx and user32.IsWindow(hwnd):
                    count += 1
        if count <= 0:
            return 0
        layout, info = self.layout_engine.choose_layout(count)
        return self._layout_capacity(layout, info)

    def _restore_windows_to_slots(self, restored):
        """Place restored windows back into their saved slots."""
        with self._tiling_lock:
            with self.lock:
                grid_snapshot = dict(self.window_mgr.grid_state)

            need_full_retile = False
            for hwnd, mon_idx, col, row in restored:
                if mon_idx >= len(self.monitors_cache):
                    continue
                if not user32.IsWindow(hwnd):
                    continue
                count = self._get_layout_count_for_monitor(mon_idx)
                if count <= 0:
                    continue

                layout, info = self.layout_engine.choose_layout(count)
                positions, grid_coords = self.layout_engine.calculate_positions(
                    self.monitors_cache[mon_idx], count, self.gap, self.edge_padding, layout, info
                )
                pos_map = dict(zip(grid_coords, positions))
                target = (col, row)
                if target not in pos_map:
                    need_full_retile = True
                    break

                # If the monitor currently has more windows than the last known layout capacity
                # (e.g. because another window appeared while this one was minimized), restoring
                # "in place" cannot be guaranteed without overlaps -> do a single full retile.
                mon_windows = [
                    (h, c, r)
                    for h, (m, c, r) in grid_snapshot.items()
                    if m == mon_idx and user32.IsWindow(h)
                ]
                if len(mon_windows) > count:
                    need_full_retile = True
                    break

                # If the saved slot is already occupied, resolve by moving the occupant(s) to
                # free slot(s) instead of doing a full retile (prevents "swap" surprises).
                conflicts = [
                    other_hwnd
                    for other_hwnd, other_col, other_row in mon_windows
                    if other_hwnd != hwnd and other_col == col and other_row == row
                ]
                if conflicts:
                    try:
                        restored_title = win32gui.GetWindowText(hwnd)[:60]
                    except Exception:
                        restored_title = ""
                    log(f"[RESTORE] Slot ({col},{row}) occupied for '{restored_title}' -> resolving")
                    used_slots = {(c, r) for _, c, r in mon_windows}
                    free_slots = [coord for coord in grid_coords if coord not in used_slots]
                    if len(free_slots) < len(conflicts):
                        log(f"[RESTORE] Not enough free slots ({len(free_slots)}) for {len(conflicts)} conflicts")
                        need_full_retile = True
                        break

                    # Move each conflicting window to a free slot.
                    for other_hwnd in conflicts:
                        new_slot = free_slots.pop(0)
                        try:
                            other_title = win32gui.GetWindowText(other_hwnd)[:60]
                        except Exception:
                            other_title = ""
                        log(f"[RESTORE] Moving '{other_title}' -> {new_slot}")
                        with self.lock:
                            self.window_mgr.grid_state[other_hwnd] = (mon_idx, new_slot[0], new_slot[1])
                        grid_snapshot[other_hwnd] = (mon_idx, new_slot[0], new_slot[1])
                        x2, y2, w2, h2 = pos_map[new_slot]
                        self.window_mgr.force_tile_resizable(other_hwnd, x2, y2, w2, h2)
                        time.sleep(0.01)

                # Finally, put the restored window back in its saved slot.
                x, y, w, h = pos_map[target]
                self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
            if need_full_retile:
                # Bypass the short "grace" window so the conflict is resolved immediately.
                with self.lock:
                    self.ignore_retile_until = 0.0
                self.smart_tile_with_restore()
    
    def apply_grid_state(self):
        """Reapply all saved grid positions physically."""
        with self._tiling_lock:
            with self.lock:
                if not self.window_mgr.grid_state:
                    return
                
                # Remove dead windows
                for hwnd in list(self.window_mgr.grid_state.keys()):
                    if not user32.IsWindow(hwnd):
                        self.window_mgr.grid_state.pop(hwnd, None)
                
                if not self.window_mgr.grid_state:
                    return
                
                grid_snapshot = dict(self.window_mgr.grid_state)
                monitors_snapshot = list(self.monitors_cache)
                gap = self.gap
                edge_padding = self.edge_padding
            
            # Group by monitor (snapshot)
            wins_by_mon = {}
            for hwnd, (mon_idx, col, row) in grid_snapshot.items():
                if mon_idx >= len(monitors_snapshot) or not user32.IsWindow(hwnd):
                    continue
                wins_by_mon.setdefault(mon_idx, []).append((hwnd, col, row))
            
            # Process each monitor
            for mon_idx, windows in wins_by_mon.items():
                monitor_rect = monitors_snapshot[mon_idx]
                count = len(windows)
                layout, info = self.layout_engine.choose_layout(count)
                
                # Calculate positions
                positions, coord_list = self.layout_engine.calculate_positions(
                    monitor_rect, count, gap, edge_padding, layout, info
                )
                
                pos_dict = dict(zip(coord_list, positions))
                
                # Apply positions
                for hwnd, col, row in windows:
                    key = (col, row)
                    if key in pos_dict:
                        x, y, w, h = pos_dict[key]
                        self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                        time.sleep(0.008)
            
            time.sleep(0.03)
    
    def force_immediate_retile(self):
        """Force immediate re-tile (bypass grace delays)."""
        if self.swap_mode_lock:
            return
        
        log("\n[FORCE RETILE] Ctrl+Alt+R → Immediate full re-tile")
        self.ignore_retile_until = 0
        self.last_visible_count = 0
        self.last_known_count = 0
        
        try:
            winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS | winsound.SND_ASYNC)
        except Exception:
            pass
        
        self.smart_tile_with_restore()
    
    # ==========================================================================
    # BORDER MANAGEMENT
    # ==========================================================================
    
    def update_active_border(self):
        """Manage colored DWM borders for visual feedback."""
        # Swap mode: red border
        if self.swap_mode_lock and self.window_mgr.selected_hwnd:
            if user32.IsWindow(self.window_mgr.selected_hwnd):
                set_window_border(self.window_mgr.selected_hwnd, 0x000000FF)
            return
        
        active = user32.GetForegroundWindow()
        
        # Get atomic copy of grid_state
        with self.lock:
            is_tiled = active in self.window_mgr.grid_state
        
        # Active window is tiled → green border
        if is_tiled and user32.IsWindow(active):
            self.window_mgr.last_active_hwnd = active
            
            if self.window_mgr.current_hwnd != active:
                if self.window_mgr.current_hwnd and user32.IsWindow(self.window_mgr.current_hwnd):
                    set_window_border(self.window_mgr.current_hwnd, None)
                
                self.window_mgr.apply_border(active, 0x0000FF00)
                self.window_mgr.current_hwnd = active
            else:
                set_window_border(self.window_mgr.current_hwnd, 0x0000FF00)
            
            return
        
        # Maintain border on last tiled window
        if self.window_mgr.last_active_hwnd:
            if user32.IsWindow(self.window_mgr.last_active_hwnd):
                # Get atomic check
                with self.lock:
                    last_still_tiled = self.window_mgr.last_active_hwnd in self.window_mgr.grid_state
                
                if last_still_tiled:
                    set_window_border(self.window_mgr.last_active_hwnd, 0x0000FF00)
                    self.window_mgr.current_hwnd = self.window_mgr.last_active_hwnd
    
    # ==========================================================================
    # SWAP MODE
    # ==========================================================================
    
    def enter_swap_mode(self):
        """Enter swap mode: red border + arrow keys."""
        self.swap_mode_lock = True
        time.sleep(0.25)
        
        # Get atomic check
        with self.lock:
            grid_empty = len(self.window_mgr.grid_state) == 0

        # Force quick update if grid_state is empty
        if grid_empty:
            visible_windows = self.window_mgr.get_visible_windows(
                self.monitors_cache, self.overlay_hwnd
            )
            
            assignments = []
            per_monitor_idx = {}
            for hwnd, title, rect in visible_windows:
                # Determine the physical monitor
                w_center_x = (rect.left + rect.right) // 2
                w_center_y = (rect.top + rect.bottom) // 2
                
                mon_idx = 0
                for i, (mx, my, mw, mh) in enumerate(self.monitors_cache):
                    if mx <= w_center_x < mx + mw and my <= w_center_y < my + mh:
                        mon_idx = i
                        break
                
                idx = per_monitor_idx.get(mon_idx, 0)
                per_monitor_idx[mon_idx] = idx + 1
                
                col = idx % 10
                row = idx // 10
                assignments.append((hwnd, (mon_idx, col, row)))
            
            # Get atomic modification
            with self.lock:
                for hwnd, pos in assignments:
                    if hwnd not in self.window_mgr.grid_state:
                        self.window_mgr.grid_state[hwnd] = pos
            
            log(f"[SWAP] grid_state rebuilt with {len(self.window_mgr.grid_state)} windows")
            
            # Get atomic re-check
            with self.lock:
                grid_still_empty = len(self.window_mgr.grid_state) == 0
            
            if grid_still_empty:
                log("[SWAP] No tiled windows. Press Ctrl+Alt+T or Ctrl+Alt+R first")
                self.swap_mode_lock = False
                return
        
        # Get smart selection with lock
        candidate = None
        with self.lock:
            if (self.window_mgr.last_active_hwnd and 
                user32.IsWindow(self.window_mgr.last_active_hwnd) and 
                self.window_mgr.last_active_hwnd in self.window_mgr.grid_state):
                candidate = self.window_mgr.last_active_hwnd
            elif user32.GetForegroundWindow() in self.window_mgr.grid_state:
                candidate = user32.GetForegroundWindow()
            else:
                candidate = next(iter(self.window_mgr.grid_state.keys()), None)
        
        if not candidate:
            log("[SWAP] No valid window to select")
            self.swap_mode_lock = False
            return
        
        self.window_mgr.selected_hwnd = candidate
        time.sleep(0.05)
        set_window_border(self.window_mgr.selected_hwnd, 0x000000FF)
        
        title = win32gui.GetWindowText(self.window_mgr.selected_hwnd)[:50]
        log(f"\n[SWAP] ✓ Activated - Selected: '{title}'")
        log("━" * 60)
        log("  DIRECT SWAP with arrow keys:")
        log("    ← → ↑ ↓  : Swap with adjacent window")
        log("    Enter     : Confirm selection")
        log("    Ctrl+Alt+S : Exit swap mode")
        log("━" * 60)
        
        self.register_swap_hotkeys()
        self.update_tray_menu()
    
    def navigate_swap(self, direction):
        """Handle arrow key in swap mode."""
        if not self.swap_mode_lock or not self.window_mgr.selected_hwnd:
            log("[SWAP] Mode not active or no window selected")
            return
        
        log(f"[SWAP] Attempting to swap {direction}...")
        target = self._find_window_in_direction(self.window_mgr.selected_hwnd, direction)
        
        if target:
            if self._swap_windows(self.window_mgr.selected_hwnd, target):
                time.sleep(0.04)
                set_window_border(self.window_mgr.selected_hwnd, None)
                time.sleep(0.04)
                set_window_border(self.window_mgr.selected_hwnd, 0x000000FF)
                
                title = win32gui.GetWindowText(self.window_mgr.selected_hwnd)[:50]
                log(f"[SWAP] ✓ '{title}' swapped {direction}")
                user32.SetForegroundWindow(self.window_mgr.selected_hwnd)
            else:
                log(f"[SWAP] ✗ Swap failed")
        else:
            log(f"[SWAP] ✗ No window in {direction} direction")
    
    def _find_window_in_direction(self, from_hwnd, direction):
        """Find closest tiled window in specified direction."""
        
        # Get atomic copy of necessary data
        with self.lock:
            if from_hwnd not in self.window_mgr.grid_state:
                return None
            
            mon_idx = self.window_mgr.grid_state[from_hwnd][0]
            
            # Copy the list of windows from the same monitor
            windows_snapshot = [
                (hwnd, mon, col, row) 
                for hwnd, (mon, col, row) in self.window_mgr.grid_state.items()
                if hwnd != from_hwnd and mon == mon_idx and user32.IsWindow(hwnd)
            ]
        
        try:
            from_rect = wintypes.RECT()
            if not user32.GetWindowRect(from_hwnd, ctypes.byref(from_rect)):
                return None
            
            fx1, fy1 = from_rect.left, from_rect.top
            fx2, fy2 = from_rect.right, from_rect.bottom
            fcx, fcy = (fx1 + fx2) // 2, (fy1 + fy2) // 2
            from_width = fx2 - fx1
            from_height = fy2 - fy1
            
            best_hwnd = None
            best_distance = float('inf')
            
            # Iterate over the snapshot (no race condition)
            for hwnd, m, _, _ in windows_snapshot:
                rect = wintypes.RECT()
                if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                    continue
                
                x1, y1, x2, y2 = rect.left, rect.top, rect.right, rect.bottom
                cx, cy = (x1 + x2) // 2, (y1 + y2) // 2
                dx, dy = cx - fcx, cy - fcy
                
                # Direction filtering
                if direction == "right" and dx <= 30: continue
                if direction == "left" and dx >= -30: continue
                if direction == "down" and dy <= 30: continue
                if direction == "up" and dy >= -30: continue
                
                # Dynamic overlap threshold
                overlap_x = max(0, min(fx2, x2) - max(fx1, x1))
                overlap_y = max(0, min(fy2, y2) - max(fy1, y1))
                
                if direction in ("left", "right"):
                    min_overlap_y = max(20, int(from_height * 0.20))
                    if overlap_y < min_overlap_y:
                        continue
                else:
                    min_overlap_x = max(20, int(from_width * 0.20))
                    if overlap_x < min_overlap_x:
                        continue
                
                # Euclidean distance
                distance = (dx * dx) + (dy * dy)
                alignment_bonus = (overlap_x if direction in ("up", "down") else overlap_y) * 10
                score = distance - alignment_bonus
                
                if score < best_distance:
                    best_distance = score
                    best_hwnd = hwnd
            
            return best_hwnd
        
        except Exception as e:
            log(f"[ERROR] _find_window_in_direction: {e}")
            return None

    def _swap_windows(self, hwnd1, hwnd2):
        """Swap two windows' grid positions and physically move them."""
        with self._tiling_lock:
            with self.lock:
                if hwnd1 not in self.window_mgr.grid_state or hwnd2 not in self.window_mgr.grid_state:
                    return False
            
            try:
                set_window_border(hwnd1, None)
                set_window_border(hwnd2, None)
                time.sleep(0.05)
                
                rect1 = wintypes.RECT()
                rect2 = wintypes.RECT()
                
                if not user32.GetWindowRect(hwnd1, ctypes.byref(rect1)):
                    return False
                if not user32.GetWindowRect(hwnd2, ctypes.byref(rect2)):
                    return False
                
                lb1, tb1, rb1, bb1 = get_frame_borders(hwnd1)
                lb2, tb2, rb2, bb2 = get_frame_borders(hwnd2)
                
                x1 = rect1.left + lb1
                y1 = rect1.top + tb1
                w1 = rect1.right - rect1.left - lb1 - rb1
                h1 = rect1.bottom - rect1.top - tb1 - bb1
                
                x2 = rect2.left + lb2
                y2 = rect2.top + tb2
                w2 = rect2.right - rect2.left - lb2 - rb2
                h2 = rect2.bottom - rect2.top - tb2 - bb2
                
                title1 = win32gui.GetWindowText(hwnd1)[:40]
                title2 = win32gui.GetWindowText(hwnd2)[:40]
                log(f"[SWAP] '{title1}' ↔ '{title2}'")
                
                # Swap in grid_state with lock
                with self.lock:
                    self.window_mgr.grid_state[hwnd1], self.window_mgr.grid_state[hwnd2] = \
                        self.window_mgr.grid_state[hwnd2], self.window_mgr.grid_state[hwnd1]
                
                # Physical swap
                self.window_mgr.force_tile_resizable(hwnd1, x2, y2, w2, h2)
                time.sleep(0.04)
                self.window_mgr.force_tile_resizable(hwnd2, x1, y1, w1, h1)
                time.sleep(0.04)
                
                return True
            
            except Exception as e:
                log(f"[ERROR] _swap_windows: {e}")
                return False
    
    def exit_swap_mode(self):
        """Exit swap mode and restore normal borders."""
        if not self.swap_mode_lock:
            return
        
        log("[SWAP] Clearing borders...")
        if self.window_mgr.selected_hwnd and user32.IsWindow(self.window_mgr.selected_hwnd):
            set_window_border(self.window_mgr.selected_hwnd, None)
        
        self.window_mgr.selected_hwnd = None
        time.sleep(0.06)
        
        self.unregister_swap_hotkeys()
        self.window_mgr.current_hwnd = None
        
        # Restore green border with lock
        active = user32.GetForegroundWindow()
        
        with self.lock:
            is_tiled = (active and 
                    user32.IsWindowVisible(active) and 
                    active in self.window_mgr.grid_state)
        
        if is_tiled:
            self.window_mgr.apply_border(active, 0x0000FF00)
            log(f"[SWAP] Green border restored")
        
        self.swap_mode_lock = False
        log("[SWAP] ✓ Deactivated\n")
        self.update_tray_menu()
    
    # ==========================================================================
    # DRAG & DROP
    # ==========================================================================
    
    def create_overlay_window(self):
        """Create transparent overlay for drag preview."""
        if self.overlay_hwnd:
            return self.overlay_hwnd
        
        class_name = "SmartGridOverlay"
        
        DefWindowProc = ctypes.windll.user32.DefWindowProcW
        DefWindowProc.argtypes = [wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM]
        DefWindowProc.restype = wintypes.LPARAM
        
        WNDPROCTYPE = ctypes.WINFUNCTYPE(
            wintypes.LPARAM, wintypes.HWND, ctypes.c_uint, wintypes.WPARAM, wintypes.LPARAM
        )

        @WNDPROCTYPE
        def wnd_proc(hwnd, msg, wparam, lparam):
            brush = None
            pen = None
            old_brush = None
            old_pen = None
            hdc = None
            
            try:
                if msg == win32con.WM_PAINT:
                    class PAINTSTRUCT(ctypes.Structure):
                        _fields_ = [
                            ("hdc", wintypes.HDC), ("fErase", wintypes.BOOL),
                            ("rcPaint", wintypes.RECT), ("fRestore", wintypes.BOOL),
                            ("fIncUpdate", wintypes.BOOL), ("rgbReserved", ctypes.c_char * 32)
                        ]
                    
                    ps = PAINTSTRUCT()
                    hdc = user32.BeginPaint(hwnd, ctypes.byref(ps))
                    
                    try:
                        if self.preview_rect:
                            _, _, w, h = self.preview_rect
                            
                            # Use cached GDI objects if available
                            if self.overlay_brush is None:
                                self.overlay_brush = win32gui.CreateSolidBrush(win32api.RGB(100, 149, 237))
                            if self.overlay_pen is None:
                                self.overlay_pen = win32gui.CreatePen(win32con.PS_SOLID, 4, win32api.RGB(65, 105, 225))
                            
                            old_brush = win32gui.SelectObject(hdc, self.overlay_brush)
                            old_pen = win32gui.SelectObject(hdc, self.overlay_pen)
                            
                            win32gui.Rectangle(hdc, 2, 2, w - 2, h - 2)
                    
                    finally:
                        if old_brush:
                            win32gui.SelectObject(hdc, old_brush)
                        if old_pen:
                            win32gui.SelectObject(hdc, old_pen)
                        
                        user32.EndPaint(hwnd, ctypes.byref(ps))
                    
                    return 0
                
                if msg == win32con.WM_DESTROY:
                    return 0
                
                return DefWindowProc(hwnd, msg, wparam, lparam)
            
            except Exception as e:
                log(f"[ERROR] overlay wnd_proc: {e}")
                # Cleanup on error
                if hdc and old_brush:
                    try:
                        win32gui.SelectObject(hdc, old_brush)
                    except Exception:
                        pass
                if hdc and old_pen:
                    try:
                        win32gui.SelectObject(hdc, old_pen)
                    except Exception:
                        pass
                return 0
        
        try:
            # Keep reference to prevent garbage collection
            self._wnd_proc_ref = wnd_proc
            
            wc = win32gui.WNDCLASS()
            wc.lpfnWndProc = self._wnd_proc_ref
            wc.lpszClassName = class_name
            wc.hCursor = win32gui.LoadCursor(0, win32con.IDC_ARROW)
            
            try:
                win32gui.RegisterClass(wc)
            except Exception:
                pass
            
            self.overlay_hwnd = win32gui.CreateWindowEx(
                win32con.WS_EX_LAYERED | win32con.WS_EX_TRANSPARENT |
                win32con.WS_EX_TOPMOST | win32con.WS_EX_TOOLWINDOW,
                class_name, "SmartGrid Preview", win32con.WS_POPUP,
                0, 0, 1, 1, 0, 0, 0, None
            )
            
            win32gui.SetLayeredWindowAttributes(
                self.overlay_hwnd, 0, int(255 * 0.3), win32con.LWA_ALPHA
            )
        
        except Exception as e:
            log(f"[ERROR] create_overlay_window: {e}")
        
        return self.overlay_hwnd
    
    def show_snap_preview(self, x, y, w, h):
        """Show blue snap preview rectangle."""
        if not self.overlay_hwnd:
            self.create_overlay_window()
        
        self.preview_rect = (x, y, w, h)
        
        try:
            win32gui.SetWindowPos(
                self.overlay_hwnd, win32con.HWND_TOPMOST,
                int(x), int(y), int(w), int(h),
                win32con.SWP_SHOWWINDOW | win32con.SWP_NOACTIVATE
            )
            win32gui.InvalidateRect(self.overlay_hwnd, None, True)
            win32gui.UpdateWindow(self.overlay_hwnd)
        except Exception as e:
            log(f"[ERROR] show_snap_preview: {e}")
    
    def hide_snap_preview(self):
        """Hide snap preview overlay."""
        self.preview_rect = None
        if self.overlay_hwnd:
            try:
                win32gui.ShowWindow(self.overlay_hwnd, win32con.SW_HIDE)
            except Exception:
                pass
    
    def calculate_target_rect(self, source_hwnd, cursor_pos):
        """Calculate target snap rectangle for drag & drop."""
        
        # Atomic verification
        with self.lock:
            if source_hwnd not in self.window_mgr.grid_state:
                return None
        
        try:
            cx, cy = cursor_pos
            
            # Find target monitor
            target_mon_idx = 0
            for i, (mx, my, mw, mh) in enumerate(self.monitors_cache):
                if mx <= cx < mx + mw and my <= cy < my + mh:
                    target_mon_idx = i
                    break
            
            monitor_rect = self.monitors_cache[target_mon_idx]
            mon_x, mon_y, mon_w, mon_h = monitor_rect
            
            # Atomic copy of window list
            with self.lock:
                wins_on_mon = [
                    h for h, (m, _, _) in self.window_mgr.grid_state.items()
                    if m == target_mon_idx and user32.IsWindow(h) and h != source_hwnd
                ]
                
                # Also copy maxc and maxr to avoid a second iteration
                maxc = max((c for h,(m,c,r) in self.window_mgr.grid_state.items() 
                        if m == target_mon_idx), default=0)
                maxr = max((r for h,(m,c,r) in self.window_mgr.grid_state.items() 
                        if m == target_mon_idx), default=0)
            
            count = len(wins_on_mon) + 1
            layout, info = self.layout_engine.choose_layout(count)
            
            # Calculate which cell cursor is in
            x = y = w = h = 0
            
            if layout == "master_stack":
                master_w = (mon_w - 2*self.edge_padding - self.gap) * 3 // 5
                master_right = mon_x + self.edge_padding + master_w + self.gap//2
                
                if cx < master_right:
                    x = mon_x + self.edge_padding
                    y = mon_y + self.edge_padding
                    w = master_w
                    h = mon_h - 2*self.edge_padding
                else:
                    sw = mon_w - 2*self.edge_padding - master_w - self.gap
                    sh = (mon_h - 2*self.edge_padding - self.gap) // 2
                    mid = mon_y + mon_h // 2
                    
                    x = mon_x + self.edge_padding + master_w + self.gap
                    y = mon_y + self.edge_padding + (sh + self.gap if cy >= mid else 0)
                    w = sw
                    h = sh
            
            elif layout == "side_by_side":
                cw = (mon_w - 2*self.edge_padding - self.gap) // 2
                x = mon_x + self.edge_padding + (0 if cx < mon_x + mon_w//2 else cw + self.gap)
                y = mon_y + self.edge_padding
                w = cw
                h = mon_h - 2*self.edge_padding
            
            elif layout == "full":
                x = mon_x + self.edge_padding
                y = mon_y + self.edge_padding
                w = mon_w - 2*self.edge_padding
                h = mon_h - 2*self.edge_padding
            
            else:  # grid
                cols, rows = info if info else (2, 2)
                # Use previously copied values
                cols = max(cols, maxc + 1)
                rows = max(rows, maxr + 1)
                
                cw = (mon_w - 2*self.edge_padding - self.gap*(cols-1)) // cols
                ch = (mon_h - 2*self.edge_padding - self.gap*(rows-1)) // rows
                
                relx = cx - mon_x - self.edge_padding
                rely = cy - mon_y - self.edge_padding
                col = min(max(0, relx // (cw + self.gap)), cols-1)
                row = min(max(0, rely // (ch + self.gap)), rows-1)
                
                x = mon_x + self.edge_padding + col * (cw + self.gap)
                y = mon_y + self.edge_padding + row * (ch + self.gap)
                w = cw
                h = ch
            
            # Frame border compensation
            lb, tb, rb, bb = get_frame_borders(source_hwnd)
            return (int(x - lb), int(y - tb), int(w + lb + rb), int(h + tb + bb))
        
        except Exception as e:
            log(f"[ERROR] calculate_target_rect: {e}")
            return None
    
    def handle_snap_drop(self, source_hwnd, cursor_pos):
        """Handle window drop during drag."""
        
        # Initial verification with lock
        with self.lock:
            if source_hwnd not in self.window_mgr.grid_state:
                return
            old_pos = self.window_mgr.grid_state[source_hwnd]
        
        try:
            cx, cy = cursor_pos
            
            # Find target monitor
            target_mon_idx = 0
            for i, (mx, my, mw, mh) in enumerate(self.monitors_cache):
                if mx <= cx < mx + mw and my <= cy < my + mh:
                    target_mon_idx = i
                    break
            
            monitor_rect = self.monitors_cache[target_mon_idx]
            mon_x, mon_y, mon_w, mon_h = monitor_rect
            
            # Atomic copy
            with self.lock:
                wins_on_mon = [
                    h for h, (m, _, _) in self.window_mgr.grid_state.items()
                    if m == target_mon_idx and user32.IsWindow(h) and h != source_hwnd
                ]
                
                max_c = max((c for h,(m,c,r) in self.window_mgr.grid_state.items() 
                            if m == target_mon_idx), default=0)
                max_r = max((r for h,(m,c,r) in self.window_mgr.grid_state.items() 
                            if m == target_mon_idx), default=0)
            
            count = len(wins_on_mon) + 1
            layout, info = self.layout_engine.choose_layout(count)
            
            # Calculate target cell
            target_col = target_row = 0
            
            if layout == "master_stack":
                master_width = (mon_w - 2*self.edge_padding - self.gap) * 3 // 5
                master_right = mon_x + self.edge_padding + master_width + self.gap//2
                
                if cx < master_right:
                    target_col, target_row = 0, 0
                else:
                    mid_y = mon_y + mon_h // 2
                    target_col = 1
                    target_row = 0 if cy < mid_y else 1
            
            elif layout == "side_by_side":
                target_col = 0 if cx < mon_x + mon_w // 2 else 1
                target_row = 0
            
            elif layout == "full":
                target_col, target_row = 0, 0
            
            else:  # grid
                cols, rows = info if info else (2, 2)
                # Use copied values
                cols = max(cols, max_c + 1)
                rows = max(rows, max_r + 1)
                
                cell_w = (mon_w - 2*self.edge_padding - self.gap*(cols-1)) // cols
                cell_h = (mon_h - 2*self.edge_padding - self.gap*(rows-1)) // rows
                
                rel_x = cx - mon_x - self.edge_padding
                rel_y = cy - mon_y - self.edge_padding
                target_col = min(max(0, rel_x // (cell_w + self.gap)), cols - 1)
                target_row = min(max(0, rel_y // (cell_h + self.gap)), rows - 1)
            
            new_pos = (target_mon_idx, target_col, target_row)
            
            if old_pos == new_pos:
                self.apply_grid_state()
                return
            
            # Check if target cell is occupied (atomic)
            target_hwnd = None
            with self.lock:
                for h, pos in self.window_mgr.grid_state.items():
                    if pos == new_pos and h != source_hwnd and user32.IsWindow(h):
                        target_hwnd = h
                        break
                
                # Atomic modification
                if target_hwnd:
                    log(f"[SNAP] SWAP with '{win32gui.GetWindowText(target_hwnd)[:40]}'")
                    self.window_mgr.grid_state[source_hwnd] = new_pos
                    self.window_mgr.grid_state[target_hwnd] = old_pos
                else:
                    log(f"[SNAP] MOVE to cell ({target_col},{target_row}) on monitor {target_mon_idx+1}")
                    self.window_mgr.grid_state[source_hwnd] = new_pos
            
            self.apply_grid_state()
        
        except Exception as e:
            log(f"[ERROR] handle_snap_drop: {e}")
    
    def start_drag_snap_monitor(self):
        """Background thread: drag detection with live preview."""
        was_down = False
        candidate_hwnd = None
        candidate_start = None
        drag_hwnd = None
        drag_start = None
        preview_active = False
        last_valid_rect = None

        WM_NCHITTEST = 0x0084
        HTCLIENT = 1
        HTCAPTION = 2
        
        while not self._stop_event.is_set():
            try:
                down = win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000
                
                # Mouse down
                if down and not was_down:
                    pt = win32api.GetCursorPos()
                    hwnd = win32gui.WindowFromPoint(pt)
                    
                    if self.overlay_hwnd and hwnd == self.overlay_hwnd:
                        hwnd = user32.GetForegroundWindow()
                    
                    # Climb to top-level
                    try:
                        GA_ROOT = 2
                        top = ctypes.windll.user32.GetAncestor(hwnd, GA_ROOT)
                        if top:
                            hwnd = top
                    except Exception:
                        try:
                            for _ in range(16):
                                parent = win32gui.GetParent(hwnd)
                                if not parent:
                                    break
                                hwnd = parent
                        except Exception:
                            pass
                    
                    if not hwnd or not user32.IsWindowVisible(hwnd):
                        hwnd = user32.GetForegroundWindow()
                    
                    # Check maximized
                    if hwnd and get_window_state(hwnd) == 'maximized':
                        was_down = True
                        continue
                    
                    candidate_hwnd = None
                    candidate_start = None

                    # Only consider drag from title bar (prevents false drags when clicking
                    # buttons like maximize/minimize/close, and avoids "retile" jitter).
                    if hwnd and user32.IsWindowVisible(hwnd):
                        try:
                            lparam = win32api.MAKELONG(pt[0] & 0xFFFF, pt[1] & 0xFFFF)
                            hit = win32gui.SendMessage(hwnd, WM_NCHITTEST, 0, lparam)
                        except Exception:
                            hit = None

                        if hit in (HTCAPTION, HTCLIENT):
                            with self.lock:
                                is_in_grid = hwnd in self.window_mgr.grid_state
                            if is_in_grid:
                                candidate_hwnd = hwnd
                                candidate_start = pt

                # Candidate drag: wait for movement threshold before engaging
                elif down and candidate_hwnd and not drag_hwnd:
                    cursor_pos = win32api.GetCursorPos()
                    dx = abs(cursor_pos[0] - candidate_start[0]) if candidate_start else 0
                    dy = abs(cursor_pos[1] - candidate_start[1]) if candidate_start else 0

                    if dx > DRAG_THRESHOLD or dy > DRAG_THRESHOLD:
                        drag_hwnd = candidate_hwnd
                        drag_start = candidate_start
                        candidate_hwnd = None
                        candidate_start = None
                        preview_active = True
                        self.drag_drop_lock = True
                
                # Drag in progress
                elif down and drag_hwnd:
                    cursor_pos = win32api.GetCursorPos()
                    
                    if drag_start:
                        dx = abs(cursor_pos[0] - drag_start[0])
                        dy = abs(cursor_pos[1] - drag_start[1])
                        
                        if preview_active:
                            target_rect = self.calculate_target_rect(drag_hwnd, cursor_pos)
                            if target_rect:
                                last_valid_rect = target_rect
                                self.show_snap_preview(*target_rect)
                            elif last_valid_rect:
                                self.show_snap_preview(*last_valid_rect)
                            else:
                                self.hide_snap_preview()
                
                # Mouse up
                elif not down and was_down:
                    if drag_hwnd:
                        self.hide_snap_preview()
                        cursor_pos = win32api.GetCursorPos()
                        # If the user used Windows "Aero Snap" to maximize during the drag,
                        # do NOT treat it as a grid move (it would move other windows).
                        if get_window_state(drag_hwnd) != 'maximized':
                            self.handle_snap_drop(drag_hwnd, cursor_pos)
                        else:
                            log("[DRAG] Drop ignored (window maximized)")
                        self.drag_drop_lock = False
                        drag_hwnd = None
                        drag_start = None
                        preview_active = False
                        last_valid_rect = None
                    candidate_hwnd = None
                    candidate_start = None
                
                was_down = down
                time.sleep(1.0 / DRAG_MONITOR_FPS)  # ~60 FPS
            
            except Exception as e:
                log(f"[ERROR] drag_snap_monitor: {e}")
                self.hide_snap_preview()
                self.drag_drop_lock = False
                drag_hwnd = None
                drag_start = None
                preview_active = False
                candidate_hwnd = None
                candidate_start = None
                time.sleep(0.1)
    
    # ==========================================================================
    # WORKSPACES
    # ==========================================================================
    
    def save_workspace(self, monitor_idx):
        """Save current workspace state."""
        if monitor_idx not in self.workspaces:
            return
        
        ws = self.current_workspace[monitor_idx]
        self.workspaces[monitor_idx][ws] = {}
        
        # Atomic copy of grid_state
        with self.lock:
            grid_snapshot = dict(self.window_mgr.grid_state)
            minimized_snapshot = dict(self.window_mgr.minimized_windows)
            maximized_snapshot = dict(self.window_mgr.maximized_windows)
        
        # Save normal windows
        for hwnd, (mon, col, row) in grid_snapshot.items():
            if mon != monitor_idx or not user32.IsWindow(hwnd):
                continue
            
            try:
                rect = wintypes.RECT()
                if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                    continue
                
                lb, tb, rb, bb = get_frame_borders(hwnd)
                x = rect.left + lb
                y = rect.top + tb
                w = rect.right - rect.left - lb - rb
                h = rect.bottom - rect.top - tb - bb
                
                self.workspaces[monitor_idx][ws][hwnd] = {
                    'pos': (x, y, w, h),
                    'grid': (col, row),
                    'state': 'normal'
                }
            except Exception as e:
                log(f"[ERROR] save_workspace (normal): {e}")
        
        # Save minimized windows
        for hwnd, (mon, col, row) in minimized_snapshot.items():
            if mon == monitor_idx and user32.IsWindow(hwnd):
                self.workspaces[monitor_idx][ws][hwnd] = {
                    'pos': (0, 0, 800, 600),
                    'grid': (col, row),
                    'state': 'minimized'
                }
        
        # Save maximized windows
        for hwnd, (mon, col, row) in maximized_snapshot.items():
            if mon == monitor_idx and user32.IsWindow(hwnd):
                self.workspaces[monitor_idx][ws][hwnd] = {
                    'pos': (0, 0, 800, 600),
                    'grid': (col, row),
                    'state': 'maximized'
                }
        
        log(f"[WS] ✓ Workspace {ws+1} saved ({len(self.workspaces[monitor_idx][ws])} windows)")
    
    def load_workspace(self, monitor_idx, ws_idx):
        """Load workspace and restore positions."""
        if monitor_idx not in self.workspaces or ws_idx >= len(self.workspaces[monitor_idx]):
            return
        
        layout = self.workspaces[monitor_idx][ws_idx]
        
        if not layout:
            log(f"[WS] Workspace {ws_idx+1} is empty")
            return
        
        log(f"[WS] Loading workspace {ws_idx+1}...")
        self.ignore_retile_until = time.time() + 2.0
        
        grid_updates = {}
        minimized_updates = {}
        maximized_updates = {}
        
        for hwnd, data in layout.items():
            if not user32.IsWindow(hwnd):
                continue
            
            try:
                x, y, w, h = data['pos']
                col, row = data['grid']
                saved_state = data.get('state', 'normal')
                
                if saved_state == 'minimized':
                    user32.ShowWindowAsync(hwnd, win32con.SW_MINIMIZE)
                    minimized_updates[hwnd] = (monitor_idx, col, row)
                elif saved_state == 'maximized':
                    user32.ShowWindowAsync(hwnd, win32con.SW_MAXIMIZE)
                    maximized_updates[hwnd] = (monitor_idx, col, row)
                else:
                    if win32gui.IsIconic(hwnd):
                        user32.ShowWindowAsync(hwnd, SW_RESTORE)
                        time.sleep(0.08)
                    if not user32.IsWindowVisible(hwnd):
                        user32.ShowWindowAsync(hwnd, SW_SHOWNORMAL)
                        time.sleep(0.08)
                    grid_updates[hwnd] = (monitor_idx, col, row)
                
                time.sleep(0.015)
            
            except Exception as e:
                log(f"[ERROR] load_workspace: {e}")
        
        with self.lock:
            for hwnd, pos in minimized_updates.items():
                self.window_mgr.minimized_windows[hwnd] = pos
                self.window_mgr.grid_state.pop(hwnd, None)
                self.window_mgr.maximized_windows.pop(hwnd, None)
            
            for hwnd, pos in maximized_updates.items():
                self.window_mgr.maximized_windows[hwnd] = pos
                self.window_mgr.grid_state.pop(hwnd, None)
                self.window_mgr.minimized_windows.pop(hwnd, None)
            
            for hwnd, pos in grid_updates.items():
                self.window_mgr.grid_state[hwnd] = pos
                self.window_mgr.minimized_windows.pop(hwnd, None)
                self.window_mgr.maximized_windows.pop(hwnd, None)
        
        time.sleep(0.15)
        self.smart_tile_with_restore()
        log(f"[WS] ✓ Workspace {ws_idx+1} restored")
        self.ignore_retile_until = time.time() + 2.0
    
    def ws_switch(self, ws_idx):
        """Switch to specified workspace."""
        self.workspace_switching_lock = True
        time.sleep(0.25)
        
        try:
            mon = self.current_monitor_index
            
            if mon not in self.workspaces:
                log(f"[WS] ✗ Monitor {mon} not initialized")
                return
            
            if ws_idx == self.current_workspace.get(mon, 0):
                log(f"[WS] Already on workspace {ws_idx+1}")
                return
            
            was_active = self.is_active
            self.is_active = False
            
            # Save current
            self.save_workspace(mon)
            
            # Hide current windows
            hidden = 0
            with self.lock:
                for hwnd, (mon_idx, col, row) in list(self.window_mgr.grid_state.items()):
                    if mon_idx == mon and user32.IsWindow(hwnd):
                        user32.ShowWindowAsync(hwnd, win32con.SW_HIDE)
                        del self.window_mgr.grid_state[hwnd]
                        hidden += 1
            
            log(f"[WS] Hidden {hidden} windows from workspace {self.current_workspace.get(mon, 0)+1}")
            
            # Switch
            self.current_workspace[mon] = ws_idx
            
            # Load new
            time.sleep(0.1)
            self.load_workspace(mon, ws_idx)
            
            # Update count
            visible_windows = self.window_mgr.get_visible_windows(
                self.monitors_cache, self.overlay_hwnd
            )
            self.last_visible_count = len(visible_windows)
            with self.lock:
                self.last_known_count = (len(self.window_mgr.grid_state) +
                                         len(self.window_mgr.minimized_windows) +
                                         len(self.window_mgr.maximized_windows))
            
            time.sleep(0.2)
            self.is_active = was_active
            
            log(f"[WS] ✓ Switched to workspace {ws_idx+1}")
        
        except Exception as e:
            log(f"[ERROR] ws_switch: {e}")
        
        finally:
            self.workspace_switching_lock = False
    
    def move_current_workspace_to_next_monitor(self):
        """Move all windows from current workspace to next monitor."""
        self.move_monitor_lock = True
        time.sleep(0.25)
        
        try:
            if not self.window_mgr.grid_state:
                log("[MOVE] Nothing to move")
                return
            
            if len(self.monitors_cache) <= 1:
                log("[MOVE] Only one monitor")
                return
            
            old_mon = self.current_monitor_index
            old_ws = self.current_workspace.get(old_mon, 0)
            self.save_workspace(old_mon)
            
            new_mon = (self.current_monitor_index + 1) % len(self.monitors_cache)
            self.current_monitor_index = new_mon
            
            monitor_rect = self.monitors_cache[new_mon]
            mon_x, mon_y, mon_w, mon_h = monitor_rect
            
            with self.lock:
                grid_snapshot = list(self.window_mgr.grid_state.items())
            
            windows_to_move = [
                (hwnd, pos) for hwnd, pos in grid_snapshot
                if pos[0] == old_mon and user32.IsWindow(hwnd)
            ]
            
            if not windows_to_move:
                log(f"[MOVE] No windows to move from monitor {old_mon+1}")
                return
            
            count = len(windows_to_move)
            layout, info = self.layout_engine.choose_layout(count)
            
            time.sleep(0.05)
            
            positions, grid_coords = self.layout_engine.calculate_positions(
                monitor_rect, count, self.gap, self.edge_padding, layout, info
            )
            
            new_grid = {}
            log(f"\n[MOVE] Moving workspace {old_ws+1} from monitor {old_mon+1} → {new_mon+1}")
            
            with self._tiling_lock:
                for i, (hwnd, (old_idx, col, row)) in enumerate(windows_to_move):
                    x, y, w, h = positions[i]
                    self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                    
                    # Assign grid position
                    if layout == "full":
                        new_grid[hwnd] = (new_mon, 0, 0)
                    elif layout == "side_by_side":
                        new_col = 0 if i == 0 else 1
                        new_grid[hwnd] = (new_mon, new_col, 0)
                    elif layout == "master_stack":
                        if i == 0:
                            new_grid[hwnd] = (new_mon, 0, 0)
                        elif i == 1:
                            new_grid[hwnd] = (new_mon, 1, 0)
                        elif i == 2:
                            new_grid[hwnd] = (new_mon, 1, 1)
                    else:  # grid
                        cols, rows = info
                        new_grid[hwnd] = (new_mon, i % cols, i // cols)
                    
                    title = win32gui.GetWindowText(hwnd)
                    log(f"   -> {title[:50]}")
                    time.sleep(0.015)
                
                # Update grid_state with lock
                with self.lock:
                    for hwnd, _ in windows_to_move:
                        self.window_mgr.grid_state.pop(hwnd, None)
                    self.window_mgr.grid_state.update(new_grid)
            
            time.sleep(0.15)
            
            active = user32.GetForegroundWindow()
            if active and user32.IsWindowVisible(active):
                self.window_mgr.apply_border(active, 0x0000FF00)
            
            self.current_workspace[new_mon] = old_ws
            self.save_workspace(new_mon)
            
            log(f"[MOVE] ✓ Workspace {old_ws+1} now on monitor {new_mon+1}")
        
        except Exception as e:
            log(f"[ERROR] move_workspace: {e}")
        
        finally:
            self.move_monitor_lock = False
    
    # ==========================================================================
    # FLOATING WINDOW TOGGLE
    # ==========================================================================
    
    def toggle_floating_selected(self):
        """Toggle float/tile for currently selected window."""
        hwnd = self.window_mgr.user_selected_hwnd
        
        if not hwnd or not user32.IsWindow(hwnd) or not user32.IsWindowVisible(hwnd):
            winsound.MessageBeep(0xFFFFFFFF)
            log("[FLOAT] No valid window selected")
            return
        
        try:
            title = win32gui.GetWindowText(hwnd)
            class_name = win32gui.GetClassName(hwnd)
            default_useful = is_useful_window(title, class_name)
            
            if hwnd in self.window_mgr.override_windows:
                self.window_mgr.override_windows.remove(hwnd)
                log(f"[OVERRIDE] {title[:60]} → rollback to ({'tile' if default_useful else 'float'})")
                if default_useful:
                    restore_slot = self.window_mgr.float_restore_slots.get(hwnd)
                    if restore_slot:
                        mon_idx, col, row = restore_slot
                        if 0 <= mon_idx < len(self.monitors_cache):
                            with self.lock:
                                if hwnd not in self.window_mgr.grid_state:
                                    self.window_mgr.grid_state[hwnd] = (mon_idx, col, row)
                    self.smart_tile_with_restore()
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)
            else:
                self.window_mgr.override_windows.add(hwnd)
                log(f"[OVERRIDE] {title[:60]} → override to ({'float' if default_useful else 'tile'})")
                if default_useful:
                    with self.lock:
                        if hwnd in self.window_mgr.grid_state:
                            self.window_mgr.float_restore_slots[hwnd] = self.window_mgr.grid_state[hwnd]
                            self.window_mgr.grid_state.pop(hwnd, None)
                    set_window_border(hwnd, None)
                elif not default_useful:
                    self.smart_tile_with_restore()
                winsound.PlaySound("SystemExclamation", winsound.SND_ALIAS | winsound.SND_ASYNC)
            
            self.update_tray_menu()
        
        except Exception as e:
            log(f"[ERROR] toggle_floating_selected: {e}")
    
    # ==========================================================================
    # SETTINGS
    # ==========================================================================
    
    def show_settings_dialog(self):
        """Show settings dialog to modify GAP and EDGE_PADDING."""
        old_ignore = self.ignore_retile_until
        self.ignore_retile_until = float('inf')  # Block tiling
        
        try:
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
            
            title = tk.Label(dialog, text="SmartGrid Settings", font=("Arial", 12, "bold"))
            title.pack(pady=10)
            
            # GAP
            gap_frame = tk.LabelFrame(dialog, text="Gap Between Windows", padx=10, pady=10)
            gap_frame.pack(fill=tk.X, padx=15, pady=5)
            
            gap_options = ["2px", "4px", "6px", "8px", "10px", "12px", "16px", "20px"]
            gap_values = [2, 4, 6, 8, 10, 12, 16, 20]
            current_gap_idx = gap_values.index(self.gap) if self.gap in gap_values else 3
            
            gap_var = tk.StringVar(value=gap_options[current_gap_idx])
            gap_dropdown = ttk.Combobox(gap_frame, textvariable=gap_var, values=gap_options,
                                        state="readonly", width=15)
            gap_dropdown.pack(side=tk.LEFT, padx=5)
            
            gap_label = tk.Label(gap_frame, text=f"(current: {self.gap}px)", font=("Arial", 9, "italic"))
            gap_label.pack(side=tk.LEFT, padx=10)
            
            # EDGE_PADDING
            padding_frame = tk.LabelFrame(dialog, text="Edge Padding (margin)", padx=10, pady=10)
            padding_frame.pack(fill=tk.X, padx=15, pady=5)
            
            padding_options = ["0px", "4px", "8px", "12px", "16px", "20px", "30px", "40px", "50px", "80px", "100px"]
            padding_values = [0, 4, 8, 12, 16, 20, 30, 40, 50, 80, 100]
            current_padding_idx = padding_values.index(self.edge_padding) if self.edge_padding in padding_values else 2
            
            padding_var = tk.StringVar(value=padding_options[current_padding_idx])
            padding_dropdown = ttk.Combobox(padding_frame, textvariable=padding_var,
                                            values=padding_options, state="readonly", width=15)
            padding_dropdown.pack(side=tk.LEFT, padx=5)
            
            padding_label = tk.Label(padding_frame, text=f"(current: {self.edge_padding}px)",
                                    font=("Arial", 9, "italic"))
            padding_label.pack(side=tk.LEFT, padx=10)
            
            # Buttons
            button_frame = tk.Frame(dialog)
            button_frame.pack(pady=15)
            
            def apply_and_close():
                gap_str = gap_var.get().replace("px", "")
                padding_str = padding_var.get().replace("px", "")
                
                # Close BEFORE modifying
                dialog.destroy()
                root.destroy()
                
                # Wait for Tkinter to fully close
                time.sleep(0.3)
                
                # Now modify
                self.gap = int(gap_str)
                self.edge_padding = int(padding_str)
                self.window_mgr.gap = self.gap
                self.window_mgr.edge_padding = self.edge_padding
                
                log(f"[SETTINGS] GAP={self.gap}px, EDGE_PADDING={self.edge_padding}px")
                
                # Apply
                self.apply_new_settings()
            
            def cancel_and_close():
                dialog.destroy()
                root.destroy()
            
            def reset_defaults():
                gap_var.set(gap_options[3])
                padding_var.set(padding_options[2])
            
            tk.Button(button_frame, text="Apply", command=apply_and_close, width=12,
                    bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
            tk.Button(button_frame, text="Reset", command=reset_defaults, width=12).pack(side=tk.LEFT, padx=5)
            tk.Button(button_frame, text="Cancel", command=cancel_and_close, width=12).pack(side=tk.LEFT, padx=5)
            
            info = tk.Label(dialog, text="Changes apply immediately on next retile cycle",
                            font=("Arial", 8), fg="gray")
            info.pack(pady=5)
            
            def on_close():
                self.ignore_retile_until = old_ignore
                dialog.destroy()
                root.destroy()
            
            dialog.protocol("WM_DELETE_WINDOW", on_close)
            
            root.mainloop()
            
        except Exception as e:
            log(f"[ERROR] show_settings_dialog: {e}")
        
        finally:
            # Always restore
            self.ignore_retile_until = old_ignore
    
    def apply_new_settings(self):
        """Apply new GAP/EDGE_PADDING settings."""
        # Wait for all Tkinter windows to close
        time.sleep(0.5)
        
        # Clean up dead windows
        self.window_mgr.cleanup_dead_windows()
        self.window_mgr.cleanup_ghost_windows()
        
        # Temporarily block drag & drop
        self.drag_drop_lock = True
        
        # Reset the counters
        self.ignore_retile_until = 0
        self.last_visible_count = 0
        
        log("[SETTINGS] Applying new settings...")
        
        # Re-tile with the new settings
        self.smart_tile_with_restore()
        
        time.sleep(0.2)
        
        # Reactivate
        self.drag_drop_lock = False
        
        log("[SETTINGS] ✓ Settings applied successfully")
    
    # ==========================================================================
    # HOTKEYS & SYSTRAY
    # ==========================================================================
    
    def register_hotkeys(self):
        """Register global hotkeys."""
        hotkeys = [
            (HOTKEY_TOGGLE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('T')),
            (HOTKEY_RETILE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('R')),
            (HOTKEY_QUIT, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('Q')),
            (HOTKEY_MOVE_MONITOR, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('M')),
            (HOTKEY_SWAP_MODE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('S')),
            (HOTKEY_WS1, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('1')),
            (HOTKEY_WS2, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('2')),
            (HOTKEY_WS3, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('3')),
            (HOTKEY_FLOAT_TOGGLE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('F')),
        ]
        
        failed = []
        for hk_id, mod, key in hotkeys:
            try:
                if not user32.RegisterHotKey(None, hk_id, mod, key):
                    failed.append(hk_id)
                    log(f"[WARN] Failed to register hotkey {hk_id}")
            except Exception as e:
                failed.append(hk_id)
                log(f"[ERROR] Hotkey {hk_id}: {e}")
        
        if failed:
            log(f"[WARN] {len(failed)} hotkeys failed to register")
        else:
            log("[HOTKEYS] All hotkeys registered successfully")
    
    def unregister_hotkeys(self):
        """Unregister all hotkeys."""
        for hk in (HOTKEY_TOGGLE, HOTKEY_RETILE, HOTKEY_QUIT, HOTKEY_MOVE_MONITOR,
                   HOTKEY_SWAP_MODE, HOTKEY_WS1, HOTKEY_WS2, HOTKEY_WS3, HOTKEY_FLOAT_TOGGLE):
            try:
                user32.UnregisterHotKey(None, hk)
            except Exception:
                pass
    
    def register_swap_hotkeys(self):
        """Register swap mode arrow keys."""
        try:
            user32.RegisterHotKey(None, HOTKEY_SWAP_LEFT, 0, win32con.VK_LEFT)
            user32.RegisterHotKey(None, HOTKEY_SWAP_RIGHT, 0, win32con.VK_RIGHT)
            user32.RegisterHotKey(None, HOTKEY_SWAP_UP, 0, win32con.VK_UP)
            user32.RegisterHotKey(None, HOTKEY_SWAP_DOWN, 0, win32con.VK_DOWN)
            user32.RegisterHotKey(None, HOTKEY_SWAP_CONFIRM, 0, win32con.VK_RETURN)
        except Exception as e:
            log(f"[ERROR] register_swap_hotkeys: {e}")
    
    def unregister_swap_hotkeys(self):
        """Unregister swap mode hotkeys."""
        for hk in (HOTKEY_SWAP_LEFT, HOTKEY_SWAP_RIGHT, HOTKEY_SWAP_UP, 
                   HOTKEY_SWAP_DOWN, HOTKEY_SWAP_CONFIRM):
            try:
                user32.UnregisterHotKey(None, hk)
            except Exception:
                pass
    
    def create_tray_menu(self):
        """Create systray context menu."""
        return Menu(
            MenuItem(
                f"Tiling: {'ON' if self.is_active else 'OFF'}",
                lambda: (self.toggle_persistent(), self.update_tray_menu()),
                checked=lambda item: self.is_active
            ),
            MenuItem('Force Re-tile All Windows (Ctrl+Alt+R)',
                     lambda: threading.Thread(target=self.force_immediate_retile, daemon=True).start()),
            MenuItem(
                f"Swap Mode: {'ON' if self.swap_mode_lock else 'OFF'} (Ctrl+Alt+S)",
                lambda: user32.PostThreadMessageW(self.main_thread_id, CUSTOM_TOGGLE_SWAP, 0, 0),
                checked=lambda item: self.swap_mode_lock
            ),
            MenuItem('Move Workspace to Next Monitor (Ctrl+Alt+M)',
                     lambda: threading.Thread(target=self.move_current_workspace_to_next_monitor, daemon=True).start()),
            MenuItem('Toggle Floating Selected Window (Ctrl+Alt+F)',
                     lambda: threading.Thread(target=self.toggle_floating_selected, daemon=True).start()),
            Menu.SEPARATOR,
            MenuItem('Workspaces', Menu(
                MenuItem('Switch to Workspace 1 (Ctrl+Alt+1)', lambda: self.ws_switch(0)),
                MenuItem('Switch to Workspace 2 (Ctrl+Alt+2)', lambda: self.ws_switch(1)),
                MenuItem('Switch to Workspace 3 (Ctrl+Alt+3)', lambda: self.ws_switch(2)),
            )),
            Menu.SEPARATOR,
            MenuItem('Settings (Gap & Padding)',
                     lambda: threading.Thread(target=self.show_settings_dialog, daemon=True).start()),
            MenuItem('Hotkeys Cheatsheet', lambda: threading.Thread(target=show_hotkeys_tooltip, daemon=True).start()),
            MenuItem('Quit SmartGrid (Ctrl+Alt+Q)', self.on_quit_from_tray)
        )
    
    def update_tray_menu(self):
        """Refresh systray menu."""
        if self.tray_icon:
            self.tray_icon.menu = self.create_tray_menu()
            self.tray_icon.update_menu()
    
    def toggle_persistent(self):
        """Toggle persistent auto-tiling mode."""
        self.is_active = not self.is_active
        log(f"\n[SMARTGRID] Persistent mode: {'ON' if self.is_active else 'OFF'}")
        
        if self.is_active:
            with self.lock:
                self.window_mgr.grid_state.clear()
                self.last_visible_count = 0
                self.last_known_count = 0
            self.smart_tile_with_restore()
        else:
            with self.lock:
                for hwnd in list(self.window_mgr.grid_state.keys()):
                    if user32.IsWindow(hwnd):
                        set_window_border(hwnd, None)
                self.window_mgr.grid_state.clear()
        
        self.update_tray_menu()
    
    def on_quit_from_tray(self, icon=None, item=None):
        """Quit from systray."""
        log("[TRAY] Quit requested")
        
        try:
            # Stop the systray icon
            if self.tray_icon:
                self.tray_icon.stop()
            
            # Clean up resources
            self.cleanup()
            
            # Cleanup overlay GDI objects
            if self.overlay_brush:
                try:
                    win32gui.DeleteObject(self.overlay_brush)
                except:
                    pass
            if self.overlay_pen:
                try:
                    win32gui.DeleteObject(self.overlay_pen)
                except:
                    pass
            
            # Destroy overlay window
            if self.overlay_hwnd:
                try:
                    win32gui.DestroyWindow(self.overlay_hwnd)
                except:
                    pass
            
            # SEND WM_QUIT TO MAIN THREAD
            user32.PostThreadMessageW(
                self.main_thread_id, 
                win32con.WM_QUIT, 
                0, 
                0
            )
            
            log("[TRAY] WM_QUIT posted to main thread")
        
        except Exception as e:
            log(f"[ERROR] on_quit_from_tray: {e}")
            # Forcer la sortie en dernier recours
            os._exit(0)
    
    def cleanup(self):
        """Cleanup before exit."""
        try:
            self._stop_event.set()
            if self.window_mgr.current_hwnd:
                set_window_border(self.window_mgr.current_hwnd, None)
            
            if self.swap_mode_lock:
                self.exit_swap_mode()
            
            self.unregister_hotkeys()
        except Exception as e:
            log(f"[ERROR] cleanup: {e}")
    
    # ==========================================================================
    # THREADS & MAIN LOOP
    # ==========================================================================
    
    def monitor_loop(self):
        """Background loop: auto-retile + border tracking + monitor detection."""
        last_monitor_count = len(self.monitors_cache)
        
        while not self._stop_event.is_set():
            try:
                # Monitor configuration change detection
                current_monitors = get_monitors()
                if len(current_monitors) != last_monitor_count:
                    log(f"[MONITOR] Configuration changed: {last_monitor_count} → {len(current_monitors)}")
                    self.monitors_cache = current_monitors
                    self._init_workspaces()
                    if self.is_active:
                        self.smart_tile_with_restore()
                    last_monitor_count = len(current_monitors)
                
                # Track user selection
                fg = user32.GetForegroundWindow()
                if fg and user32.IsWindow(fg) and user32.IsWindowVisible(fg):
                    title = win32gui.GetWindowText(fg)
                    class_name = win32gui.GetClassName(fg)
                    if is_useful_window(title, class_name):
                        self.window_mgr.user_selected_hwnd = fg
                
                self.update_active_border()
                
                if (self.swap_mode_lock or self.drag_drop_lock or 
                    self.move_monitor_lock or self.workspace_switching_lock):
                    time.sleep(0.1)
                    continue
                
                if self.is_active:
                    restored, minimized_moved, _maximized_moved = self._sync_window_state_changes()
                    if restored:
                        self._restore_windows_to_slots(restored)

                    # Lightweight cleanup (safe during maximize freeze)
                    self.window_mgr.cleanup_dead_windows()

                    # Hard rule: while ANY window is maximized, do not auto-retile/reflow.
                    # This prevents "background retiles" from moving other windows (Hyprland-like).
                    with self.lock:
                        has_maximized = any(
                            user32.IsWindow(hwnd) for hwnd in self.window_mgr.maximized_windows.keys()
                        )
                    if has_maximized and not self._maximize_freeze_active:
                        log("[FREEZE] Maximize detected -> auto-retile paused")
                    elif not has_maximized and self._maximize_freeze_active:
                        log("[FREEZE] Maximize cleared -> auto-retile resumed")
                    self._maximize_freeze_active = has_maximized
                    if has_maximized:
                        # Keep the counters in sync to avoid a retile storm when unmaximizing.
                        visible_windows = self.window_mgr.get_visible_windows(
                            self.monitors_cache, self.overlay_hwnd
                        )
                        current_count = len(visible_windows)
                        with self.lock:
                            known_hwnds = (set(self.window_mgr.grid_state.keys()) |
                                           set(self.window_mgr.minimized_windows.keys()) |
                                           set(self.window_mgr.maximized_windows.keys()))
                        self.last_visible_count = current_count
                        self.last_known_count = len(known_hwnds)
                        time.sleep(0.06)
                        continue

                    if minimized_moved:
                        # Minimizing a tiled window reduces the effective count; do a single
                        # reflow so the remaining windows fill the layout.
                        self.smart_tile_with_restore()

                    visible_windows = self.window_mgr.get_visible_windows(
                        self.monitors_cache, self.overlay_hwnd
                    )
                    current_count = len(visible_windows)
                    visible_hwnds = {hwnd for hwnd, _, _ in visible_windows}
                    counts_by_monitor = self._count_visible_by_monitor(visible_windows)

                    layout_change = False
                    for mon_idx, visible_count in counts_by_monitor.items():
                        if visible_count <= 0:
                            continue
                        # Use the last known layout capacity to avoid "retile storms" when
                        # windows are temporarily minimized/maximized (visible_count changes,
                        # but the intended grid layout should remain stable).
                        prev_sig = self.layout_signature.get(mon_idx)
                        if prev_sig is None:
                            continue
                        expected_count = self._get_layout_count_for_monitor(mon_idx)
                        if expected_count <= 0:
                            expected_count = visible_count
                        layout, info = self.layout_engine.choose_layout(expected_count)
                        sig = (layout, info)
                        if prev_sig != sig:
                            layout_change = True
                            break

                    with self.lock:
                        known_hwnds = (set(self.window_mgr.grid_state.keys()) |
                                       set(self.window_mgr.minimized_windows.keys()) |
                                       set(self.window_mgr.maximized_windows.keys()))
                    known_count = len(known_hwnds)
                    new_windows = [h for h in visible_hwnds if h not in known_hwnds]
                    
                    # Debounced retiling
                    now = time.time()
                    if now >= self.ignore_retile_until and current_count > 0:
                        should_retile = False
                        if new_windows:
                            should_retile = True
                        elif known_count < self.last_known_count:
                            should_retile = True
                        elif layout_change:
                            should_retile = True

                        if should_retile and now - self.last_retile_time >= RETILE_DEBOUNCE:
                            log(f"[AUTO-RETILE] {self.last_visible_count} → {current_count} windows")
                            self.smart_tile_with_restore()
                            self.last_visible_count = current_count
                            self.last_known_count = known_count
                            self.last_retile_time = now
                            time.sleep(0.2)
                        elif not should_retile:
                            self.last_visible_count = current_count
                            self.last_known_count = known_count
                
                time.sleep(0.06)
            
            except Exception as e:
                log(f"[ERROR] monitor_loop: {e}")
                time.sleep(0.5)
    
    def message_loop(self):
        """Main message loop for hotkeys."""
        msg = wintypes.MSG()
        while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) > 0:
            try:
                if msg.message == win32con.WM_HOTKEY:
                    if msg.wParam == HOTKEY_TOGGLE:
                        self.toggle_persistent()
                    elif msg.wParam == HOTKEY_RETILE:
                        threading.Thread(target=self.force_immediate_retile, daemon=True).start()
                    elif msg.wParam == HOTKEY_MOVE_MONITOR:
                        threading.Thread(target=self.move_current_workspace_to_next_monitor, daemon=True).start()
                    elif msg.wParam == HOTKEY_QUIT:
                        log("[HOTKEY] Ctrl+Alt+Q pressed - Quitting...")
                        self.on_quit_from_tray()  # Reuse the same function
                        break
                    elif msg.wParam == HOTKEY_SWAP_MODE:
                        if self.swap_mode_lock:
                            self.exit_swap_mode()
                        else:
                            self.enter_swap_mode()
                    elif self.swap_mode_lock:
                        if msg.wParam == HOTKEY_SWAP_LEFT:
                            self.navigate_swap("left")
                        elif msg.wParam == HOTKEY_SWAP_RIGHT:
                            self.navigate_swap("right")
                        elif msg.wParam == HOTKEY_SWAP_UP:
                            self.navigate_swap("up")
                        elif msg.wParam == HOTKEY_SWAP_DOWN:
                            self.navigate_swap("down")
                        elif msg.wParam == HOTKEY_SWAP_CONFIRM:
                            self.exit_swap_mode()
                    elif msg.wParam == HOTKEY_WS1:
                        self.ws_switch(0)
                    elif msg.wParam == HOTKEY_WS2:
                        self.ws_switch(1)
                    elif msg.wParam == HOTKEY_WS3:
                        self.ws_switch(2)
                    elif msg.wParam == HOTKEY_FLOAT_TOGGLE:
                        self.toggle_floating_selected()
                
                elif msg.message == CUSTOM_TOGGLE_SWAP:
                    if self.swap_mode_lock:
                        self.exit_swap_mode()
                    else:
                        self.enter_swap_mode()
                
                # HANDLE WM_QUIT EXPLICITLY
                elif msg.message == win32con.WM_QUIT:
                    log("[MAIN] WM_QUIT received - Exiting message loop")
                    break
                
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            
            except Exception as e:
                log(f"[ERROR] message_loop: {e}")
        
        # CLEANUP AFTER LOOP EXIT
        log("[MAIN] Message loop ended - Final cleanup")
        self.cleanup()
        log("[EXIT] SmartGrid stopped.")
    
    def start(self):
        """Start all background threads."""
        threading.Thread(target=self.monitor_loop, daemon=True).start()
        threading.Thread(target=self.start_drag_snap_monitor, daemon=True).start()
        log("[MAIN] Background threads started")

# ==============================================================================
# UTILITY FUNCTIONS (Global helpers for UI)
# ==============================================================================

def show_hotkeys_tooltip():
    """Show hotkeys notification."""
    try:
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
            0x40
        )
    except Exception as e:
        log(f"[ERROR] show_hotkeys_tooltip: {e}")

# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================

# Global instance (needed for Win32 callbacks)
_app_instance = None

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
    
    # Create application instance
    _app_instance = SmartGrid()
    app = _app_instance
    
    # Start background threads
    app.start()
    
    # Register hotkeys
    app.register_hotkeys()
    
    # Create systray icon
    app.tray_icon = Icon(
        "SmartGrid",
        create_icon_image(),
        "SmartGrid - Tiling Window Manager",
        menu=app.create_tray_menu()
    )
    app.update_tray_menu()
    
    def setup(icon):
        icon.visible = True
    
    app.tray_icon.run_detached(setup=setup)
    log("[MAIN] Systray launched")
    
    # Main message loop
    try:
        app.message_loop()
    except KeyboardInterrupt:
        log("[EXIT] Keyboard interrupt")
    except Exception as e:
        log(f"[ERROR] main: {e}")
    finally:
        app.cleanup()
        log("[EXIT] SmartGrid stopped.")
        time.sleep(0.2)  # Allow threads to finish
        os._exit(0)  # Force stop all threads
