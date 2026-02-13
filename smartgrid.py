"""
SmartGrid – Advanced Windows Tiling Manager
Features: intelligent layouts, drag & drop snap, swap mode, workspaces per monitor,
          colored DWM borders, multi-monitor, systray, hotkeys
Author: C0sm0cats (2025)
"""

import os
import ctypes
import time
import threading
import winsound
import bisect
import math
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
ANIMATION_DURATION = 0.08
ANIMATION_FPS = 60
DRAG_MONITOR_FPS = 60  # Reduced from 200
RETILE_DEBOUNCE = 0.05  # 50ms between auto-retiles
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
HOTKEY_LAYOUT_PICKER = 9012

# Custom messages
CUSTOM_TOGGLE_SWAP = 0x9000
CUSTOM_OPEN_LAYOUT_PICKER = 0x9001
CUSTOM_OPEN_SETTINGS = 0x9002

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
        "smartgrid settings", "smartgrid layout manager", "tk"
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

def animate_window_move(
    hwnd,
    target_x,
    target_y,
    target_w,
    target_h,
    duration=ANIMATION_DURATION,
    fps=ANIMATION_FPS,
    effect="smoothstep",
):
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
        
        fps = max(1, int(fps))
        duration = max(0.0, float(duration))
        if duration <= 0.0:
            return False
        # Number of frames
        frames = max(1, int(duration * fps))
        
        # Interpolation with easing
        for i in range(1, frames + 1):
            t = i / frames
            if effect == "linear":
                ease = t
            elif effect == "ease_in":
                ease = t ** 3
            elif effect == "ease_in_out":
                ease = 0.5 * (1.0 - math.cos(math.pi * t))
            elif effect == "ease_out":
                ease = 1.0 - ((1.0 - t) ** 3)
            elif effect == "expo_out":
                ease = 1.0 if t >= 1.0 else (1.0 - (2.0 ** (-10.0 * t)))
            elif effect == "back_out":
                c1 = 1.70158
                c3 = c1 + 1.0
                p = t - 1.0
                ease = 1.0 + (c3 * (p ** 3)) + (c1 * (p ** 2))
            elif effect == "elastic_out":
                if t <= 0.0 or t >= 1.0:
                    ease = t
                else:
                    c4 = (2.0 * math.pi) / 3.0
                    ease = (2.0 ** (-10.0 * t)) * math.sin((t * 10.0 - 0.75) * c4) + 1.0
            elif effect == "spring_out":
                if t <= 0.0:
                    ease = 0.0
                elif t >= 1.0:
                    ease = 1.0
                else:
                    # Damped spring response (distinct from elastic: less aggressive, more "physical").
                    zeta = 0.32
                    omega0 = 10.0
                    omega_d = omega0 * math.sqrt(max(1e-6, 1.0 - zeta * zeta))
                    expo = math.exp(-zeta * omega0 * t)
                    sin_scale = zeta / math.sqrt(max(1e-6, 1.0 - zeta * zeta))
                    ease = 1.0 - expo * (
                        math.cos(omega_d * t) + sin_scale * math.sin(omega_d * t)
                    )
            elif effect == "crit_damped":
                if t <= 0.0:
                    ease = 0.0
                elif t >= 1.0:
                    ease = 1.0
                else:
                    # Critically damped response: fast settle, no overshoot.
                    omega = 10.0
                    ease = 1.0 - math.exp(-omega * t) * (1.0 + omega * t)
            elif effect == "bounce_out":
                n1 = 7.5625
                d1 = 2.75
                if t < 1.0 / d1:
                    ease = n1 * t * t
                elif t < 2.0 / d1:
                    p = t - 1.5 / d1
                    ease = n1 * p * p + 0.75
                elif t < 2.5 / d1:
                    p = t - 2.25 / d1
                    ease = n1 * p * p + 0.9375
                else:
                    p = t - 2.625 / d1
                    ease = n1 * p * p + 0.984375
            elif effect == "arc_wave":
                ease = 0.5 * (1.0 - math.cos(math.pi * t))
            else:  # smoothstep (default)
                ease = t * t * (3 - 2 * t)

            x = start_x + (target_x - start_x) * ease
            y = start_y + (target_y - start_y) * ease
            if effect == "arc_wave":
                # Curved "fly-in" path for a visibly distinct premium effect.
                travel = math.hypot(target_x - start_x, target_y - start_y)
                arc_amp = max(16.0, min(90.0, travel * 0.12))
                direction = -1.0 if target_y >= start_y else 1.0
                y += direction * math.sin(math.pi * t) * arc_amp
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
            
            time.sleep(1.0 / fps)
        
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
        self.animation_enabled = True
        self.animation_duration = ANIMATION_DURATION
        self.animation_fps = ANIMATION_FPS
        self.animation_effect = "crit_damped"
        self.tile_timeout = TILE_TIMEOUT
        self.max_tile_retries = MAX_TILE_RETRIES
        
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
        self.last_cleanup_minimized_moved = 0
        
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
        minimized_moved = 0
        maximized_moved = 0
        
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
                    minimized_moved += 1
                    continue
                if state == 'maximized':
                    self.maximized_windows[hwnd] = self.grid_state[hwnd]
                    self.grid_state.pop(hwnd, None)
                    maximized_moved += 1
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
            self.last_cleanup_minimized_moved = minimized_moved
        
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
            
            if animate and self.animation_enabled:
                animated = animate_window_move(
                    hwnd,
                    x,
                    y,
                    w,
                    h,
                    duration=self.animation_duration,
                    fps=self.animation_fps,
                    effect=self.animation_effect,
                )
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
            
            for attempt in range(max(1, int(self.max_tile_retries))):
                if time.time() - start_time > max(0.2, float(self.tile_timeout)):
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
        self.is_active = True
        self.swap_mode_lock = False
        self.move_monitor_lock = False
        self.workspace_switching_lock = False
        self.drag_drop_lock = False
        self.ignore_retile_until = 0.0
        self.last_visible_count = 0
        self.last_retile_time = 0.0
        self.retile_debounce = RETILE_DEBOUNCE
        self.last_known_count = 0
        self.layout_signature = {}
        self.layout_capacity = {}
        self.workspace_layout_signature = {}  # (monitor_idx, ws_idx) -> (layout, info)
        self._maximize_freeze_active = False
        self.compact_on_minimize = True
        self.compact_on_close = True
        self._pending_compact_minimize = False
        self._pending_compact_close = False
        self.window_state_ws = {}  # hwnd -> workspace index when cached in min/max maps
        # hwnd -> minimize snapshot (for exact full-grid restore when context matches)
        self.minimize_restore_snapshots = {}
        
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
        self._layout_picker_lock = threading.Lock()
        self._layout_picker_open = False
        self._layout_picker_hwnd = None
        
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

    def _reconcile_workspaces_after_monitor_change(self, new_monitors):
        """Preserve workspace state by monitor index when display count changes."""
        new_count = len(new_monitors)
        new_workspaces = {}
        new_current_workspace = {}

        with self.lock:
            old_workspaces = dict(self.workspaces)
            old_current_workspace = dict(self.current_workspace)
            old_layout_signature = dict(self.layout_signature)
            old_layout_capacity = dict(self.layout_capacity)
            old_workspace_layout_signature = dict(self.workspace_layout_signature)

            for i in range(new_count):
                ws_maps = old_workspaces.get(i)
                if isinstance(ws_maps, list) and len(ws_maps) == 3:
                    rebuilt = []
                    for ws_map in ws_maps:
                        rebuilt.append(dict(ws_map) if isinstance(ws_map, dict) else {})
                    new_workspaces[i] = rebuilt
                else:
                    new_workspaces[i] = [{}, {}, {}]

                ws_idx = old_current_workspace.get(i, 0)
                new_current_workspace[i] = ws_idx if ws_idx in (0, 1, 2) else 0

            self.monitors_cache = list(new_monitors)
            self.workspaces = new_workspaces
            self.current_workspace = new_current_workspace

            if self.current_monitor_index >= new_count:
                self.current_monitor_index = max(0, new_count - 1)

            self.layout_signature = {
                mon: sig for mon, sig in old_layout_signature.items()
                if mon < new_count
            }
            self.layout_capacity = {
                mon: cap for mon, cap in old_layout_capacity.items()
                if mon < new_count
            }
            self.workspace_layout_signature = {
                (mon, ws): sig
                for (mon, ws), sig in old_workspace_layout_signature.items()
                if mon < new_count and ws in (0, 1, 2)
            }

    # ==========================================================================
    # LAYOUT HELPERS (Manual layout picker)
    # ==========================================================================

    def _layout_label(self, layout, info):
        if layout == "full":
            return "Full"
        if layout == "side_by_side":
            return "Side-by-side"
        if layout == "master_stack":
            return "Master Stack"
        if layout == "grid":
            cols, rows = info if info else (2, 2)
            return f"Grid {cols}x{rows}"
        return layout

    def _get_layout_presets(self):
        return [
            ("Full", ("full", None)),
            ("Side-by-side", ("side_by_side", None)),
            ("Master Stack", ("master_stack", None)),
            ("Grid 2x2", ("grid", (2, 2))),
            ("Grid 3x2", ("grid", (3, 2))),
            ("Grid 3x3", ("grid", (3, 3))),
            ("Grid 4x3", ("grid", (4, 3))),
            ("Grid 5x3", ("grid", (5, 3))),
        ]

    def _get_monitor_index_for_rect(self, rect):
        w_center_x = (rect.left + rect.right) // 2
        w_center_y = (rect.top + rect.bottom) // 2
        mon_idx = 0
        for i, (mx, my, mw, mh) in enumerate(self.monitors_cache):
            if mx <= w_center_x < mx + mw and my <= w_center_y < my + mh:
                mon_idx = i
                break
        return mon_idx

    def _build_window_descriptor(self, hwnd):
        try:
            title = win32gui.GetWindowText(hwnd)
        except Exception:
            title = ""
        process = get_process_name(hwnd)
        return {"title": title, "process": process}

    def _get_window_choices_for_monitor(self, mon_idx):
        """Return list of (hwnd, descriptor) for windows on a monitor across all workspaces."""
        choices = []
        seen = set()

        def _add_choice(hwnd):
            if not hwnd or hwnd in seen or not user32.IsWindow(hwnd):
                return
            seen.add(hwnd)
            choices.append((hwnd, self._build_window_descriptor(hwnd)))

        visible = self.window_mgr.get_visible_windows(self.monitors_cache, self.overlay_hwnd)
        for hwnd, _title, rect in visible:
            if self._get_monitor_index_for_rect(rect) != mon_idx:
                continue
            _add_choice(hwnd)

        with self.lock:
            minimized_snapshot = dict(self.window_mgr.minimized_windows)
            maximized_snapshot = dict(self.window_mgr.maximized_windows)
            grid_snapshot = dict(self.window_mgr.grid_state)
            ws_snapshot = []
            for ws_map in self.workspaces.get(mon_idx, []):
                ws_snapshot.append(dict(ws_map) if isinstance(ws_map, dict) else {})

        # Runtime tracked windows on this monitor.
        for hwnd, (m, _c, _r) in minimized_snapshot.items():
            if m == mon_idx:
                _add_choice(hwnd)
        for hwnd, (m, _c, _r) in maximized_snapshot.items():
            if m == mon_idx:
                _add_choice(hwnd)
        for hwnd, (m, _c, _r) in grid_snapshot.items():
            if m == mon_idx:
                _add_choice(hwnd)

        # Include windows referenced by any workspace map on this monitor.
        for ws_map in ws_snapshot:
            for hwnd in ws_map.keys():
                _add_choice(hwnd)

        if not choices:
            def enum(hwnd, _):
                try:
                    if not user32.IsWindowVisible(hwnd):
                        return True
                    state = get_window_state(hwnd)
                    if state not in ("normal", "maximized"):
                        return True
                    rect = wintypes.RECT()
                    if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                        return True
                    w = rect.right - rect.left
                    h = rect.bottom - rect.top
                    if w <= MIN_WINDOW_WIDTH or h <= MIN_WINDOW_HEIGHT:
                        return True
                    if self._get_monitor_index_for_rect(rect) != mon_idx:
                        return True
                    title_buf = ctypes.create_unicode_buffer(256)
                    user32.GetWindowTextW(hwnd, title_buf, 256)
                    title = title_buf.value or ""
                    class_name = win32gui.GetClassName(hwnd)
                    if not is_useful_window(title, class_name, hwnd):
                        return True
                    if hwnd in seen:
                        return True
                    choices.append((hwnd, self._build_window_descriptor(hwnd)))
                    seen.add(hwnd)
                except Exception:
                    pass
                return True

            try:
                user32.EnumWindows(
                    ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)(enum), 0
                )
            except Exception:
                pass

        return choices

    def _apply_manual_layout(self, mon_idx, layout, info, assignments, target_ws=None, activate_target=True):
        """Apply manual layout with explicit slot assignments."""
        capacity = self._layout_capacity(layout, info)
        if not assignments or len(assignments) > capacity:
            return False

        selected_hwnds = set(assignments.values())
        with self.lock:
            active_ws = self.current_workspace.get(mon_idx, 0)
        if target_ws is None or target_ws not in (0, 1, 2):
            target_ws = active_ws
        with self.lock:
            self.workspace_layout_signature[(mon_idx, target_ws)] = (layout, info)

        workspace_layout = {
            hwnd: {"pos": (0, 0, 800, 600), "grid": (int(col), int(row)), "state": "normal"}
            for (col, row), hwnd in assignments.items()
        }

        # Editing an inactive workspace only updates its map; visual apply happens on switch.
        if (not activate_target) or (target_ws != active_ws):
            with self.lock:
                ws_list = self.workspaces.get(mon_idx, [])
                if 0 <= target_ws < len(ws_list):
                    ws_list[target_ws] = workspace_layout
            return True

        monitor_rect = self.monitors_cache[mon_idx]
        positions, grid_coords = self.layout_engine.calculate_positions(
            monitor_rect, capacity, self.gap, self.edge_padding, layout, info
        )
        pos_map = dict(zip(grid_coords, positions))

        # Ensure selected windows are visible before tiling.
        for hwnd in selected_hwnds:
            state = get_window_state(hwnd)
            if state in ("minimized", "maximized"):
                try:
                    user32.ShowWindowAsync(hwnd, SW_RESTORE)
                except Exception:
                    pass
            elif state == "hidden":
                try:
                    user32.ShowWindowAsync(hwnd, SW_SHOWNORMAL)
                except Exception:
                    pass

        visible = self.window_mgr.get_visible_windows(self.monitors_cache, self.overlay_hwnd)
        visible_on_mon = []
        for hwnd, _title, rect in visible:
            if self._get_monitor_index_for_rect(rect) == mon_idx:
                visible_on_mon.append(hwnd)

        with self._tiling_lock:
            with self.lock:
                for hwnd in selected_hwnds:
                    self.window_mgr.override_windows.discard(hwnd)
                # Hide non-selected visible windows on this monitor, but do not persist
                # manager-specific exclusions in override_windows.
                for hwnd in visible_on_mon:
                    if hwnd not in selected_hwnds:
                        pos = self.window_mgr.grid_state.get(hwnd)
                        if pos:
                            self.window_mgr.minimized_windows[hwnd] = pos
                        self.window_mgr.maximized_windows.pop(hwnd, None)
                        self.window_state_ws[hwnd] = active_ws
                        try:
                            user32.ShowWindowAsync(hwnd, win32con.SW_MINIMIZE)
                        except Exception:
                            pass
                    else:
                        self.window_mgr.override_windows.discard(hwnd)
                        self.window_state_ws.pop(hwnd, None)
                        self.window_mgr.minimized_windows.pop(hwnd, None)

                # Remove any non-selected windows from grid_state on this monitor.
                for hwnd, (m, c, r) in list(self.window_mgr.grid_state.items()):
                    if m == mon_idx and hwnd not in selected_hwnds:
                        self.window_mgr.grid_state.pop(hwnd, None)
                        if hwnd not in self.window_mgr.minimized_windows:
                            self.window_mgr.minimized_windows[hwnd] = (m, c, r)
                        self.window_state_ws[hwnd] = active_ws

                # Apply assigned slots.
                for (col, row), hwnd in assignments.items():
                    self.window_mgr.grid_state[hwnd] = (mon_idx, col, row)
                    self.window_state_ws.pop(hwnd, None)
                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    self.window_mgr.maximized_windows.pop(hwnd, None)
                ws_list = self.workspaces.get(mon_idx, [])
                if 0 <= target_ws < len(ws_list):
                    ws_list[target_ws] = workspace_layout

            for (col, row), hwnd in assignments.items():
                if (col, row) in pos_map:
                    x, y, w, h = pos_map[(col, row)]
                    self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                    time.sleep(0.01)

        self.layout_signature[mon_idx] = (layout, info)
        self.layout_capacity[mon_idx] = capacity
        self.ignore_retile_until = time.time() + 0.3
        return True
    
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
                self._backfill_window_state_ws_locked()
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

                prev_sig = self.layout_signature.get(mon_idx)
                layout, info = self.layout_engine.choose_layout(effective_count)
                capacity = self._layout_capacity(layout, info)
                layout_changed = prev_sig is not None and prev_sig != (layout, info)
                self.layout_signature[mon_idx] = (layout, info)
                self.layout_capacity[mon_idx] = capacity
                active_ws = self.current_workspace.get(mon_idx, 0)
                self.workspace_layout_signature[(mon_idx, active_ws)] = (layout, info)
                self._tile_monitor(
                    mon_idx,
                    windows,
                    new_grid,
                    layout,
                    info,
                    capacity,
                    reserved_slots=reserved_slots,
                    compact_after_restore=layout_changed,
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
            state_ws_snapshot = dict(self.window_state_ws)
        
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
                mon_idx = physical_mon_idx
                col, row = 0, 0
                active_ws = self.current_workspace.get(mon_idx, 0)
                origin_ws = state_ws_snapshot.get(hwnd)
                slot = self._get_workspace_slot(mon_idx, active_ws, hwnd)
                can_restore = slot is not None and (origin_ws is None or origin_ws == active_ws)
                if can_restore:
                    col, row = slot
                    log(f"[RESTORE] Restoring minimized window to ({col},{row}): {title[:40]}")
                with self.lock:
                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    self.window_state_ws.pop(hwnd, None)
            elif hwnd in maximized_snapshot:
                mon_idx = physical_mon_idx
                col, row = 0, 0
                active_ws = self.current_workspace.get(mon_idx, 0)
                origin_ws = state_ws_snapshot.get(hwnd)
                slot = self._get_workspace_slot(mon_idx, active_ws, hwnd)
                can_restore = slot is not None and (origin_ws is None or origin_ws == active_ws)
                if can_restore:
                    col, row = slot
                    log(f"[RESTORE] Restoring maximized window to ({col},{row}): {title[:40]}")
                with self.lock:
                    self.window_mgr.maximized_windows.pop(hwnd, None)
                    self.window_state_ws.pop(hwnd, None)
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
    
    def _tile_monitor(
        self,
        mon_idx,
        windows,
        new_grid,
        layout=None,
        info=None,
        capacity=None,
        reserved_slots=None,
        compact_after_restore=False,
    ):
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

        # Layout changed: keep restore-first behavior, then compact holes.
        if compact_after_restore:
            coord_order = {coord: idx for idx, coord in enumerate(grid_coords)}
            tiled = []
            for hwnd, (m, c, r) in new_grid.items():
                coord = (c, r)
                if m != mon_idx or coord not in coord_order:
                    continue
                tiled.append((coord_order[coord], hwnd, coord))

            if len(tiled) > 1:
                tiled.sort(key=lambda item: item[0])
                target_coords = grid_coords[:len(tiled)]

                for target_idx, (_old_idx, hwnd, old_coord) in enumerate(tiled):
                    target_coord = target_coords[target_idx]
                    if old_coord == target_coord:
                        continue

                    x, y, w, h = pos_map[target_coord]
                    self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                    new_grid[hwnd] = (mon_idx, target_coord[0], target_coord[1])
                    log(
                        f"   ↻ COMPACT ({old_coord[0]},{old_coord[1]}) -> "
                        f"({target_coord[0]},{target_coord[1]})"
                    )
                    time.sleep(0.01)

    def _sync_window_state_changes(self):
        """Track min/max/restore state transitions for stable retile decisions."""
        restored = []
        minimized_moved = 0
        maximized_moved = 0
        with self.lock:
            # Cleanup stale state markers.
            for hwnd in list(self.window_state_ws.keys()):
                if not user32.IsWindow(hwnd):
                    self.window_state_ws.pop(hwnd, None)
            for hwnd in list(self.minimize_restore_snapshots.keys()):
                if not user32.IsWindow(hwnd):
                    self.minimize_restore_snapshots.pop(hwnd, None)

            # Move minimized/maximized windows out of grid_state (keep their slot)
            for hwnd, (mon, col, row) in list(self.window_mgr.grid_state.items()):
                if not user32.IsWindow(hwnd):
                    continue
                state = get_window_state(hwnd)
                if state == 'minimized':
                    # Capture full monitor snapshot before removing the minimized window from grid_state.
                    # This allows restoring all windows to their exact pre-minimize slots if context matches.
                    snapshot_slots = {
                        h: (c, r)
                        for h, (m, c, r) in self.window_mgr.grid_state.items()
                        if m == mon and user32.IsWindow(h)
                    }
                    self.minimize_restore_snapshots[hwnd] = {
                        "monitor": mon,
                        "workspace": self.current_workspace.get(mon, 0),
                        "slots": snapshot_slots,
                        "layout": self.layout_signature.get(mon),
                        "captured_at": time.time(),
                    }
                    self.window_mgr.minimized_windows[hwnd] = (mon, col, row)
                    self.window_state_ws[hwnd] = self.current_workspace.get(mon, 0)
                    self.window_mgr.grid_state.pop(hwnd, None)
                    minimized_moved += 1
                elif state == 'maximized':
                    self.minimize_restore_snapshots.pop(hwnd, None)
                    self.window_mgr.maximized_windows[hwnd] = (mon, col, row)
                    self.window_state_ws[hwnd] = self.current_workspace.get(mon, 0)
                    self.window_mgr.grid_state.pop(hwnd, None)
                    maximized_moved += 1

            # Restore minimized windows that returned to normal
            for hwnd, (mon, col, row) in list(self.window_mgr.minimized_windows.items()):
                if not user32.IsWindow(hwnd):
                    self.minimize_restore_snapshots.pop(hwnd, None)
                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    continue
                state = get_window_state(hwnd)
                if state == 'normal' and user32.IsWindowVisible(hwnd):
                    snapshot = self.minimize_restore_snapshots.pop(hwnd, None)
                    # Restore to the exact slot captured at minimize time.
                    # This mirrors maximize restore behavior and avoids slot drift/swap.
                    target_mon = mon
                    if target_mon < 0 or target_mon >= len(self.monitors_cache):
                        rect = wintypes.RECT()
                        if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                            target_mon = self._get_monitor_index_for_rect(rect)
                        else:
                            target_mon = 0

                    active_ws = self.current_workspace.get(target_mon, 0)
                    origin_ws = self.window_state_ws.get(hwnd)
                    if origin_ws is not None and origin_ws != active_ws:
                        # Window restored from another workspace context: do not force old slot.
                        self.window_mgr.minimized_windows.pop(hwnd, None)
                        self.window_state_ws.pop(hwnd, None)
                        continue

                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    self.window_state_ws.pop(hwnd, None)
                    self.window_mgr.grid_state[hwnd] = (target_mon, col, row)
                    restored.append((hwnd, target_mon, col, row, snapshot))
                elif state == 'maximized':
                    self.minimize_restore_snapshots.pop(hwnd, None)
                    self.window_mgr.maximized_windows[hwnd] = (mon, col, row)
                    self.window_state_ws[hwnd] = self.current_workspace.get(mon, 0)
                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    maximized_moved += 1

            # Restore maximized windows that returned to normal
            for hwnd, (mon, col, row) in list(self.window_mgr.maximized_windows.items()):
                if not user32.IsWindow(hwnd):
                    self.window_mgr.maximized_windows.pop(hwnd, None)
                    continue
                state = get_window_state(hwnd)
                if state == 'normal' and user32.IsWindowVisible(hwnd):
                    # Critical: restore to the exact slot captured at maximize time.
                    # Do not recalculate from workspace maps here, otherwise stale
                    # workspace snapshots can cause unintended slot swaps.
                    target_mon = mon
                    if target_mon < 0 or target_mon >= len(self.monitors_cache):
                        # Monitor topology changed while maximized: fallback to current physical monitor.
                        rect = wintypes.RECT()
                        if user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                            target_mon = self._get_monitor_index_for_rect(rect)
                        else:
                            target_mon = 0
                    self.window_mgr.maximized_windows.pop(hwnd, None)
                    self.window_state_ws.pop(hwnd, None)
                    self.window_mgr.grid_state[hwnd] = (target_mon, col, row)
                    restored.append((hwnd, target_mon, col, row))
                elif state == 'minimized':
                    self.window_mgr.minimized_windows[hwnd] = (mon, col, row)
                    self.window_state_ws[hwnd] = self.current_workspace.get(mon, 0)
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
        count = 0
        with self.lock:
            for hwnd, (m, _, _) in self.window_mgr.grid_state.items():
                if m == mon_idx and user32.IsWindow(hwnd):
                    count += 1
        if count <= 0:
            return 0
        layout, info = self.layout_engine.choose_layout(count)
        return self._layout_capacity(layout, info)

    def _find_workspace_owner(self, mon_idx, hwnd, prefer_ws=None):
        """Return workspace index that references hwnd on monitor mon_idx, else None."""
        ws_list = self.workspaces.get(mon_idx, [])
        if prefer_ws is not None and 0 <= prefer_ws < len(ws_list):
            if hwnd in ws_list[prefer_ws]:
                return prefer_ws
        for ws_idx, ws_map in enumerate(ws_list):
            if hwnd in ws_map:
                return ws_idx
        return None

    def _get_workspace_slot(self, mon_idx, ws_idx, hwnd):
        """Return (col,row) from saved workspace map if present and valid."""
        ws_list = self.workspaces.get(mon_idx, [])
        if ws_idx is None or ws_idx < 0 or ws_idx >= len(ws_list):
            return None
        data = ws_list[ws_idx].get(hwnd)
        if not isinstance(data, dict):
            return None
        grid = data.get("grid")
        if isinstance(grid, (list, tuple)) and len(grid) == 2:
            try:
                return int(grid[0]), int(grid[1])
            except Exception:
                return None
        return None

    def _backfill_window_state_ws_locked(self):
        """Fill missing workspace markers for windows cached in min/max maps."""
        for hwnd, (mon, _, _) in self.window_mgr.minimized_windows.items():
            if hwnd not in self.window_state_ws:
                active_ws = self.current_workspace.get(mon, 0)
                owner_ws = self._find_workspace_owner(mon, hwnd, prefer_ws=active_ws)
                self.window_state_ws[hwnd] = owner_ws if owner_ws is not None else active_ws
        for hwnd, (mon, _, _) in self.window_mgr.maximized_windows.items():
            if hwnd not in self.window_state_ws:
                active_ws = self.current_workspace.get(mon, 0)
                owner_ws = self._find_workspace_owner(mon, hwnd, prefer_ws=active_ws)
                self.window_state_ws[hwnd] = owner_ws if owner_ws is not None else active_ws

    def _backfill_window_state_ws(self):
        with self.lock:
            self._backfill_window_state_ws_locked()

    def _restore_windows_to_slots(self, restored):
        """Place restored windows back into their saved slots."""
        with self._tiling_lock:
            with self.lock:
                grid_snapshot = dict(self.window_mgr.grid_state)
                layout_sig_snapshot = dict(self.layout_signature)
                ws_layout_sig_snapshot = dict(self.workspace_layout_signature)
                current_ws_snapshot = dict(self.current_workspace)

            def _apply_snapshot_restore_if_possible(mon_idx, snapshot):
                if not isinstance(snapshot, dict):
                    return False
                snap_mon = snapshot.get("monitor")
                if snap_mon != mon_idx:
                    return False
                snap_ws = snapshot.get("workspace")
                if snap_ws != self.current_workspace.get(mon_idx, 0):
                    return False

                slots = snapshot.get("slots")
                if not isinstance(slots, dict) or not slots:
                    return False

                current_mon_windows = {
                    h for h, (m, _c, _r) in grid_snapshot.items()
                    if m == mon_idx and user32.IsWindow(h)
                }
                snapshot_windows = {h for h in slots.keys() if user32.IsWindow(h)}
                if not snapshot_windows or current_mon_windows != snapshot_windows:
                    return False

                layout_sig = snapshot.get("layout")
                if not layout_sig:
                    layout_sig = self.layout_engine.choose_layout(len(snapshot_windows))
                layout, info = layout_sig
                capacity = self._layout_capacity(layout, info)
                positions, grid_coords = self.layout_engine.calculate_positions(
                    self.monitors_cache[mon_idx], capacity, self.gap, self.edge_padding, layout, info
                )
                pos_map = dict(zip(grid_coords, positions))
                coord_order = {coord: i for i, coord in enumerate(grid_coords)}

                restored_slots = {}
                used_coords = set()
                for h in snapshot_windows:
                    coord = slots.get(h)
                    if not isinstance(coord, (list, tuple)) or len(coord) != 2:
                        return False
                    try:
                        c = int(coord[0])
                        r = int(coord[1])
                    except Exception:
                        return False
                    key = (c, r)
                    if key not in pos_map or key in used_coords:
                        return False
                    used_coords.add(key)
                    restored_slots[h] = key

                ordered_hwnds = sorted(
                    snapshot_windows,
                    key=lambda h: coord_order.get(restored_slots[h], 10_000),
                )

                with self.lock:
                    for h in ordered_hwnds:
                        c, r = restored_slots[h]
                        self.window_mgr.grid_state[h] = (mon_idx, c, r)
                        grid_snapshot[h] = (mon_idx, c, r)

                for h in ordered_hwnds:
                    c, r = restored_slots[h]
                    x, y, w, hh = pos_map[(c, r)]
                    self.window_mgr.force_tile_resizable(h, x, y, w, hh)
                    time.sleep(0.006)
                return True

            def _get_pos_map_for_exact_slot(mon_idx, col, row):
                """Resolve a monitor position map that contains (col,row), preferring remembered layout."""
                if mon_idx < 0 or mon_idx >= len(self.monitors_cache):
                    return None

                candidates = []
                sig = layout_sig_snapshot.get(mon_idx)
                if sig:
                    candidates.append(sig)

                ws_idx = current_ws_snapshot.get(mon_idx, 0)
                ws_sig = ws_layout_sig_snapshot.get((mon_idx, ws_idx))
                if ws_sig and ws_sig not in candidates:
                    candidates.append(ws_sig)

                mon_window_count = sum(
                    1 for h, (m, _c, _r) in grid_snapshot.items()
                    if m == mon_idx and user32.IsWindow(h)
                )
                if mon_window_count > 0:
                    auto_sig = self.layout_engine.choose_layout(mon_window_count)
                    if auto_sig not in candidates:
                        candidates.append(auto_sig)

                # Fallbacks for dense layouts where count can be transient during staggered unmaximize.
                for fallback_sig in [
                    ("grid", (5, 3)),
                    ("grid", (4, 3)),
                    ("grid", (3, 3)),
                    ("grid", (3, 2)),
                    ("grid", (2, 2)),
                    ("master_stack", None),
                    ("side_by_side", None),
                    ("full", None),
                ]:
                    if fallback_sig not in candidates:
                        candidates.append(fallback_sig)

                target = (col, row)
                for layout, info in candidates:
                    cap = self._layout_capacity(layout, info)
                    positions, grid_coords = self.layout_engine.calculate_positions(
                        self.monitors_cache[mon_idx], cap, self.gap, self.edge_padding, layout, info
                    )
                    if target not in set(grid_coords):
                        continue
                    return dict(zip(grid_coords, positions))
                return None

            def _get_inner_rect(hwnd):
                """Return client-like rect (x, y, w, h) corrected from extended frame borders."""
                rect = wintypes.RECT()
                if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                    return None
                lb, tb, rb, bb = get_frame_borders(hwnd)
                return (
                    rect.left + lb,
                    rect.top + tb,
                    rect.right - rect.left - lb - rb,
                    rect.bottom - rect.top - tb - bb,
                )

            def _slot_is_drifted(hwnd, target_x, target_y, target_w, target_h, tol_pos=8, tol_size=8):
                cur = _get_inner_rect(hwnd)
                if not cur:
                    return False
                cur_x, cur_y, cur_w, cur_h = cur
                return (
                    abs(cur_x - target_x) > tol_pos
                    or abs(cur_y - target_y) > tol_pos
                    or abs(cur_w - target_w) > tol_size
                    or abs(cur_h - target_h) > tol_size
                )

            def _repack_monitor_without_holes(mon_idx, preferred_slots=None):
                """Pack monitor windows into earliest slots to avoid restore holes."""
                if mon_idx < 0 or mon_idx >= len(self.monitors_cache):
                    return False

                mon_windows = [
                    (h, c, r)
                    for h, (m, c, r) in grid_snapshot.items()
                    if m == mon_idx and user32.IsWindow(h)
                ]
                if not mon_windows:
                    return False

                count = len(mon_windows)
                layout, info = self.layout_engine.choose_layout(count)
                capacity = self._layout_capacity(layout, info)
                positions, grid_coords = self.layout_engine.calculate_positions(
                    self.monitors_cache[mon_idx], capacity, self.gap, self.edge_padding, layout, info
                )
                if not grid_coords:
                    return False

                packed_coords = list(grid_coords[:count])
                coord_order = {coord: i for i, coord in enumerate(grid_coords)}
                packed_set = set(packed_coords)
                pos_map = dict(zip(grid_coords, positions))
                present_hwnds = {h for h, _c, _r in mon_windows}
                preferred_slots = preferred_slots or {}

                assigned = {}
                for h, coord in preferred_slots.items():
                    if h not in present_hwnds:
                        continue
                    if not isinstance(coord, (list, tuple)) or len(coord) != 2:
                        continue
                    try:
                        key = (int(coord[0]), int(coord[1]))
                    except Exception:
                        continue
                    if key in packed_set and key not in assigned.values():
                        assigned[h] = key

                available_coords = [coord for coord in packed_coords if coord not in assigned.values()]

                def _coord_idx(c, r):
                    return coord_order.get((c, r), 10_000)

                remaining = [
                    (h, c, r)
                    for h, c, r in mon_windows
                    if h not in assigned
                ]
                remaining.sort(key=lambda item: (_coord_idx(item[1], item[2]), item[0]))
                for idx, (h, _c, _r) in enumerate(remaining):
                    if idx >= len(available_coords):
                        break
                    assigned[h] = available_coords[idx]

                ordered_hwnds = sorted(
                    assigned.keys(),
                    key=lambda h: coord_order.get(assigned[h], 10_000),
                )

                with self.lock:
                    self.layout_signature[mon_idx] = (layout, info)
                    self.layout_capacity[mon_idx] = capacity
                    for h in ordered_hwnds:
                        c, r = assigned[h]
                        self.window_mgr.grid_state[h] = (mon_idx, c, r)
                        grid_snapshot[h] = (mon_idx, c, r)

                for h in ordered_hwnds:
                    c, r = assigned[h]
                    x, y, w, hh = pos_map[(c, r)]
                    self.window_mgr.force_tile_resizable(h, x, y, w, hh)
                    time.sleep(0.006)
                return True

            need_full_retile = False
            snapshot_restored_monitors = set()
            fallback_restore_slots = {}
            for item in restored:
                snapshot = None
                from_minimize_restore = len(item) >= 5
                if len(item) >= 5:
                    hwnd, mon_idx, col, row, snapshot = item
                else:
                    hwnd, mon_idx, col, row = item
                if mon_idx >= len(self.monitors_cache):
                    continue
                if not user32.IsWindow(hwnd):
                    continue
                if mon_idx in snapshot_restored_monitors:
                    continue

                if snapshot and _apply_snapshot_restore_if_possible(mon_idx, snapshot):
                    snapshot_restored_monitors.add(mon_idx)
                    continue

                # Maximized -> normal restore: enforce exact saved slot directly using
                # remembered monitor/workspace layout, not transient visible-count layout.
                if not from_minimize_restore:
                    pos_map_exact = _get_pos_map_for_exact_slot(mon_idx, col, row)
                    target_exact = (col, row)
                    if pos_map_exact and target_exact in pos_map_exact:
                        with self.lock:
                            self.window_mgr.grid_state[hwnd] = (mon_idx, col, row)
                        grid_snapshot[hwnd] = (mon_idx, col, row)
                        x, y, w, h = pos_map_exact[target_exact]
                        self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                        # Some apps apply a delayed self-resize right after unmaximize.
                        # Re-check briefly and clamp again only when drift is detected.
                        for settle_delay in (0.05, 0.14):
                            if not user32.IsWindow(hwnd):
                                break
                            if get_window_state(hwnd) != "normal":
                                break
                            time.sleep(settle_delay)
                            if not _slot_is_drifted(hwnd, x, y, w, h):
                                break
                            log(
                                f"[RESTORE] Re-clamp maximize restore hwnd={hwnd} "
                                f"slot=({col},{row})"
                            )
                            self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                        continue
                    # If no exact map is available, avoid aggressive fallback/repack that can
                    # shuffle slots; keep current slot hint and wait for next stable pass.
                    log(f"[RESTORE] Deferred exact slot map for maximized restore hwnd={hwnd} ({col},{row})")
                    continue

                # Only minimized restores need post-repack fallback.
                # Maximized -> normal should keep exact slots and must not trigger compaction-style repacks.
                if from_minimize_restore:
                    fallback_restore_slots.setdefault(mon_idx, {})[hwnd] = (col, row)

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
                with self.lock:
                    grid_snapshot = dict(self.window_mgr.grid_state)

            for mon_idx, preferred in fallback_restore_slots.items():
                if mon_idx in snapshot_restored_monitors:
                    continue
                _repack_monitor_without_holes(mon_idx, preferred_slots=preferred)

    def _compact_grid_after_minimize(self):
        """Fill earliest empty slots by moving windows from the end of the layout."""
        with self._tiling_lock:
            with self.lock:
                grid_snapshot = dict(self.window_mgr.grid_state)
                layout_signature = dict(self.layout_signature)
                layout_capacity = dict(self.layout_capacity)

            # Hybrid behavior: if the window count implies a different layout, retile fully.
            counts_by_monitor = {}
            for hwnd, (mon_idx, _, _) in grid_snapshot.items():
                if user32.IsWindow(hwnd):
                    counts_by_monitor[mon_idx] = counts_by_monitor.get(mon_idx, 0) + 1
            for mon_idx, count in counts_by_monitor.items():
                if count <= 0:
                    continue
                desired_layout, desired_info = self.layout_engine.choose_layout(count)
                current_sig = layout_signature.get(mon_idx)
                if current_sig is None or current_sig != (desired_layout, desired_info):
                    with self.lock:
                        self.ignore_retile_until = 0.0
                    self.smart_tile_with_restore()
                    return

            for mon_idx, (layout, info) in layout_signature.items():
                capacity = layout_capacity.get(mon_idx, 0)
                if capacity <= 0 or mon_idx >= len(self.monitors_cache):
                    continue

                positions, grid_coords = self.layout_engine.calculate_positions(
                    self.monitors_cache[mon_idx],
                    capacity,
                    self.gap,
                    self.edge_padding,
                    layout,
                    info,
                )
                pos_map = dict(zip(grid_coords, positions))
                order_index = {coord: i for i, coord in enumerate(grid_coords)}

                slot_to_hwnd = {}
                for hwnd, (m, col, row) in grid_snapshot.items():
                    coord = (col, row)
                    if m != mon_idx or coord not in pos_map:
                        continue
                    if not user32.IsWindow(hwnd):
                        continue
                    slot_to_hwnd[coord] = hwnd

                if not slot_to_hwnd:
                    continue

                empty_indices = [
                    order_index[coord]
                    for coord in grid_coords
                    if coord not in slot_to_hwnd
                ]
                if not empty_indices:
                    continue

                empty_indices.sort()
                filled_indices = sorted(order_index[coord] for coord in slot_to_hwnd.keys())

                while empty_indices and filled_indices and filled_indices[-1] > empty_indices[0]:
                    donor_idx = filled_indices.pop(-1)
                    target_idx = empty_indices.pop(0)

                    donor_coord = grid_coords[donor_idx]
                    target_coord = grid_coords[target_idx]
                    hwnd = slot_to_hwnd.pop(donor_coord)

                    x, y, w, h = pos_map[target_coord]
                    self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                    with self.lock:
                        self.window_mgr.grid_state[hwnd] = (mon_idx, target_coord[0], target_coord[1])
                    slot_to_hwnd[target_coord] = hwnd
                    bisect.insort(filled_indices, target_idx)

    def _compact_grid_after_close(self):
        """Compact grid after a window closes (hybrid layout change)."""
        self._compact_grid_after_minimize()

    def _run_deferred_compactions(self):
        """Apply deferred compact operations once locks/freeze allow it."""
        with self.lock:
            has_maximized = any(
                user32.IsWindow(hwnd) for hwnd in self.window_mgr.maximized_windows.keys()
            )
            do_minimize = self.compact_on_minimize and self._pending_compact_minimize
            do_close = self.compact_on_close and self._pending_compact_close

        if has_maximized or not (do_minimize or do_close):
            return False

        # Prioritize minimize compaction when both are pending.
        if do_minimize:
            log("[AUTO-COMPACT] running deferred minimize compaction")
            self._compact_grid_after_minimize()
        else:
            log("[AUTO-COMPACT] running deferred close compaction")
            self._compact_grid_after_close()

        now = time.time()
        visible_windows = self.window_mgr.get_visible_windows(
            self.monitors_cache, self.overlay_hwnd
        )
        current_count = len(visible_windows)
        with self.lock:
            known_hwnds = (
                set(self.window_mgr.grid_state.keys())
                | set(self.window_mgr.minimized_windows.keys())
                | set(self.window_mgr.maximized_windows.keys())
            )
            if do_minimize:
                self._pending_compact_minimize = False
            if do_close:
                self._pending_compact_close = False

        self.last_visible_count = current_count
        self.last_known_count = len(known_hwnds)
        self.last_retile_time = now
        return True

    def _sync_manual_cross_monitor_moves(self):
        """
        Detect tiled windows manually dragged to another monitor and trigger a retile.
        This covers plain OS drags that bypass SmartGrid drag/snap drop handling.
        """
        try:
            # Avoid reassigning while the user is still holding the mouse button.
            if win32api.GetAsyncKeyState(win32con.VK_LBUTTON) & 0x8000:
                return False
        except Exception:
            pass

        with self.lock:
            grid_snapshot = list(self.window_mgr.grid_state.items())

        moved = []
        touched_monitors = set()
        for hwnd, (mon_idx, col, row) in grid_snapshot:
            if not user32.IsWindow(hwnd):
                continue
            if get_window_state(hwnd) != "normal":
                continue
            rect = wintypes.RECT()
            if not user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                continue
            physical_mon = self._get_monitor_index_for_rect(rect)
            if physical_mon != mon_idx:
                moved.append((hwnd, mon_idx, physical_mon, col, row))
                touched_monitors.add(mon_idx)
                touched_monitors.add(physical_mon)

        if not moved:
            return False

        with self.lock:
            for hwnd, old_mon, new_mon, col, row in moved:
                cur = self.window_mgr.grid_state.get(hwnd)
                if not cur:
                    continue
                cur_mon, cur_col, cur_row = cur
                if cur_mon != old_mon:
                    continue
                # Keep slot hint; tiler will remap safely if slot is invalid on target layout.
                self.window_mgr.grid_state[hwnd] = (new_mon, cur_col if cur_col is not None else col, cur_row if cur_row is not None else row)

            for mon in touched_monitors:
                self.layout_signature.pop(mon, None)
                self.layout_capacity.pop(mon, None)
            self.ignore_retile_until = 0.0

        log(f"[MONITOR-DRIFT] Reassigned {len(moved)} window(s) after manual cross-monitor drag")
        self.smart_tile_with_restore()
        return True
    
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

                    # Only consider drag from title-like areas. Some modern apps report
                    # HTCLIENT on custom title bars (Teams/Electron/Chromium), so allow
                    # HTCLIENT only in a constrained top region and exclude the right-side
                    # caption controls zone.
                    if hwnd and user32.IsWindowVisible(hwnd):
                        try:
                            lparam = win32api.MAKELONG(pt[0] & 0xFFFF, pt[1] & 0xFFFF)
                            hit = win32gui.SendMessage(hwnd, WM_NCHITTEST, 0, lparam)
                        except Exception:
                            hit = None

                        allow_drag_source = False
                        if hit == HTCAPTION:
                            allow_drag_source = True
                        elif hit == HTCLIENT:
                            try:
                                rect = win32gui.GetWindowRect(hwnd)
                                ww = max(1, rect[2] - rect[0])
                                wh = max(1, rect[3] - rect[1])
                                rel_x = pt[0] - rect[0]
                                rel_y = pt[1] - rect[1]

                                # Heuristic title bar zone (top strip only).
                                title_zone_h = max(28, min(72, int(wh * 0.14)))
                                in_title_zone = 0 <= rel_y <= title_zone_h

                                # Exclude right controls region (min/max/close).
                                controls_zone_w = max(120, min(220, int(ww * 0.20)))
                                in_controls_zone = (rel_x >= ww - controls_zone_w) and (rel_y <= title_zone_h + 10)

                                allow_drag_source = in_title_zone and not in_controls_zone
                            except Exception:
                                allow_drag_source = False

                        if allow_drag_source:
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
            state_ws_snapshot = dict(self.window_state_ws)
        
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
            if mon != monitor_idx or not user32.IsWindow(hwnd):
                continue
            origin_ws = state_ws_snapshot.get(hwnd)
            if origin_ws is not None and origin_ws != ws:
                continue
            if origin_ws is None:
                owner_ws = self._find_workspace_owner(monitor_idx, hwnd, prefer_ws=ws)
                if owner_ws is not None and owner_ws != ws:
                    continue
            if hwnd in self.workspaces[monitor_idx][ws]:
                continue
            self.workspaces[monitor_idx][ws][hwnd] = {
                'pos': (0, 0, 800, 600),
                'grid': (col, row),
                'state': 'minimized'
            }
        
        # Save maximized windows
        for hwnd, (mon, col, row) in maximized_snapshot.items():
            if mon != monitor_idx or not user32.IsWindow(hwnd):
                continue
            origin_ws = state_ws_snapshot.get(hwnd)
            if origin_ws is not None and origin_ws != ws:
                continue
            if origin_ws is None:
                owner_ws = self._find_workspace_owner(monitor_idx, hwnd, prefer_ws=ws)
                if owner_ws is not None and owner_ws != ws:
                    continue
            if hwnd in self.workspaces[monitor_idx][ws]:
                continue
            self.workspaces[monitor_idx][ws][hwnd] = {
                'pos': (0, 0, 800, 600),
                'grid': (col, row),
                'state': 'maximized'
            }

        with self.lock:
            current_sig = self.layout_signature.get(monitor_idx)
            if current_sig is not None:
                self.workspace_layout_signature[(monitor_idx, ws)] = current_sig
        
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
        # Short guard while toggling window visibility during workspace load.
        self.ignore_retile_until = time.time() + 0.35
        
        grid_updates = {}
        minimized_updates = {}
        maximized_updates = {}
        
        for hwnd, data in layout.items():
            if not user32.IsWindow(hwnd):
                continue
            
            try:
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
                self.window_state_ws[hwnd] = ws_idx
                self.window_mgr.grid_state.pop(hwnd, None)
                self.window_mgr.maximized_windows.pop(hwnd, None)
            
            for hwnd, pos in maximized_updates.items():
                self.window_mgr.maximized_windows[hwnd] = pos
                self.window_state_ws[hwnd] = ws_idx
                self.window_mgr.grid_state.pop(hwnd, None)
                self.window_mgr.minimized_windows.pop(hwnd, None)
            
            for hwnd, pos in grid_updates.items():
                self.window_mgr.grid_state[hwnd] = pos
                self.window_state_ws.pop(hwnd, None)
                self.window_mgr.minimized_windows.pop(hwnd, None)
                self.window_mgr.maximized_windows.pop(hwnd, None)
        
        time.sleep(0.15)
        # Force one immediate retile for the target workspace now that grid_state is ready.
        self.ignore_retile_until = 0.0
        self.smart_tile_with_restore()
        log(f"[WS] ✓ Workspace {ws_idx+1} restored")
        # Keep a small debounce window only.
        self.ignore_retile_until = time.time() + 0.35
    
    def ws_switch(self, ws_idx):
        """Switch to specified workspace."""
        self.workspace_switching_lock = True
        time.sleep(0.25)
        was_active = self.is_active
        
        try:
            mon = self.current_monitor_index
            
            if mon not in self.workspaces:
                log(f"[WS] ✗ Monitor {mon} not initialized")
                return
            
            if ws_idx == self.current_workspace.get(mon, 0):
                log(f"[WS] Already on workspace {ws_idx+1}")
                return
            
            self.is_active = False
            # Flush pending min/max transitions before persisting workspace map.
            self._sync_window_state_changes()
            
            # Save current
            self.save_workspace(mon)
            
            # Hide and clear all runtime states from current workspace on this monitor.
            hidden = 0
            with self.lock:
                source_ws = self.current_workspace.get(mon, 0)
                source_map = dict(self.workspaces.get(mon, [{}, {}, {}])[source_ws])
                for hwnd in list(source_map.keys()):
                    if user32.IsWindow(hwnd) and user32.IsWindowVisible(hwnd):
                        user32.ShowWindowAsync(hwnd, win32con.SW_HIDE)
                        hidden += 1
                    self.window_mgr.grid_state.pop(hwnd, None)
                    self.window_mgr.minimized_windows.pop(hwnd, None)
                    self.window_mgr.maximized_windows.pop(hwnd, None)
                    self.window_state_ws.pop(hwnd, None)
            
            log(f"[WS] Hidden {hidden} windows from workspace {self.current_workspace.get(mon, 0)+1}")
            
            # Switch
            self.current_workspace[mon] = ws_idx
            # Workspace changed: drop monitor-local layout caches to avoid stale restore math.
            self.layout_signature.pop(mon, None)
            self.layout_capacity.pop(mon, None)
            
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
            
            log(f"[WS] ✓ Switched to workspace {ws_idx+1}")
        
        except Exception as e:
            log(f"[ERROR] ws_switch: {e}")
        
        finally:
            self.is_active = was_active
            self.workspace_switching_lock = False
    
    def move_current_workspace_to_next_monitor(self):
        """Move all windows from current workspace to next monitor."""
        self.move_monitor_lock = True
        time.sleep(0.25)
        
        try:
            if len(self.monitors_cache) <= 1:
                log("[MOVE] Only one monitor")
                try:
                    ctypes.windll.user32.MessageBoxW(
                        0,
                        "Move Workspace to Next Monitor is unavailable:\nonly one monitor is detected.",
                        "SmartGrid",
                        0x40 | 0x00010000 | 0x00040000
                    )
                except Exception:
                    pass
                return
            
            old_mon = self.current_monitor_index
            old_ws = self.current_workspace.get(old_mon, 0)
            new_mon = (old_mon + 1) % len(self.monitors_cache)

            # Flush pending transitions and snapshot current workspace first.
            self._sync_window_state_changes()
            self.save_workspace(old_mon)

            with self.lock:
                source_ws_map = dict(self.workspaces.get(old_mon, [{}])[old_ws])

            moved_hwnds = [hwnd for hwnd in source_ws_map.keys() if user32.IsWindow(hwnd)]
            if not moved_hwnds:
                log(f"[MOVE] No windows to move from monitor {old_mon+1} workspace {old_ws+1}")
                return

            # Separate normal/minimized/maximized so the move includes all workspace states.
            normal_hwnds = []
            minimized_hwnds = []
            maximized_hwnds = []
            for hwnd in moved_hwnds:
                state = get_window_state(hwnd)
                if state == "minimized":
                    minimized_hwnds.append(hwnd)
                elif state == "maximized":
                    maximized_hwnds.append(hwnd)
                else:
                    normal_hwnds.append(hwnd)

            monitor_rect = self.monitors_cache[new_mon]
            new_grid = {}  # hwnd -> (new_mon, col, row) for normal windows moved now.
            log(f"\n[MOVE] Moving workspace {old_ws+1} from monitor {old_mon+1} → {new_mon+1}")

            with self._tiling_lock:
                if normal_hwnds:
                    count = len(normal_hwnds)
                    layout, info = self.layout_engine.choose_layout(count)
                    positions, grid_coords = self.layout_engine.calculate_positions(
                        monitor_rect, count, self.gap, self.edge_padding, layout, info
                    )
                    for i, hwnd in enumerate(normal_hwnds):
                        x, y, w, h = positions[i]
                        self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                        col, row = grid_coords[i]
                        new_grid[hwnd] = (new_mon, col, row)
                        title = win32gui.GetWindowText(hwnd)
                        log(f"   -> {title[:50]}")
                        time.sleep(0.015)

                # Transactional state update.
                with self.lock:
                    # Remove moved windows from runtime caches on old monitor.
                    for hwnd in moved_hwnds:
                        self.window_mgr.grid_state.pop(hwnd, None)
                        self.window_mgr.minimized_windows.pop(hwnd, None)
                        self.window_mgr.maximized_windows.pop(hwnd, None)

                    # Reinsert moved normal windows on new monitor.
                    self.window_mgr.grid_state.update(new_grid)

                    # Keep moved minimized/maximized windows tracked on destination monitor.
                    for hwnd in minimized_hwnds:
                        data = source_ws_map.get(hwnd, {})
                        grid = data.get("grid", (0, 0)) if isinstance(data, dict) else (0, 0)
                        col, row = (grid[0], grid[1]) if isinstance(grid, (list, tuple)) and len(grid) == 2 else (0, 0)
                        self.window_mgr.minimized_windows[hwnd] = (new_mon, int(col), int(row))
                        self.window_state_ws[hwnd] = old_ws
                    for hwnd in maximized_hwnds:
                        data = source_ws_map.get(hwnd, {})
                        grid = data.get("grid", (0, 0)) if isinstance(data, dict) else (0, 0)
                        col, row = (grid[0], grid[1]) if isinstance(grid, (list, tuple)) and len(grid) == 2 else (0, 0)
                        self.window_mgr.maximized_windows[hwnd] = (new_mon, int(col), int(row))
                        self.window_state_ws[hwnd] = old_ws
                    for hwnd in normal_hwnds:
                        self.window_state_ws.pop(hwnd, None)

                    # Move workspace snapshot: source -> destination (same ws index), without global save pollution.
                    moved_payload = {}
                    for hwnd in moved_hwnds:
                        data = source_ws_map.get(hwnd)
                        if not isinstance(data, dict):
                            continue
                        cloned = dict(data)
                        if hwnd in new_grid:
                            _m, c, r = new_grid[hwnd]
                            cloned["grid"] = (c, r)
                            cloned["state"] = "normal"
                        moved_payload[hwnd] = cloned

                    # Remove moved hwnds from any other workspace map to avoid duplicates.
                    for ws_idx, ws_map in enumerate(self.workspaces.get(old_mon, [])):
                        if ws_idx == old_ws or not isinstance(ws_map, dict):
                            continue
                        for hwnd in moved_hwnds:
                            ws_map.pop(hwnd, None)
                    for ws_idx, ws_map in enumerate(self.workspaces.get(new_mon, [])):
                        if not isinstance(ws_map, dict):
                            continue
                        for hwnd in moved_hwnds:
                            ws_map.pop(hwnd, None)

                    self.workspaces[old_mon][old_ws] = {}
                    self.workspaces[new_mon][old_ws] = moved_payload
                    moved_sig = self.workspace_layout_signature.pop((old_mon, old_ws), None)
                    if moved_sig is not None:
                        self.workspace_layout_signature[(new_mon, old_ws)] = moved_sig

                    # Invalidate stale monitor layout caches after move.
                    self.layout_signature.pop(old_mon, None)
                    self.layout_capacity.pop(old_mon, None)
                    self.layout_signature.pop(new_mon, None)
                    self.layout_capacity.pop(new_mon, None)

                    # Update active workspace on destination monitor, then switch current monitor.
                    self.current_workspace[new_mon] = old_ws
                    self.current_monitor_index = new_mon

            time.sleep(0.15)

            # Reflow destination monitor once after transaction.
            self.ignore_retile_until = 0.0
            self.smart_tile_with_restore()

            active = user32.GetForegroundWindow()
            if active and user32.IsWindowVisible(active):
                self.window_mgr.apply_border(active, 0x0000FF00)

            # Refresh counters/debounce to avoid immediate storm.
            visible_windows = self.window_mgr.get_visible_windows(
                self.monitors_cache, self.overlay_hwnd
            )
            self.last_visible_count = len(visible_windows)
            with self.lock:
                self.last_known_count = (
                    len(self.window_mgr.grid_state)
                    + len(self.window_mgr.minimized_windows)
                    + len(self.window_mgr.maximized_windows)
                )
            now = time.time()
            self.last_retile_time = now
            self.ignore_retile_until = max(self.ignore_retile_until, now + 0.35)

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

    def _center_tk_window(self, window, width, height, monitor_idx=None):
        """Center a Tk/Toplevel window on the target monitor work area."""
        try:
            window.update_idletasks()
            w = int(width)
            h = int(height)

            x = y = None
            if monitor_idx is not None:
                monitors = self.monitors_cache or get_monitors()
                if 0 <= monitor_idx < len(monitors):
                    mx, my, mw, mh = monitors[monitor_idx]
                    x = int(mx + (mw - w) // 2)
                    y = int(my + (mh - h) // 2)

            if x is None or y is None:
                screen_w = window.winfo_screenwidth()
                screen_h = window.winfo_screenheight()
                x = max(0, (screen_w - w) // 2)
                y = max(0, (screen_h - h) // 2)

            # Tk geometry expects signed offsets without a redundant '+' for negatives.
            x_part = f"+{x}" if x >= 0 else f"{x}"
            y_part = f"+{y}" if y >= 0 else f"{y}"
            window.geometry(f"{w}x{h}{x_part}{y_part}")
        except Exception:
            # Non-fatal: fallback to caller's existing geometry.
            pass
    
    def show_settings_dialog(self):
        """Show settings dialog to modify layout, speed, animation, and compact options."""
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
            dialog.attributes('-topmost', True)
            dialog.resizable(True, True)
            self._center_tk_window(dialog, 500, 680, monitor_idx=self.current_monitor_index)
            
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

            # Retile speed and stability
            speed_frame = tk.LabelFrame(dialog, text="Retile Speed & Stability", padx=10, pady=10)
            speed_frame.pack(fill=tk.X, padx=15, pady=5)

            debounce_values = [50, 100, 200, 350, 500, 750, 1000, 1500, 2000, 3000]
            current_debounce_ms = int(round(float(self.retile_debounce) * 1000.0))
            current_debounce_idx = min(
                range(len(debounce_values)),
                key=lambda i: abs(debounce_values[i] - current_debounce_ms)
            )
            debounce_idx_var = tk.IntVar(value=current_debounce_idx)
            debounce_display_var = tk.StringVar()
            debounce_profile_var = tk.StringVar()

            def set_debounce_label():
                idx = max(0, min(len(debounce_values) - 1, int(debounce_idx_var.get())))
                val = debounce_values[idx]
                if val <= 150:
                    profile = "Very Fast"
                elif val <= 500:
                    profile = "Fast"
                elif val <= 1000:
                    profile = "Balanced"
                elif val <= 2000:
                    profile = "Stable"
                else:
                    profile = "Very Stable"
                debounce_display_var.set(f"{val} ms")
                debounce_profile_var.set(profile)
                debounce_idx_var.set(idx)

            def change_debounce(delta):
                debounce_idx_var.set(int(debounce_idx_var.get()) + delta)
                set_debounce_label()

            debounce_row = tk.Frame(speed_frame)
            debounce_row.pack(fill=tk.X)
            tk.Label(debounce_row, text="Retile Debounce", width=18, anchor="w").pack(side=tk.LEFT, padx=2)
            tk.Button(debounce_row, text="-", width=2, command=lambda: change_debounce(-1)).pack(side=tk.LEFT, padx=(2, 1))
            tk.Label(debounce_row, textvariable=debounce_display_var, width=9, anchor="center").pack(side=tk.LEFT, padx=2)
            tk.Button(debounce_row, text="+", width=2, command=lambda: change_debounce(1)).pack(side=tk.LEFT, padx=(1, 6))
            tk.Label(debounce_row, textvariable=debounce_profile_var, width=12, anchor="w", fg="#555555").pack(side=tk.LEFT, padx=(2, 0))
            set_debounce_label()
            tk.Label(
                speed_frame,
                text="Lower = faster retile after events, Higher = fewer retiles.",
                font=("Arial", 8),
                fg="gray",
                anchor="w",
            ).pack(fill=tk.X, padx=2, pady=(2, 6))

            timeout_options = ["1.0s", "1.5s", "2.0s", "2.5s", "3.0s", "3.5s"]
            timeout_values = [1.0, 1.5, 2.0, 2.5, 3.0, 3.5]
            current_timeout_idx = min(
                range(len(timeout_values)),
                key=lambda i: abs(timeout_values[i] - float(self.window_mgr.tile_timeout))
            )
            timeout_var = tk.StringVar(value=timeout_options[current_timeout_idx])
            speed_row2 = tk.Frame(speed_frame)
            speed_row2.pack(fill=tk.X)
            tk.Label(speed_row2, text="Timeout", width=10, anchor="w").pack(side=tk.LEFT, padx=(2, 2))
            ttk.Combobox(
                speed_row2, textvariable=timeout_var, values=timeout_options,
                state="readonly", width=8
            ).pack(side=tk.LEFT, padx=4)

            retries_options = ["5", "8", "10", "12", "15", "20"]
            retries_values = [5, 8, 10, 12, 15, 20]
            current_retries_idx = min(
                range(len(retries_values)),
                key=lambda i: abs(retries_values[i] - int(self.window_mgr.max_tile_retries))
            )
            retries_var = tk.StringVar(value=retries_options[current_retries_idx])
            tk.Label(speed_row2, text="Retries", width=8, anchor="w").pack(side=tk.LEFT, padx=(16, 2))
            ttk.Combobox(
                speed_row2, textvariable=retries_var, values=retries_options,
                state="readonly", width=6
            ).pack(side=tk.LEFT, padx=4)

            # Retile effect / animation
            anim_frame = tk.LabelFrame(dialog, text="Retile Effect", padx=10, pady=10)
            anim_frame.pack(fill=tk.X, padx=15, pady=5)

            anim_enabled_var = tk.BooleanVar(value=bool(self.window_mgr.animation_enabled))
            tk.Checkbutton(
                anim_frame,
                text="Enable animated retile",
                variable=anim_enabled_var,
            ).pack(anchor="w", pady=(0, 5))

            anim_opts_frame = tk.Frame(anim_frame)
            anim_opts_frame.pack(fill=tk.X)

            effect_frame = tk.Frame(anim_opts_frame)
            effect_frame.pack(fill=tk.X)
            effect_map = {
                "Critically Damped": "crit_damped",
                "Spring": "spring_out",
                "Wave (Arc)": "arc_wave",
            }
            reverse_effect_map = {v: k for k, v in effect_map.items()}
            current_effect_label = reverse_effect_map.get(
                getattr(self.window_mgr, "animation_effect", "crit_damped"),
                "Critically Damped",
            )
            effect_var = tk.StringVar(value=current_effect_label)
            tk.Label(effect_frame, text="Effect", width=12, anchor="w").pack(side=tk.LEFT, padx=2)
            ttk.Combobox(
                effect_frame,
                textvariable=effect_var,
                values=list(effect_map.keys()),
                state="readonly",
                width=18,
            ).pack(side=tk.LEFT, padx=4)

            duration_values_ms = [80, 120, 160, 200, 250, 300, 400, 500, 700]
            current_duration_ms = int(round(float(self.window_mgr.animation_duration) * 1000.0))
            current_duration_idx = min(
                range(len(duration_values_ms)),
                key=lambda i: abs(duration_values_ms[i] - current_duration_ms)
            )
            duration_idx_var = tk.IntVar(value=current_duration_idx)
            duration_display_var = tk.StringVar()
            duration_profile_var = tk.StringVar()

            def set_duration_label():
                idx = max(0, min(len(duration_values_ms) - 1, int(duration_idx_var.get())))
                val = duration_values_ms[idx]
                if val <= 120:
                    profile = "Very Fast"
                elif val <= 200:
                    profile = "Fast"
                elif val <= 300:
                    profile = "Balanced"
                elif val <= 500:
                    profile = "Smooth"
                else:
                    profile = "Cinematic"
                duration_display_var.set(f"{val} ms")
                duration_profile_var.set(profile)
                duration_idx_var.set(idx)

            def change_duration(delta):
                duration_idx_var.set(int(duration_idx_var.get()) + delta)
                set_duration_label()

            duration_frame = tk.Frame(anim_opts_frame)
            duration_frame.pack(fill=tk.X, pady=(4, 0))
            tk.Label(duration_frame, text="Duration", width=12, anchor="w").pack(side=tk.LEFT, padx=2)
            tk.Button(duration_frame, text="-", width=2, command=lambda: change_duration(-1)).pack(side=tk.LEFT, padx=(2, 1))
            tk.Label(duration_frame, textvariable=duration_display_var, width=9, anchor="center").pack(side=tk.LEFT, padx=2)
            tk.Button(duration_frame, text="+", width=2, command=lambda: change_duration(1)).pack(side=tk.LEFT, padx=(1, 6))
            tk.Label(duration_frame, textvariable=duration_profile_var, width=10, anchor="w", fg="#555555").pack(side=tk.LEFT, padx=(2, 0))
            set_duration_label()

            fps_values = [24, 30, 45, 60, 75, 90, 120, 144]
            current_fps = max(1, int(self.window_mgr.animation_fps))
            current_fps_idx = min(
                range(len(fps_values)),
                key=lambda i: abs(fps_values[i] - current_fps)
            )
            fps_idx_var = tk.IntVar(value=current_fps_idx)
            fps_display_var = tk.StringVar()
            fps_profile_var = tk.StringVar()

            def set_fps_label():
                idx = max(0, min(len(fps_values) - 1, int(fps_idx_var.get())))
                val = fps_values[idx]
                if val <= 30:
                    profile = "Battery"
                elif val <= 60:
                    profile = "Balanced"
                elif val <= 90:
                    profile = "Smooth"
                elif val <= 120:
                    profile = "High"
                else:
                    profile = "Ultra"
                fps_display_var.set(str(val))
                fps_profile_var.set(profile)
                fps_idx_var.set(idx)

            def change_fps(delta):
                fps_idx_var.set(int(fps_idx_var.get()) + delta)
                set_fps_label()

            fps_frame = tk.Frame(anim_opts_frame)
            fps_frame.pack(fill=tk.X, pady=(4, 0))
            tk.Label(fps_frame, text="FPS", width=12, anchor="w").pack(side=tk.LEFT, padx=2)
            tk.Button(fps_frame, text="-", width=2, command=lambda: change_fps(-1)).pack(side=tk.LEFT, padx=(2, 1))
            tk.Label(fps_frame, textvariable=fps_display_var, width=5, anchor="center").pack(side=tk.LEFT, padx=2)
            tk.Button(fps_frame, text="+", width=2, command=lambda: change_fps(1)).pack(side=tk.LEFT, padx=(1, 6))
            tk.Label(fps_frame, textvariable=fps_profile_var, width=10, anchor="w", fg="#555555").pack(side=tk.LEFT, padx=(2, 0))
            set_fps_label()

            def sync_anim_controls(*_):
                if anim_enabled_var.get():
                    if not anim_opts_frame.winfo_manager():
                        anim_opts_frame.pack(fill=tk.X)
                else:
                    if anim_opts_frame.winfo_manager():
                        anim_opts_frame.pack_forget()

            anim_enabled_var.trace_add("write", sync_anim_controls)
            sync_anim_controls()

            # Auto-compact options moved from tray into settings.
            compact_frame = tk.LabelFrame(dialog, text="Auto-Compact", padx=10, pady=10)
            compact_frame.pack(fill=tk.X, padx=15, pady=5)
            compact_min_var = tk.BooleanVar(value=bool(self.compact_on_minimize))
            compact_close_var = tk.BooleanVar(value=bool(self.compact_on_close))
            tk.Checkbutton(
                compact_frame,
                text="Auto-compact on minimize",
                variable=compact_min_var,
            ).pack(anchor="w")
            tk.Checkbutton(
                compact_frame,
                text="Auto-compact on close",
                variable=compact_close_var,
            ).pack(anchor="w")
            
            # Buttons
            button_frame = tk.Frame(dialog)
            button_frame.pack(pady=15)
            
            def apply_and_close():
                gap_str = gap_var.get().replace("px", "")
                padding_str = padding_var.get().replace("px", "")
                debounce_s = debounce_values[int(debounce_idx_var.get())] / 1000.0
                timeout_s = float(timeout_var.get().replace("s", ""))
                retries_n = int(retries_var.get())
                anim_duration_s = duration_values_ms[int(duration_idx_var.get())] / 1000.0
                anim_fps_n = fps_values[int(fps_idx_var.get())]
                anim_effect_key = effect_map.get(effect_var.get(), "crit_damped")
                
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
                self.retile_debounce = debounce_s
                self.window_mgr.tile_timeout = timeout_s
                self.window_mgr.max_tile_retries = retries_n
                self.window_mgr.animation_enabled = bool(anim_enabled_var.get())
                self.window_mgr.animation_duration = anim_duration_s
                self.window_mgr.animation_fps = anim_fps_n
                self.window_mgr.animation_effect = anim_effect_key
                self.compact_on_minimize = bool(compact_min_var.get())
                self.compact_on_close = bool(compact_close_var.get())
                
                log(
                    f"[SETTINGS] GAP={self.gap}px EDGE_PADDING={self.edge_padding}px "
                    f"RETILE_DEBOUNCE={self.retile_debounce:.2f}s TILE_TIMEOUT={self.window_mgr.tile_timeout:.1f}s "
                    f"RETRIES={self.window_mgr.max_tile_retries} "
                    f"ANIM={'ON' if self.window_mgr.animation_enabled else 'OFF'} "
                    f"EFFECT={self.window_mgr.animation_effect} "
                    f"DUR={self.window_mgr.animation_duration:.2f}s FPS={self.window_mgr.animation_fps} "
                    f"COMPACT_MIN={'ON' if self.compact_on_minimize else 'OFF'} "
                    f"COMPACT_CLOSE={'ON' if self.compact_on_close else 'OFF'}"
                )
                self.update_tray_menu()
                
                # Apply
                self.apply_new_settings()
            
            def cancel_and_close():
                dialog.destroy()
                root.destroy()
            
            def reset_defaults():
                gap_var.set(gap_options[3])
                padding_var.set(padding_options[2])
                debounce_idx_var.set(
                    min(range(len(debounce_values)), key=lambda i: abs(debounce_values[i] - 50))
                )
                set_debounce_label()
                timeout_var.set("2.0s")
                retries_var.set("10")
                anim_enabled_var.set(True)
                effect_var.set("Critically Damped")
                duration_idx_var.set(
                    min(range(len(duration_values_ms)), key=lambda i: abs(duration_values_ms[i] - 80))
                )
                set_duration_label()
                fps_idx_var.set(
                    min(range(len(fps_values)), key=lambda i: abs(fps_values[i] - 60))
                )
                set_fps_label()
                compact_min_var.set(True)
                compact_close_var.set(True)
                sync_anim_controls()
            
            tk.Button(button_frame, text="Apply", command=apply_and_close, width=12,
                    bg="#4CAF50", fg="white").pack(side=tk.LEFT, padx=5)
            tk.Button(button_frame, text="Reset", command=reset_defaults, width=12).pack(side=tk.LEFT, padx=5)
            tk.Button(button_frame, text="Cancel", command=cancel_and_close, width=12).pack(side=tk.LEFT, padx=5)
            
            info = tk.Label(dialog, text="Changes apply immediately on next retile cycle",
                            font=("Arial", 8), fg="gray")
            info.pack(pady=5)

            # Fit dialog to full content so footer buttons are never clipped.
            dialog.update_idletasks()
            req_w = dialog.winfo_reqwidth() + 20
            req_h = dialog.winfo_reqheight() + 20
            screen_w = dialog.winfo_screenwidth()
            screen_h = dialog.winfo_screenheight()
            target_w = max(500, min(screen_w - 60, req_w))
            target_h = max(680, min(screen_h - 80, req_h))
            self._center_tk_window(dialog, target_w, target_h, monitor_idx=self.current_monitor_index)
            dialog.minsize(500, 680)
            
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
        """Apply updated settings and force one clean retile."""
        # Wait for all Tkinter windows to close
        time.sleep(0.5)
        
        # Clean up dead windows
        self.window_mgr.cleanup_dead_windows()
        self._backfill_window_state_ws()
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
            (HOTKEY_LAYOUT_PICKER, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('P')),
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
            if HOTKEY_LAYOUT_PICKER in failed:
                try:
                    ctypes.windll.user32.MessageBoxW(
                        0,
                        "Ctrl+Alt+P hotkey failed to register.\n"
                        "It may already be used by another program.\n\n"
                        "You can still open the Layout Picker from the systray menu.",
                        "SmartGrid Hotkey",
                        0x30
                    )
                except Exception:
                    pass
        else:
            log("[HOTKEYS] All hotkeys registered successfully")
    
    def unregister_hotkeys(self):
        """Unregister all hotkeys."""
        for hk in (HOTKEY_TOGGLE, HOTKEY_RETILE, HOTKEY_QUIT, HOTKEY_MOVE_MONITOR,
                   HOTKEY_SWAP_MODE, HOTKEY_WS1, HOTKEY_WS2, HOTKEY_WS3,
                   HOTKEY_FLOAT_TOGGLE, HOTKEY_LAYOUT_PICKER):
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
            MenuItem('Layout Manager (Ctrl+Alt+P)',
                     lambda: user32.PostThreadMessageW(self.main_thread_id, CUSTOM_OPEN_LAYOUT_PICKER, 0, 0)),
            Menu.SEPARATOR,
            MenuItem('Workspaces', Menu(
                MenuItem('Switch to Workspace 1 (Ctrl+Alt+1)',
                         lambda: threading.Thread(target=self.ws_switch, args=(0,), daemon=True).start()),
                MenuItem('Switch to Workspace 2 (Ctrl+Alt+2)',
                         lambda: threading.Thread(target=self.ws_switch, args=(1,), daemon=True).start()),
                MenuItem('Switch to Workspace 3 (Ctrl+Alt+3)',
                         lambda: threading.Thread(target=self.ws_switch, args=(2,), daemon=True).start()),
            )),
            Menu.SEPARATOR,
            MenuItem('Settings',
                     lambda: user32.PostThreadMessageW(self.main_thread_id, CUSTOM_OPEN_SETTINGS, 0, 0)),
            MenuItem('Open Recycle Bin',
                     lambda: threading.Thread(target=self.open_recycle_bin, daemon=True).start()),
            MenuItem('Hotkeys Cheatsheet', lambda: threading.Thread(target=show_hotkeys_tooltip, daemon=True).start()),
            MenuItem('Quit SmartGrid (Ctrl+Alt+Q)', self.on_quit_from_tray)
        )
    
    def open_recycle_bin(self):
        """Open Windows Recycle Bin from systray menu."""
        try:
            os.startfile("shell:RecycleBinFolder")
        except Exception as e:
            log(f"[ERROR] open_recycle_bin: {e}")
            try:
                winsound.MessageBeep(0xFFFFFFFF)
            except Exception:
                pass

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
                # Purge all runtime tiling state so next ON starts clean.
                known_hwnds = (
                    set(self.window_mgr.grid_state.keys())
                    | set(self.window_mgr.minimized_windows.keys())
                    | set(self.window_mgr.maximized_windows.keys())
                )
                for hwnd in known_hwnds:
                    if user32.IsWindow(hwnd):
                        set_window_border(hwnd, None)
                self.window_mgr.grid_state.clear()
                self.window_mgr.minimized_windows.clear()
                self.window_mgr.maximized_windows.clear()
                self.window_state_ws.clear()
                self.last_visible_count = 0
                self.last_known_count = 0
        
        self.update_tray_menu()

    def toggle_compact_on_minimize(self):
        """Toggle compact-on-minimize behavior."""
        self.compact_on_minimize = not self.compact_on_minimize
        log(f"[SMARTGRID] Auto-compact on minimize: {'ON' if self.compact_on_minimize else 'OFF'}")
        # Prevent systray menu windows from triggering an immediate auto-retile.
        now = time.time()
        self.ignore_retile_until = max(self.ignore_retile_until, now + 0.35)
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
        self.last_retile_time = now
        self.update_tray_menu()

    def toggle_compact_on_close(self):
        """Toggle compact-on-close behavior."""
        self.compact_on_close = not self.compact_on_close
        log(f"[SMARTGRID] Auto-compact on close: {'ON' if self.compact_on_close else 'OFF'}")
        # Prevent systray menu windows from triggering an immediate auto-retile.
        now = time.time()
        self.ignore_retile_until = max(self.ignore_retile_until, now + 0.35)
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
        self.last_retile_time = now
        self.update_tray_menu()

    def show_layout_picker(self):
        """Manual layout picker (choose layout + assign windows)."""
        with self._layout_picker_lock:
            if self._layout_picker_open:
                return
            self._layout_picker_open = True

        old_ignore = self.ignore_retile_until
        self.ignore_retile_until = float('inf')
        self.drag_drop_lock = True
        try:
            mon_idx = self.current_monitor_index
            with self.lock:
                active_ws_at_open = self.current_workspace.get(mon_idx, 0)
                active_grid_count = sum(
                    1
                    for hwnd, (m, _c, _r) in self.window_mgr.grid_state.items()
                    if m == mon_idx and user32.IsWindow(hwnd)
                )
                current_layout_sig = self.layout_signature.get(mon_idx)
            current_layout = (
                self.layout_engine.choose_layout(active_grid_count)
                if active_grid_count > 0
                else current_layout_sig
            )

            import tkinter as tk
            import tkinter.font as tkfont
            from tkinter import ttk, messagebox

            root = tk.Tk()
            root.withdraw()
            root.attributes("-topmost", True)

            dialog = tk.Toplevel(root)
            dialog.title("SmartGrid Layout Manager")
            dialog.geometry("660x620")
            dialog.attributes("-topmost", True)
            dialog.resizable(False, False)
            with self._layout_picker_lock:
                self._layout_picker_hwnd = int(dialog.winfo_id())

            accent = "#3A7BD5"
            accent_dark = "#2E64AE"
            accent_hover = "#2F6FD0"
            section_bg = "#F5F5F7"

            header = tk.Label(
                dialog,
                text="Layout Manager",
                font=("Arial", 12, "bold"),
                bg=accent,
                fg="white",
            )
            header.pack(fill=tk.X, pady=(8, 6))

            layout_frame = tk.LabelFrame(
                dialog,
                text="🧩 Choose Layout to Customize",
                padx=10,
                pady=10,
                font=("Arial", 9, "bold"),
                bg=section_bg,
            )
            layout_frame.pack(fill=tk.X, padx=12, pady=6)

            layout_presets = self._get_layout_presets()
            layout_labels = [label for label, _ in layout_presets]
            ws_labels = [f"Workspace {i+1}" for i in range(3)]
            target_ws_var = tk.StringVar(value=ws_labels[active_ws_at_open])

            def get_target_ws_index():
                try:
                    return ws_labels.index(target_ws_var.get().strip())
                except Exception:
                    return active_ws_at_open

            def get_default_label_for_ws(ws_idx):
                default = layout_labels[0]
                with self.lock:
                    remembered = self.workspace_layout_signature.get((mon_idx, ws_idx))
                    if remembered is None and ws_idx == self.current_workspace.get(mon_idx, active_ws_at_open):
                        remembered = current_layout
                    ws_list = self.workspaces.get(mon_idx, [])
                    ws_map = {}
                    if 0 <= ws_idx < len(ws_list) and isinstance(ws_list[ws_idx], dict):
                        ws_map = dict(ws_list[ws_idx])

                layout_sig = remembered
                if layout_sig is None:
                    ws_count = sum(1 for hwnd in ws_map.keys() if user32.IsWindow(hwnd))
                    if ws_count > 0:
                        layout_sig = self.layout_engine.choose_layout(ws_count)

                if layout_sig:
                    cur_label = self._layout_label(*layout_sig)
                    if cur_label in layout_labels:
                        return cur_label
                return default

            layout_var = tk.StringVar(value=get_default_label_for_ws(active_ws_at_open))

            note_frame = tk.Frame(layout_frame, bg=section_bg)
            note_frame.pack(fill=tk.X, padx=4, pady=(0, 6))
            tk.Label(
                note_frame,
                text="This section applies only to the active monitor.",
                font=("Arial", 8),
                fg="gray",
                bg=section_bg,
                anchor="w",
            ).pack(side=tk.LEFT)

            monitor_badge = tk.Label(
                note_frame,
                text=f"CURRENT MONITOR {mon_idx + 1} · WS{active_ws_at_open + 1}",
                font=("Arial", 8, "bold"),
                bg="#EFEFEF",
                fg="#444444",
                padx=6,
                pady=1,
                relief="solid",
                bd=1,
            )
            monitor_badge.pack(side=tk.LEFT, padx=(8, 0))

            tk.Label(
                note_frame,
                text="Target:",
                font=("Arial", 8, "bold"),
                fg="#555555",
                bg=section_bg,
            ).pack(side=tk.LEFT, padx=(8, 3))
            target_ws_combo = ttk.Combobox(
                note_frame,
                textvariable=target_ws_var,
                values=ws_labels,
                state="readonly",
                width=11,
            )
            target_ws_combo.pack(side=tk.LEFT)

            current_badge = tk.Label(
                note_frame,
                text="CURRENT",
                font=("Arial", 8, "bold"),
                bg="#E7F1FF",
                fg=accent_dark,
                padx=6,
                pady=1,
                relief="solid",
                bd=1,
            )
            current_badge.pack(side=tk.LEFT, padx=(6, 0))

            layout_top = tk.Frame(layout_frame, bg=section_bg)
            layout_top.pack(fill=tk.X)
            layout_top.columnconfigure(0, weight=0)
            layout_top.columnconfigure(1, weight=0)
            layout_top.columnconfigure(2, weight=1)

            layout_combo = ttk.Combobox(
                layout_top,
                textvariable=layout_var,
                values=layout_labels,
                state="readonly",
                width=24,
            )
            layout_combo.grid(row=0, column=0, padx=5, sticky="w")
            tk.Frame(layout_top, bg=section_bg).grid(row=0, column=2, sticky="ew")
            current_badge.pack_forget()

            layout_canvas = tk.Canvas(
                layout_frame,
                height=120,
                bg="white",
                highlightthickness=1,
                highlightbackground="#cfcfcf",
            )
            layout_canvas.pack(fill=tk.X, pady=(6, 2))

            font = tkfont.nametofont("TkDefaultFont")
            screen_w = dialog.winfo_screenwidth()
            display_labels = []
            label_to_hwnd = {}
            hwnd_to_label = {}
            combo_width = 30
            dialog_size = {"width": 620}

            def refresh_window_choices():
                nonlocal display_labels, label_to_hwnd, hwnd_to_label, combo_width
                window_choices = self._get_window_choices_for_monitor(mon_idx)
                display_labels = []
                label_to_hwnd = {}
                hwnd_to_label = {}
                for idx, (hwnd, desc) in enumerate(window_choices, start=1):
                    title = desc.get("title") or "(untitled)"
                    proc = desc.get("process") or "unknown"
                    label = f"{idx}. {title} [{proc}]"
                    display_labels.append(label)
                    label_to_hwnd[label] = hwnd
                    hwnd_to_label[hwnd] = label

                max_label_len = max((len(label) for label in display_labels), default=40)
                max_label_px = max((font.measure(label) for label in display_labels), default=320)
                slot_label_px = font.measure("Slot 99 (9,9)") + 20
                combo_px = max_label_px + 40
                combo_width = max(30, max_label_len + 2)
                dialog_width = max(620, slot_label_px + combo_px + 120)
                dialog_width = min(screen_w - 80, dialog_width)
                if dialog_width > dialog_size["width"]:
                    dialog_size["width"] = dialog_width
                    cur_h = dialog.winfo_height() or 620
                    self._center_tk_window(dialog, dialog_size["width"], cur_h, monitor_idx=mon_idx)
                return bool(display_labels)

            refresh_window_choices()
            self._center_tk_window(dialog, dialog_size["width"], 620, monitor_idx=mon_idx)

            slots_frame = tk.LabelFrame(
                dialog,
                text="🧷 Assign windows/apps to slots",
                padx=10,
                pady=10,
                font=("Arial", 9, "bold"),
                bg=section_bg,
            )
            slots_frame.pack(fill=tk.X, expand=False, padx=12, pady=6)

            slot_vars = []
            slot_widgets = []
            apply_btn = None
            poll_shutdown_job = None
            preview_job = None
            picker_closing = False

            def resize_dialog_to_content():
                dialog.update_idletasks()
                req_h = dialog.winfo_reqheight()
                target_height = max(340, req_h + 10)
                self._center_tk_window(dialog, dialog_size["width"], target_height, monitor_idx=mon_idx)

            def draw_layout_preview(positions, grid_coords, selected_coords=None):
                layout_canvas.delete("all")
                layout_canvas.update_idletasks()
                cw = layout_canvas.winfo_width()
                ch = layout_canvas.winfo_height()
                if cw < 10:
                    cw = layout_canvas.winfo_reqwidth() or 220
                if ch < 10:
                    ch = layout_canvas.winfo_reqheight() or 120
                pad = 6
                selected_coords = selected_coords or set()

                if not positions:
                    return
                min_x = min(x for x, y, w, h in positions)
                min_y = min(y for x, y, w, h in positions)
                max_x = max(x + w for x, y, w, h in positions)
                max_y = max(y + h for x, y, w, h in positions)
                span_x = max(1, max_x - min_x)
                span_y = max(1, max_y - min_y)
                scale = min((cw - 2 * pad) / span_x, (ch - 2 * pad) / span_y)

                for i, (pos, coord) in enumerate(zip(positions, grid_coords), start=1):
                    x, y, w, h = pos
                    sx = pad + (x - min_x) * scale
                    sy = pad + (y - min_y) * scale
                    sw = w * scale
                    sh = h * scale
                    fill = "#D7E6FA" if coord in selected_coords else "white"
                    layout_canvas.create_rectangle(sx, sy, sx + sw, sy + sh, outline="#4a4a4a", fill=fill)
                    layout_canvas.create_text(sx + sw / 2, sy + sh / 2, text=str(i), fill="#555", font=("Arial", 8))

            def close_picker():
                nonlocal poll_shutdown_job, preview_job, picker_closing
                if picker_closing:
                    return
                picker_closing = True
                try:
                    if poll_shutdown_job is not None and dialog.winfo_exists():
                        dialog.after_cancel(poll_shutdown_job)
                        poll_shutdown_job = None
                except Exception:
                    pass
                try:
                    if preview_job is not None and layout_canvas.winfo_exists():
                        layout_canvas.after_cancel(preview_job)
                        preview_job = None
                except Exception:
                    pass
                try:
                    if dialog.winfo_exists():
                        dialog.destroy()
                except Exception:
                    pass
                try:
                    if root.winfo_exists():
                        root.destroy()
                except Exception:
                    pass

            def poll_shutdown():
                nonlocal poll_shutdown_job
                # If a global quit is requested (e.g. from systray), close picker promptly.
                try:
                    if picker_closing:
                        return
                    if not dialog.winfo_exists():
                        return
                    if self._stop_event.is_set():
                        close_picker()
                        return
                    poll_shutdown_job = dialog.after(120, poll_shutdown)
                except Exception:
                    pass

            def quit_from_picker(_event=None):
                close_picker()
                threading.Thread(target=self.on_quit_from_tray, daemon=True).start()
                return "break"

            def add_hover(button, base_color, hover_color):
                def _enter(_):
                    if str(button["state"]) != "disabled":
                        button.config(bg=hover_color)

                def _leave(_):
                    if str(button["state"]) != "disabled":
                        button.config(bg=base_color)

                button.bind("<Enter>", _enter)
                button.bind("<Leave>", _leave)

            def update_current_badge():
                with self.lock:
                    grid_count = sum(
                        1
                        for hwnd, (m, _c, _r) in self.window_mgr.grid_state.items()
                        if m == mon_idx and user32.IsWindow(hwnd)
                    )
                    sig = self.layout_signature.get(mon_idx)
                    ws_num = self.current_workspace.get(mon_idx, active_ws_at_open) + 1
                monitor_badge.config(text=f"CURRENT MONITOR {mon_idx + 1} · WS{ws_num}")
                cur_layout = self.layout_engine.choose_layout(grid_count) if grid_count > 0 else sig
                if not cur_layout:
                    if current_badge.winfo_manager():
                        current_badge.pack_forget()
                    return
                cur_label = self._layout_label(*cur_layout)
                current_badge.config(text=f"CURRENT · {cur_label}")
                if layout_var.get() == cur_label:
                    if not current_badge.winfo_manager():
                        current_badge.pack(side=tk.LEFT, padx=(6, 0))
                else:
                    if current_badge.winfo_manager():
                        current_badge.pack_forget()

            def update_apply_button_label():
                if apply_btn is None:
                    return
                target_ws = get_target_ws_index()
                with self.lock:
                    active_ws = self.current_workspace.get(mon_idx, active_ws_at_open)
                if target_ws == active_ws:
                    apply_btn.config(text="Apply Layout")
                else:
                    apply_btn.config(text="Apply & Switch")

            def update_apply_state():
                if apply_btn is None:
                    return
                update_apply_button_label()
                filled = sum(1 for _coord, var in slot_vars if var.get().strip())
                apply_btn.config(state=tk.NORMAL if filled > 0 else tk.DISABLED)

            def get_current_grid_prefill(layout, info, grid_coords):
                """Return {(col,row): label} for the target workspace."""
                target_ws = get_target_ws_index()
                with self.lock:
                    ws_list = self.workspaces.get(mon_idx, [])
                    ws_map = {}
                    if 0 <= target_ws < len(ws_list) and isinstance(ws_list[target_ws], dict):
                        ws_map = dict(ws_list[target_ws])
                    active_ws = self.current_workspace.get(mon_idx, active_ws_at_open)
                    sig = self.layout_signature.get(mon_idx)
                    grid_items = [
                        (hwnd, c, r)
                        for hwnd, (m, c, r) in self.window_mgr.grid_state.items()
                        if m == mon_idx and user32.IsWindow(hwnd)
                    ]

                valid_coords = set(grid_coords)
                prefill = {}

                # For the active workspace, prefer current runtime grid when layout matches.
                if target_ws == active_ws:
                    filtered = [(hwnd, c, r) for hwnd, c, r in grid_items if user32.IsWindow(hwnd)]
                    if sig is not None:
                        current_layout = sig
                    else:
                        current_layout = self.layout_engine.choose_layout(len(filtered)) if filtered else None

                    if current_layout == (layout, info):
                        for hwnd, c, r in filtered:
                            coord = (c, r)
                            if coord not in valid_coords:
                                continue
                            label = hwnd_to_label.get(hwnd)
                            if label:
                                prefill[coord] = label
                        if prefill:
                            return prefill

                # Fallback: use stored workspace map.
                for hwnd, data in ws_map.items():
                    if not user32.IsWindow(hwnd) or not isinstance(data, dict):
                        continue
                    grid = data.get("grid")
                    if not isinstance(grid, (list, tuple)) or len(grid) != 2:
                        continue
                    try:
                        c = int(grid[0])
                        r = int(grid[1])
                    except Exception:
                        continue
                    coord = (c, r)
                    if coord not in valid_coords:
                        continue
                    label = hwnd_to_label.get(hwnd)
                    if label:
                        prefill[coord] = label
                return prefill

            def rebuild_slots(*_):
                nonlocal combo_width, preview_job
                for w in slots_frame.winfo_children():
                    w.destroy()
                slot_vars.clear()
                slot_widgets.clear()
                no_windows_available = not display_labels

                sel_label = layout_var.get()
                layout, info = dict(layout_presets).get(sel_label, (None, None))
                if layout is None:
                    sel_idx = layout_combo.current()
                    if sel_idx is not None and 0 <= sel_idx < len(layout_presets):
                        layout, info = layout_presets[sel_idx][1]
                    else:
                        layout, info = ("full", None)

                capacity = self._layout_capacity(layout, info)
                positions, grid_coords = self.layout_engine.calculate_positions(
                    self.monitors_cache[mon_idx], capacity, self.gap, self.edge_padding, layout, info
                )

                if no_windows_available:
                    tk.Label(
                        slots_frame,
                        text="No windows detected on this monitor. Slots are shown for preview.",
                        font=("Arial", 9, "italic"),
                        fg="gray",
                        bg=section_bg,
                        anchor="w",
                    ).pack(anchor="w", pady=(0, 4))

                for i, (col, row) in enumerate(grid_coords, start=1):
                    row_frame = tk.Frame(slots_frame, bg=section_bg)
                    row_frame.pack(fill=tk.X, pady=2)
                    row_frame.columnconfigure(1, weight=1)

                    tk.Label(
                        row_frame,
                        text=f"Slot {i} ({col},{row})",
                        width=16,
                        anchor="w",
                        bg=section_bg,
                    ).grid(row=0, column=0, sticky="w")

                    var = tk.StringVar()
                    combo = ttk.Combobox(
                        row_frame,
                        textvariable=var,
                        values=display_labels if display_labels else [],
                        state="readonly" if display_labels else "disabled",
                        height=15,
                        width=combo_width,
                    )
                    combo.grid(row=0, column=1, sticky="ew", padx=5)

                    # Prevent accidental slot reassignment from mouse wheel over the field.
                    combo.bind("<MouseWheel>", lambda _e: "break")
                    combo.bind("<Button-4>", lambda _e: "break")  # Linux scroll up
                    combo.bind("<Button-5>", lambda _e: "break")  # Linux scroll down

                    proc_label = tk.Label(
                        row_frame,
                        text="",
                        fg=accent,
                        bg=section_bg,
                        font=("Arial", 8, "bold"),
                        anchor="w",
                    )
                    proc_label.grid(row=0, column=2, sticky="e", padx=6)

                    slot_vars.append(((col, row), var))
                    slot_widgets.append((var, combo, proc_label))

                def update_proc_label(label_text, target_label):
                    if not label_text:
                        target_label.config(text="")
                        return
                    if "[" in label_text and "]" in label_text:
                        proc = label_text[label_text.rfind("[") + 1: label_text.rfind("]")]
                        target_label.config(text=proc)
                    else:
                        target_label.config(text="")

                def get_selected_coords():
                    return {coord for coord, var in slot_vars if var.get().strip()}

                def refresh_options(*_):
                    for v, c, pl in slot_widgets:
                        c["values"] = display_labels
                        update_proc_label(v.get().strip(), pl)
                    update_apply_state()
                    draw_layout_preview(positions, grid_coords, get_selected_coords())

                def on_select(current_var):
                    label = current_var.get().strip()
                    if not label:
                        return
                    for v, _c, _pl in slot_widgets:
                        if v is not current_var and v.get() == label:
                            v.set("")
                    refresh_options()

                def on_combo_selected(current_var, current_proc_label):
                    on_select(current_var)
                    update_proc_label(current_var.get(), current_proc_label)

                for v, c, pl in slot_widgets:
                    c.bind("<<ComboboxSelected>>", lambda _e, var=v, pl=pl: on_combo_selected(var, pl))

                # Auto-prefill from currently active tiled grid when selected layout matches it.
                prefill_by_coord = get_current_grid_prefill(layout, info, grid_coords)
                for coord, var in slot_vars:
                    label = prefill_by_coord.get(coord)
                    if label:
                        var.set(label)

                refresh_options()
                draw_layout_preview(positions, grid_coords, get_selected_coords())
                try:
                    if preview_job is not None and layout_canvas.winfo_exists():
                        layout_canvas.after_cancel(preview_job)
                except Exception:
                    pass
                preview_job = layout_canvas.after(
                    0,
                    lambda p=positions, g=grid_coords: draw_layout_preview(p, g, get_selected_coords())
                )
                update_current_badge()
                resize_dialog_to_content()

            def apply_layout():
                sel_label = layout_var.get()
                layout, info = dict(layout_presets).get(sel_label, (None, None))
                if layout is None:
                    sel_idx = layout_combo.current()
                    if sel_idx is not None and 0 <= sel_idx < len(layout_presets):
                        layout, info = layout_presets[sel_idx][1]
                    else:
                        layout, info = ("full", None)

                assignments = {}
                used = set()
                for (col, row), var in slot_vars:
                    label = var.get().strip()
                    if not label:
                        continue
                    hwnd = label_to_hwnd.get(label)
                    if hwnd is None:
                        messagebox.showwarning("SmartGrid", "Invalid window selection.")
                        return
                    if hwnd in used:
                        messagebox.showwarning("SmartGrid", "Duplicate window selected.")
                        return
                    used.add(hwnd)
                    assignments[(col, row)] = hwnd

                if not assignments:
                    messagebox.showwarning("SmartGrid", "Select at least one slot to apply.")
                    return

                target_ws = get_target_ws_index()
                with self.lock:
                    active_ws = self.current_workspace.get(mon_idx, active_ws_at_open)
                apply_now = target_ws == active_ws

                if not self._apply_manual_layout(
                    mon_idx,
                    layout,
                    info,
                    assignments,
                    target_ws=target_ws,
                    activate_target=apply_now,
                ):
                    messagebox.showwarning("SmartGrid", "Failed to apply layout.")
                    return
                close_picker()
                if not apply_now:
                    threading.Thread(target=self.ws_switch, args=(target_ws,), daemon=True).start()

            apply_btn = tk.Button(
                layout_top,
                text="Apply Layout",
                command=apply_layout,
                width=16,
                bg=accent,
                fg="white",
                activebackground=accent_dark,
                activeforeground="white",
            )
            apply_btn.grid(row=0, column=1, padx=6, sticky="w")
            add_hover(apply_btn, accent, accent_hover)

            layout_combo.bind("<<ComboboxSelected>>", rebuild_slots)
            layout_var.trace_add("write", lambda *_: rebuild_slots())
            def on_target_ws_selected(_event=None):
                target_ws = get_target_ws_index()
                target_label = get_default_label_for_ws(target_ws)
                if layout_var.get() != target_label:
                    layout_var.set(target_label)
                else:
                    rebuild_slots()

            target_ws_combo.bind("<<ComboboxSelected>>", on_target_ws_selected)
            refresh_window_choices()
            rebuild_slots()
            update_apply_state()

            btn_frame = tk.Frame(dialog)
            btn_frame.pack(pady=8)
            tk.Button(btn_frame, text="Cancel", command=close_picker, width=10).pack(side=tk.LEFT, padx=6)
            # Recompute final size after footer buttons are created, so Cancel is never clipped.
            resize_dialog_to_content()

            dialog.protocol("WM_DELETE_WINDOW", close_picker)
            dialog.bind("<Control-Alt-q>", quit_from_picker)
            dialog.bind("<Control-Alt-Q>", quit_from_picker)
            poll_shutdown_job = dialog.after(120, poll_shutdown)
            dialog.mainloop()

        except Exception as e:
            try:
                import tkinter as tk
                from tkinter import messagebox
                err_root = tk.Tk()
                err_root.withdraw()
                err_root.attributes("-topmost", True)
                messagebox.showerror("SmartGrid Layout Manager", f"Layout Manager failed to open:\n{e}")
                err_root.destroy()
            except Exception:
                pass
            log(f"[ERROR] show_layout_picker: {e}")
        finally:
            self.ignore_retile_until = old_ignore
            self.drag_drop_lock = False
            with self._layout_picker_lock:
                self._layout_picker_open = False
                self._layout_picker_hwnd = None
    
    def on_quit_from_tray(self, icon=None, item=None):
        """Quit from systray."""
        log("[TRAY] Quit requested")
        
        try:
            # If the layout picker is open, ask Windows to close it first so
            # the main thread can return from Tk mainloop.
            with self._layout_picker_lock:
                picker_hwnd = self._layout_picker_hwnd
            if picker_hwnd and user32.IsWindow(picker_hwnd):
                try:
                    user32.PostMessageW(picker_hwnd, win32con.WM_CLOSE, 0, 0)
                except Exception:
                    pass

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
                    self._reconcile_workspaces_after_monitor_change(current_monitors)
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
                    self._backfill_window_state_ws()
                    cleanup_minimized_moved = self.window_mgr.last_cleanup_minimized_moved
                    if cleanup_minimized_moved > 0:
                        minimized_moved += cleanup_minimized_moved

                    if self._sync_manual_cross_monitor_moves():
                        visible_windows = self.window_mgr.get_visible_windows(
                            self.monitors_cache, self.overlay_hwnd
                        )
                        self.last_visible_count = len(visible_windows)
                        with self.lock:
                            known_hwnds = (
                                set(self.window_mgr.grid_state.keys())
                                | set(self.window_mgr.minimized_windows.keys())
                                | set(self.window_mgr.maximized_windows.keys())
                            )
                        self.last_known_count = len(known_hwnds)
                        self.last_retile_time = time.time()
                        time.sleep(0.08)
                        continue

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
                        if minimized_moved and self.compact_on_minimize:
                            with self.lock:
                                self._pending_compact_minimize = True
                        with self.lock:
                            known_hwnds_pre = (
                                set(self.window_mgr.grid_state.keys())
                                | set(self.window_mgr.minimized_windows.keys())
                                | set(self.window_mgr.maximized_windows.keys())
                            )
                        if self.compact_on_close and len(known_hwnds_pre) < self.last_known_count:
                            with self.lock:
                                self._pending_compact_close = True
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

                    if self._run_deferred_compactions():
                        time.sleep(0.08)
                        continue

                    if minimized_moved:
                        # Minimizing a tiled window reduces the effective count; do a single
                        # reflow so the remaining windows fill the layout.
                        if self.compact_on_minimize:
                            self._compact_grid_after_minimize()
                            with self.lock:
                                self._pending_compact_minimize = False
                        else:
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
                        has_minimized_windows = len(self.window_mgr.minimized_windows) > 0
                    known_count = len(known_hwnds)
                    new_windows = [h for h in visible_hwnds if h not in known_hwnds]
                    closed_windows = known_count < self.last_known_count

                    # Robust fallback: if a minimize transition was missed by state sync,
                    # still compact when we detect a "hole" pattern.
                    now = time.time()
                    minimize_hole_detected = (
                        self.compact_on_minimize and
                        not minimized_moved and
                        has_minimized_windows and
                        current_count < self.last_visible_count and
                        known_count == self.last_known_count and
                        not new_windows
                    )
                    if minimize_hole_detected and now >= self.ignore_retile_until:
                        log(f"[AUTO-COMPACT] minimize fallback {self.last_visible_count} → {current_count} windows")
                        self._compact_grid_after_minimize()
                        self.last_visible_count = current_count
                        self.last_known_count = known_count
                        self.last_retile_time = now
                        time.sleep(0.08)
                        continue
                    
                    # Debounced retiling
                    if now >= self.ignore_retile_until and current_count > 0:
                        should_retile = False
                        if new_windows:
                            should_retile = True
                        elif known_count < self.last_known_count:
                            should_retile = True
                        elif layout_change:
                            should_retile = True

                        if should_retile and now - self.last_retile_time >= self.retile_debounce:
                            if self.compact_on_close and closed_windows and not new_windows:
                                log(f"[AUTO-RETILE] close compaction {self.last_visible_count} → {current_count} windows")
                                self._compact_grid_after_close()
                                with self.lock:
                                    self._pending_compact_close = False
                            else:
                                log(f"[AUTO-RETILE] {self.last_visible_count} → {current_count} windows")
                                self.smart_tile_with_restore()
                            self.last_visible_count = current_count
                            self.last_known_count = known_count
                            self.last_retile_time = now
                            time.sleep(0.2)
                        elif should_retile and self.compact_on_close and closed_windows and not new_windows:
                            with self.lock:
                                self._pending_compact_close = True
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
                        threading.Thread(target=self.ws_switch, args=(0,), daemon=True).start()
                    elif msg.wParam == HOTKEY_WS2:
                        threading.Thread(target=self.ws_switch, args=(1,), daemon=True).start()
                    elif msg.wParam == HOTKEY_WS3:
                        threading.Thread(target=self.ws_switch, args=(2,), daemon=True).start()
                    elif msg.wParam == HOTKEY_FLOAT_TOGGLE:
                        self.toggle_floating_selected()
                    elif msg.wParam == HOTKEY_LAYOUT_PICKER:
                        self.show_layout_picker()
                
                elif msg.message == CUSTOM_TOGGLE_SWAP:
                    if self.swap_mode_lock:
                        self.exit_swap_mode()
                    else:
                        self.enter_swap_mode()
                elif msg.message == CUSTOM_OPEN_LAYOUT_PICKER:
                    self.show_layout_picker()
                elif msg.message == CUSTOM_OPEN_SETTINGS:
                    self.show_settings_dialog()
                
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
        import tkinter as tk
        import tkinter.font as tkfont

        left_col_width = 32

        sections = [
            (
                "MAIN HOTKEYS",
                [
                    ("Ctrl+Alt+T", "Toggle tiling on/off"),
                    ("Ctrl+Alt+R", "Force re-tile all windows now"),
                    ("Ctrl+Alt+S", "Enter Swap Mode (red border + arrows)"),
                    ("Ctrl+Alt+M", "Move workspace to next monitor"),
                    ("Ctrl+Alt+F", "Toggle Floating Selected Window"),
                    ("Ctrl+Alt+P", "Layout Manager"),
                ],
                "title_main",
            ),
            (
                "SETTINGS",
                [
                    ("Gap / Edge Padding", "Window spacing and margins"),
                    ("Retile Debounce", "Auto-retile responsiveness"),
                    ("Timeout / Retries", "Robustness for stubborn windows"),
                    ("Animated Retile", "Enable/disable animation"),
                    ("Effect", "Critically Damped / Spring / Wave (Arc)"),
                    ("Duration / FPS", "Animation feel and smoothness"),
                    ("Auto-Compact on Minimize", "Fills empty slot"),
                    ("Auto-Compact on Close", "Fills empty slot"),
                ],
                "title_settings",
            ),
            (
                "WORKSPACES",
                [
                    ("Ctrl+Alt+1/2/3", "Switch workspace"),
                ],
                "title_workspaces",
            ),
            (
                "EXIT",
                [
                    ("Ctrl+Alt+Q", "Quit"),
                ],
                "title_exit",
            ),
        ]

        # Exact content line count: section title + rows, plus one blank line between sections.
        content_lines = sum(1 + len(rows) for _title, rows, _tag in sections) + (len(sections) - 1)
        text_height = max(16, min(42, content_lines + 1))

        def format_row(left, right):
            return f"{left:<{left_col_width}} -> {right}\n"

        preview_lines = []
        for idx, (title, rows, _title_tag) in enumerate(sections):
            preview_lines.append(title)
            for left, right in rows:
                preview_lines.append(format_row(left, right).rstrip("\n"))
            if idx < len(sections) - 1:
                preview_lines.append("")
        max_line_chars = max((len(line) for line in preview_lines), default=80)
        text_width = max(78, max_line_chars + 2)

        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        dialog = tk.Toplevel(root)
        dialog.title("SmartGrid Hotkeys")
        dialog.attributes("-topmost", True)
        dialog.resizable(False, False)

        text = tk.Text(
            dialog,
            width=text_width,
            height=text_height,
            wrap="none",
            font=("Consolas", 10),
            padx=10,
            pady=10,
            bd=1,
            relief="solid",
        )
        text.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 8))

        body_font = tkfont.Font(family="Consolas", size=10)
        title_font = tkfont.Font(family="Segoe UI", size=10, weight="bold")
        text.tag_configure("body", font=body_font, foreground="#1F1F1F")
        text.tag_configure("title_main", font=title_font, foreground="#0D47A1")
        text.tag_configure("title_settings", font=title_font, foreground="#1B5E20")
        text.tag_configure("title_workspaces", font=title_font, foreground="#E65100")
        text.tag_configure("title_exit", font=title_font, foreground="#B71C1C")

        for idx, (title, rows, title_tag) in enumerate(sections):
            text.insert("end", f"{title}\n", title_tag)
            for left, right in rows:
                text.insert("end", format_row(left, right), "body")
            if idx < len(sections) - 1:
                text.insert("end", "\n", "body")

        text.configure(state="disabled")

        button_row = tk.Frame(dialog)
        button_row.pack(fill=tk.X, padx=10, pady=(0, 10))

        def close_dialog(_event=None):
            try:
                dialog.destroy()
            except Exception:
                pass
            try:
                root.destroy()
            except Exception:
                pass
            return "break"

        close_btn = tk.Button(button_row, text="Close", width=12, command=close_dialog)
        close_btn.pack()

        dialog.protocol("WM_DELETE_WINDOW", close_dialog)
        dialog.bind("<Escape>", close_dialog)

        dialog.update_idletasks()
        w = dialog.winfo_width()
        h = dialog.winfo_height()
        sw = dialog.winfo_screenwidth()
        sh = dialog.winfo_screenheight()
        x = max(0, (sw - w) // 2)
        y = max(0, (sh - h) // 2)
        dialog.geometry(f"{w}x{h}+{x}+{y}")

        dialog.focus_force()
        close_btn.focus_set()
        root.mainloop()
    except Exception as e:
        log(f"[ERROR] show_hotkeys_tooltip: {e}")
        try:
            ctypes.windll.user32.MessageBoxW(
                0,
                "SmartGrid Hotkeys\n\nFailed to open advanced cheatsheet popup.",
                "SmartGrid Hotkeys",
                0x30,
            )
        except Exception:
            pass

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
    print("Ctrl+Alt+P     → Layout Manager")
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
