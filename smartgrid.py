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

# Animation settings
ANIMATION_DURATION = 0.20
ANIMATION_FPS = 60

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
            if w < 400 and h < 300:
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
        
        # Interpolation with easing (easeOutCubic for smooth effect)
        for i in range(1, frames + 1):
            t = i / frames
            # Smooth
            # ease = 1 - pow(1 - t, 3)  # easeOutCubic
            
            # Very smooth (light bounce)
            # ease = 1 - pow(1 - t, 4)  # easeOutQuart

            # Linear (constant speed)
            # ease = t

            # Bounce effect
            ease = 1 - (1 - t) * (1 - t) * (1 - 2 * t)

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
        
        # Active/selected windows
        self.current_hwnd = None  # Green border
        self.selected_hwnd = None  # Red border (swap mode)
        self.last_active_hwnd = None
        self.user_selected_hwnd = None
        
        # Thread safety
        self.lock = threading.Lock()
    
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
                
                if w <= 180 or h <= 180:
                    return True
                
                title_buf = ctypes.create_unicode_buffer(256)
                user32.GetWindowTextW(hwnd, title_buf, 256)
                title = title_buf.value or ""
                class_name = win32gui.GetClassName(hwnd)
                
                if overlay_hwnd and hwnd == overlay_hwnd:
                    return True
                
                # Check override (float toggle)
                useful = is_useful_window(title, class_name, hwnd=hwnd)
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
        """Remove dead windows from grid_state (STEP 1: fast cleanup)."""
        dead_windows = []
        
        with self.lock:
            for hwnd in list(self.grid_state.keys()):
                if not user32.IsWindow(hwnd):
                    dead_windows.append(hwnd)
                    self.grid_state.pop(hwnd, None)
                    continue
                
                state = get_window_state(hwnd)
                if state in ('minimized', 'maximized'):
                    self.grid_state.pop(hwnd, None)
                    continue
                
                if not user32.IsWindowVisible(hwnd):
                    self.grid_state.pop(hwnd, None)
                    continue
        
        if dead_windows:
            log(f"[CLEAN] Removed {len(dead_windows)} dead windows")
        
        return len(dead_windows)
    
    def cleanup_ghost_windows(self):
        """Remove ghost windows (STEP 2: zombie-like windows)."""
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
            
            # Fallback : classic method without animation
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
        
        # Multi-monitor & workspaces
        self.monitors_cache = []
        self.current_monitor_index = 0
        self.workspaces = {}
        self.current_workspace = {}
        
        # Overlay & UI
        self.overlay_hwnd = None
        self.preview_rect = None
        self.tray_icon = None
        
        # Threading
        self.lock = threading.Lock()
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
        if time.time() < self.ignore_retile_until:
            return
        
        with self.lock:
            self.ignore_retile_until = time.time() + 0.3
            
            # Cleanup
            self.window_mgr.cleanup_dead_windows()
            self.window_mgr.cleanup_ghost_windows()
            
            visible_windows = self.window_mgr.get_visible_windows(
                self.monitors_cache, self.overlay_hwnd
            )
            
            if not visible_windows:
                log("[TILE] No windows detected.")
                return
            
            # Separate windows by monitor
            wins_by_monitor = self._group_windows_by_monitor(visible_windows)
            
            # Process each monitor
            new_grid = {}
            for mon_idx, windows in wins_by_monitor.items():
                if mon_idx >= len(self.monitors_cache):
                    continue
                
                self._tile_monitor(mon_idx, windows, new_grid)
            
            self.window_mgr.grid_state = new_grid
            time.sleep(0.06)
    
    def _group_windows_by_monitor(self, visible_windows):
        """Group windows by their assigned monitor, preserving saved positions."""
        wins_by_monitor = {}
        
        for hwnd, title, rect in visible_windows:
            win_class = get_window_class(hwnd)
            
            # Check for saved position
            if hwnd in self.window_mgr.minimized_windows:
                mon_idx, col, row = self.window_mgr.minimized_windows[hwnd]
                del self.window_mgr.minimized_windows[hwnd]
                log(f"[RESTORE] Restoring minimized window to ({col},{row}): {title[:40]}")
            elif hwnd in self.window_mgr.maximized_windows:
                mon_idx, col, row = self.window_mgr.maximized_windows[hwnd]
                del self.window_mgr.maximized_windows[hwnd]
                log(f"[RESTORE] Restoring maximized window to ({col},{row}): {title[:40]}")
            elif hwnd in self.window_mgr.grid_state:
                mon_idx, col, row = self.window_mgr.grid_state[hwnd]
            else:
                mon_idx, col, row = 0, 0, 0
            
            wins_by_monitor.setdefault(mon_idx, []).append(
                (hwnd, title, rect, col, row, win_class)
            )
        
        return wins_by_monitor
    
    def _tile_monitor(self, mon_idx, windows, new_grid):
        """Tile windows on a specific monitor."""
        monitor_rect = self.monitors_cache[mon_idx]
        count = len(windows)
        
        layout, info = self.layout_engine.choose_layout(count)
        log(f"\n[TILE] Monitor {mon_idx+1}: {count} windows -> {layout} layout")
        
        # Calculate positions
        positions, grid_coords = self.layout_engine.calculate_positions(
            monitor_rect, count, self.gap, self.edge_padding, layout, info
        )
        
        pos_map = dict(zip(grid_coords, positions))
        
        # Phase 1: Restore saved positions
        assigned = set()
        unassigned_windows = []
        
        for hwnd, title, rect, saved_col, saved_row, win_class in windows:
            target_coords = (saved_col, saved_row)
            
            if target_coords in pos_map and target_coords not in assigned:
                x, y, w, h = pos_map[target_coords]
                self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                new_grid[hwnd] = (mon_idx, saved_col, saved_row)
                assigned.add(target_coords)
                log(f"   ✓ RESTORED to ({saved_col},{saved_row}): {title[:50]} [{win_class}]")
                time.sleep(0.015)
            else:
                unassigned_windows.append((hwnd, title, rect, 0, 0, win_class))
        
        # Phase 2: Assign remaining windows
        available_positions = [coord for coord in grid_coords if coord not in assigned]
        
        for i, (hwnd, title, rect, _, _, win_class) in enumerate(unassigned_windows):
            if i < len(available_positions):
                col, row = available_positions[i]
                x, y, w, h = pos_map[(col, row)]
                self.window_mgr.force_tile_resizable(hwnd, x, y, w, h)
                new_grid[hwnd] = (mon_idx, col, row)
                log(f"   → NEW position ({col},{row}): {title[:50]} [{win_class}]")
                time.sleep(0.015)
    
    def apply_grid_state(self):
        """Reapply all saved grid positions physically."""
        with self.lock:
            if not self.window_mgr.grid_state:
                return
            
            # Remove dead windows
            for hwnd in list(self.window_mgr.grid_state.keys()):
                if not user32.IsWindow(hwnd):
                    self.window_mgr.grid_state.pop(hwnd, None)
            
            if not self.window_mgr.grid_state:
                return
            
            # Group by monitor
            wins_by_mon = {}
            for hwnd, (mon_idx, col, row) in self.window_mgr.grid_state.items():
                wins_by_mon.setdefault(mon_idx, []).append((hwnd, col, row))
            
            # Process each monitor
            for mon_idx, windows in wins_by_mon.items():
                if mon_idx >= len(self.monitors_cache):
                    continue
                
                monitor_rect = self.monitors_cache[mon_idx]
                count = len(windows)
                layout, info = self.layout_engine.choose_layout(count)
                
                # Calculate positions
                positions, coord_list = self.layout_engine.calculate_positions(
                    monitor_rect, count, self.gap, self.edge_padding, layout, info
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
        
        # Active window is tiled → green border
        if active in self.window_mgr.grid_state and user32.IsWindow(active):
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
                if self.window_mgr.last_active_hwnd in self.window_mgr.grid_state:
                    set_window_border(self.window_mgr.last_active_hwnd, 0x0000FF00)
                    self.window_mgr.current_hwnd = self.window_mgr.last_active_hwnd
    
    # ==========================================================================
    # SWAP MODE
    # ==========================================================================
    
    def enter_swap_mode(self):
        """Enter swap mode: red border + arrow keys."""
        self.swap_mode_lock = True
        time.sleep(0.25)
        
        # Force quick update if grid_state is empty
        if not self.window_mgr.grid_state:
            visible_windows = self.window_mgr.get_visible_windows(
                self.monitors_cache, self.overlay_hwnd
            )
            for hwnd, title, _ in visible_windows:
                if hwnd not in self.window_mgr.grid_state:
                    self.window_mgr.grid_state[hwnd] = (0, 0, 0)
            
            log(f"[SWAP] grid_state rebuilt with {len(self.window_mgr.grid_state)} windows")
            if not self.window_mgr.grid_state:
                log("[SWAP] No tiled windows. Press Ctrl+Alt+T or Ctrl+Alt+R first")
                self.swap_mode_lock = False
                return
        
        # Smart selection
        candidate = None
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
        if from_hwnd not in self.window_mgr.grid_state:
            return None
        
        mon_idx = self.window_mgr.grid_state[from_hwnd][0]
        
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
            
            for hwnd, (m, _, _) in self.window_mgr.grid_state.items():
                if hwnd == from_hwnd or m != mon_idx or not user32.IsWindow(hwnd):
                    continue
                
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
                
                # Dynamic overlap threshold (20% of source window)
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
            
            # Swap in grid_state
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
        
        # Restore green border
        active = user32.GetForegroundWindow()
        if active and user32.IsWindowVisible(active) and active in self.window_mgr.grid_state:
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
                    
                    if self.preview_rect:
                        _, _, w, h = self.preview_rect
                        brush = win32gui.CreateSolidBrush(win32api.RGB(100, 149, 237))
                        pen = win32gui.CreatePen(win32con.PS_SOLID, 4, win32api.RGB(65, 105, 225))
                        
                        old_brush = win32gui.SelectObject(hdc, brush)
                        old_pen = win32gui.SelectObject(hdc, pen)
                        
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
                log(f"[ERROR] overlay wnd_proc: {e}")
                return 0
        
        try:
            wc = win32gui.WNDCLASS()
            wc.lpfnWndProc = wnd_proc
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
            
            # Get windows on target monitor
            wins_on_mon = [
                h for h, (m, _, _) in self.window_mgr.grid_state.items()
                if m == target_mon_idx and user32.IsWindow(h) and h != source_hwnd
            ]
            
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
                maxc = max((c for h,(m,c,r) in self.window_mgr.grid_state.items() 
                           if m == target_mon_idx), default=0)
                maxr = max((r for h,(m,c,r) in self.window_mgr.grid_state.items() 
                           if m == target_mon_idx), default=0)
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
        if source_hwnd not in self.window_mgr.grid_state:
            return
        
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
            
            wins_on_mon = [
                h for h, (m, _, _) in self.window_mgr.grid_state.items()
                if m == target_mon_idx and user32.IsWindow(h) and h != source_hwnd
            ]
            
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
                max_c = max((c for h,(m,c,r) in self.window_mgr.grid_state.items() 
                            if m == target_mon_idx), default=0)
                max_r = max((r for h,(m,c,r) in self.window_mgr.grid_state.items() 
                            if m == target_mon_idx), default=0)
                cols = max(cols, max_c + 1)
                rows = max(rows, max_r + 1)
                
                cell_w = (mon_w - 2*self.edge_padding - self.gap*(cols-1)) // cols
                cell_h = (mon_h - 2*self.edge_padding - self.gap*(rows-1)) // rows
                
                rel_x = cx - mon_x - self.edge_padding
                rel_y = cy - mon_y - self.edge_padding
                target_col = min(max(0, rel_x // (cell_w + self.gap)), cols - 1)
                target_row = min(max(0, rel_y // (cell_h + self.gap)), rows - 1)
            
            old_pos = self.window_mgr.grid_state[source_hwnd]
            new_pos = (target_mon_idx, target_col, target_row)
            
            if old_pos == new_pos:
                self.apply_grid_state()
                return
            
            # Check if target cell is occupied
            target_hwnd = None
            for h, pos in self.window_mgr.grid_state.items():
                if pos == new_pos and h != source_hwnd and user32.IsWindow(h):
                    target_hwnd = h
                    break
            
            with self.lock:
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
        drag_hwnd = None
        drag_start = None
        preview_active = False
        last_valid_rect = None
        
        while True:
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
                            while True:
                                parent = win32gui.GetParent(hwnd)
                                if not parent:
                                    break
                                hwnd = parent
                        except Exception:
                            pass
                    
                    if not hwnd or not user32.IsWindowVisible(hwnd):
                        hwnd = user32.GetForegroundWindow()
                    
                    # Check maximized
                    if hwnd:
                        style = user32.GetWindowLongW(hwnd, GWL_STYLE)
                        if style & WS_MAXIMIZE:
                            was_down = True
                            continue
                    
                    if hwnd and hwnd in self.window_mgr.grid_state and user32.IsWindowVisible(hwnd):
                        self.drag_drop_lock = True
                        time.sleep(0.25)
                        drag_hwnd = hwnd
                        drag_start = pt
                        preview_active = False
                
                # Drag in progress
                elif down and drag_hwnd:
                    cursor_pos = win32api.GetCursorPos()
                    
                    if drag_start:
                        dx = abs(cursor_pos[0] - drag_start[0])
                        dy = abs(cursor_pos[1] - drag_start[1])
                        
                        if (dx > 1 or dy > 1) and not preview_active:
                            preview_active = True
                        
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
                elif not down and was_down and drag_hwnd:
                    self.hide_snap_preview()
                    cursor_pos = win32api.GetCursorPos()
                    moved = False
                    
                    if drag_start:
                        dx = abs(cursor_pos[0] - drag_start[0])
                        dy = abs(cursor_pos[1] - drag_start[1])
                        moved = (dx > 10 or dy > 10)
                    
                    if moved:
                        self.handle_snap_drop(drag_hwnd, cursor_pos)
                    else:
                        self.apply_grid_state()
                    
                    self.drag_drop_lock = False
                    drag_hwnd = None
                    drag_start = None
                    preview_active = False
                
                was_down = down
                time.sleep(0.005)
            
            except Exception as e:
                log(f"[ERROR] drag_snap_monitor: {e}")
                self.hide_snap_preview()
                self.drag_drop_lock = False
                drag_hwnd = None
                drag_start = None
                preview_active = False
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
        
        # Save normal windows
        for hwnd, (mon, col, row) in self.window_mgr.grid_state.items():
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
        for hwnd, (mon, col, row) in self.window_mgr.minimized_windows.items():
            if mon == monitor_idx and user32.IsWindow(hwnd):
                self.workspaces[monitor_idx][ws][hwnd] = {
                    'pos': (0, 0, 800, 600),
                    'grid': (col, row),
                    'state': 'minimized'
                }
        
        # Save maximized windows
        for hwnd, (mon, col, row) in self.window_mgr.maximized_windows.items():
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
        
        for hwnd, data in layout.items():
            if not user32.IsWindow(hwnd):
                continue
            
            try:
                x, y, w, h = data['pos']
                col, row = data['grid']
                saved_state = data.get('state', 'normal')
                
                if saved_state == 'minimized':
                    user32.ShowWindowAsync(hwnd, win32con.SW_MINIMIZE)
                    self.window_mgr.minimized_windows[hwnd] = (monitor_idx, col, row)
                elif saved_state == 'maximized':
                    user32.ShowWindowAsync(hwnd, win32con.SW_MAXIMIZE)
                    self.window_mgr.maximized_windows[hwnd] = (monitor_idx, col, row)
                else:
                    if win32gui.IsIconic(hwnd):
                        user32.ShowWindowAsync(hwnd, SW_RESTORE)
                        time.sleep(0.08)
                    if not user32.IsWindowVisible(hwnd):
                        user32.ShowWindowAsync(hwnd, SW_SHOWNORMAL)
                        time.sleep(0.08)
                    self.window_mgr.grid_state[hwnd] = (monitor_idx, col, row)
                
                time.sleep(0.015)
            
            except Exception as e:
                log(f"[ERROR] load_workspace: {e}")
        
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
            
            windows_to_move = [
                (hwnd, pos) for hwnd, pos in self.window_mgr.grid_state.items()
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
            
            # Update grid_state
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
                    self.smart_tile_with_restore()
                winsound.PlaySound("SystemAsterisk", winsound.SND_ALIAS | winsound.SND_ASYNC)
            else:
                self.window_mgr.override_windows.add(hwnd)
                log(f"[OVERRIDE] {title[:60]} → override to ({'float' if default_useful else 'tile'})")
                if default_useful:
                    if hwnd in self.window_mgr.grid_state:
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
        # LOCK to prevent tiling during settings
        settings_lock = True
        
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
            
            # 🔒 Prevent tiling of this window
            dialog.update()  # Force update to get the hwnd
            
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
                
                # Close BEFORE modifying the settings
                dialog.destroy()
                root.destroy()
                
                # Wait for Tkinter to fully close
                time.sleep(0.3)
                
                # Now we can modify
                self.gap = int(gap_str)
                self.edge_padding = int(padding_str)
                self.window_mgr.gap = self.gap
                self.window_mgr.edge_padding = self.edge_padding
                
                log(f"[SETTINGS] GAP={self.gap}px, EDGE_PADDING={self.edge_padding}px")
                
                # Apply the new settings
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
            
            # Prevent re-tiling while the dialog is open
            old_ignore = self.ignore_retile_until
            self.ignore_retile_until = time.time() + 999999  # Block tiling
            
            def on_close():
                self.ignore_retile_until = old_ignore  # Restore
                dialog.destroy()
                root.destroy()
            
            dialog.protocol("WM_DELETE_WINDOW", on_close)
            
            root.mainloop()
            
            # Restore tiling
            self.ignore_retile_until = old_ignore
        
        except Exception as e:
            log(f"[ERROR] show_settings_dialog: {e}")
            self.ignore_retile_until = 0  # Restore in case of an error
    
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
        try:
            user32.RegisterHotKey(None, HOTKEY_TOGGLE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('T'))
            user32.RegisterHotKey(None, HOTKEY_RETILE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('R'))
            user32.RegisterHotKey(None, HOTKEY_QUIT, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('Q'))
            user32.RegisterHotKey(None, HOTKEY_MOVE_MONITOR, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('M'))
            user32.RegisterHotKey(None, HOTKEY_SWAP_MODE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('S'))
            user32.RegisterHotKey(None, HOTKEY_WS1, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('1'))
            user32.RegisterHotKey(None, HOTKEY_WS2, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('2'))
            user32.RegisterHotKey(None, HOTKEY_WS3, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('3'))
            user32.RegisterHotKey(None, HOTKEY_FLOAT_TOGGLE, win32con.MOD_CONTROL | win32con.MOD_ALT, ord('F'))
        except Exception as e:
            log(f"[ERROR] register_hotkeys: {e}")
    
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
            self.window_mgr.grid_state.clear()
            self.smart_tile_with_restore()
        else:
            for hwnd in list(self.window_mgr.grid_state.keys()):
                if user32.IsWindow(hwnd):
                    set_window_border(hwnd, None)
            self.window_mgr.grid_state.clear()
        
        self.update_tray_menu()
    
    def on_quit_from_tray(self, icon, item):
        """Quit from systray."""
        log("[TRAY] Quit requested")
        
        if self.tray_icon:
            self.tray_icon.stop()
        
        self.cleanup()
        os._exit(0)
    
    def cleanup(self):
        """Cleanup before exit."""
        try:
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
        """Background loop: auto-retile + border tracking."""
        while True:
            try:
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
                    # Lightweight cleanup
                    dead_count = self.window_mgr.cleanup_dead_windows()
                    
                    visible_windows = self.window_mgr.get_visible_windows(
                        self.monitors_cache, self.overlay_hwnd
                    )
                    current_count = len(visible_windows)
                    
                    if time.time() >= self.ignore_retile_until and current_count > 0:
                        if current_count != self.last_visible_count:
                            log(f"[AUTO-RETILE] {self.last_visible_count} → {current_count} windows")
                            self.smart_tile_with_restore()
                            self.last_visible_count = current_count
                        time.sleep(0.2)
                
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
                
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
            
            except Exception as e:
                log(f"[ERROR] message_loop: {e}")
    
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
    log("[MAIN] All hotkeys registered")
    
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
    finally:
        app.cleanup()
        log("[EXIT] SmartGrid stopped.")
        os._exit(0)