"""Explicit Win32 ABI. Importing this module never loads DLLs on other platforms."""
from __future__ import annotations
import ctypes as C
import sys
from contextlib import contextmanager

BOOL = LONG = C.c_int32
DWORD = UINT = C.c_uint32
WORD = C.c_uint16
BYTE = C.c_uint8
HANDLE = HWND = HMONITOR = HDC = HINSTANCE = C.c_void_p
LPARAM = LRESULT = C.c_ssize_t
WPARAM = ULONG_PTR = C.c_size_t
WCHAR = C.c_wchar
CALLBACK = getattr(C, 'WINFUNCTYPE', C.CFUNCTYPE)

class POINT(C.Structure):
    _fields_ = [('x', LONG), ('y', LONG)]
class RECT(C.Structure):
    _fields_ = [('left', LONG), ('top', LONG), ('right', LONG), ('bottom', LONG)]
class DWM_THUMBNAIL_PROPERTIES(C.Structure):
    _fields_ = [('dwFlags', DWORD), ('rcDestination', RECT), ('rcSource', RECT),
                ('opacity', BYTE), ('fVisible', BOOL), ('fSourceClientAreaOnly', BOOL)]
class MONITORINFOEX(C.Structure):
    _fields_ = [('cbSize', DWORD), ('rcMonitor', RECT), ('rcWork', RECT), ('dwFlags', DWORD), ('szDevice', WCHAR * 32)]
class DISPLAY_DEVICE(C.Structure):
    _fields_ = [('cb', DWORD), ('DeviceName', WCHAR * 32), ('DeviceString', WCHAR * 128), ('StateFlags', DWORD), ('DeviceID', WCHAR * 128), ('DeviceKey', WCHAR * 128)]
class WINDOWPLACEMENT(C.Structure):
    _fields_ = [('length', UINT), ('flags', UINT), ('showCmd', UINT), ('ptMinPosition', POINT), ('ptMaxPosition', POINT), ('rcNormalPosition', RECT)]
class MINMAXINFO(C.Structure):
    _fields_ = [('ptReserved', POINT), ('ptMaxSize', POINT), ('ptMaxPosition', POINT), ('ptMinTrackSize', POINT), ('ptMaxTrackSize', POINT)]
class MSG(C.Structure):
    _fields_ = [('hwnd', HWND), ('message', UINT), ('wParam', WPARAM), ('lParam', LPARAM), ('time', DWORD), ('pt', POINT), ('lPrivate', DWORD)]
WNDPROC = CALLBACK(LRESULT, HWND, UINT, WPARAM, LPARAM)
WINEVENTPROC = CALLBACK(None, HANDLE, DWORD, HWND, LONG, LONG, DWORD, DWORD)
ENUMPROC = CALLBACK(BOOL, HWND, LPARAM)
MONITORPROC = CALLBACK(BOOL, HMONITOR, HDC, C.POINTER(RECT), LPARAM)
class WNDCLASSEX(C.Structure):
    _fields_ = [('cbSize', UINT), ('style', UINT), ('lpfnWndProc', WNDPROC), ('cbClsExtra', C.c_int), ('cbWndExtra', C.c_int), ('hInstance', HINSTANCE), ('hIcon', HANDLE), ('hCursor', HANDLE), ('hbrBackground', HANDLE), ('lpszMenuName', C.c_wchar_p), ('lpszClassName', C.c_wchar_p), ('hIconSm', HANDLE)]
class FILETIME(C.Structure):
    _fields_ = [('low', DWORD), ('high', DWORD)]
class SHFILEINFO(C.Structure):
    _fields_ = [('hIcon', HANDLE), ('iIcon', C.c_int), ('dwAttributes', DWORD), ('szDisplayName', WCHAR * 260), ('szTypeName', WCHAR * 80)]
class ICONINFO(C.Structure):
    _fields_ = [('fIcon', BOOL), ('xHotspot', DWORD), ('yHotspot', DWORD), ('hbmMask', HANDLE), ('hbmColor', HANDLE)]
class BITMAP(C.Structure):
    _fields_ = [('bmType', LONG), ('bmWidth', LONG), ('bmHeight', LONG), ('bmWidthBytes', LONG), ('bmPlanes', WORD), ('bmBitsPixel', WORD), ('bmBits', C.c_void_p)]
class BITMAPINFOHEADER(C.Structure):
    _fields_ = [('biSize', DWORD), ('biWidth', LONG), ('biHeight', LONG), ('biPlanes', WORD), ('biBitCount', WORD), ('biCompression', DWORD), ('biSizeImage', DWORD), ('biXPelsPerMeter', LONG), ('biYPelsPerMeter', LONG), ('biClrUsed', DWORD), ('biClrImportant', DWORD)]

class Win32:
    def __init__(self):
        if sys.platform != 'win32':
            raise RuntimeError('WindowsBackend requires an interactive Windows desktop')
        self.user32 = C.WinDLL('user32', use_last_error=True)
        self.kernel32 = C.WinDLL('kernel32', use_last_error=True)
        self.dwmapi = C.WinDLL('dwmapi', use_last_error=True)
        self.shell32 = C.WinDLL('shell32', use_last_error=True)
        self.shcore = C.WinDLL('shcore', use_last_error=True)
        self.gdi32 = C.WinDLL('gdi32', use_last_error=True)
        self.ole32 = C.WinDLL('ole32', use_last_error=True)
        p = C.POINTER
        signatures = {
            'user32': {
                'EnumWindows': (BOOL, [ENUMPROC, LPARAM]), 'EnumChildWindows': (BOOL, [HWND, ENUMPROC, LPARAM]), 'EnumDisplayMonitors': (BOOL, [HDC, p(RECT), MONITORPROC, LPARAM]),
                'GetMonitorInfoW': (BOOL, [HMONITOR, p(MONITORINFOEX)]), 'EnumDisplayDevicesW': (BOOL, [C.c_wchar_p, DWORD, p(DISPLAY_DEVICE), DWORD]),
                'MonitorFromWindow': (HMONITOR, [HWND, DWORD]), 'GetWindowRect': (BOOL, [HWND, p(RECT)]),
                'IsWindow': (BOOL, [HWND]), 'IsWindowVisible': (BOOL, [HWND]), 'IsIconic': (BOOL, [HWND]), 'IsZoomed': (BOOL, [HWND]), 'IsHungAppWindow': (BOOL, [HWND]), 'MoveWindow': (BOOL, [HWND, C.c_int, C.c_int, C.c_int, C.c_int, BOOL]),
                'GetWindow': (HWND, [HWND, UINT]), 'GetAncestor': (HWND, [HWND, UINT]),
                'GetWindowThreadProcessId': (DWORD, [HWND, p(DWORD)]), 'GetWindowTextLengthW': (C.c_int, [HWND]),
                'SetWindowTextW': (BOOL, [HWND, C.c_wchar_p]), 'GetWindowTextW': (C.c_int, [HWND, C.c_wchar_p, C.c_int]), 'GetClassNameW': (C.c_int, [HWND, C.c_wchar_p, C.c_int]),
                'GetWindowPlacement': (BOOL, [HWND, p(WINDOWPLACEMENT)]), 'SetWindowPlacement': (BOOL, [HWND, p(WINDOWPLACEMENT)]),
                'SetWindowPos': (BOOL, [HWND, HWND, C.c_int, C.c_int, C.c_int, C.c_int, UINT]),
                'ShowWindowAsync': (BOOL, [HWND, C.c_int]), 'SetForegroundWindow': (BOOL, [HWND]), 'GetForegroundWindow': (HWND, []),
                'PostMessageW': (BOOL, [HWND, UINT, WPARAM, LPARAM]),
                'SendMessageTimeoutW': (LRESULT, [HWND, UINT, WPARAM, LPARAM, UINT, UINT, p(ULONG_PTR)]),
                'SetWinEventHook': (HANDLE, [DWORD, DWORD, HANDLE, WINEVENTPROC, DWORD, DWORD, DWORD]), 'UnhookWinEvent': (BOOL, [HANDLE]),
                'RegisterHotKey': (BOOL, [HWND, C.c_int, UINT, UINT]), 'UnregisterHotKey': (BOOL, [HWND, C.c_int]),
                'GetMessageW': (BOOL, [p(MSG), HWND, UINT, UINT]), 'PeekMessageW': (BOOL, [p(MSG), HWND, UINT, UINT, UINT]),
                'TranslateMessage': (BOOL, [p(MSG)]), 'DispatchMessageW': (LRESULT, [p(MSG)]), 'PostThreadMessageW': (BOOL, [DWORD, UINT, WPARAM, LPARAM]),
                'RegisterClassExW': (WORD, [p(WNDCLASSEX)]), 'UnregisterClassW': (BOOL, [C.c_wchar_p, HINSTANCE]),
                'CreateWindowExW': (HWND, [DWORD, C.c_wchar_p, C.c_wchar_p, DWORD, C.c_int, C.c_int, C.c_int, C.c_int, HWND, HANDLE, HINSTANCE, C.c_void_p]),
                'DestroyWindow': (BOOL, [HWND]), 'DefWindowProcW': (LRESULT, [HWND, UINT, WPARAM, LPARAM]),
                'SetLayeredWindowAttributes': (BOOL, [HWND, DWORD, BYTE, DWORD]), 'GetDC': (HDC, [HWND]), 'ReleaseDC': (C.c_int, [HWND, HDC]),
                'DestroyIcon': (BOOL, [HANDLE]), 'GetIconInfo': (BOOL, [HANDLE, p(ICONINFO)]),
                'GetCursorPos': (BOOL, [p(POINT)]), 'GetAsyncKeyState': (C.c_short, [C.c_int]),
            },
            'kernel32': {
                'OpenProcess': (HANDLE, [DWORD, BOOL, DWORD]), 'CloseHandle': (BOOL, [HANDLE]),
                'QueryFullProcessImageNameW': (BOOL, [HANDLE, DWORD, C.c_wchar_p, p(DWORD)]),
                'GetProcessTimes': (BOOL, [HANDLE, p(FILETIME), p(FILETIME), p(FILETIME), p(FILETIME)]),
                'GetApplicationUserModelId': (LONG, [HANDLE, p(UINT), C.c_wchar_p]),
                'GetCurrentThreadId': (DWORD, []), 'GetModuleHandleW': (HANDLE, [C.c_wchar_p]),
                'CreateMutexW': (HANDLE, [C.c_void_p, BOOL, C.c_wchar_p]),
            },
            'dwmapi': {'DwmGetWindowAttribute': (LONG, [HWND, DWORD, C.c_void_p, DWORD]), 'DwmSetWindowAttribute': (LONG, [HWND, DWORD, C.c_void_p, DWORD]),
                'DwmGetColorizationColor': (LONG, [p(DWORD),p(BOOL)]),
                'DwmRegisterThumbnail': (LONG, [HWND, HWND, p(HANDLE)]),
                'DwmUpdateThumbnailProperties': (LONG, [HANDLE, p(DWM_THUMBNAIL_PROPERTIES)]),
                'DwmUnregisterThumbnail': (LONG, [HANDLE])},
            'ole32': {'CoInitializeEx': (LONG, [C.c_void_p, DWORD]), 'CoUninitialize': (None, []), 'CoTaskMemFree': (None, [C.c_void_p])},
            'shell32': {'SHParseDisplayName': (LONG, [C.c_wchar_p, C.c_void_p, p(C.c_void_p), DWORD, p(DWORD)]), 'ShellExecuteW': (C.c_ssize_t, [HWND, C.c_wchar_p, C.c_wchar_p, C.c_wchar_p, C.c_wchar_p, C.c_int]), 'SHGetFileInfoW': (ULONG_PTR, [C.c_wchar_p, DWORD, p(SHFILEINFO), UINT, UINT])},
            'shcore': {'GetDpiForMonitor': (LONG, [HMONITOR, C.c_int, p(UINT), p(UINT)])},
            'gdi32': {'GetObjectW': (C.c_int, [HANDLE, C.c_int, C.c_void_p]), 'GetDIBits': (C.c_int, [HDC, HANDLE, UINT, UINT, C.c_void_p, C.c_void_p, UINT]), 'DeleteObject': (BOOL, [HANDLE]), 'CreateSolidBrush': (HANDLE, [DWORD])},
        }
        for dll, entries in signatures.items():
            for name, (result, args) in entries.items():
                fn = getattr(getattr(self, dll), name)
                fn.restype, fn.argtypes = result, args
                setattr(self, name, fn)
        for base in ('GetWindowLong', 'SetWindowLong'):
            name = base + ('PtrW' if C.sizeof(C.c_void_p) == 8 else 'W')
            fn = getattr(self.user32, name)
            fn.restype = C.c_ssize_t if C.sizeof(C.c_void_p) == 8 else LONG
            fn.argtypes = [HWND, C.c_int] + ([fn.restype] if base == 'SetWindowLong' else [])
            setattr(self, base, fn)
        for name, result, args in [('SetProcessDpiAwarenessContext', BOOL, [HANDLE]), ('SetThreadDpiAwarenessContext', HANDLE, [HANDLE]), ('GetDpiForWindow', UINT, [HWND])]:
            fn = getattr(self.user32, name, None)
            if fn is not None:
                fn.restype, fn.argtypes = result, args
            setattr(self, name, fn)

    @contextmanager
    def physical_coordinates(self):
        old = self.SetThreadDpiAwarenessContext(HANDLE(-4)) if self.SetThreadDpiAwarenessContext else None
        try:
            yield
        finally:
            if old:
                self.SetThreadDpiAwarenessContext(old)

    def check(self, result, operation):
        if not result:
            raise OSError(C.get_last_error(), operation)
        return result
