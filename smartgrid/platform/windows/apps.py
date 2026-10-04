"""Stable executable/AUMID identities, Start menu catalogue and cached native icons."""
from __future__ import annotations
import ctypes as C
import hashlib
import json
import logging
import ntpath
import os
from pathlib import Path
import re
import struct
import subprocess
from .api import DWORD, UINT, FILETIME, SHFILEINFO, ICONINFO, BITMAP, BITMAPINFOHEADER
from smartgrid.core.models import ApplicationRef

log = logging.getLogger(__name__)

def application_id(executable='', aumid=''):
    if aumid:
        return 'aumid:' + aumid.casefold()
    return 'exe:' + ntpath.normpath(executable).casefold() if executable else ''

DOCUMENTS = ('.txt', '.pdf', '.chm', '.htm', '.html', '.url', '.rtf', '.md', '.ini', '.log', '.lnk')
_UNINSTALL = re.compile(r'(uninstall|d\u00e9sinstall|desinstall|deinstall)', re.IGNORECASE)


def catalogue_entry_wanted(name, target):
    """Start menu entries that are applications, not uninstallers, help or web links."""
    base = ntpath.basename(target).lower()
    if _UNINSTALL.search(name) or base.startswith('unins') or base.startswith('uninst'):
        return False
    return True


class AppCatalogue:
    def __init__(self, api, cache_dir=None):
        self.api = api
        self.cache_dir = Path(cache_dir or Path(os.environ.get('LOCALAPPDATA', Path.home())) / 'SmartGrid' / 'icons')
        self.apps = {}
        self.process_cache = {}

    def process_identity(self, pid):
        handle = self.api.OpenProcess(0x1000, False, pid)
        if not handle:
            return ApplicationRef(f'pid:{pid}', f'Process {pid}'), ''
        try:
            times = [FILETIME() for _ in range(4)]
            creation = ''
            if self.api.GetProcessTimes(handle, *(C.byref(item) for item in times)):
                creation = str((times[0].high << 32) | times[0].low)
            key = (pid, creation)
            if key in self.process_cache:
                return self.process_cache[key], creation
            path = C.create_unicode_buffer(32768)
            size = DWORD(len(path))
            executable = path.value if self.api.QueryFullProcessImageNameW(handle, 0, path, C.byref(size)) else ''
            length = UINT(0)
            aumid = ''
            if self.api.GetApplicationUserModelId(handle, C.byref(length), None) == 122 and length.value:
                buf = C.create_unicode_buffer(length.value)
                if self.api.GetApplicationUserModelId(handle, C.byref(length), buf) == 0:
                    aumid = buf.value
            identifier = application_id(executable, aumid) or f'pid:{pid}'
            app = self.apps.get(identifier) or ApplicationRef(identifier, ntpath.splitext(ntpath.basename(executable))[0] or f'Process {pid}', executable, aumid)
            self.apps[identifier] = app
            if creation:
                self.process_cache[key] = app
            # Bound a long-running desktop session's process cache.
            if len(self.process_cache) > 2048:
                self.process_cache = {key: app}
            return app, creation
        finally:
            self.api.CloseHandle(handle)

    def discover(self):
        """Start menu catalogue: Get-StartApps plus Start menu shortcuts, deduplicated.

        Desktop apps listed by Get-StartApps with a path AppID
        ({known-folder-GUID}\app.exe) become executable identities, so that a
        catalogue entry and a running window of the same app share one id.
        """
        # All calls are bounded; no command strings originate in user/archive data.
        script = r"""
$ErrorActionPreference='SilentlyContinue'; [Console]::OutputEncoding=[System.Text.UTF8Encoding]::new(); $out=@();
Get-StartApps | ForEach-Object { $out += [PSCustomObject]@{name=$_.Name; aumid=$_.AppID; executable=''} };
$ws=New-Object -ComObject WScript.Shell;
@([Environment]::GetFolderPath('StartMenu'),[Environment]::GetFolderPath('CommonStartMenu')) | ForEach-Object {
 Get-ChildItem $_ -Filter *.lnk -Recurse | ForEach-Object {
  $s=$ws.CreateShortcut($_.FullName); if($s.TargetPath -and (Test-Path $s.TargetPath)) {
   $out += [PSCustomObject]@{name=$_.BaseName; executable=$s.TargetPath; aumid=''}
  }
 }
}; ConvertTo-Json -InputObject @($out) -Compress
"""
        try:
            result = subprocess.run(['powershell.exe', '-NoLogo', '-NoProfile', '-NonInteractive', '-Command', script], capture_output=True, timeout=60, creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            if result.returncode:
                raise RuntimeError('Start menu catalogue query failed')
            rows = json.loads(result.stdout.decode('utf-8-sig', errors='replace') or '[]')
            entries = {}
            for row in rows if isinstance(rows, list) else [rows]:
                name = str(row.get('name') or '').strip()
                executable = str(row.get('executable') or '')
                aumid = str(row.get('aumid') or '')
                if not name or not catalogue_entry_wanted(name, executable or aumid):
                    continue
                if aumid and not executable:
                    resolved = self.resolve_path_appid(aumid)
                    if resolved is not None:
                        executable, aumid = resolved, ''
                # Only executable shortcuts are applications; never run documents or links.
                if executable and not executable.lower().endswith(('.exe', '.com')):
                    continue
                if aumid and ('://' in aumid or aumid.lower().endswith(DOCUMENTS)):
                    continue
                identifier = application_id(executable, aumid)
                if not identifier:
                    continue
                # Prefer the executable identity of a desktop app that also has an
                # explicit Start menu AppID: running windows report the executable.
                key = name.casefold()
                previous = entries.get(key)
                if previous and (previous[1] or '!' in previous[2] or not executable):
                    continue
                entries[key] = (name, executable, aumid, identifier)
            seen = set()
            for name, executable, aumid, identifier in entries.values():
                if identifier in seen:
                    continue
                seen.add(identifier)
                icon = self.icon(executable or ('shell:AppsFolder\\' + aumid))
                self.apps[identifier] = ApplicationRef(identifier, name, executable, aumid, icon, True)
        except (OSError, subprocess.TimeoutExpired, ValueError, RuntimeError, AttributeError, TypeError):
            log.exception('Application catalogue incomplete; discovered running applications retained')
        return sorted(self.apps.values(), key=lambda app: (app.name.casefold(), app.id))

    def resolve_path_appid(self, appid):
        """{known-folder-GUID}\\relative\\app.exe or an absolute path to an executable path."""
        match = re.fullmatch(r'\{([0-9a-fA-F-]{36})\}\\(.+)', appid)
        if match:
            folder = self.known_folder(match.group(1))
            if not folder:
                return None
            path = ntpath.join(folder, match.group(2))
        elif re.match(r'^[a-zA-Z]:\\', appid):
            path = appid
        else:
            return None
        return path if os.path.isfile(path) else None

    def known_folder(self, guid):
        try:
            from .virtual_desktops import GUID
            shell32 = C.WinDLL('shell32')
            ole32 = C.WinDLL('ole32')
            shell32.SHGetKnownFolderPath.argtypes = [C.POINTER(GUID), DWORD, C.c_void_p, C.POINTER(C.c_wchar_p)]
            shell32.SHGetKnownFolderPath.restype = C.c_long
            ole32.CoTaskMemFree.argtypes = [C.c_void_p]
            path = C.c_wchar_p()
            if shell32.SHGetKnownFolderPath(C.byref(GUID.parse(guid)), 0, None, C.byref(path)) != 0:
                return None
            try:
                return path.value
            finally:
                ole32.CoTaskMemFree(C.cast(path, C.c_void_p))
        except (OSError, AttributeError, ValueError):
            return None

    def icon(self, executable):
        if not executable:
            return ''
        target = self.cache_dir / (hashlib.sha256(executable.casefold().encode()).hexdigest() + '.ico')
        if target.exists():
            return str(target)
        info = SHFILEINFO()
        pidl = C.c_void_p()
        attributes = DWORD()
        source = executable
        flags = 0x100
        com_result = self.api.CoInitializeEx(None, 2)
        if com_result < 0 and com_result != -2147417850:  # existing MTA apartment is valid
            return ''
        try:
            if executable.startswith('shell:'):
                if self.api.SHParseDisplayName(executable, None, C.byref(pidl), 0, C.byref(attributes)) != 0:
                    return ''
                source = C.cast(pidl, C.c_wchar_p)
                flags |= 8
            found = self.api.SHGetFileInfoW(source, 0, C.byref(info), C.sizeof(info), flags)
        finally:
            if pidl:
                self.api.CoTaskMemFree(pidl)
            if com_result >= 0:
                self.api.CoUninitialize()
        if not found:
            return ''
        icon = ICONINFO()
        try:
            if not self.api.GetIconInfo(info.hIcon, C.byref(icon)) or not icon.hbmColor:
                return ''
            bitmap = BITMAP()
            if not self.api.GetObjectW(icon.hbmColor, C.sizeof(bitmap), C.byref(bitmap)):
                return ''
            width, height = bitmap.bmWidth, abs(bitmap.bmHeight)
            if not (0 < width <= 256 and 0 < height <= 256):
                return ''
            header = BITMAPINFOHEADER(C.sizeof(BITMAPINFOHEADER), width, height, 1, 32, 0, width * height * 4, 0, 0, 0, 0)
            pixels = C.create_string_buffer(width * height * 4)
            dc = self.api.GetDC(None)
            try:
                if self.api.GetDIBits(dc, icon.hbmColor, 0, height, pixels, C.byref(header), 0) != height:
                    return ''
            finally:
                self.api.ReleaseDC(None, dc)
            mask_stride = ((width + 31) // 32) * 4
            mask_header = BITMAPINFOHEADER(40, width, height, 1, 1, 0, mask_stride * height, 0, 0, 2, 0)
            mask_info = C.create_string_buffer(bytes(mask_header) + b'\0\0\0\0\xff\xff\xff\0')
            mask = C.create_string_buffer(mask_stride * height)
            dc = self.api.GetDC(None)
            try:
                self.api.GetDIBits(dc, icon.hbmMask, 0, height, mask, mask_info, 0)
            finally:
                self.api.ReleaseDC(None, dc)
            header.biHeight = height * 2
            data = bytes(header) + pixels.raw + mask.raw
            self.cache_dir.mkdir(parents=True, exist_ok=True)
            target.write_bytes(struct.pack('<HHH', 0, 1, 1) + struct.pack('<BBBBHHII', width % 256, height % 256, 0, 0, 1, 32, len(data), 22) + data)
            return str(target)
        except OSError:
            log.exception('Could not cache application icon')
            return ''
        finally:
            if icon.hbmColor:
                self.api.DeleteObject(icon.hbmColor)
            if icon.hbmMask:
                self.api.DeleteObject(icon.hbmMask)
            if info.hIcon:
                self.api.DestroyIcon(info.hIcon)
