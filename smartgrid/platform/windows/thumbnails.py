"""Optional live DWM thumbnails, owned by the Studio's Qt thread.

No PrintWindow, screen capture, activation or client messages. An unsupported
source falls back to its ordinary app card. Handles are scoped to a destination
top-level window and the source process generation, never just a reused HWND.
"""
import ctypes as C
from .api import HANDLE, RECT, DWM_THUMBNAIL_PROPERTIES


class ThumbnailManager:
    def __init__(self, api):
        self.api=api
        self.entries={}

    def clear(self):
        for handle,_,_ in self.entries.values():
            self.api.DwmUnregisterThumbnail(handle)
        self.entries.clear()

    def sync(self, destination, targets):
        """targets maps slot -> (WindowRef, physical destination Rect)."""
        active=set()
        for key in self.entries.keys()-targets.keys():
            self.api.DwmUnregisterThumbnail(self.entries.pop(key)[0])
        for key,(source,rect) in targets.items():
            identity=(destination,source)
            entry=self.entries.get(key)
            if entry and entry[1]!=identity:
                self.api.DwmUnregisterThumbnail(self.entries.pop(key)[0])
                entry=None
            if entry is None:
                handle=HANDLE()
                if self.api.DwmRegisterThumbnail(destination,source.hwnd,C.byref(handle))!=0 or not handle.value:
                    continue
                entry=(handle.value,identity,None)
            handle,_,previous_rect=entry
            if previous_rect!=rect:
                properties=DWM_THUMBNAIL_PROPERTIES()
                properties.dwFlags=0x01|0x04|0x08|0x10
                properties.rcDestination=RECT(rect.x,rect.y,rect.right,rect.bottom)
                properties.opacity=255
                properties.fVisible=True
                properties.fSourceClientAreaOnly=False
                if self.api.DwmUpdateThumbnailProperties(handle,C.byref(properties))!=0:
                    self.api.DwmUnregisterThumbnail(handle)
                    self.entries.pop(key,None)
                    continue
            self.entries[key]=(handle,identity,rect)
            active.add(key)
        return active
