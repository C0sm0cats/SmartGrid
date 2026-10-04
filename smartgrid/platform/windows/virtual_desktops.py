"""Documented IVirtualDesktopManager access, COM ownership scoped to one scan.

Explorer's read-only registry value is a fallback for an empty desktop. No
undocumented COM interface, workspace switch or window transfer is used.
"""
from contextlib import contextmanager
import ctypes as C
import sys
import os
import uuid
from .api import CALLBACK,BOOL,LONG,HWND


class GUID(C.Structure):
    _fields_=[('data1',C.c_uint32),('data2',C.c_uint16),('data3',C.c_uint16),('data4',C.c_ubyte*8)]
    @classmethod
    def parse(cls,value): return cls.from_buffer_copy(uuid.UUID(value).bytes_le)
    def text(self):
        value=uuid.UUID(bytes_le=bytes(self))
        return str(value) if value.int else None


def scoped_display(display_id,desktop):
    return display_id.split('@desktop:',1)[0]+('@desktop:'+desktop if desktop else '')


class DesktopQuery:
    def __init__(self,pointer):
        self.pointer=pointer
        self.table=C.cast(pointer,C.POINTER(C.POINTER(C.c_void_p))).contents
    def call(self,index,result,*types): return CALLBACK(result,C.c_void_p,*types)(self.table[index])
    def desktop_id(self,hwnd):
        if not hwnd: return None
        value=GUID()
        return value.text() if self.call(4,LONG,HWND,C.POINTER(GUID))(self.pointer,hwnd,C.byref(value))==0 else None
    def on_current(self,hwnd):
        value=BOOL()
        result=self.call(3,LONG,HWND,C.POINTER(BOOL))(self.pointer,hwnd,C.byref(value))
        return bool(value.value) if result==0 else None


class VirtualDesktops:
    def __init__(self):
        self.available=False
        if sys.platform!='win32': return
        self.ole=C.WinDLL('ole32')
        self.ole.CoInitializeEx.argtypes=[C.c_void_p,C.c_uint32];self.ole.CoInitializeEx.restype=LONG
        self.ole.CoUninitialize.argtypes=[];self.ole.CoUninitialize.restype=None
        self.ole.CoCreateInstance.argtypes=[C.POINTER(GUID),C.c_void_p,C.c_uint32,C.POINTER(GUID),C.POINTER(C.c_void_p)]
        self.ole.CoCreateInstance.restype=LONG
        self.kernel=C.WinDLL('kernel32')
        self.kernel.ProcessIdToSessionId.argtypes=[C.c_uint32,C.POINTER(C.c_uint32)]
        self.kernel.ProcessIdToSessionId.restype=BOOL
        self.available=True

    @contextmanager
    def query(self):
        if not self.available:
            yield None;return
        result=self.ole.CoInitializeEx(None,2)
        initialized=result in (0,1)
        if result not in (0,1,-2147417850): # RPC_E_CHANGED_MODE: reuse caller apartment.
            yield None;return
        pointer=C.c_void_p()
        try:
            cls=GUID.parse('c5e0cdca-7b6e-41b2-9fc4-d93975cc467b')
            iid=GUID.parse('a5cd92ff-29be-454c-8d04-d82879fb3f1b')
            result=self.ole.CoCreateInstance(C.byref(cls),None,1,C.byref(iid),C.byref(pointer))
            yield DesktopQuery(pointer) if result==0 and pointer.value else None
        finally:
            if pointer.value:
                query=DesktopQuery(pointer);query.call(2,C.c_uint32)(pointer)
            if initialized: self.ole.CoUninitialize()

    def _explorer_value(self,name):
        """First 16-byte-multiple binary value from the session key, then the global key."""
        import winreg
        base=r'Software\Microsoft\Windows\CurrentVersion\Explorer'
        session=C.c_uint32()
        keys=[base+r'\VirtualDesktops']
        if self.kernel.ProcessIdToSessionId(os.getpid(),C.byref(session)):
            keys.insert(0,base+rf'\SessionInfo\{session.value}\VirtualDesktops')
        for key_name in keys:
            try:
                with winreg.OpenKey(winreg.HKEY_CURRENT_USER,key_name) as key:
                    raw,_=winreg.QueryValueEx(key,name)
                    if isinstance(raw,bytes) and raw and len(raw)%16==0: return raw
            except OSError: continue
        return None

    def resolve(self,candidates=()):
        """Return (desktop GUID or None, identified).

        A single Windows desktop is always identified: Windows writes no
        CurrentVirtualDesktop value until several desktops have been used.
        """
        if not self.available: return None,True
        # Registry identifies an empty desktop; pinned windows may report their
        # original desktop, so the Explorer value takes precedence when present.
        try:
            raw=self._explorer_value('CurrentVirtualDesktop')
            value=GUID.from_buffer_copy(raw[:16]).text() if raw and len(raw)==16 else None
            if value: return value,True
            desktops=self._explorer_value('VirtualDesktopIDs')
            if not desktops or len(desktops)==16:
                return (GUID.from_buffer_copy(desktops).text() if desktops else None),True
        except (OSError,ValueError,ImportError): pass
        with self.query() as query:
            if query:
                for hwnd in (candidates() if callable(candidates) else candidates):
                    if hwnd and query.on_current(hwnd):
                        value=query.desktop_id(hwnd)
                        if value: return value,True
        return None,False

    def current(self,foreground):
        return self.resolve([foreground])[0]
