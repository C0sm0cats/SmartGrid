import ctypes as C
from .api import Win32

class InstanceGuard:
    """Per-login-session instance guard; close the handle to release the guard."""
    def __init__(self, name='SmartGrid', api=None):
        self.api = api or Win32()
        self.handle = self.api.CreateMutexW(None, False, 'Local\\' + name)
        self.api.check(self.handle, 'CreateMutexW')
        self.acquired = C.get_last_error() != 183
        if not self.acquired:
            self.close()

    def close(self):
        if self.handle:
            self.api.CloseHandle(self.handle)
            self.handle = None

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()
