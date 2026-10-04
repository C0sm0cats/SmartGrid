"""Serviced native owner thread; commands and cleanup always execute on owner."""
from __future__ import annotations
import ctypes as C
import logging
import queue
import threading
from concurrent.futures import Future
from .api import MSG, WNDCLASSEX, WNDPROC

log = logging.getLogger(__name__)
WM_COMMAND_QUEUE = 0x8001

class MessageThread:
    def __init__(self, api, name='SmartGrid native', window_callback=None):
        self.api, self.name, self.window_callback = api, name, window_callback
        self.commands = queue.Queue()
        self.ready = threading.Event()
        self.thread = None
        self.thread_id = 0
        self.hwnd = None
        self.error = None
        self._closing = False
        self._lock = threading.RLock()
        self._class_name = f'SmartGridOwner_{id(self)}'
        self._cleanup = []

    def start(self):
        with self._lock:
            if self.thread and self.thread.is_alive():
                return
            self.thread = threading.Thread(target=self._run, name=self.name, daemon=False)
            self.thread.start()
        if not self.ready.wait(5):
            raise TimeoutError('Native message owner failed to start')
        if self.error:
            raise self.error

    def _run(self):
        api = self.api
        try:
            with api.physical_coordinates():
                self.thread_id = api.GetCurrentThreadId()
                self._wndproc = WNDPROC(self._window_message)
                wc = WNDCLASSEX()
                wc.cbSize = C.sizeof(wc)
                wc.lpfnWndProc = self._wndproc
                wc.hInstance = api.GetModuleHandleW(None)
                wc.lpszClassName = self._class_name
                api.check(api.RegisterClassExW(C.byref(wc)), 'RegisterClassExW')
                # Invisible top-level window receives display/settings broadcasts.
                self.hwnd = api.CreateWindowExW(0x80, self._class_name, '', 0x80000000, 0, 0, 0, 0, None, None, wc.hInstance, None)
                api.check(self.hwnd, 'CreateWindowExW')
                self.ready.set()
                msg = MSG()
                while True:
                    status = api.GetMessageW(C.byref(msg), None, 0, 0)
                    if status == -1:
                        raise OSError(C.get_last_error(), 'GetMessageW')
                    if status == 0:
                        break
                    if msg.message == WM_COMMAND_QUEUE:
                        self._drain()
                    else:
                        self.on_message(msg)
                        api.TranslateMessage(C.byref(msg))
                        api.DispatchMessageW(C.byref(msg))
        except BaseException as error:
            self.error = error
            log.exception('Native owner failed')
            self.ready.set()
        finally:
            for callback in reversed(self._cleanup):
                try:
                    callback()
                except Exception:
                    log.exception('Native cleanup failed')
            if self.hwnd:
                api.DestroyWindow(self.hwnd)
                self.hwnd = None
            api.UnregisterClassW(self._class_name, api.GetModuleHandleW(None))
            while not self.commands.empty():
                _, future = self.commands.get()
                if not future.done():
                    future.set_exception(RuntimeError('Native message owner stopped'))

    def _window_message(self, hwnd, message, wparam, lparam):
        try:
            if self.window_callback:
                result = self.window_callback(hwnd, message, wparam, lparam)
                if result is not None:
                    return result
        except Exception:
            log.exception('Native message callback failed')
        return self.api.DefWindowProcW(hwnd, message, wparam, lparam)

    def on_message(self, message):
        pass

    def _drain(self):
        while not self.commands.empty():
            callback, future = self.commands.get()
            if future.set_running_or_notify_cancel():
                try:
                    future.set_result(callback())
                except BaseException as error:
                    future.set_exception(error)

    def call(self, callback, timeout=5):
        if threading.current_thread() is self.thread:
            return callback()
        with self._lock:
            if self._closing:
                raise RuntimeError('Native message owner is closing')
            self.start()
            future = Future()
            self.commands.put((callback, future))
            self.api.check(self.api.PostThreadMessageW(self.thread_id, WM_COMMAND_QUEUE, 0, 0), 'PostThreadMessageW')
        return future.result(timeout=timeout)

    def stop(self):
        with self._lock:
            if not self.thread or not self.thread.is_alive() or self._closing:
                return
            self._closing = True
            self.api.check(self.api.PostThreadMessageW(self.thread_id, 0x12, 0, 0), 'PostThreadMessageW(WM_QUIT)')
        if threading.current_thread() is not self.thread:
            self.thread.join(5)
            if self.thread.is_alive():
                raise TimeoutError('Native owner did not stop')
