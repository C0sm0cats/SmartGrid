"""One serialized worker, queued results and no native work on Qt's UI thread."""
from concurrent.futures import ThreadPoolExecutor
from PySide6.QtCore import QObject, Signal, Slot, Qt


class AppBridge(QObject):
    changed = Signal()
    guides = Signal()
    error = Signal(str)
    status = Signal(str)
    busy_changed = Signal(bool)
    ui_action = Signal(str)
    _completed = Signal(object, object, object)

    def __init__(self, controller):
        super().__init__()
        self.controller = controller
        self.executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="smartgrid-command")
        self.pending = 0
        self.closed = False
        self._completed.connect(self._finish, Qt.ConnectionType.QueuedConnection)
        controller.subscribe(lambda *args: self.changed.emit() if not self.closed else None)
        if hasattr(controller, 'subscribe_guides'):
            controller.subscribe_guides(lambda: self.guides.emit() if not self.closed else None)
        if hasattr(controller, 'set_ui_action_handler'):
            controller.set_ui_action_handler(lambda action: self.ui_action.emit(action))

    def submit(self, action, *args, on_success=None, message="", **kwargs):
        if self.closed:
            return None
        fn = getattr(self.controller, action) if isinstance(action, str) else action
        self.pending += 1
        self.busy_changed.emit(True)
        future = self.executor.submit(fn, *args, **kwargs)
        def done(result):
            try:
                value = result.result()
            except Exception as exc:
                self._completed.emit(None, exc, None)
            else:
                self._completed.emit(value, None, (on_success, message))
        future.add_done_callback(done)
        return future

    @Slot(object, object, object)
    def _finish(self, value, exception, completion):
        self.pending = max(0, self.pending - 1)
        self.busy_changed.emit(bool(self.pending))
        if exception:
            self.error.emit(str(exception) or type(exception).__name__)
        elif completion:
            callback, message = completion
            if callback:
                try:
                    callback(value)
                except Exception as exc:
                    self.error.emit(str(exc))
            if message:
                self.status.emit(message)
        self.changed.emit()

    def shutdown(self, wait=True):
        self.closed = True
        self.executor.shutdown(wait=wait, cancel_futures=True)


def bridge_for(controller):
    bridge = getattr(controller, '_qt_bridge', None)
    if bridge is None:
        bridge = AppBridge(controller)
        controller._qt_bridge = bridge
    return bridge
