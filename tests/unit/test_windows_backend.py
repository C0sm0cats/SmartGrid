import ctypes as C
from contextlib import nullcontext
import sys
from unittest import TestCase
from smartgrid.core.models import Rect, ApplicationRef
from smartgrid.platform.windows.api import RECT, WINDOWPLACEMENT, MSG, HWND, LPARAM, DWORD, Win32
from smartgrid.platform.windows.backend import WindowsBackend, eligibility, animation_rect
from smartgrid.platform.windows.apps import application_id
from smartgrid.platform.windows.hotkeys import parse_shortcut
from smartgrid.platform.fake import FakeBackend

class PolicyTests(TestCase):
    def test_passive_restore_does_not_request_foreground_activation(self):
        from types import SimpleNamespace
        calls=[]
        backend=WindowsBackend.__new__(WindowsBackend)
        backend.alive=lambda hwnd:True
        backend.api=SimpleNamespace(ShowWindowAsync=lambda h,cmd:calls.append((h,cmd)) or True,
            IsIconic=lambda h:False,IsZoomed=lambda h:False,IsWindowVisible=lambda h:True)
        self.assertTrue(backend.show(42))
        self.assertEqual(calls,[(42,4)])

    def test_firefox_and_normal_titles_have_no_substring_filter(self):
        for name in ['MozillaWindowClass', 'Chrome_WidgetWin_1', 'TkTopLevel', 'ConsoleWindowClass']:
            allowed, reason = eligibility(visible=True, cloaked=False, owner=0, style=0x00cf0000, exstyle=0, class_name=name)
            self.assertTrue(allowed, reason)

    def test_popup_and_other_desktop_reasons(self):
        values = dict(visible=True, cloaked=False, owner=0, style=0, exstyle=0, class_name='Normal')
        for key, value in [('cloaked', True), ('owner', 17), ('style', 0x40000000), ('exstyle', 0x80)]:
            allowed, reason = eligibility(**dict(values, **{key:value}))
            self.assertFalse(allowed)
            self.assertTrue(reason)
        self.assertTrue(eligibility(**dict(values, owner=1, exstyle=0x40000))[0])

    def test_native_abi_fixed_width_and_pointer_fields(self):
        self.assertEqual(C.sizeof(DWORD), 4)
        self.assertEqual(C.sizeof(RECT), 16)
        self.assertEqual(C.sizeof(WINDOWPLACEMENT), 44)
        self.assertEqual(C.sizeof(HWND), C.sizeof(C.c_void_p))
        self.assertEqual(C.sizeof(LPARAM), C.sizeof(C.c_void_p))
        self.assertEqual(MSG.wParam.offset, 16 if C.sizeof(HWND) == 8 else 8)

    def test_import_safe_and_clear_platform_failure(self):
        if sys.platform != 'win32':
            with self.assertRaisesRegex(RuntimeError, 'Windows'):
                Win32()

    def test_stable_application_identity(self):
        self.assertEqual(application_id(r'C:\Apps\FIREFOX.EXE'), application_id(r'c:\apps\firefox.exe'))
        self.assertEqual(application_id('host.exe', 'Package!App'), 'aumid:package!app')

    def test_hotkey_parser(self):
        self.assertEqual(parse_shortcut('Ctrl+Alt+Left'), (0x4003, 0x25))
        self.assertEqual(parse_shortcut('Win+Shift+F24'), (0x400c, 0x87))
        with self.assertRaises(ValueError):
            parse_shortcut('Ctrl+Imaginary')

class PlacementTests(TestCase):
    def backend(self):
        backend = WindowsBackend.__new__(WindowsBackend)
        class API:
            physical_coordinates = staticmethod(nullcontext)
            IsIconic = staticmethod(lambda hwnd: False)
            IsZoomed = staticmethod(lambda hwnd: False)
        backend.api = API()
        backend.current = Rect(5, 5, 200, 200)
        backend.window_rect = lambda hwnd: backend.current
        backend.alive = lambda hwnd: True
        backend.min_size = lambda hwnd: (100, 100)
        return backend

    def test_constraint_failure_does_not_force_style(self):
        backend = self.backend()
        backend._place_once = lambda *args: self.fail('Must not attempt constrained placement')
        result = backend.place_result(1, Rect(0, 0, 20, 20))
        self.assertFalse(result.success)
        self.assertIn('minimum', result.reason)
        self.assertEqual(result.attempts, 0)

    def test_effects_end_exactly_and_respect_application_minimums(self):
        source, target = Rect(100, 100, 600, 400), Rect(500, 300, 100, 100)
        for effect in ('crit_damped', 'spring', 'curved', 'linear'):
            self.assertEqual(animation_rect(source, target, 0, effect), source)
            self.assertEqual(animation_rect(source, target, 1, effect), target)
            for frame in range(1, 100):
                rect = animation_rect(source, target, frame / 100, effect, (100, 100))
                self.assertGreaterEqual(rect.width, 100)
                self.assertGreaterEqual(rect.height, 100)
        self.assertNotEqual(animation_rect(source, target, .5, 'curved'), animation_rect(source, target, .5, 'linear'))

    def test_success_and_cancel_use_same_verification(self):
        backend = self.backend()
        def place(hwnd, rect):
            backend.current = rect
            return True
        backend._place_once = place
        self.assertTrue(backend.place_result(1, Rect(-400, 0, 150, 100)).success)
        self.assertEqual(backend.place_result(1, Rect(0, 0, 150, 100), cancel=lambda:True).reason, 'Cancelled')

    def test_refusing_app_is_bounded_and_reported(self):
        backend = self.backend()
        backend._place_once = lambda hwnd, rect: True
        result = backend.place_result(1, Rect(0, 0, 150, 100), timeout=.05, retries=20)
        self.assertFalse(result.success)
        self.assertLessEqual(result.attempts, 1)
        self.assertEqual(result.actual, backend.current)

class FakeTests(TestCase):
    def test_restoration_and_handle_reuse(self):
        backend = FakeBackend(apps=[ApplicationRef('firefox', 'Firefox')])
        window = backend.add_window('firefox')
        hwnd = window.ref.hwnd
        before = backend.snapshot(hwnd)
        self.assertTrue(backend.place(hwnd, Rect(0, 0, 500, 500)))
        backend.minimize(hwnd)
        self.assertTrue(backend.restore(hwnd, before))
        self.assertEqual(backend.windows[hwnd].rect, window.rect)
        backend.remove_window(hwnd)
        backend.add_window('firefox', hwnd=hwnd)
        self.assertFalse(backend.restore(hwnd, before))

    def test_launch_can_arrive_late_and_events_stop(self):
        backend = FakeBackend(apps=[ApplicationRef('a', 'Application')])
        events = []
        backend.start_events(events.append)
        backend.auto_launch = False
        self.assertTrue(backend.launch(backend.apps['a']))
        self.assertFalse(backend.windows)
        window = backend.add_window('a')
        self.assertEqual(events[-1]['type'], 'created')
        backend.minimums[window.ref.hwnd] = (400, 300)
        self.assertFalse(backend.place(window.ref.hwnd, Rect(0, 0, 200, 200)))
        backend.stop()
        backend.remove_window(window.ref.hwnd)
        self.assertEqual(len(events), 1)

class OwnerThreadTests(TestCase):
    def transport(self):
        import queue
        import threading
        from smartgrid.platform.windows.api import MSG
        class API:
            def __init__(self):
                self.messages = queue.Queue()
                self.owners = []
                self.hotkeys = {}
            physical_coordinates = staticmethod(nullcontext)
            GetCurrentThreadId = staticmethod(threading.get_ident)
            GetModuleHandleW = staticmethod(lambda name: 1)
            RegisterClassExW = staticmethod(lambda wc: 1)
            CreateWindowExW = staticmethod(lambda *args: 0x123456789)
            DefWindowProcW = staticmethod(lambda *args: 0)
            UnregisterClassW = staticmethod(lambda *args: 1)
            TranslateMessage = DispatchMessageW = staticmethod(lambda *args: 1)
            def check(self, result, name):
                if not result:
                    raise RuntimeError(name)
                return result
            def DestroyWindow(self, hwnd):
                self.owners.append(('destroy', threading.get_ident(), hwnd))
                return 1
            def PostThreadMessageW(self, thread_id, message, wp, lp):
                self.messages.put((message, wp, lp))
                return 1
            def GetMessageW(self, pointer, hwnd, lo, hi):
                message, wp, lp = self.messages.get(timeout=2)
                if message == 0x12:
                    return 0
                native = C.cast(pointer, C.POINTER(MSG)).contents
                native.message, native.wParam, native.lParam = message, wp, lp
                return 1
            def RegisterHotKey(self, hwnd, identifier, modifiers, key):
                self.owners.append(('register', threading.get_ident(), identifier))
                if key == ord('X'):
                    return 0
                self.hotkeys[identifier] = (modifiers, key)
                return 1
            def UnregisterHotKey(self, hwnd, identifier):
                self.owners.append(('unregister', threading.get_ident(), identifier))
                self.hotkeys.pop(identifier, None)
                return 1
        return API()

    def test_hotkey_conflicts_callback_and_cleanup_stay_on_owner(self):
        import threading
        from smartgrid.platform.windows.hotkeys import HotkeyManager
        api = self.transport()
        callback = threading.Event()
        actions = []
        def handle(action):
            actions.append(action)
            callback.set()
        manager = HotkeyManager(handle, api)
        try:
            conflicts = manager.register({'left':'Ctrl+Left', 'bad':'Ctrl+X', 'duplicate':'Ctrl+Left'})
            self.assertEqual(set(conflicts), {'bad', 'duplicate'})
            api.PostThreadMessageW(manager.thread_id, 0x312, 1, 0)
            self.assertTrue(callback.wait(1))
            self.assertEqual(actions, ['left'])
            self.assertFalse(manager.process_message(0x312, 999, 0))
        finally:
            manager.stop()
        manager.stop()
        self.assertFalse(api.hotkeys)
        self.assertFalse(manager.thread.is_alive())
        self.assertEqual({operation[1] for operation in api.owners}, {manager.thread_id})
        self.assertEqual(api.owners[-1][2], 0x123456789)

    def test_owner_callback_failure_is_propagated_without_killing_pump(self):
        from smartgrid.platform.windows.messages import MessageThread
        owner = MessageThread(self.transport())
        try:
            with self.assertRaisesRegex(ValueError, 'input'):
                owner.call(lambda: (_ for _ in ()).throw(ValueError('input')))
            self.assertEqual(owner.call(lambda: 42), 42)
        finally:
            owner.stop()

class DiscoveryTests(TestCase):
    def backend(self):
        from types import SimpleNamespace
        from unittest.mock import Mock
        backend=WindowsBackend.__new__(WindowsBackend)
        backend._generations={}
        backend._generation_counter=0
        backend._discovery_records={}
        from smartgrid.platform.windows.virtual_desktops import VirtualDesktops
        backend.virtual_desktops=VirtualDesktops()
        backend.desktop_id=None
        backend._monitor_ids={1:'display'}
        backend._monitor_areas={1:Rect(0,0,1920,1080)}
        backend.visible=set(range(1001,1011))
        backend.handles=list(range(1,1011))
        def enumerate_windows(callback,data):
            for hwnd in backend.handles:
                callback(hwnd,data)
            return True
        def pid(hwnd,pointer):
            pointer._obj.value=999999
            return 1
        def class_name(hwnd,buffer,length):
            buffer.value='NormalApp'
            return 9
        def title(hwnd,buffer,length):
            buffer.value='Window'
            return 6
        backend.api=SimpleNamespace(physical_coordinates=nullcontext,EnumWindows=enumerate_windows,
            GetWindowThreadProcessId=pid,IsWindowVisible=lambda h:h in backend.visible,
            GetClassNameW=class_name,GetWindowLong=lambda h,i:0x00cf0000 if i==-16 else 0,
            GetWindow=lambda h,c:0,GetWindowTextLengthW=lambda h:6,GetWindowTextW=title,
            MonitorFromWindow=lambda h,c:1,IsIconic=lambda h:False,IsZoomed=lambda h:False,
            check=lambda value,op:None)
        backend.catalogue=SimpleNamespace(process_identity=Mock(return_value=(ApplicationRef('app','App'),'creation')))
        backend._attribute=lambda *args:0
        backend.window_rect=lambda h:Rect(10,10,600,400)
        return backend

    def test_hidden_helpers_are_skipped_and_process_queried_once_per_scan(self):
        backend=self.backend()
        windows=backend.discover_windows()
        self.assertEqual(len(windows),10)
        self.assertTrue(all(w.eligible for w in windows))
        backend.catalogue.process_identity.assert_called_once_with(999999)
        backend.discover_windows()
        self.assertEqual(backend.catalogue.process_identity.call_count,2)

    def test_previously_visible_app_remains_tracked_when_hidden(self):
        backend=self.backend()
        original=backend.discover_windows()[0]
        backend.visible.remove(original.ref.hwnd)
        hidden=next(w for w in backend.discover_windows() if w.ref.hwnd==original.ref.hwnd)
        self.assertEqual(hidden.ref,original.ref)
        self.assertFalse(hidden.eligible)
        self.assertEqual(hidden.exclusion_reason,'Not visible')


class AccentTests(TestCase):
    def test_registry_accent_is_abgr(self):
        from smartgrid.platform.windows.backend import abgr_to_hex
        self.assertEqual(abgr_to_hex(0xffd77800),'#0078d7')


class StackAboveTests(TestCase):
    def backend(self, above):
        from types import SimpleNamespace
        backend = WindowsBackend.__new__(WindowsBackend)
        calls = []
        backend.api = SimpleNamespace(GetWindow=lambda hwnd, cmd: above,
                                      SetWindowPos=lambda *args: calls.append(args) or 1)
        return backend, calls

    def test_outline_already_above_its_window_is_not_raised_over_other_windows(self):
        backend, calls = self.backend(above=50)
        self.assertTrue(backend.stack_above(50, 10))
        self.assertEqual(calls, [])

    def test_outline_is_inserted_just_below_the_window_covering_its_target(self):
        backend, calls = self.backend(above=70)
        backend.stack_above(50, 10)
        self.assertEqual(calls[0][1], 70)


class ForcedResizeTests(TestCase):
    def test_fixed_size_window_gets_a_resizable_frame(self):
        from types import SimpleNamespace
        backend = WindowsBackend.__new__(WindowsBackend)
        styles, calls = {'value': 0x00C00000 | 0x01000000}, []
        backend.api = SimpleNamespace(GetWindowLong=lambda hwnd, index: styles['value'],
                                      SetWindowLong=lambda hwnd, index, value: styles.update(value=value),
                                      SetWindowPos=lambda *args: calls.append(args) or 1)
        self.assertTrue(backend.force_resizable(10))
        self.assertTrue(styles['value'] & 0x00040000)
        self.assertFalse(styles['value'] & 0x01000000)
        self.assertEqual(calls[0][-1] & 0x0020, 0x0020)
        self.assertFalse(backend.force_resizable(10))
