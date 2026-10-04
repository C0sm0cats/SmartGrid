"""COM ABI and namespace regressions; these are not native desktop tests."""
import ctypes as C
from unittest import TestCase
from unittest.mock import Mock
from smartgrid.platform.windows.virtual_desktops import GUID,DesktopQuery,VirtualDesktops,scoped_display
from smartgrid.platform.windows.api import CALLBACK,LONG,BOOL,HWND

class DesktopTests(TestCase):
    def test_guid_abi_and_display_namespaces(self):
        text='d0e87b71-7370-44f8-8ab8-4f25a9a9a991'
        self.assertEqual(C.sizeof(GUID),16)
        self.assertEqual(GUID.parse(text).text(),text)
        self.assertIsNone(GUID().text())
        self.assertEqual(scoped_display('monitor:a','desk1'),'monitor:a@desktop:desk1')
        self.assertEqual(scoped_display('monitor:a@desktop:desk1','desk2'),'monitor:a@desktop:desk2')

    def test_documented_com_slots_and_apartment_release(self):
        text='d0e87b71-7370-44f8-8ab8-4f25a9a9a991';released=[]
        @CALLBACK(C.c_uint32,C.c_void_p)
        def release(this): released.append(True);return 0
        @CALLBACK(LONG,C.c_void_p,HWND,C.POINTER(BOOL))
        def current(this,hwnd,result): result.contents.value=1;return 0
        @CALLBACK(LONG,C.c_void_p,HWND,C.POINTER(GUID))
        def desktop(this,hwnd,result): C.memmove(result,C.byref(GUID.parse(text)),16);return 0
        table=(C.c_void_p*6)()
        for index,fn in ((2,release),(3,current),(4,desktop)): table[index]=C.cast(fn,C.c_void_p).value
        obj=C.pointer(C.cast(table,C.POINTER(C.c_void_p)));ptr=C.cast(obj,C.c_void_p)
        manager=VirtualDesktops();manager.available=True;manager.ole=Mock()
        manager.ole.CoInitializeEx.return_value=0
        def create(cls,outer,context,iid,result): result._obj.value=ptr.value;return 0
        manager.ole.CoCreateInstance.side_effect=create
        with manager.query() as query:
            self.assertTrue(query.on_current(100))
            self.assertEqual(query.desktop_id(100),text)
        self.assertEqual(released,[True]);manager.ole.CoUninitialize.assert_called_once()
        manager.ole.CoInitializeEx.return_value=-2147417850
        manager.ole.CoUninitialize.reset_mock()
        with manager.query() as query: self.assertTrue(query.on_current(100))
        manager.ole.CoUninitialize.assert_not_called()


class DesktopResolutionTests(TestCase):
    def manager(self,values,query=None):
        from contextlib import contextmanager
        manager=VirtualDesktops();manager.available=True
        manager._explorer_value=lambda name:values.get(name)
        @contextmanager
        def scoped(): yield query
        manager.query=scoped
        return manager

    def test_single_desktop_without_registry_is_identified(self):
        self.assertEqual(self.manager({}).resolve([1]),(None,True))
        guid='5f6c2a3e-0000-4000-8000-000000000001'
        only=GUID.parse(guid)
        self.assertEqual(self.manager({'VirtualDesktopIDs':bytes(only)}).resolve(),(guid,True))

    def test_registry_current_desktop_wins(self):
        guid='5f6c2a3e-0000-4000-8000-000000000002'
        self.assertEqual(self.manager({'CurrentVirtualDesktop':bytes(GUID.parse(guid))}).resolve(),(guid,True))

    def test_several_desktops_use_a_window_or_stay_unidentified(self):
        from unittest.mock import Mock
        ids=bytes(GUID.parse('5f6c2a3e-0000-4000-8000-000000000003'))*2
        query=Mock();query.on_current=lambda hwnd:hwnd==7;query.desktop_id=lambda hwnd:'desk'
        self.assertEqual(self.manager({'VirtualDesktopIDs':ids},query).resolve(lambda:[0,5,7]),('desk',True))
        self.assertEqual(self.manager({'VirtualDesktopIDs':ids},None).resolve([7]),(None,False))
