import ctypes as C
from types import SimpleNamespace
from unittest import TestCase
from smartgrid.core.models import Rect,WindowRef
from smartgrid.platform.windows.api import HANDLE,DWM_THUMBNAIL_PROPERTIES
from smartgrid.platform.windows.thumbnails import ThumbnailManager


class ThumbnailTests(TestCase):
    def make(self,fail_update=False):
        calls=[]
        def register(destination,source,pointer):
            calls.append(('register',destination,source))
            C.cast(pointer,C.POINTER(HANDLE))[0]=HANDLE(100+len(calls))
            return 0
        def update(handle,pointer):
            p=C.cast(pointer,C.POINTER(DWM_THUMBNAIL_PROPERTIES)).contents
            calls.append(('update',p.rcDestination.left,p.rcDestination.top,p.rcDestination.right,p.rcDestination.bottom,p.fVisible))
            return -1 if fail_update else 0
        return ThumbnailManager(SimpleNamespace(DwmRegisterThumbnail=register,DwmUpdateThumbnailProperties=update,
            DwmUnregisterThumbnail=lambda h:calls.append(('unregister',h)))),calls

    def test_reuse_unchanged_geometry_and_cleanup_removed_sources(self):
        manager,calls=self.make()
        targets={0:(WindowRef(5,6,'original'),Rect(20,30,180,90))}
        self.assertEqual(manager.sync(10,targets),{0})
        self.assertEqual(calls[1],('update',20,30,200,120,1))
        manager.sync(10,targets)
        self.assertEqual(len(calls),2)
        manager.sync(10,{})
        self.assertEqual(calls[-1][0],'unregister')
        self.assertFalse(manager.entries)

    def test_handle_reuse_and_destination_changes_unregister_before_registering(self):
        manager,calls=self.make()
        manager.sync(10,{0:(WindowRef(5,6,'original'),Rect(0,0,10,10))})
        manager.sync(10,{0:(WindowRef(5,6,'replacement'),Rect(0,0,10,10))})
        self.assertEqual([c[0] for c in calls],['register','update','unregister','register','update'])
        manager.sync(20,{0:(WindowRef(5,6,'replacement'),Rect(0,0,10,10))})
        self.assertEqual(calls[-2],('register',20,5))
        manager.clear()
        self.assertFalse(manager.entries)

    def test_failed_thumbnail_falls_back_and_releases_its_handle(self):
        manager,calls=self.make(fail_update=True)
        self.assertEqual(manager.sync(10,{0:(WindowRef(5,6,'source'),Rect(0,0,10,10))}),set())
        self.assertEqual(calls[-1][0],'unregister')
        self.assertFalse(manager.entries)
        self.assertEqual(C.sizeof(DWM_THUMBNAIL_PROPERTIES),48)
