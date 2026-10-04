from dataclasses import asdict
import math
import unittest
from smartgrid.core.models import (ApplicationRef, Assignment, Display, Draft,
    LayoutTemplate, Rect, Settings, SpaceProfile, Tile, WindowRecord, WindowRef)
from smartgrid.core.reconcile import (compact_assignments, memberships,
                                     reconcile_assignments, reconcile_profile)


def window(hwnd, app='app', display='display', **kwargs):
    return WindowRecord(WindowRef(hwnd, hwnd+100),app,str(hwnd),Rect(0,0,100,100),display,**kwargs)


class ModelTests(unittest.TestCase):
    def test_every_contract_model_round_trips_nested_state(self):
        profile = SpaceProfile('display',2,'custom',tiles=[Tile('first',0,0,1,1)],assignments=[Assignment('app',1,True)])
        for value in (Rect(-10,-5,20,10), Tile('tile',0,0,1,1),
                      Display('display','Display',Rect(0,0,1920,1080)),
                      ApplicationRef('app','App'),WindowRef(1,100,'generation'),window(1),
                      Assignment('app',None,True),LayoutTemplate('template','Template'),
                      profile,Draft(profile,True),Settings()):
            self.assertEqual(type(value).from_dict(asdict(value)),value)
        self.assertTrue(Rect(0,0,10,10).contains(9,9))
        self.assertFalse(Rect(0,0,10,10).contains(10,10))

    def test_reference_defaults_motion_and_legacy_settings(self):
        defaults=Settings()
        self.assertEqual((defaults.gap,defaults.padding,defaults.master_ratio),(12,12,.6))
        self.assertEqual(Settings.from_dict({'gap':8,'padding':8}).gap,8)
        self.assertEqual(Settings(animation_speed='fast').visual_duration(),90)
        self.assertEqual(Settings(animation_speed='custom',animation_duration=300).visual_duration(),300)
        self.assertEqual(Settings(animations=False).visual_duration(),0)
        for data in ({'master_ratio':.1},{'animation_speed':'invalid'},{'border_color':'blue'}):
            with self.assertRaises(ValueError): Settings.from_dict(data)

    def test_mutable_defaults_are_independent(self):
        a,b = Settings(),Settings()
        a.margins['top'] = 99
        a.excluded_apps.append('app')
        self.assertEqual(b.margins['top'],12)
        self.assertEqual(b.excluded_apps,[])
        a,b = SpaceProfile('display',0),SpaceProfile('display',0)
        a.assignments.append(Assignment('app'))
        self.assertEqual(b.assignments,[])

    def test_invalid_persisted_shapes_types_and_values_rejected(self):
        for cls, data in ((Rect,dict(x=True,y=0,width=1,height=1)),
            (Tile,dict(id='a',x=0,y=0,width=math.nan,height=1)),
            (Assignment,dict(app_id='',window_id=1)),
            (Assignment,dict(app_id='a',window_id=True)),
            (Assignment,dict(app_id='a',pinned='yes')),
            (SpaceProfile,dict(display_id='d',space=3)),
            (SpaceProfile,dict(display_id='d',space=0,assignments={})),
            (SpaceProfile,dict(display_id='d',space=0,preset='nonsense')),
            (SpaceProfile,dict(display_id='d',space=0,preset='custom',tiles=[])),
            (LayoutTemplate,dict(id='t',name='Test',preset='1x1',assignments=[dict(app_id='a'),dict(app_id='b')])),
            (Settings,dict(gap=-1)), (Settings,dict(theme='invalid')),
            (Settings,dict(animations='yes')), (Settings,dict(tile_timeout=math.inf)),
            (Settings,dict(accent='blue')), (Settings,dict(margins={'top':1})),
            (Settings,dict(included_apps=[7])), (Settings,dict(new_unknown_setting=1))):
            with self.subTest(cls=cls,data=data), self.assertRaises(ValueError):
                cls.from_dict(data)
        for shape in ([],None,'hello'):
            with self.assertRaises(ValueError):
                Settings.from_dict(shape)


class ReconcileTests(unittest.TestCase):
    def test_pins_and_pending_destinations_stay_stable_under_compaction(self):
        original = [None,Assignment('closed',99,True),None,Assignment('app',1),Assignment('pending')]
        result = reconcile_assignments(original,[window(1)],compact=True)
        self.assertEqual(result,[Assignment('app',1),Assignment('closed',None,True),None,None,Assignment('pending')])
        self.assertEqual(original[1].window_id,99)
        no_compact = reconcile_assignments(original,[window(1)],compact=False)
        self.assertEqual(no_compact[3].window_id,1)
        self.assertIsNone(no_compact[0])
        self.assertEqual(compact_assignments(no_compact),result)

    def test_matching_duplicate_apps_protects_live_bindings_and_is_deterministic(self):
        original = [Assignment('app'),Assignment('app',20),Assignment('app',99,True)]
        result = reconcile_assignments(original,[window(30),window(20),window(10)])
        self.assertEqual([a.window_id for a in result],[10,20,30])
        self.assertEqual(result,reconcile_assignments(original,[window(10),window(30),window(20)]))
        duplicate = reconcile_assignments([Assignment('app',20),Assignment('app',20,True)],[window(20)])
        self.assertEqual(duplicate,[Assignment('app',20),Assignment('app',None,True)])

    def test_close_holes_and_pending_cancellation(self):
        original = [Assignment('app',1),Assignment('app',2),Assignment('pending')]
        result = reconcile_assignments(original,[window(2)],compact=False,preserve_pending=False)
        self.assertEqual(result,[None,Assignment('app',2),None])
        result = reconcile_assignments(original,[window(2)],compact=True)
        self.assertEqual(result,[Assignment('app',2),None,Assignment('pending')])

    def test_inclusion_filters_minimization_and_cross_display_memberships(self):
        live = [window(1,state='minimized'),window(2,floating=True),window(3,eligible=False),window(4,display='other')]
        profile = SpaceProfile('display',0,assignments=[None])
        result = reconcile_profile(profile,live,include_unassigned=True)
        self.assertEqual(result.assignments,[Assignment('app',1)])
        other = SpaceProfile('display',1,assignments=[Assignment('app',1,True)])
        self.assertEqual(memberships([result,other]),{1:{('display',0),('display',1)}})
        self.assertEqual(profile.assignments,[None])


class AppDisplayNameTests(unittest.TestCase):
    def test_uncatalogued_ids_get_readable_names(self):
        from smartgrid.core.models import app_display_name
        self.assertEqual(app_display_name(r'exe:c:\program files\git\usr\bin\mintty.exe'),'mintty')
        self.assertEqual(app_display_name('aumid:microsoft.windowsnotepad_8wekyb3d8bbwe!app'),'windowsnotepad')
        self.assertEqual(app_display_name('browser'),'browser')
