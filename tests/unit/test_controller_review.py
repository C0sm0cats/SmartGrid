"""Desktop behavior regressions discovered while integrating the pure core."""
from copy import deepcopy
from dataclasses import asdict
import json
from pathlib import Path
import tempfile
import unittest

from smartgrid.core.controller import Controller
from smartgrid.core.models import (ApplicationRef, Assignment, Display, LayoutTemplate,
    Rect, SpaceProfile, Tile, WindowRecord, WindowRef)
from smartgrid.platform.fake import FakeBackend
from smartgrid.storage.repository import Repository


def record(hwnd, app='app', display='d1', state='normal'):
    return WindowRecord(WindowRef(hwnd,hwnd,f'generation:{hwnd}'),app,str(hwnd),
                        Rect(30+hwnd,40+hwnd,640,480),display,state)


class ControllerReviewTests(unittest.TestCase):
    def test_pending_launch_survives_process_reconstruction_without_automatic_movement(self):
        controller,backend=self.pending_template('d2')
        self.assertTrue(controller._pending)
        saved=Path(self.temporary.name)/'session.json'
        self.assertNotIn('window_id',saved.read_text())
        resumed=Controller(backend,Repository(self.temporary.name))
        self.assertTrue(resumed._pending);self.assertFalse(resumed.running)
        arrived=backend.add_window('slow');original=arrived.rect
        resumed.refresh();self.assertEqual(backend.windows[arrived.ref.hwnd].rect,original)
        resumed.start()
        self.assertEqual(resumed.profile('d2',0).assignments[0].window_id,arrived.ref.hwnd)
        self.assertEqual(backend.windows[arrived.ref.hwnd].display_id,'d2')
        resumed.stop()
        self.assertEqual(Repository(self.temporary.name).load_session()['pending'],[])

    def test_slow_launch_destination_remains_after_warning_threshold(self):
        controller,backend=self.pending_template()
        for pending in controller._pending.values(): pending['created']-=121
        controller.refresh();self.assertTrue(controller._pending)
        arrived=backend.add_window('slow');controller.refresh()
        self.assertEqual(controller.profile('d1',0).assignments[0].window_id,arrived.ref.hwnd)

    def test_desktop_namespaces_preserve_groups_and_unknown_desktop_stops_movement(self):
        from dataclasses import replace
        controller,backend=self.make([record(1)])
        backend.desktop_id='first';backend.desktop_known=True
        backend.displays=[replace(d,id=d.id+'@desktop:first') for d in self.displays]
        backend.windows[1]=replace(backend.windows[1],display_id='d1@desktop:first')
        controller.refresh(auto=False);controller.start()
        first=controller.profile('d1@desktop:first',0)
        first.preset='custom';first.tiles=[Tile('all',0,0,1,1)]
        controller._save()
        backend.desktop_id='second'
        backend.displays=[replace(d,id=d.id+'@desktop:second') for d in self.displays]
        backend.windows[1]=replace(backend.windows[1],eligible=False)
        backend.windows[2]=replace(record(2,'other'),display_id='d1@desktop:second')
        controller.refresh()
        self.assertEqual(controller.profile('d1@desktop:first',0).preset,'custom')
        self.assertEqual(controller.profile('d1@desktop:second',0).assignments[0].window_id,2)
        self.assertFalse(controller.history.can_undo)
        old_window=controller.profile('d1@desktop:first',0).assignments[0].window_id
        controller.set_assignment('d1@desktop:first',1,0,Assignment('app',None,True))
        controller.arrange();controller.undo()
        self.assertEqual(controller.profile('d1@desktop:first',0).assignments[0].window_id,old_window)
        self.assertTrue(controller.draft('d1@desktop:first',1).dirty)
        operations=len(backend.operations);backend.desktop_known=False
        with self.assertRaises(ValueError): controller.arrange()
        self.assertEqual(len(backend.operations),operations)
        loaded=Repository(self.temporary.name).load_layouts()[1]
        self.assertIn(('d1@desktop:first',0),loaded)


    def test_explicit_window_actions_reject_recycled_hwnd_and_close_normally(self):
        from dataclasses import replace
        controller,backend=self.make([record(1),record(2)])
        controller.start();reference=backend.windows[1].ref
        backend.windows[1]=replace(backend.windows[1],ref=WindowRef(1,999,'replacement'))
        self.assertFalse(controller.window_action('close',reference))
        self.assertIn(1,backend.windows)
        reference=backend.windows[2].ref
        self.assertTrue(controller.window_action('maximize',reference))
        self.assertEqual(backend.windows[2].state,'maximized')
        self.assertTrue(controller.window_action('maximize',reference))
        self.assertEqual(backend.windows[2].state,'normal')
        self.assertTrue(controller.window_action('close',reference))
        self.assertNotIn(2,backend.windows)

    def test_always_floating_rule_restores_and_rule_removal_retiles(self):
        from dataclasses import replace
        controller,backend=self.make([record(1),record(2,'other')])
        original=backend.windows[1].rect
        controller.start()
        controller.update_settings(replace(controller.settings,excluded_apps=['app']))
        self.assertEqual(backend.windows[1].rect,original)
        self.assertTrue(controller._window(1).floating)
        self.assertFalse(controller._eligible(controller._window(1)))
        controller.update_settings(replace(controller.settings,excluded_apps=[]))
        controller.arrange()
        self.assertTrue(controller._eligible(controller._window(1)))
        self.assertNotEqual(backend.windows[1].rect,original)


    def test_hidden_system_lifecycle_does_not_retile_application_windows(self):
        from dataclasses import replace
        controller,backend=self.make([record(1)])
        controller.start()
        operations=len(backend.operations)
        backend.windows[99]=replace(record(99),eligible=False,exclusion_reason='Cloaked or on another Windows virtual desktop')
        controller.refresh()
        del backend.windows[99]
        controller.refresh()
        self.assertEqual(len(backend.operations),operations)

    def test_resize_preview_uses_targeted_rect_and_caches_client_minimums(self):
        from dataclasses import replace
        from unittest.mock import Mock
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.windows[1]=replace(backend.windows[1],rect=Rect(8,8,1000,1064))
        backend.discover_windows=Mock(side_effect=AssertionError('No desktop scan during preview'))
        backend.min_size=Mock(return_value=(80,80))
        for _ in range(5):
            controller._preview_native_drag(1)
        self.assertEqual(backend.min_size.call_count,2)
        self.assertTrue(controller.preview_rectangles)

    def test_classic_layout_grows_and_shrinks_and_toggle_restores(self):
        controller,backend=self.make([record(1)])
        original=deepcopy(backend.windows[1])
        controller.profile('d1',0).preset='5x5'
        controller.handle_action('toggle')
        self.assertEqual(len(controller.resolved_rects('d1',0)),1)
        self.assertGreater(backend.windows[1].rect.width,1800)
        for hwnd in range(2,7):
            backend.windows[hwnd]=record(hwnd)
        controller.refresh()
        self.assertEqual(len(controller.resolved_rects('d1',0)),6)
        for hwnd in range(2,7):
            del backend.windows[hwnd]
        controller.refresh()
        self.assertEqual(len(controller.resolved_rects('d1',0)),1)
        controller.handle_action('toggle')
        self.assertFalse(controller.running)
        self.assertEqual(backend.windows[1].rect,original.rect)

    def test_empty_desktop_activation_and_stop_hotkey_do_not_quit_application(self):
        from smartgrid.core.controller import DEFAULT_HOTKEYS
        controller,backend=self.make()
        controller.start()
        self.assertEqual(controller.last_result.placed,[])
        self.assertEqual(DEFAULT_HOTKEYS['stop'],'Ctrl+Alt+Q')
        self.assertEqual(DEFAULT_HOTKEYS['redo'],'Ctrl+Alt+Y')
        controller.handle_action('stop')
        self.assertFalse(controller.running)
        self.assertFalse(controller._closed)

    def test_saved_classic_layout_collapses_empty_unpinned_tiles(self):
        controller,backend=self.make([record(1)])
        controller.templates=[LayoutTemplate('t','Sparse','5x5',assignments=[None,None,Assignment('app',1)])]
        controller.restore_template('t','d1',0)
        self.assertEqual(len(controller.profile('d1',0).assignments),1)
        self.assertGreater(backend.windows[1].rect.width,1800)

    def setUp(self):
        self.temporary = tempfile.TemporaryDirectory()
        self.addCleanup(self.temporary.cleanup)
        self.displays = [Display('d1','Primary',Rect(0,0,1920,1080),primary=True),
                         Display('d2','Secondary',Rect(1920,0,1600,900))]

    def make(self, windows=(), apps=()):
        backend = FakeBackend(self.displays, windows, apps)
        controller = Controller(backend,Repository(self.temporary.name))
        return controller,backend

    def test_compaction_keeps_absent_pin_at_original_slot(self):
        controller,backend = self.make([record(1)])
        controller.profiles[('d1',0)] = SpaceProfile('d1',0,assignments=[None,None,None,Assignment('missing',None,True)])
        controller.start()
        self.assertEqual(controller.profile('d1',0).assignments[3],Assignment('missing',None,True))
        self.assertEqual(controller.profile('d1',0).assignments[0].window_id,1)

    def test_duplicate_app_reservation_cannot_steal_explicit_binding(self):
        controller,backend = self.make([record(1),record(2)])
        profile = SpaceProfile('d1',0,assignments=[Assignment('app',None,True),Assignment('app',1)])
        controller.profiles[('d1',0)] = profile
        controller.start()
        self.assertEqual(profile.assignments[1].window_id,1)
        self.assertEqual(profile.assignments[0].window_id,2)

    def pending_template(self, display='d1'):
        controller,backend = self.make(apps=[ApplicationRef('slow','Slow app')])
        backend.auto_launch = False
        controller.templates = [LayoutTemplate('template','Slow layout','1x1',assignments=[Assignment('slow',None,True)])]
        controller.restore_template('template',display,0)
        self.assertEqual(backend.launches,['slow'])
        return controller,backend

    def test_late_application_arrives_on_target_secondary_display(self):
        controller,backend = self.pending_template('d2')
        arrived = backend.add_window('slow')
        controller.refresh(auto=False)
        assignment = controller.profile('d2',0).assignments[0]
        self.assertEqual(assignment.window_id,arrived.ref.hwnd)
        self.assertEqual(backend.windows[arrived.ref.hwnd].display_id,'d2')
        self.assertEqual(backend.windows[arrived.ref.hwnd].rect,controller.resolved_rects('d2',0)[0])

    def test_pause_does_not_minimize_late_active_context_application(self):
        controller,backend = self.pending_template()
        controller.pause()
        arrived = backend.add_window('slow')
        original = deepcopy(arrived)
        controller.refresh(auto=False)
        self.assertEqual(backend.windows[arrived.ref.hwnd].state,'normal')
        self.assertEqual(backend.windows[arrived.ref.hwnd].rect,original.rect)
        controller.resume()
        self.assertEqual(backend.windows[arrived.ref.hwnd].rect,controller.resolved_rects('d1',0)[0])

    def test_late_application_after_clear_cannot_repopulate_context(self):
        controller,backend = self.pending_template()
        controller.clear_space('d1',0)
        controller.apply_drafts()
        arrived = backend.add_window('slow')
        before=deepcopy(arrived.rect)
        controller.refresh(auto=True)
        self.assertEqual(backend.windows[arrived.ref.hwnd].rect,before)
        self.assertTrue(all(a is None for a in controller.profile('d1',0).assignments))
        self.assertEqual(backend.windows[arrived.ref.hwnd].state,'normal')

    def test_failed_geometry_edit_preserves_previous_draft(self):
        controller,backend = self.make([record(1),record(2)])
        draft = controller.draft('d1',0)
        draft.profile.preset = '2x1'
        draft.profile.assignments = [Assignment('app',1),Assignment('app',2)]
        draft.dirty = True
        previous = deepcopy(draft)
        with self.assertRaises(ValueError):
            controller.set_draft_geometry('d1',0,[Tile('only',0,0,1,1)])
        self.assertEqual(controller.draft('d1',0),previous)

    def test_stop_restores_parked_and_originally_minimized_windows(self):
        originals = [record(1),record(2,state='minimized')]
        controller,backend = self.make(originals)
        controller.start()
        controller.switch_space('d1',1)
        self.assertEqual(backend.windows[1].state,'minimized')
        controller.pause()
        controller.stop()
        for original in originals:
            self.assertEqual(backend.windows[original.ref.hwnd],original)

    def test_malformed_archive_is_atomic_without_desktop_actions(self):
        controller,backend = self.make([record(1)])
        controller.templates = [LayoutTemplate('existing','Existing')]
        controller.profile('d1',0).assignments = [Assignment('app',1)]
        previous_templates = deepcopy(controller.templates)
        previous_profiles = deepcopy(controller.profiles)
        path = Path(self.temporary.name)/'invalid-archive.json'
        path.write_text(json.dumps({'format':'smartgrid-layouts','version':1,
            'templates':[asdict(LayoutTemplate('valid','Valid')),{'id':'bad','name':'Bad','preset':'custom','tiles':[]}]}))
        with self.assertRaises(ValueError):
            controller.import_archive(path)
        self.assertEqual(controller.templates,previous_templates)
        self.assertEqual(controller.profiles,previous_profiles)
        self.assertEqual(backend.operations,[])
        self.assertEqual(backend.launches,[])

    def test_wrong_layout_container_recovers_without_crashing_startup(self):
        path = Path(self.temporary.name)/'layouts.json'
        path.write_text(json.dumps({'version':1,'templates':None,'profiles':[]}))
        repository = Repository(self.temporary.name)
        templates,profiles = repository.load_layouts()
        self.assertEqual(templates,[])
        self.assertEqual(profiles,{})
        self.assertTrue(repository.errors)


    def test_undo_after_close_does_not_reinject_dead_window_references(self):
        controller,backend = self.make([record(1),record(2)])
        controller.start()
        controller.arrange('d1')
        backend.remove_window(1)
        controller.refresh(auto=False)
        controller.undo()
        refs = {a.window_id for profile in controller.profiles.values()
                for a in profile.assignments if a and a.window_id}
        self.assertNotIn(1,refs)
        self.assertIn(2,refs)

    def test_stop_retry_restores_windows_after_transient_failure(self):
        original = record(1)
        controller,backend = self.make([original])
        controller.start()
        backend.failures.add(1)
        self.assertFalse(controller.stop())
        self.assertNotEqual(backend.windows[1].rect,original.rect)
        backend.failures.clear()
        self.assertTrue(controller.stop())
        self.assertEqual(backend.windows[1].rect,original.rect)

    def test_import_strips_runtime_handles_from_foreign_archive(self):
        controller,backend = self.make([record(1)])
        foreign = LayoutTemplate('foreign','Foreign','1x1',assignments=[Assignment('app',1)])
        path = Path(self.temporary.name)/'foreign.json'
        path.write_text(json.dumps({'format':'smartgrid-layouts','version':1,'templates':[asdict(foreign)]}))
        controller.import_archive(path)
        self.assertIsNone(controller.templates[-1].assignments[0].window_id)
        self.assertEqual(backend.operations,[])

    def test_explicit_and_custom_overflow_survive_storage_without_losing_geometry(self):
        from smartgrid.core.models import Tile
        for preset,tiles in [('2x2',[]),('custom',[Tile('one',0,0,1,1)])]:
            with self.subTest(preset=preset):
                controller,backend=self.make([record(i,app=f'app{i}') for i in range(1,31)])
                controller.profiles[('d1',0)]=SpaceProfile('d1',0,preset=preset,tiles=tiles)
                controller.start()
                profile=controller.profile('d1',0)
                self.assertEqual(len(profile.assignments),30)
                self.assertEqual(profile.preset,preset)
                self.assertEqual(profile.tiles,tiles)
                self.assertEqual(len(controller.last_result.placed),30)
                reloaded=Controller(backend,Repository(self.temporary.name))
                self.assertEqual(reloaded.profile('d1',0).tiles,tiles)
                self.assertEqual(len(reloaded.profile('d1',0).assignments),30)
                self.assertFalse(reloaded.repository.errors)

    def test_persisted_invalid_settings_retained_before_next_save(self):
        from smartgrid.core.models import Settings
        path=Path(self.temporary.name)/'settings.json'
        original='{"version":1,"settings":{"gap":-7}}'
        path.write_text(original)
        repository=Repository(self.temporary.name)
        repository.load_settings()
        repository.save_settings(Settings())
        backups=list(path.parent.glob('settings.json.*.invalid'))
        self.assertEqual(len(backups),1)
        self.assertEqual(backups[0].read_text(),original)

    def test_applying_named_context_does_not_commit_other_dirty_context(self):
        controller,backend=self.make([record(1)])
        controller.set_assignment('d1',0,0,Assignment('app',1))
        controller.set_assignment('d1',1,0,Assignment('later',None,True))
        controller.apply_drafts(contexts=[('d1',0)])
        self.assertTrue(controller.draft('d1',1).dirty)
        self.assertFalse(controller.profile('d1',1).assignments)

    def test_floating_pin_keeps_reference_and_restores_without_launching_duplicate(self):
        controller,backend=self.make([record(1)])
        controller.profiles[('d1',0)]=SpaceProfile('d1',0,assignments=[Assignment('app',1,True)])
        controller.start()
        controller.toggle_float(1)
        self.assertEqual(controller.profile('d1',0).assignments[0].window_id,1)
        self.assertIn(1,controller._floating)
        controller.refresh(auto=False)
        self.assertTrue(controller.windows[0].floating)
        self.assertEqual(controller.reserved_slots()[0]['state'],'FLOATING')
        controller.activate_reserved('d1',0,0)
        self.assertNotIn(1,controller._floating)
        self.assertFalse(backend.launches)
        self.assertEqual(backend.windows[1].rect,controller.resolved_rects('d1',0)[0])

    def test_cancelled_native_interaction_clears_preview_without_changing_membership(self):
        controller,backend=self.make([record(1)])
        controller.start()
        original=deepcopy(controller.profile('d1',0))
        controller._handle_event({'type':'move_start','hwnd':1})
        controller.preview_rectangles=[('d1',Rect(10,10,50,50))]
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(controller.preview_rectangles,[])
        self.assertEqual(controller.profile('d1',0),original)

    def test_read_only_diagnostic_does_not_backup_invalid_config(self):
        path=Path(self.temporary.name)/'settings.json'
        path.write_text('{"settings":{"gap":-2}}')
        repository=Repository(self.temporary.name,read_only=True)
        repository.load_settings()
        self.assertFalse(list(path.parent.glob('*.invalid')))
        self.assertEqual(path.read_text(),'{"settings":{"gap":-2}}')

    def test_undo_clear_restores_delayed_launch_destination(self):
        controller,backend=self.pending_template()
        controller.clear_space('d1',0)
        controller.apply_drafts()
        controller.undo()
        arrived=backend.add_window('slow')
        controller.refresh(auto=True)
        self.assertEqual(controller.profile('d1',0).assignments[0].window_id,arrived.ref.hwnd)
        self.assertNotIn(arrived.ref.hwnd,controller._floating)

    def test_template_preserves_native_resized_geometry(self):
        controller,backend=self.make([record(1)])
        profile=SpaceProfile('d1',0,preset='2x1',assignments=[Assignment('app',1)],
            resize_tiles=[Tile('left',0,0,.7,1),Tile('right',.7,0,.3,1)])
        template=controller.save_profile_template('Resized',profile)
        self.assertEqual(template.preset,'custom')
        self.assertEqual(template.tiles,profile.resize_tiles)
        self.assertIsNone(template.assignments[0].window_id)

    def test_activation_tiles_maximized_windows_and_other_open_apps(self):
        controller,backend=self.make([record(1,state='maximized'),record(2,app='other')])
        original=deepcopy(backend.windows)
        controller.start()
        self.assertCountEqual(controller.last_result.placed,[1,2])
        self.assertEqual(backend.windows[1].state,'normal')
        self.assertEqual(len(controller.profile('d1',0).assignments),2)
        controller.stop()
        self.assertEqual(backend.windows[1].state,'maximized')
        self.assertEqual(backend.windows[1].rect,original[1].rect)

    def test_explicit_arrange_restores_maximized_app_but_automatic_reflow_preserves_it(self):
        from dataclasses import replace
        controller,backend=self.make([record(1),record(2,app='other')])
        controller.start()
        backend.windows[1]=replace(backend.windows[1],state='maximized')
        operations=len(backend.operations)
        controller.refresh(auto=True)
        self.assertEqual(backend.windows[1].state,'maximized')
        self.assertFalse(any(o[0]=='place' for o in backend.operations[operations:]))
        controller.arrange('d1')
        self.assertCountEqual(controller.last_result.placed,[1,2])
        self.assertEqual(backend.windows[1].state,'normal')

    def test_unchanged_periodic_discovery_does_not_rebuild_ui(self):
        controller,backend=self.make([record(1)])
        notifications=[]
        controller.subscribe(lambda:notifications.append(True))
        for _ in range(20):
            controller.refresh(auto=False)
        self.assertFalse(notifications)
        backend.windows[1].title='Changed title'
        controller.refresh(auto=False)
        self.assertEqual(len(notifications),1)

    def test_applying_unchanged_empty_draft_does_not_pretend_tiling_is_active(self):
        controller,backend=self.make([record(1)])
        controller.draft('d1',0)
        controller.apply_drafts()
        self.assertFalse(controller.running)
        self.assertFalse(backend.operations)
        self.assertIn('Arrange windows',controller.status)


class RestartBindingTests(unittest.TestCase):
    def test_restart_forgets_closed_regular_apps_but_keeps_pins_and_rebinds_open_windows(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','Display',Rect(0,0,1920,1080),primary=True)
            apps=[ApplicationRef(name,name) for name in ('open','closed','pinned')]
            repository=Repository(directory)
            repository.save_layouts([],{('d1',0):SpaceProfile('d1',0,assignments=[
                Assignment('closed'),Assignment('open'),Assignment('pinned',None,True)])})
            backend=FakeBackend([display],[record(1,'open')],apps)
            controller=Controller(backend,Repository(directory))
            assignments=controller.profile('d1',0).assignments
            self.assertEqual([(a.app_id,a.window_id,a.pinned) for a in assignments if a],
                             [('open',1,False),('pinned',None,True)])


class CollapsedResizeTests(unittest.TestCase):
    def test_collapsed_linked_resize_is_dropped_for_the_normal_layout(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','Display',Rect(0,0,2057,1059),primary=True)
            backend=FakeBackend([display],[record(1),record(2),record(3)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            profile=controller.profile('d1',0)
            # Logged on Windows: 147 px column, then a 3 px high tile.
            profile.resize_tiles=[Tile('0',0,0,.0773,1),Tile('1',.0773,0,.9227,.0102),Tile('2',.0773,.0102,.9227,.9898)]
            controller._reflow_all()
            self.assertEqual(profile.resize_tiles,[])
            heights=sorted(backend.windows[h].rect.height for h in (1,2,3))
            self.assertGreater(heights[0],400)


class GuideStateTests(unittest.TestCase):
    def test_swap_hints_and_space_event_follow_the_controller(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','Display',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            self.assertIsNone(controller.swap_hints())
            backend.focused=1
            controller._swap_hwnd=1;controller._swap_snapshot=controller._snapshot()
            display_id,rect,arrows=controller.swap_hints()
            self.assertEqual((display_id,[(d,t) for d,t,x,y,_ in arrows]),('d1',[('right',1)]))
            controller._swap_snapshot=None
            controller.switch_space('d1',1)
            self.assertEqual(controller.space_event[:2],('d1',1))
            self.assertEqual(controller.status,'Space 2 · Display 1')


class TrayFocusTests(unittest.TestCase):
    def test_tray_menu_keeps_the_focused_window_outline(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','Display',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            backend.foreground_hwnd=1;controller._update_border()
            self.assertEqual(controller._last_border,1)
            backend.foreground_hwnd=999;backend.is_shell_surface=lambda hwnd:hwnd==999
            controller._update_border()
            self.assertEqual(controller._last_border,1)


class PinnedStateTests(unittest.TestCase):
    def test_pinned_tile_states(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','Display',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1,'editor'),record(2,'other')],[ApplicationRef('closed','Closed')])
            controller=Controller(backend,Repository(directory));controller.start()
            profile=controller.profile('d1',0)
            profile.assignments=[Assignment('editor',1,True),Assignment('other',2),Assignment('closed',None,True)]
            backend.windows[1].state='minimized';controller.refresh(auto=False)
            states={s['app_id']:s['state'] for s in controller.reserved_slots()}
            self.assertEqual(states,{'editor':'MINIMIZED','closed':'CLOSED'})
            # Another window of the same app tiled in this space: the pin is still CLOSED.
            profile.assignments[2]=Assignment('other',None,True)
            self.assertEqual({s['app_id']:s['state'] for s in controller.reserved_slots()}['other'],'CLOSED')
            # The same app in another space of the display: OPEN ELSEWHERE.
            profile.assignments[1]=None
            controller.profile('d1',1).assignments=[Assignment('other',2)]
            self.assertEqual({s['app_id']:s['state'] for s in controller.reserved_slots()}['other'],'OPEN ELSEWHERE')


class HotkeyMigrationTests(unittest.TestCase):
    def test_retired_focus_shortcuts_are_replaced_but_custom_ones_kept(self):
        from smartgrid.core.models import Settings
        with tempfile.TemporaryDirectory() as directory:
            repository=Repository(directory)
            repository.save_settings(Settings(hotkeys={'focus_left':'Ctrl+Meta+Left','focus_up':'Ctrl+Shift+Up'}))
            controller=Controller(FakeBackend([Display('d1','D',Rect(0,0,1920,1080),primary=True)],[],[]),Repository(directory))
            self.assertEqual(controller.settings.hotkeys['focus_left'],'Ctrl+Alt+Left')
            self.assertEqual(controller.settings.hotkeys['focus_up'],'Ctrl+Shift+Up')
            self.assertEqual(Repository(directory).load_settings().hotkeys['focus_left'],'Ctrl+Alt+Left')


class ForceResizeSettingTests(unittest.TestCase):
    def test_minimum_size_is_ignored_only_when_forcing(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(i) for i in range(1,5)],[])
            backend.minimums[1]=(1500,900)
            controller=Controller(backend,Repository(directory))
            self.assertTrue(controller.settings.force_resize)
            controller.start()
            self.assertIn(1,controller.last_result.placed)
            controller.stop()
            controller.settings.force_resize=False
            controller.start()
            self.assertIn(1,controller.last_result.failures)


class OverflowCompactionTests(unittest.TestCase):
    def test_overflowing_custom_layout_compacts_like_auto(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(i) for i in range(1,8)],[])
            controller=Controller(backend,Repository(directory))
            profile=controller.profile('d1',0)
            profile.preset='custom';profile.tiles=[Tile('a',0,0,1,1)]
            controller.start()
            backend.windows[5].state='minimized';controller.refresh()
            occupants=[a.window_id if a else None for a in controller.profile('d1',0).assignments]
            self.assertEqual(occupants,[1,2,3,4,6,7])

    def test_fitting_custom_layout_keeps_positions(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(i) for i in range(1,4)],[])
            controller=Controller(backend,Repository(directory))
            profile=controller.profile('d1',0)
            profile.preset='custom';profile.tiles=[Tile('a',0,0,.5,1),Tile('b',.5,0,.5,.5),Tile('c',.5,.5,.5,.5)]
            controller.start()
            before=backend.windows[3].rect
            backend.windows[2].state='minimized';controller.refresh()
            self.assertEqual(backend.windows[3].rect,before)


class SwapArrowTests(unittest.TestCase):
    def arrows(self, count, focused, preset='auto', tiles=()):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(i) for i in range(1,count+1)],[])
            controller=Controller(backend,Repository(directory))
            profile=controller.profile('d1',0);profile.preset=preset;profile.tiles=list(tiles)
            controller.start()
            controller._swap_hwnd=focused;controller._swap_snapshot=controller._snapshot()
            return sorted((d,t) for d,t,x,y,_ in controller.swap_hints()[2])

    def test_focus_layout_large_tile_shows_both_right_neighbours(self):
        self.assertEqual(self.arrows(3,1),[('right',1),('right',2)])

    def test_grid_middle_tile_shows_all_four_sides(self):
        self.assertEqual(self.arrows(9,5),[('down',7),('left',3),('right',5),('up',1)])

    def test_custom_layout_with_t_junction(self):
        tiles=[Tile('a',0,0,.5,.5),Tile('b',.5,0,.5,1),Tile('c',0,.5,.25,.5),Tile('d',.25,.5,.25,.5)]
        self.assertEqual(self.arrows(4,2,'custom',tiles),[('left',0),('left',3)])
        self.assertEqual(self.arrows(4,3,'custom',tiles),[('right',3),('up',0)])

    def swap(self, count, focused, direction):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(i) for i in range(1,count+1)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            backend.foreground_hwnd=focused;controller.begin_swap()
            controller.swap_direction(direction)
            return [a.window_id if a else None for a in controller.profile('d1',0).assignments]

    def test_arrow_keys_never_swap_beyond_a_neighbour(self):
        # Master with two stacked windows: nothing above or below the master.
        self.assertEqual(self.swap(3,1,'up'),[1,2,3])
        self.assertEqual(self.swap(3,1,'down'),[1,2,3])
        # 5 windows in 3x2: the bottom-middle window has no window on its right.
        self.assertEqual(self.swap(5,5,'right'),[1,2,3,4,5])
        self.assertEqual(self.swap(5,5,'left'),[1,2,3,5,4])
        self.assertEqual(self.swap(5,5,'up'),[1,5,3,4,2])

    def test_pinned_windows_are_swap_neighbours(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(i) for i in range(1,5)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            for assignment in controller.profile('d1',0).assignments: assignment.pinned=True
            backend.foreground_hwnd=1;controller.begin_swap()
            self.assertEqual(sorted((d,t) for d,t,x,y,_ in controller.swap_hints()[2]),[('down',2),('right',1)])
            controller.swap_direction('right')
            self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[2,1,3,4])

    def test_every_arrow_key_target_has_an_arrow(self):
        from smartgrid.core.geometry import auto_preset, capacity, resolve_layout, edge_neighbors, swap_neighbor
        for count in range(2,10):
            preset=auto_preset(count)
            rects=resolve_layout(Rect(0,0,1920,1040),capacity(preset),preset=preset)
            for index in range(count):
                for direction in ('left','right','up','down'):
                    target=swap_neighbor(rects,index,direction,set(range(count)))
                    shown=[i for i,_ in edge_neighbors(rects,index,direction,set(range(count)))]
                    self.assertEqual(target is None,not shown)
                    if target is not None: self.assertIn(target,shown)


class OverlayApplicationTests(unittest.TestCase):
    def test_added_overlay_applications_float_while_the_overlay_list_is_on(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            window=controller._window(2)
            controller.settings.overlay_apps_added=[window.app_id]
            self.assertTrue(controller._ruled_out(window))
            controller.settings.builtin_exclusions=False
            self.assertFalse(controller._ruled_out(window))


class SeveralPinsTests(unittest.TestCase):
    def setup(self, directory, count):
        display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
        backend=FakeBackend([display],[record(i,'chrome') for i in range(1,count+1)]+[record(9,'editor')],[ApplicationRef('chrome','Chrome')])
        controller=Controller(backend,Repository(directory));controller.start()
        return backend,controller

    def test_several_windows_of_one_application_keep_their_own_pinned_tile(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.setup(directory,3)
            profile=controller.profile('d1',0)
            for i in (0,2): profile.assignments[i].pinned=True
            pins=[(i,a.window_id) for i,a in enumerate(profile.assignments) if a and a.pinned]
            controller.refresh()
            self.assertEqual([(i,a.window_id) for i,a in enumerate(controller.profile('d1',0).assignments) if a and a.pinned],pins)
            self.assertEqual(len({w for _,w in pins}),2)

    def test_a_closed_pinned_window_leaves_a_closed_card_not_its_sibling(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.setup(directory,2)
            profile=controller.profile('d1',0)
            for a in profile.assignments[:2]: a.pinned=True
            closed=profile.assignments[1].window_id
            del backend.windows[closed];controller.refresh()
            assignments=controller.profile('d1',0).assignments
            self.assertIsNotNone(assignments[0].window_id)
            self.assertIsNone(assignments[1].window_id)
            self.assertEqual([(s['index'],s['state']) for s in controller.reserved_slots()],[(1,'CLOSED')])
            # Its card opens a new window for that tile; the sibling stays in place.
            sibling=assignments[0].window_id
            controller.activate_reserved('d1',0,1)
            assignments=controller.profile('d1',0).assignments
            self.assertEqual(backend.launches,['chrome'])
            backend.add_window('chrome',hwnd=20);controller.refresh()
            assignments=controller.profile('d1',0).assignments
            self.assertEqual(assignments[0].window_id,sibling)
            self.assertNotIn(assignments[1].window_id,(None,sibling))


class ForcedResizeTests(unittest.TestCase):
    def test_tiles_forced_below_their_minimum_can_still_be_resized(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(i) for i in range(1,5)],[])
            for i in range(1,5): backend.minimums[i]=(1200,700)
            controller=Controller(backend,Repository(directory));controller.start()
            before=backend.windows[1].rect
            after=Rect(before.x,before.y,before.width+150,before.height)
            controller._native_resize(1,before,after,None)
            tiles=controller.profile('d1',0).resize_tiles
            self.assertTrue(tiles)
            self.assertGreater(tiles[0].width,.5)


class SmoothDragTests(unittest.TestCase):
    def make(self, directory):
        display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
        backend=FakeBackend([display],[record(1),record(2,'other')],[])
        controller=Controller(backend,Repository(directory));controller.start()
        return controller,backend

    def test_drag_start_does_not_rescan_the_desktop(self):
        from unittest.mock import Mock
        with tempfile.TemporaryDirectory() as directory:
            controller,backend=self.make(directory)
            backend.discover_windows=Mock(side_effect=AssertionError('No desktop scan at drag start'))
            controller._handle_event({'type':'move_start','hwnd':1})
            self.assertIn(1,controller._interacting)

    def test_a_window_growing_during_a_move_is_still_a_move(self):
        from dataclasses import replace
        with tempfile.TemporaryDirectory() as directory:
            controller,backend=self.make(directory)
            backend.gesture_kind=lambda hwnd:'move'
            guides=[];changes=[]
            controller.subscribe_guides(lambda:guides.append(1));controller.subscribe(lambda:changes.append(1))
            controller._handle_event({'type':'move_start','hwnd':1})
            # The application enforces its minimum size as the move starts.
            target=backend.windows[2].rect
            backend.windows[1]=replace(backend.windows[1],rect=Rect(target.x+20,target.y+10,target.width+300,target.height))
            controller._preview_native_drag(1)
            self.assertFalse(controller.preview_rectangles)
            self.assertTrue(controller.drag_guide[2].startswith('Swap'))
            self.assertTrue(guides);self.assertFalse(changes)
            controller._handle_event({'type':'move_end','hwnd':1})
            self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[2,1])
            self.assertFalse(controller.profile('d1',0).resize_tiles)

    def test_location_events_are_coalesced_per_window(self):
        with tempfile.TemporaryDirectory() as directory:
            controller,backend=self.make(directory)
            while not controller._queue.empty(): controller._queue.get_nowait()
            for i in range(50): controller._enqueue_event({'type':'location','hwnd':1,'n':i})
            controller._enqueue_event({'type':'location','hwnd':2,'n':0})
            self.assertEqual(controller._queue.qsize(),2)
            self.assertEqual(controller._pending_locations[1]['n'],49)


class FloatCenteringTests(unittest.TestCase):
    def test_a_window_made_floating_keeps_its_size_and_is_centred(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2)],[])
            size=backend.windows[2].rect
            controller=Controller(backend,Repository(directory));controller.start()
            controller.toggle_float(2)
            rect=backend.windows[2].rect
            self.assertEqual((rect.width,rect.height),(size.width,size.height))
            self.assertEqual((rect.x+rect.width//2,rect.y+rect.height//2),(960,540))


class FloatingMoveTests(unittest.TestCase):
    def test_moving_a_floating_window_is_not_a_drop_and_it_returns_to_its_tile(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(i) for i in range(1,7)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            tile=backend.windows[2].rect
            controller.toggle_float(2)
            controller._handle_event({'type':'move_start','hwnd':2})
            self.assertNotIn(2,controller._interacting)
            backend.windows[2]=backend.windows[2].__class__(**{**backend.windows[2].__dict__,'rect':Rect(1400,700,500,300)})
            controller._handle_event({'type':'location','hwnd':2})
            self.assertIsNone(controller.drag_guide)
            controller._handle_event({'type':'move_end','hwnd':2})
            self.assertEqual([a.window_id if a else None for a in controller.profile('d1',0).assignments],[1,None,3,4,5,6])
            controller.toggle_float(2)
            self.assertEqual(backend.windows[2].rect,tile)
