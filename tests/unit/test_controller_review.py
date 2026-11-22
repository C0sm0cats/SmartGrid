"""Desktop behavior regressions discovered while integrating the pure core."""
from copy import deepcopy
from dataclasses import asdict, replace
import json
from pathlib import Path
import tempfile
import time
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

    def test_slow_launch_warning_names_the_application_and_clears_when_it_arrives(self):
        controller,backend=self.pending_template('d2')
        controller.refresh();self.assertEqual(controller.last_error,'')
        for pending in controller._pending.values(): pending['created']-=121
        controller.refresh()
        self.assertIn('Slow app (Display 2 · Space 1 · tile 1)',controller.last_error)
        controller._last_error_at-=controller.ERROR_LIFETIME+1
        controller.refresh();self.assertIn('Slow app',controller.last_error)
        backend.add_window('slow');controller.refresh()
        self.assertEqual(controller.last_error,'')

    def test_slow_launch_warning_clears_when_its_tile_is_released(self):
        controller,backend=self.pending_template()
        for pending in controller._pending.values(): pending['created']-=121
        controller.refresh();self.assertIn('Slow app',controller.last_error)
        controller._cancel_pending()
        self.assertEqual(controller.last_error,'')

    def test_window_minimized_by_the_user_stays_minimized_after_stop_and_restart(self):
        controller,backend=self.make([record(1),record(2),record(3,'deskflow')])
        controller.start()
        backend.minimize(3);controller.refresh()
        self.assertNotIn(3,[a.window_id for a in controller.profile('d1',0).assignments if a])
        controller.stop()
        self.assertEqual(backend.windows[3].state,'minimized')
        restarted=Controller(backend,Repository(self.temporary.name));restarted.start()
        self.assertNotIn(3,[a.window_id for a in restarted.profile('d1',0).assignments if a])
        self.assertEqual(backend.windows[3].state,'minimized')

    def test_window_hidden_to_the_notification_area_stays_hidden_after_stop(self):
        controller,backend=self.make([record(1),record(2),record(3,'deskflow')])
        controller.start()
        backend.windows[3]=replace(backend.windows[3],eligible=False,exclusion_reason='Not visible');controller.refresh()
        self.assertNotIn(3,[a.window_id for a in controller.profile('d1',0).assignments if a])
        controller.stop()
        self.assertEqual(backend.windows[3].state,'hidden')

    def test_always_floating_keeps_a_minimized_window_minimized(self):
        controller,backend=self.make([record(1),record(2),record(3,'deskflow')])
        controller.start()
        backend.minimize(3);controller.refresh()
        settings=deepcopy(controller.settings);settings.excluded_apps=['deskflow']
        controller.update_settings(settings)
        self.assertEqual(backend.windows[3].state,'minimized')

    def test_one_off_error_expires(self):
        controller,backend=self.make([record(1)])
        controller.last_error='Windows refused the window action.'
        controller.refresh();self.assertTrue(controller.last_error)
        controller._last_error_at-=controller.ERROR_LIFETIME+1
        controller.refresh();self.assertEqual(controller.last_error,'')

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
        backend.min_size=Mock(return_value=(80,80))
        controller._handle_event({'type':'move_start','hwnd':1})
        # Asked once per window when the gesture starts, never on its beats.
        self.assertEqual(backend.min_size.call_count,2)
        backend.windows[1]=replace(backend.windows[1],rect=Rect(8,8,1000,1064))
        backend.discover_windows=Mock(side_effect=AssertionError('No desktop scan during preview'))
        for _ in range(5):
            controller._preview_native_drag(1)
        self.assertEqual(backend.min_size.call_count,2)
        self.assertTrue(controller.preview_rectangles)

    def resize_gesture(self, width_change):
        """Two side-by-side windows; window 1's right edge dragged by width_change."""
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        tiles={h:backend.windows[h].rect for h in (1,2)}
        controller._handle_event({'type':'move_start','hwnd':1})
        held=tiles[1]
        backend.windows[1]=replace(backend.windows[1],rect=Rect(held.x,held.y,held.width+width_change,held.height))
        controller._preview_native_drag(1)
        return controller,backend,tiles

    def live_requests(self, backend):
        return [op for op in backend.operations if op[0]=='place_live']

    def test_linked_resize_shows_neighbours_and_asks_nothing_of_them(self):
        controller,backend,tiles=self.resize_gesture(200)
        # Window 2 shrinks: shown on its future tile, nothing asked of it
        # before the release; the new gap, over its old size, is covered.
        self.assertEqual([p[2] for p in controller.preview_rectangles],[None,2])
        self.assertLess(controller.preview_rectangles[1][1].width,tiles[2].width)
        self.assertEqual(self.live_requests(backend),[])
        held=backend.windows[1].rect
        self.assertEqual(controller.preview_covers,[('d1',Rect(held.right,tiles[2].y,controller.settings.gap,tiles[2].height))])
        controller,backend,tiles=self.resize_gesture(-200)
        # Window 2 grows: shown larger, resized at the release; the held
        # window is clipped to its future tile: the new gap is covered.
        self.assertGreater(controller.preview_rectangles[1][1].width,tiles[2].width)
        self.assertEqual(self.live_requests(backend),[])
        held=controller.preview_rectangles[0][1]
        self.assertEqual(controller.preview_covers,[('d1',Rect(held.right,tiles[1].y,controller.settings.gap,tiles[1].height))])
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(controller.preview_covers,[])
        self.assertEqual(backend.windows[2].rect.x,tiles[2].x-200)

    def drag_gesture(self, controller, backend, hwnd, target, kind='move', driven=False):
        """Window hwnd dragged by its title bar until the pointer is over the
        tile target; the window itself has moved less (its centre may still
        be in its own tile). Returns the drop target shown."""
        backend.gesture_kind=lambda h:kind
        rect=backend.windows[hwnd].rect
        backend.cursor=(rect.right-40,rect.y+10)
        controller._handle_event({'type':'move_start','hwnd':hwnd})
        dx,dy=target.x+target.width//2-backend.cursor[0],target.y+20-backend.cursor[1]
        backend.cursor=(backend.cursor[0]+dx,backend.cursor[1]+dy)
        backend.windows[hwnd]=replace(backend.windows[hwnd],rect=Rect(rect.x+dx,rect.y+dy,rect.width,rect.height))
        if driven:
            self.assertTrue(controller.drive_preview())
        else:
            controller._handle_event({'type':'location','hwnd':hwnd})
        guide=controller.drag_guide
        controller._handle_event({'type':'move_end','hwnd':hwnd})
        return guide

    def test_drop_target_is_under_the_pointer(self):
        for driven in (False,True):
            # A new run, not the same windows found again after a restart.
            (Path(self.temporary.name)/'session.json').unlink(missing_ok=True)
            controller,backend=self.make([record(1),record(2)])
            controller.preview_driven=driven
            controller.start()
            # Window 1 is made wide, window 2 narrow.
            backend.windows[1]=replace(backend.windows[1],rect=replace(backend.windows[1].rect,width=1500))
            controller._native_resize(1,controller._tile_rects[1][1],backend.windows[1].rect,None,kind='resize')
            wide,narrow=controller._tile_rects[1][1],controller._tile_rects[2][1]
            self.assertGreater(wide.width,narrow.width*2)
            # Dragged by its title bar onto the narrow tile: its centre is
            # still in the wide one, the pointer decides.
            guide=self.drag_gesture(controller,backend,1,narrow,driven=driven)
            self.assertEqual(guide[1],narrow)
            self.assertTrue(guide[2].startswith('Swap'))
            self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[2,1])
            self.assertEqual(backend.windows[2].rect,wide)

    def test_gestures_follow_one_another(self):
        for driven in (False,True):
            # A new run, not the same windows found again after a restart.
            (Path(self.temporary.name)/'session.json').unlink(missing_ok=True)
            controller,backend=self.make([record(1),record(2)])
            controller.preview_driven=driven
            controller.start()
            def resize(hwnd,change):
                rect=backend.windows[hwnd].rect
                backend.gesture_kind=lambda h:'resize'
                backend.cursor=(rect.right,rect.y+100)
                controller._handle_event({'type':'move_start','hwnd':hwnd})
                backend.cursor=(rect.right+change,rect.y+100)
                backend.windows[hwnd]=replace(backend.windows[hwnd],rect=replace(rect,width=rect.width+change))
                self.assertTrue(controller.drive_preview() if driven else controller._preview_native_drag(hwnd) is None)
                self.assertTrue(controller.preview_rectangles)
                controller._handle_event({'type':'move_end','hwnd':hwnd})
                self.assertEqual(controller.preview_rectangles,[])
                self.assertIsNone(controller._preview_hwnd)
            # Resize, resize again, move, resize.
            resize(1,-200)
            self.assertEqual(backend.windows[2].rect,controller._tile_rects[2][1])
            resize(1,-100)
            self.assertEqual(backend.windows[1].rect.width,backend.windows[1].rect.width)
            self.assertEqual(backend.windows[2].rect.x,backend.windows[1].rect.right+controller.settings.gap)
            guide=self.drag_gesture(controller,backend,1,controller._tile_rects[2][1],driven=driven)
            self.assertTrue(guide and guide[2].startswith('Swap'))
            self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[2,1])
            self.assertEqual(controller.drag_guide,None)
            resize(2,150)
            self.assertEqual(backend.windows[1].rect.x,backend.windows[2].rect.right+controller.settings.gap)

    def test_edges_told_by_the_window_follow_the_pointer_from_the_first_pixel(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        tiles={h:backend.windows[h].rect for h in (1,2)}
        backend.gesture_kind=lambda h:'move'
        backend.grab_edges=lambda h:{'right'}
        backend.cursor=(tiles[1].right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        self.assertEqual(controller._gesture_kinds[1],'resize')
        # The pointer moved 30 px, the window not yet: the neighbour follows.
        backend.cursor=(tiles[1].right-30,500)
        controller._preview_native_drag(1)
        self.assertEqual(controller.preview_rectangles[1][1].x,tiles[2].x-30)
        self.assertEqual(controller.preview_rectangles[0][1].height,tiles[1].height)

    def limits(self, controller, hwnd, edge):
        """Where the linked resize lets hwnd's edge go: (least, most) positions."""
        from smartgrid.core.geometry import linked_resize
        profile=controller.profile('d1',0)
        index=next(i for i,a in enumerate(profile.assignments) if a and a.window_id==hwnd)
        rects=controller._profile_rects(profile)
        minimums=controller._gesture_minimums(profile,rects)
        position={'left':'x','right':'right','top':'y','bottom':'bottom'}[edge]
        return tuple(getattr(linked_resize(rects,index,edge,delta,min_sizes=minimums)[index],position) for delta in (-10**5,10**5))

    def held_edge_gesture(self, controller, backend, hwnd, edges, pointer_travel):
        """hwnd held by edges (as the window tells), the pointer going
        pointer_travel=(dx,dy) from the edge's corner; the window follows the
        pointer only as far as the clip allows."""
        rect=backend.windows[hwnd].rect
        backend.gesture_kind=lambda h:'resize'
        backend.grab_edges=lambda h:set(edges)
        px=rect.right if 'right' in edges else rect.x
        py=rect.bottom if 'bottom' in edges else rect.y
        backend.cursor=(px,py)
        controller._handle_event({'type':'move_start','hwnd':hwnd})
        clip=backend.cursor_clip
        self.assertIsNotNone(clip)
        x=min(max(px+pointer_travel[0],clip.x),clip.right-1)
        y=min(max(py+pointer_travel[1],clip.y),clip.bottom-1)
        backend.cursor=(x,y)
        dx,dy=x-px,y-py
        left=rect.x+dx if 'left' in edges else rect.x
        right=rect.right+dx if 'right' in edges else rect.right
        top=rect.y+dy if 'top' in edges else rect.y
        bottom=rect.bottom+dy if 'bottom' in edges else rect.bottom
        backend.windows[hwnd]=replace(backend.windows[hwnd],rect=Rect(left,top,right-left,bottom-top))
        controller._preview_native_drag(hwnd)
        controller._handle_event({'type':'move_end','hwnd':hwnd})
        self.assertIsNone(backend.cursor_clip)
        return clip

    def test_held_window_stops_at_the_limit_on_every_edge(self):
        from smartgrid.core.controller import FORCED_MINIMUM
        controller,backend=self.make([record(h) for h in range(1,5)])
        controller.start()
        # 2x2: window 1 top-left has neighbours right and below, 4 left and above.
        for hwnd,edge,travel,neighbour in ((1,'right',(2000,0),2),(1,'bottom',(0,2000),3),(4,'left',(-2000,0),3),(4,'top',(0,-2000),2)):
            tile=controller._tile_rects[hwnd][1]
            least,most=self.limits(controller,hwnd,edge)
            clip=self.held_edge_gesture(controller,backend,hwnd,{edge},travel)
            axis=0 if edge in ('left','right') else 1
            start=getattr(tile,{'left':'x','right':'right','top':'y','bottom':'bottom'}[edge])
            origin=(tile.right if edge=='right' else tile.x,tile.bottom if edge=='bottom' else tile.y)
            # The pointer may travel exactly as far as the edge can, both ways.
            self.assertEqual((clip.x,clip.right-1) if axis==0 else (clip.y,clip.bottom-1),
                             (origin[axis]+least-start,origin[axis]+most-start))
            # Released past the limit: the window lands exactly there, the
            # neighbour at its minimum.
            placed=backend.windows[hwnd].rect
            self.assertEqual(getattr(placed,{'left':'x','right':'right','top':'y','bottom':'bottom'}[edge]),most if edge in ('right','bottom') else least)
            size=backend.windows[neighbour].rect
            self.assertEqual(size.width if axis==0 else size.height,FORCED_MINIMUM[axis])
            controller.undo()

    def test_corner_keeps_following_the_free_axis(self):
        controller,backend=self.make([record(h) for h in range(1,5)])
        controller.start()
        tile=controller._tile_rects[1][1]
        (xl,xm),(yl,ym)=self.limits(controller,1,'right'),self.limits(controller,1,'bottom')
        # Far right, only 40 px down: x stops at its limit, y follows.
        clip=self.held_edge_gesture(controller,backend,1,{'right','bottom'},(2000,40))
        self.assertEqual((clip.x,clip.right-1),(tile.right+xl-tile.right,tile.right+xm-tile.right))
        self.assertEqual((clip.y,clip.bottom-1),(tile.bottom+yl-tile.bottom,tile.bottom+ym-tile.bottom))
        placed=backend.windows[1].rect
        self.assertEqual(placed.right,xm)
        self.assertEqual(placed.bottom,tile.bottom+40)
        self.assertEqual(backend.windows[3].rect.y,tile.bottom+40+controller.settings.gap)
        controller.undo()
        # Both axes past their limits, from the other corner of window 4.
        tile=controller._tile_rects[4][1]
        self.held_edge_gesture(controller,backend,4,{'left','top'},(-2000,-2000))
        placed=backend.windows[4].rect
        self.assertEqual((placed.x,placed.y),(self.limits(controller,4,'left')[0],self.limits(controller,4,'top')[0]))

    def test_limit_cascades_through_tiles_and_follows_the_settings(self):
        from smartgrid.core.controller import FORCED_MINIMUM
        controller,backend=self.make([record(h) for h in range(1,7)])
        controller.start()
        # 3x2: window 1's right edge pushes 2 then 3 to their minimum.
        most=self.limits(controller,1,'right')[1]
        self.held_edge_gesture(controller,backend,1,{'right'},(3000,0))
        self.assertEqual(backend.windows[1].rect.right,most)
        self.assertEqual([backend.windows[h].rect.width for h in (2,3)],[FORCED_MINIMUM[0]]*2)
        controller.undo()
        # Without forced tiling the applications' own minimums bound the edge;
        # the held window's own minimum bounds it the other way.
        controller.update_settings(replace(controller.settings,force_resize=False))
        backend.minimums[2]=(500,1);backend.minimums[1]=(400,1)
        least,most=self.limits(controller,1,'right')
        clip=self.held_edge_gesture(controller,backend,1,{'right'},(3000,0))
        self.assertEqual(backend.windows[2].rect.width,500)
        self.assertEqual(backend.windows[1].rect.right,most)
        tile=controller._tile_rects[1][1]
        self.assertEqual(clip.x,tile.right+least-tile.right)
        controller.undo()
        self.held_edge_gesture(controller,backend,1,{'right'},(-3000,0))
        self.assertEqual(backend.windows[1].rect.width,400)

    def test_edges_told_by_the_window_moving_confine_the_pointer_too(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        # The cursor on the edge (a mouse resize), the edge not told.
        backend.gesture_kind=lambda hwnd:'resize'
        tiles={h:backend.windows[h].rect for h in (1,2)}
        backend.cursor=(tiles[1].right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        self.assertIsNone(backend.cursor_clip)
        backend.cursor=(tiles[1].right+50,500)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tiles[1],width=tiles[1].width+50))
        controller._preview_native_drag(1)
        self.assertEqual(backend.cursor_clip.right-1,tiles[1].right+self.limits(controller,1,'right')[1]-tiles[1].right)
        # Freed whatever ends the gesture: the window closed, tiling stopped.
        del backend.windows[1]
        controller._handle_event({'type':'destroyed','hwnd':1})
        self.assertIsNone(backend.cursor_clip)
        backend.windows[1]=record(1);controller.refresh()
        backend.cursor=(tiles[1].right,500);backend.grab_edges=lambda h:{'right'}
        controller._handle_event({'type':'move_start','hwnd':1})
        self.assertIsNotNone(backend.cursor_clip)
        controller.stop()
        self.assertIsNone(backend.cursor_clip)

    def test_ui_drives_the_preview_without_signals(self):
        import threading
        controller,backend=self.make([record(1),record(2)])
        controller.preview_driven=True
        controller.start()
        guides=[]
        controller.subscribe_guides(lambda:guides.append(list(controller.preview_rectangles)))
        tiles={h:backend.windows[h].rect for h in (1,2)}
        controller._handle_event({'type':'move_start','hwnd':1})
        # The start is told, so the UI starts its beat; the runtime previews
        # nothing itself on the window's moves.
        self.assertEqual(guides,[[]])
        backend.windows[1]=replace(backend.windows[1],rect=replace(tiles[1],width=tiles[1].width-200))
        controller._handle_event({'type':'location','hwnd':1})
        self.assertFalse(controller.preview_rectangles)
        self.assertTrue(controller.drive_preview())
        self.assertEqual(controller.preview_rectangles[1][1].x,tiles[2].x-200)
        self.assertEqual(len(guides),1)
        # Nothing moved: nothing to draw. Runtime busy: the next frame will do.
        self.assertFalse(controller.drive_preview())
        with controller._lock:
            result=[]
            thread=threading.Thread(target=lambda:result.append(controller.drive_preview()));thread.start();thread.join()
            self.assertEqual(result,[False])
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertIsNone(controller._preview_hwnd)
        self.assertFalse(controller.drive_preview())

    def test_release_leaves_windows_already_in_their_tiles_alone(self):
        controller,backend,tiles=self.resize_gesture(-200)
        images=[]
        controller.subscribe_guides(lambda:images.append(list(controller.preview_rectangles)))
        done=len(backend.operations)
        controller._handle_event({'type':'move_end','hwnd':1})
        # Both windows followed the gesture: no placement, no settling wait.
        self.assertEqual([op for op in backend.operations[done:] if op[0]=='place'],[])
        self.assertEqual(controller._tile_rects[2][1],backend.windows[2].rect)
        # Window 2's image went as soon as it was found in its tile, before the end.
        self.assertNotIn(2,[p[2] for p in images[0]])
        self.assertTrue(controller.profile('d1',0).resize_tiles)

    def test_partly_filled_grid_keeps_its_resized_geometry(self):
        # Five windows in a 3x2 grid: the sixth tile is empty.
        controller,backend=self.make([record(h) for h in range(1,6)])
        controller.start()
        held=backend.windows[1].rect
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.windows[1]=replace(backend.windows[1],rect=replace(held,width=held.width-200))
        controller._preview_native_drag(1)
        previewed=[p[1] for p in controller.preview_rectangles]
        self.assertEqual(len(previewed),6)
        controller._handle_event({'type':'move_end','hwnd':1})
        profile=controller.profile('d1',0)
        self.assertEqual(len(profile.resize_tiles),6)
        self.assertEqual(controller._profile_rects(profile),previewed)
        self.assertEqual(backend.windows[2].rect,previewed[1])

    def test_custom_layout_with_empty_tiles_keeps_its_resized_geometry(self):
        controller,backend=self.make([record(1),record(2)])
        profile=controller.profile('d1',0)
        profile.preset='custom'
        profile.tiles=[Tile('a',0,0,.5,.5),Tile('b',.5,0,.5,.5),Tile('c',0,.5,.5,.5),Tile('d',.5,.5,.5,.5)]
        controller.start()
        held=backend.windows[1].rect
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.windows[1]=replace(backend.windows[1],rect=replace(held,width=held.width-200))
        controller._preview_native_drag(1)
        previewed=[p[1] for p in controller.preview_rectangles]
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(controller._profile_rects(controller.profile('d1',0)),previewed)
        self.assertEqual(backend.windows[2].rect,previewed[1])

    def test_an_edge_drag_moves_the_neighbours_from_its_first_pixels(self):
        controller,backend=self.make([record(1),record(2)])
        backend.gesture_kind=lambda hwnd:'resize'
        controller.start()
        tiles={h:backend.windows[h].rect for h in (1,2)}
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.windows[1]=replace(backend.windows[1],rect=replace(tiles[1],width=tiles[1].width-3))
        controller._preview_native_drag(1)
        self.assertIsNone(controller.drag_guide)
        self.assertEqual(controller.preview_rectangles[1][1].x,tiles[2].x-3)
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[2].rect.x,tiles[2].x-3)
        self.assertTrue(controller.profile('d1',0).resize_tiles)

    def test_preview_follows_the_pointer_past_a_slow_window(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        # The cursor on the edge (a mouse resize), the edge not told.
        backend.gesture_kind=lambda hwnd:'resize'
        tiles={h:backend.windows[h].rect for h in (1,2)}
        backend.cursor=(tiles[1].right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        # The pointer tells nothing before the window says which edge is held.
        backend.cursor=(tiles[1].right-60,500)
        controller._preview_native_drag(1)
        self.assertFalse(controller.preview_rectangles)
        # The window followed the first 100 px; the pointer went on 150 more
        # (a slow application, or one at its minimum size).
        backend.cursor=(tiles[1].right-100,500)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tiles[1],width=tiles[1].width-100))
        controller._preview_native_drag(1)
        self.assertEqual(controller.preview_rectangles[1][1].x,tiles[2].x-100)
        backend.cursor=(tiles[1].right-250,540)
        controller._preview_native_drag(1)
        self.assertEqual(controller.preview_rectangles[1][1].x,tiles[2].x-250)
        # Only the edge the gesture moves follows the pointer: not the bottom.
        self.assertEqual(controller.preview_rectangles[0][1].height,tiles[1].height)
        # The window catching up part of the way does not pull the preview back.
        backend.windows[1]=replace(backend.windows[1],rect=replace(tiles[1],width=tiles[1].width-150))
        controller._preview_native_drag(1)
        self.assertEqual(controller.preview_rectangles[1][1].x,tiles[2].x-250)
        # The release ends where the images showed; the held window is put there.
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[2].rect.x,tiles[2].x-250)
        self.assertEqual(backend.windows[1].rect.width,tiles[1].width-250)

    def test_release_takes_the_window_where_windows_left_it(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        tiles={h:backend.windows[h].rect for h in (1,2)}
        backend.cursor=(tiles[1].right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.cursor=(tiles[1].right-200,500)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tiles[1],width=tiles[1].width-200))
        controller._preview_native_drag(1)
        # Released with a flick: the window went 5 px further than previewed.
        backend.windows[1]=replace(backend.windows[1],rect=replace(tiles[1],width=tiles[1].width-205))
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[1].rect.width,tiles[1].width-205)
        self.assertEqual(backend.windows[2].rect.x,tiles[2].x-205)
        # Escape: Windows puts the window back although the pointer is far away.
        controller._handle_event({'type':'move_start','hwnd':1})
        held=backend.windows[1].rect
        backend.cursor=(held.right-200,500)
        backend.windows[1]=replace(backend.windows[1],rect=replace(held,width=held.width-200))
        controller._preview_native_drag(1)
        backend.windows[1]=replace(backend.windows[1],rect=held)
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[1].rect,held)
        self.assertEqual(backend.windows[2].rect.x,tiles[2].x-205)

    def test_gesture_is_previewed_between_events(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        controller.start_runtime()
        try:
            held=backend.windows[1].rect
            controller._enqueue_event({'type':'move_start','hwnd':1})
            deadline=time.monotonic()+2
            while 1 not in controller._interacting and time.monotonic()<deadline:
                time.sleep(.01)
            # The window changes without any event reaching the queue.
            backend.windows[1]=replace(backend.windows[1],rect=replace(held,width=held.width-200))
            while not controller.preview_rectangles and time.monotonic()<deadline:
                time.sleep(.01)
            self.assertTrue(controller.preview_rectangles)
        finally:
            controller.quit()

    def test_linked_resize_release_places_neighbours_and_undo_restores_them(self):
        controller,backend,tiles=self.resize_gesture(-200)
        target=controller.preview_rectangles[1][1]
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[2].rect,target)
        self.assertTrue(controller.profile('d1',0).resize_tiles)
        self.assertEqual(controller.preview_rectangles,[])
        self.assertFalse(controller.motion_events)
        controller.undo()
        self.assertEqual(backend.windows[2].rect,tiles[2])

    def test_cancelled_linked_resize_leaves_neighbours_in_their_tiles(self):
        controller,backend,tiles=self.resize_gesture(-200)
        # Escape: Windows puts the held window back where it was.
        backend.windows[1]=replace(backend.windows[1],rect=tiles[1])
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[2].rect,tiles[2])
        self.assertFalse(controller.profile('d1',0).resize_tiles)
        self.assertEqual(controller.preview_rectangles,[])

    def test_neighbours_follow_back_to_the_start_size(self):
        controller,backend,tiles=self.resize_gesture(-200)
        backend.windows[1]=replace(backend.windows[1],rect=tiles[1])
        controller._preview_native_drag(1)
        self.assertEqual(controller.preview_rectangles[1][1],tiles[2])

    def test_window_closed_during_linked_resize_ends_the_gesture(self):
        controller,backend,tiles=self.resize_gesture(-200)
        del backend.windows[1]
        controller._handle_event({'type':'destroyed','hwnd':1})
        self.assertFalse(controller._interacting)
        self.assertEqual(controller.preview_rectangles,[])
        self.assertEqual(self.live_requests(backend),[])

    def test_minimized_neighbour_is_not_shown_during_linked_resize(self):
        controller,backend=self.make([record(1),record(2),record(3)])
        controller.start()
        held=backend.windows[1].rect
        backend.windows[3]=replace(backend.windows[3],state='minimized')
        controller.refresh(auto=False)
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.windows[1]=replace(backend.windows[1],rect=Rect(held.x,held.y,held.width-200,held.height))
        controller._preview_native_drag(1)
        self.assertNotIn(3,[p[2] for p in controller.preview_rectangles])

    def test_linked_resize_takes_the_drop_target_of_its_first_pixels_away(self):
        controller,backend,tiles=self.resize_gesture(3)
        self.assertIsNotNone(controller.drag_guide)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tiles[1],width=tiles[1].width-200))
        controller._preview_native_drag(1)
        self.assertTrue(controller.preview_rectangles)
        self.assertIsNone(controller.drag_guide)

    def test_failed_arrangement_still_ends_the_linked_resize(self):
        controller,backend,tiles=self.resize_gesture(-200)
        backend.desktop_known=False
        with self.assertRaises(ValueError):
            controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(controller.preview_rectangles,[])
        self.assertEqual(backend.windows[2].rect,tiles[2])

    def test_neighbour_maximized_during_linked_resize_is_left_alone(self):
        controller,backend,tiles=self.resize_gesture(-200)
        # Maximized since the last scan: its record still says normal.
        backend.windows[2]=replace(backend.windows[2],state='maximized')
        backend.windows[1]=replace(backend.windows[1],rect=tiles[1])
        backend.discover_windows=lambda:[replace(w,state='normal') if w.ref.hwnd==2 else w for w in backend.windows.values()]
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[2].state,'maximized')

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

    def late_start(self, travel):
        """Window 1 (2x1) held by its right edge; the runtime handles the
        start when the window has already grown by travel, the event telling
        where the window and the pointer were when Windows started it."""
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        backend.grab_edges=lambda h:{'right'}
        tile=backend.windows[1].rect
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+travel))
        backend.cursor=(tile.right+travel,500)
        controller._handle_event({'type':'move_start','hwnd':1,'rect':tile,'cursor':(tile.right,500)})
        return controller,backend,tile

    def test_a_gesture_handled_late_loses_nothing_at_the_release(self):
        controller,backend,tile=self.late_start(60)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+200))
        backend.cursor=(tile.right+200,500)
        controller._preview_native_drag(1)
        self.assertEqual(controller.preview_rectangles[0][1].right,tile.right+200)
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[1].rect.right,tile.right+200)

    def test_undo_after_a_gesture_handled_late_puts_the_window_back_in_its_tile(self):
        controller,backend,tile=self.late_start(60)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+200))
        backend.cursor=(tile.right+200,500)
        controller._preview_native_drag(1)
        controller._handle_event({'type':'move_end','hwnd':1})
        controller.undo()
        self.assertEqual(backend.windows[1].rect,tile)

    def test_escape_cancels_a_gesture_handled_late(self):
        controller,backend,tile=self.late_start(40)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+300))
        backend.cursor=(tile.right+300,500)
        controller._preview_native_drag(1)
        # Escape: Windows puts the window back; the pointer stays.
        backend.windows[1]=replace(backend.windows[1],rect=tile)
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[1].rect,tile)
        self.assertFalse(controller.profile('d1',0).resize_tiles)

    def test_release_keeps_the_minimums_of_the_preview(self):
        controller,backend=self.make([record(1),record(2)])
        controller.settings.force_resize=False
        controller.start()
        backend.grab_edges=lambda h:{'right'}
        answers=iter([(1,1),(1,1)])
        # Busy at the start (no answer), then 700 px: the preview's stays.
        backend.min_size=lambda h:next(answers,(700,100) if h==2 else (1,1))
        tile=backend.windows[1].rect
        backend.cursor=(tile.right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.cursor=(tile.right+600,500)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+600))
        controller._preview_native_drag(1)
        preview=[p[1] for p in controller.preview_rectangles]
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual([backend.windows[h].rect for h in (1,2)],preview)

    def test_an_outer_edge_does_not_hold_the_pointer_still(self):
        controller,backend=self.make([record(h,f'app{h}') for h in range(1,5)])
        controller.start()
        held=max(backend.windows,key=lambda h:(backend.windows[h].rect.right,backend.windows[h].rect.bottom))
        tile=backend.windows[held].rect
        backend.grab_edges=lambda h:{'right','bottom'}
        backend.cursor=(tile.right,tile.bottom)
        controller._handle_event({'type':'move_start','hwnd':held})
        area=self.displays[0].work_area
        self.assertEqual(backend.cursor_clip,Rect(area.x,area.y,area.width,area.height))
        # Past its tile, the window is covered up to it.
        backend.windows[held]=replace(backend.windows[held],rect=replace(tile,width=tile.width-100))
        backend.cursor=(tile.right-100,tile.bottom)
        controller._preview_native_drag(held)
        self.assertIn(('d1',Rect(tile.right-100,tile.y,100,tile.height)),controller.preview_covers)

    def test_nothing_is_rescanned_during_a_gesture_and_new_windows_join_at_the_release(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        controller._suppress_events_until=0
        backend.grab_edges=lambda h:{'right'}
        tile=backend.windows[1].rect
        backend.cursor=(tile.right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        scans=[];discover=backend.discover_windows
        backend.discover_windows=lambda:(scans.append(1),discover())[1]
        backend.windows[3]=record(3,'new')
        for kind in ('created','shown','reconcile','hidden','minimized'):
            controller._last_reconcile=0
            controller._handle_event({'type':kind,'hwnd':3})
        self.assertEqual(scans,[])
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertIn(3,[a.window_id for a in controller.profile('d1',0).assignments if a])

    def test_a_window_closed_during_a_resize_does_not_change_the_resize(self):
        controller,backend=self.make([record(h,f'app{h}') for h in (1,2,3)])
        controller.start()
        controller._suppress_events_until=0
        backend.grab_edges=lambda h:{'right'}
        tile=backend.windows[1].rect
        backend.cursor=(tile.right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.cursor=(tile.right+100,500)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+100))
        controller._preview_native_drag(1)
        preview=[p[1] for p in controller.preview_rectangles]
        del backend.windows[3]
        controller._handle_event({'type':'destroyed','hwnd':3})
        controller._handle_event({'type':'move_end','hwnd':1})
        # Made for the three tiles shown; back when three windows are.
        backend.windows[3]=record(3,'app3');controller.refresh()
        self.assertEqual([backend.windows[h].rect for h in (1,2,3)],preview)

    def test_a_drag_to_another_display_survives_a_window_opened_meanwhile(self):
        controller,backend=self.make([record(1),record(2),record(3,'app','d2')])
        backend.gesture_kind=lambda hwnd:'move'
        controller.start()
        controller._suppress_events_until=0
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.windows[9]=replace(record(9,'tooltip'),eligible=False)
        controller._handle_event({'type':'created','hwnd':9})
        backend.windows[1]=replace(backend.windows[1],rect=Rect(2200,100,640,480),display_id='d2')
        backend.cursor=(2500,300)
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertIn(1,[a.window_id for a in controller.profile('d2',0).assignments if a])

    def test_shortcuts_that_move_windows_wait_for_the_release(self):
        for action in ('arrange','undo','space2','float','swap','toggle','pause'):
            controller,backend=self.make([record(1),record(2)])
            controller.start()
            backend.grab_edges=lambda h:{'right'}
            tile=backend.windows[1].rect
            backend.cursor=(tile.right,500)
            controller._handle_event({'type':'move_start','hwnd':1})
            backend.foreground_hwnd=1
            operations=len(backend.operations)
            controller.handle_action(action)
            self.assertEqual(backend.operations[operations:],[],action)
            self.assertEqual(controller.active_spaces['d1'],0,action)
            self.assertIsNone(controller._swap_snapshot,action)
            self.assertNotIn(1,controller._floating,action)
            self.assertTrue(controller.running and not controller.paused,action)

    def test_a_resize_is_the_layout_remembered_for_its_preset(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        backend.grab_edges=lambda h:{'right'}
        tile=backend.windows[1].rect
        backend.cursor=(tile.right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.cursor=(tile.right+200,500)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+200))
        controller._preview_native_drag(1)
        controller._handle_event({'type':'move_end','hwnd':1})
        profile=controller.profile('d1',0)
        self.assertEqual(controller.repository.get_variant('d1',0,profile.preset).resize_tiles,profile.resize_tiles)

    def test_a_window_opened_just_after_an_arrangement_joins_at_the_release(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        backend.grab_edges=lambda h:{'right'}
        tile=backend.windows[1].rect
        backend.cursor=(tile.right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        # Within the quiet time after an arrangement: the event is not seen.
        controller._suppress_events_until=time.monotonic()+60
        backend.windows[3]=record(3,'new')
        controller._handle_event({'type':'created','hwnd':3})
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertIn(3,[a.window_id for a in controller.profile('d1',0).assignments if a])

    def test_the_held_window_takes_no_lock_from_the_event_thread(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        controller._enqueue_event({'type':'move_start','hwnd':1})
        guarded=[];controller._guard_slots=lambda only=None:guarded.append(only)
        controller._enqueue_event({'type':'location','hwnd':1})
        self.assertEqual(guarded,[])

    def test_cover_colour_is_read_where_nothing_moves(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        backend.grab_edges=lambda h:{'right'}
        tile=backend.windows[1].rect
        backend.cursor=(tile.right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        # Read at the start, in the margin along the other tile (the only gap
        # touches the held window and its outline).
        display,strip=controller.preview_gap
        other=backend.windows[2].rect
        self.assertTrue(strip.x>=other.x and strip.right<=other.right)

    def test_keyboard_sizing_is_followed_by_the_window_and_kept(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        tile=backend.windows[1].rect
        # Alt+Space, Size: the cursor in the middle (kind unknown), then on the edge.
        backend.cursor=(tile.x+tile.width//2,tile.y+tile.height//2)
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.cursor=(tile.right+120,tile.y+tile.height//2)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+120))
        controller._preview_native_drag(1)
        self.assertEqual(controller.preview_rectangles[0][1].right,tile.right+120)
        # Held at the limit like a mouse resize, from where the pointer is.
        self.assertEqual(backend.cursor_clip.right-1,tile.right+self.limits(controller,1,'right')[1]-tile.right)
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertEqual(backend.windows[1].rect.right,tile.right+120)

    def test_pause_frees_the_pointer(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        backend.grab_edges=lambda h:{'right'}
        tile=backend.windows[1].rect
        backend.cursor=(tile.right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        self.assertIsNotNone(backend.cursor_clip)
        controller.pause()
        self.assertIsNone(backend.cursor_clip)

    def test_release_below_the_minimum_is_forced_at_once(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        backend.grab_edges=lambda h:{'right'}
        backend.min_size=lambda h:(900,100) if h==2 else (1,1)
        tile=backend.windows[1].rect
        backend.cursor=(tile.right,500)
        controller._handle_event({'type':'move_start','hwnd':1})
        backend.cursor=(tile.right+400,500)
        backend.windows[1]=replace(backend.windows[1],rect=replace(tile,width=tile.width+400))
        controller._preview_native_drag(1)
        controller._handle_event({'type':'move_end','hwnd':1})
        self.assertNotIn(2,[op[1] for op in backend.operations if op[0]=='place_live'])

    def test_the_ui_beat_never_waits_for_an_application(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        controller._interaction_minimums.clear()
        asked=[];backend.min_size=lambda h:asked.append(h) or (1,1)
        controller._driving=True
        try:
            controller._gesture_minimums(controller.profile('d1',0),controller._profile_rects(controller.profile('d1',0)))
        finally:
            controller._driving=False
        self.assertEqual(asked,[])

    def test_a_gesture_never_ended_is_ended_after_two_minutes(self):
        controller,backend=self.make([record(1),record(2)])
        controller.start()
        controller._handle_event({'type':'move_start','hwnd':1})
        controller._end_stale_gestures()
        self.assertIn(1,controller._interacting)
        controller._interaction_started[1]-=121
        # Still held: a long resize goes on.
        backend.mouse_pressed=True
        controller._end_stale_gestures()
        self.assertIn(1,controller._interacting)
        backend.mouse_pressed=False
        controller._end_stale_gestures()
        self.assertFalse(controller._interacting)

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

    def test_hidden_windows_leave_no_empty_reserved_tiles(self):
        controller,backend = self.make([record(h,f'app{h}') for h in range(1,4)])
        controller.start()
        for hwnd in (4,5):
            backend.windows[hwnd]=record(hwnd,'teams')
        controller.refresh()
        self.assertEqual(len(controller.profile('d1',0).assignments),5)
        # Teams hides its call windows instead of closing them.
        for hwnd in (4,5):
            backend.windows[hwnd]=replace(backend.windows[hwnd],eligible=False,exclusion_reason='Not visible')
        controller.refresh()
        self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[1,2,3])
        backend.windows[6]=record(6,'other')
        controller.refresh()
        del backend.windows[6],backend.windows[4],backend.windows[5]
        controller.refresh()
        self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[1,2,3])

    def test_window_excluded_by_its_new_title_frees_its_tile(self):
        controller,backend = self.make([record(1),record(2,'teams')])
        controller.start()
        backend.windows[2]=replace(backend.windows[2],title='Meeting with Alex | Microsoft Teams')
        controller.refresh()
        self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[1])
        backend.windows[2]=replace(backend.windows[2],title='Chat | Microsoft Teams')
        controller.refresh()
        self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[1,2])

    def test_small_windows_are_not_tiled_until_they_grow(self):
        controller,backend = self.make([record(1)])
        backend.windows[2]=replace(record(2,'bar'),rect=Rect(800,0,600,60))
        controller.start()
        self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[1])
        backend.windows[2]=replace(backend.windows[2],rect=Rect(800,0,600,400))
        controller.refresh()
        self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[1,2])

    def test_window_in_a_small_tile_stays_tiled(self):
        controller,backend = self.make([record(h,f'app{h}') for h in range(1,4)])
        controller.start()
        backend.windows[1]=replace(backend.windows[1],rect=Rect(0,0,170,170))
        controller.refresh()
        self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments],[1,2,3])

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

    def maximized_behind(self, count=3, order=(1,2,3), pointer=None):
        """Window 3 maximized from its own button, left behind the others."""
        controller,backend = self.make([record(h,f'app{h}') for h in range(1,count+1)])
        controller.start()
        tile = backend.windows[3].rect
        backend.z_order = list(order);backend.foreground_hwnd = order[0]
        backend.cursor = pointer or (tile.right-60,tile.y+15)
        backend.window_at = lambda point: 3
        backend.windows[3] = replace(backend.windows[3],state='maximized',rect=Rect(0,0,1920,1080))
        backend.operations.clear()
        return controller,backend,tile

    def fronted(self, backend):
        return [o for o in backend.operations if o[0] in ('activate','raise')]

    def test_window_maximized_by_its_own_button_comes_in_front_of_the_tiles(self):
        controller,backend,_ = self.maximized_behind()
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(backend.z_order[0],3);self.assertEqual(backend.foreground_hwnd,3)
        self.assertEqual(backend.windows[3].state,'maximized')
        # Once per maximization, and nothing more once in front.
        controller._handle_event({'type':'location','hwnd':3});controller._recheck_front()
        self.assertEqual(self.fronted(backend),[('activate',3)])
        # Restored, then maximized again behind: in front again.
        backend.windows[3] = replace(backend.windows[3],state='normal',rect=controller._tile_rects[3][1])
        controller._handle_event({'type':'location','hwnd':3})
        backend.windows[3] = replace(backend.windows[3],state='maximized',rect=Rect(0,0,1920,1080))
        backend.z_order = [1,2,3]
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(backend.z_order[0],3)

    def test_active_window_still_behind_is_raised_and_not_left_always_on_top(self):
        controller,backend,_ = self.maximized_behind()
        backend._to_front = lambda hwnd: None  # Activated, not raised.
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(backend.foreground_hwnd,3);self.assertEqual(backend.z_order[0],1)
        del backend._to_front
        controller._recheck_front()
        self.assertEqual(backend.z_order[0],3)
        backend.topmost.add(3)
        controller._recheck_topmost()
        self.assertEqual(backend.topmost,set())
        controller._recheck_topmost()
        self.assertIsNone(controller._topmost_check)
        # A window that never answers is not asked forever.
        backend.topmost.add(3);backend.drop_topmost = lambda hwnd: False
        controller._topmost_check = (3,0,2)
        controller._recheck_topmost();self.assertIsNotNone(controller._topmost_check)
        controller._recheck_topmost();self.assertIsNone(controller._topmost_check)

    def test_window_refused_the_foreground_never_covers_the_keyboard(self):
        controller,backend,_ = self.maximized_behind()
        backend.failures.add(3)  # Windows refuses the activation.
        controller._handle_event({'type':'location','hwnd':3})
        for _ in range(3): controller._recheck_front()
        self.assertEqual(self.fronted(backend),[('activate',3)]*4)
        self.assertEqual(backend.z_order[0],1);self.assertIsNone(controller._raise_check)

    def test_activation_is_not_asked_again_once_the_pointer_is_on_a_tile(self):
        controller,backend,_ = self.maximized_behind()
        backend.failures.add(3)
        controller._handle_event({'type':'location','hwnd':3})
        backend.window_at = lambda point: 1
        for _ in range(3): controller._recheck_front()
        # The click, and one more try at once.
        self.assertEqual(self.fronted(backend),[('activate',3)]*2)

    def test_activation_refused_once_is_asked_again(self):
        controller,backend,_ = self.maximized_behind()
        backend.failures.add(3)
        controller._handle_event({'type':'location','hwnd':3})
        backend.failures.clear()
        controller._recheck_front()
        self.assertEqual(backend.z_order[0],3);self.assertEqual(backend.foreground_hwnd,3)
        # Then the user goes back to the other window: left there.
        backend.foreground_hwnd = 1;backend.z_order = [1,3,2]
        controller._recheck_front()
        self.assertEqual(backend.foreground_hwnd,1);self.assertIsNone(controller._raise_check)

    def test_window_maximized_by_a_drag_is_not_brought_to_the_front_later(self):
        controller,backend,_ = self.maximized_behind()
        controller._interacting.add(3)
        controller._enqueue_event({'type':'location','hwnd':3})
        controller._handle_event({'type':'location','hwnd':3})
        controller._interacting.clear()
        controller._handle_event({'type':'location','hwnd':3})
        controller.refresh()
        self.assertEqual(self.fronted(backend),[])
        # Maximized before SmartGrid saw it at all: not either.
        other,backend,_ = self.maximized_behind()
        other._maximized.clear();other.windows=[w for w in other.windows if w.ref.hwnd!=3]
        other.refresh(auto=False)
        self.assertTrue(other._maximized[3]["handled"])

    def test_maximized_window_comes_to_the_front_from_the_event_thread(self):
        controller,backend,_ = self.maximized_behind()
        controller._enqueue_event({'type':'location','hwnd':3})
        self.assertEqual(backend.z_order[0],3)
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[('activate',3)])

    def busy_enqueue(self, controller, hwnd):
        """The event queued while the runtime holds the lock."""
        import threading
        with controller._lock:
            thread = threading.Thread(target=controller._enqueue_event,args=({'type':'location','hwnd':hwnd},))
            thread.start();thread.join()

    def test_pointer_gone_before_the_event_is_handled_is_no_miss(self):
        controller,backend,_ = self.maximized_behind()
        self.busy_enqueue(controller,3)
        backend.cursor = (5,5);backend.window_at = lambda point: 1
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(backend.z_order[0],3)

    def test_active_window_is_raised_at_once_whatever_the_pointer_does_next(self):
        controller,backend,_ = self.maximized_behind()
        backend.foreground_hwnd = 3
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[('raise',3)])
        # Sent behind again by its application: raised at the next checks,
        # three times at most.
        backend.cursor = (5,5)
        for _ in range(3):
            backend.z_order = [1,2,3];controller._recheck_front()
            self.assertEqual(backend.z_order[0],3)
        self.assertIsNone(controller._raise_check)

    def test_foreground_taken_back_after_its_activation_is_respected(self):
        controller,backend,_ = self.maximized_behind()
        controller._handle_event({'type':'location','hwnd':3})
        backend.foreground_hwnd = 1;backend.z_order = [1,3,2]
        controller._recheck_front()
        self.assertEqual(self.fronted(backend),[('activate',3)]);self.assertIsNone(controller._raise_check)

    def test_late_activation_is_followed_by_a_raise(self):
        controller,backend,_ = self.maximized_behind()
        backend.failures.add(3)
        controller._handle_event({'type':'location','hwnd':3})
        controller._recheck_front()
        backend.foreground_hwnd = 3
        controller._recheck_front()
        self.assertEqual(backend.z_order[0],3)

    def test_another_window_activated_before_the_event_is_handled_stays_in_front(self):
        controller,backend,_ = self.maximized_behind()
        self.busy_enqueue(controller,3)
        backend.foreground_hwnd = 2
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[]);self.assertIsNone(controller._raise_check)

    def test_restored_and_maximized_again_before_being_handled(self):
        controller,backend,tile = self.maximized_behind()
        controller._handle_event({'type':'location','hwnd':3})
        maximized = backend.windows[3]
        backend.windows[3] = replace(maximized,state='normal',rect=tile)
        controller._enqueue_event({'type':'location','hwnd':3})
        backend.windows[3] = maximized;backend.z_order = [1,2,3];backend.foreground_hwnd = 1
        controller._enqueue_event({'type':'location','hwnd':3})
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(backend.z_order[0],3)

    def test_maximized_while_another_window_is_held_comes_to_the_front_after(self):
        controller,backend,_ = self.maximized_behind()
        controller._interacting.add(1)
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[])
        controller._interacting.clear();controller.refresh()
        self.assertEqual(backend.z_order[0],3)

    def test_card_of_a_window_minimized_while_maximized_brings_it_to_the_front(self):
        controller,backend,_ = self.maximized_behind(order=(3,1,2))
        next(a for a in controller.profile('d1',0).assignments if a and a.window_id==3).pinned = True
        controller.refresh()
        backend.windows[3] = replace(backend.windows[3],state='minimized')
        controller.refresh()
        backend.z_order = [1,2,3];backend.foreground_hwnd = 1
        # Windows brings it back as it was before: maximized.
        backend.show = lambda hwnd: backend.windows.update({hwnd:replace(backend.windows[hwnd],state='maximized')}) or False
        index = next(i for i,a in enumerate(controller.profile('d1',0).assignments) if a and a.window_id==3)
        backend.operations.clear()
        controller.activate_reserved('d1',0,index)
        self.assertEqual(backend.windows[3].state,'maximized')
        self.assertEqual(backend.z_order[0],3);self.assertEqual(backend.foreground_hwnd,3)
        controller._handle_event({'type':'location','hwnd':3})
        # Recalled in front by the backend itself: not shown without activation.
        shown = []
        backend.windows[3] = replace(backend.windows[3],state='minimized');controller.refresh()
        backend.show = shown.append
        backend.recall_maximized = lambda hwnd: backend.windows.update({hwnd:replace(backend.windows[hwnd],state='maximized')}) or True
        backend.foreground_hwnd = 3
        controller.activate_reserved('d1',0,index)
        self.assertEqual(shown,[]);self.assertEqual(backend.windows[3].state,'maximized')
        backend.operations[:] = backend.operations[:1]
        controller._enqueue_event({'type':'location','hwnd':3})
        controller._handle_event({'type':'location','hwnd':3})
        for _ in range(3): controller._recheck_front() if controller._raise_check else None
        self.assertEqual(self.fronted(backend),[('activate',3)])

    def test_card_of_a_window_back_in_its_tile_gives_it_no_focus(self):
        controller,backend,_ = self.maximized_behind(order=(3,1,2))
        next(a for a in controller.profile('d1',0).assignments if a and a.window_id==3).pinned = True
        controller.refresh()
        backend.windows[3] = replace(backend.windows[3],state='minimized')
        controller.refresh()
        index = next(i for i,a in enumerate(controller.profile('d1',0).assignments) if a and a.window_id==3)
        backend.foreground_hwnd = 1;backend.operations.clear()
        controller.activate_reserved('d1',0,index)
        self.assertEqual(backend.windows[3].state,'normal')
        self.assertEqual(self.fronted(backend),[]);self.assertEqual(backend.foreground_hwnd,1)

    def pressed(self, **state):
        """The mouse pressed in the title bar of window 3, window 1 in front."""
        controller,backend = self.make([record(h,f'app{h}') for h in (1,2,3)])
        controller.start()
        tile = backend.windows[3].rect
        backend.z_order = [1,2,3];backend.foreground_hwnd = 1;backend.mouse_pressed = True
        backend.cursor = (tile.right-60,tile.y+15)
        backend.window_at = lambda point: 3
        backend.operations.clear()
        for name,value in state.items(): setattr(backend,name,value(tile) if callable(value) and name=='cursor' else value)
        return controller,backend

    def press(self, controller, then=None):
        import threading
        from unittest.mock import patch
        with patch('smartgrid.core.controller.threading.Timer') as timer:
            controller._enqueue_event({'type':'capture','hwnd':3})
        if timer.called:
            if then: then()
            controller._activate_pressed(*timer.call_args[0][2])
        return timer.called

    def test_window_pressed_on_its_title_bar_is_activated_before_it_maximizes(self):
        controller,backend = self.pressed()
        self.assertTrue(self.press(controller))
        self.assertEqual(backend.foreground_hwnd,3)
        # In front already once maximized: nothing more.
        backend.windows[3] = replace(backend.windows[3],state='maximized',rect=Rect(0,0,1920,1080))
        controller._enqueue_event({'type':'location','hwnd':3})
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[('activate',3)])

    def test_press_activating_its_window_by_itself_is_left_to_windows(self):
        controller,backend = self.pressed(foreground_hwnd=3)
        self.assertFalse(self.press(controller))
        controller,backend = self.pressed()
        self.press(controller,lambda: setattr(backend,'foreground_hwnd',3))
        self.assertEqual(self.fronted(backend),[])
        controller,backend = self.pressed()
        self.press(controller,lambda: setattr(backend,'foreground_hwnd',2))
        self.assertEqual(self.fronted(backend),[])

    def test_other_presses_never_activate_a_window(self):
        for state in (dict(mouse_pressed=False),dict(cursor=lambda tile:(tile.right-60,tile.y+200)),
                      dict(cursor=lambda tile:(5,tile.y-40)),dict(window_at=lambda point:1),
                      dict(accepts_activation=lambda hwnd:False),dict(foreground_hwnd=0)):
            controller,backend = self.pressed(**state)
            self.press(controller)
            self.assertEqual(self.fronted(backend),[],state)
        for change in (lambda c: c._floating.add(3),lambda c: setattr(c,'paused',True),
                       lambda c: c._interacting.add(1),lambda c: c._native_gestures.update({1:0})):
            controller,backend = self.pressed();change(controller)
            self.press(controller)
            self.assertEqual(self.fronted(backend),[])
        # A gesture started by this very press.
        controller,backend = self.pressed()
        self.press(controller,lambda: controller._native_gestures.update({3:0}))
        self.assertEqual(self.fronted(backend),[])

    def test_maximization_first_seen_by_a_rescan_comes_to_the_front_too(self):
        controller,backend,_ = self.maximized_behind()
        controller.refresh()
        self.assertEqual(self.fronted(backend),[('activate',3)])
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[('activate',3)])

    def test_window_already_in_front_once_maximized_is_left_alone(self):
        controller,backend,_ = self.maximized_behind(order=(3,1,2))
        controller._handle_event({'type':'location','hwnd':3})
        controller._recheck_front()
        self.assertEqual(self.fronted(backend),[])

    def test_window_maximizing_itself_in_the_background_stays_behind(self):
        controller,backend,tile = self.maximized_behind(pointer=(5,5))
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[])
        # The pointer on its content, or on another window in front of it.
        for pointer,over in (((tile.right-60,tile.y+200),3),((tile.right-60,tile.y+15),1)):
            controller._maximized.clear();backend.cursor = pointer
            backend.window_at = lambda point: over
            controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[]);self.assertIsNone(controller._raise_check)

    def test_another_window_activated_meanwhile_is_not_covered(self):
        controller,backend,_ = self.maximized_behind()
        backend.failures.add(3)
        controller._handle_event({'type':'location','hwnd':3})
        backend.foreground_hwnd = 2
        controller._recheck_front()
        self.assertEqual(self.fronted(backend),[('activate',3)])

    def test_floating_windows_and_a_stopped_grid_are_never_brought_to_the_front(self):
        controller,backend,_ = self.maximized_behind()
        controller._floating.add(3)
        controller._handle_event({'type':'location','hwnd':3})
        controller._floating.clear();controller._maximized.clear();controller.paused = True
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[])
        # A floating window in front of it stays there on its own.
        controller,backend,_ = self.maximized_behind(order=(1,3,2))
        controller._floating.add(1)
        controller._handle_event({'type':'location','hwnd':3})
        self.assertEqual(self.fronted(backend),[])

    def test_launched_application_starting_maximized_goes_into_its_tile(self):
        controller,backend = self.pending_template()
        arrived = backend.add_window('slow',state='maximized');hwnd = arrived.ref.hwnd
        controller.refresh()
        self.assertEqual(controller.profile('d1',0).assignments[0].window_id,hwnd)
        self.assertEqual(backend.windows[hwnd].state,'normal')
        self.assertEqual(controller._tile_rects[hwnd][1],backend.windows[hwnd].rect)
        self.assertNotEqual(backend.windows[hwnd].rect,arrived.rect)

    def test_launched_application_maximizing_itself_is_restored_three_more_times(self):
        controller,backend = self.pending_template()
        hwnd = backend.add_window('slow').ref.hwnd
        controller.refresh()
        for _ in range(3):
            backend.maximize(hwnd);controller.refresh()
            self.assertEqual(backend.windows[hwnd].state,'normal')
            self.assertEqual(controller._tile_rects[hwnd][1],backend.windows[hwnd].rect)
        backend.maximize(hwnd);controller.refresh()
        self.assertEqual(backend.windows[hwnd].state,'maximized')

    def test_maximize_after_the_launch_grace_is_kept(self):
        controller,backend = self.pending_template()
        hwnd = backend.add_window('slow').ref.hwnd
        controller.refresh()
        controller._launch_grace[hwnd][0] -= 6
        backend.maximize(hwnd);controller.refresh()
        self.assertEqual(backend.windows[hwnd].state,'maximized')
        self.assertFalse(controller._launch_grace)

    def test_maximized_launch_is_kept_while_paused_floating_or_not_launched(self):
        controller,backend = self.pending_template()
        other = backend.add_window('other',state='maximized').ref.hwnd
        controller.refresh()
        self.assertEqual(backend.windows[other].state,'maximized')
        backend.remove_window(other);controller.refresh()
        hwnd = backend.add_window('slow').ref.hwnd
        controller.refresh()
        controller.pause()
        backend.maximize(hwnd);controller.refresh()
        self.assertEqual(backend.windows[hwnd].state,'maximized')
        controller.paused = False;controller._floating.add(hwnd)
        controller.refresh()
        self.assertEqual(backend.windows[hwnd].state,'maximized')

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


class RestartIdentityTests(unittest.TestCase):
    def make(self, directory, pins=(False,False), displays=1):
        shown=[Display(f'd{i+1}',f'D{i+1}',Rect(1920*i,0,1920,1080),primary=not i) for i in range(displays)]
        backend=FakeBackend(shown,[record(1,'chrome'),record(2,'chrome')],[ApplicationRef('chrome','Chrome')])
        backend.auto_launch=False
        controller=Controller(backend,Repository(directory))
        controller.profiles[('d1',0)]=SpaceProfile('d1',0,assignments=[
            Assignment('chrome',2,pins[0]),Assignment('chrome',1,pins[1])])
        controller.start()
        return backend,controller

    def restart(self, backend, directory, start=True):
        controller=Controller(backend,Repository(directory))
        if start: controller.start()
        return controller

    def tiles(self, controller, space=0):
        return [(a.window_id,a.pinned) if a else None for a in controller.profile('d1',space).assignments]

    def test_windows_that_stayed_open_return_to_their_tiles(self):
        for pins in ((False,False),(True,True),(True,False),(False,True)):
            # Quit, or ended without a goodbye.
            for crash in (False,True):
                with self.subTest(pins=pins,crash=crash),tempfile.TemporaryDirectory() as directory:
                    backend,controller=self.make(directory,pins)
                    if not crash: controller.quit()
                    self.assertEqual(self.tiles(self.restart(backend,directory)),[(2,pins[0]),(1,pins[1])])

    def test_a_restart_without_start_forgets_nothing(self):
        for quits in (False,True):
            with self.subTest(quits=quits),tempfile.TemporaryDirectory() as directory:
                backend,controller=self.make(directory);controller.quit()
                for _ in range(2):
                    idle=self.restart(backend,directory,start=False)
                    self.assertEqual(self.tiles(idle),[(2,False),(1,False)])
                    if quits: idle.quit()
                self.assertEqual(self.tiles(self.restart(backend,directory)),[(2,False),(1,False)])

    def test_the_session_is_not_emptied_before_the_windows_are_found(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory);controller.quit()
            before=(Path(directory)/'session.json').read_text()
            bind=Controller._bind_restored_profiles
            try:
                def crash(*args): raise RuntimeError('crash')
                Controller._bind_restored_profiles=crash
                with self.assertRaises(RuntimeError): Controller(backend,Repository(directory))
            finally: Controller._bind_restored_profiles=bind
            self.assertEqual((Path(directory)/'session.json').read_text(),before)
            self.assertEqual(self.tiles(self.restart(backend,directory)),[(2,False),(1,False)])

    def test_a_window_closed_meanwhile_does_not_move_the_other(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,(True,True));controller.quit()
            backend.remove_window(2)
            self.assertEqual(self.tiles(self.restart(backend,directory)),[(None,True),(1,True)])

    def test_another_window_with_the_same_number_is_not_taken_for_it(self):
        for change in (dict(ref=WindowRef(2,77,'generation:2')),dict(ref=WindowRef(2,2,'other:2')),dict(app_id='note')):
            with self.subTest(change=change),tempfile.TemporaryDirectory() as directory:
                backend,controller=self.make(directory,(True,True));controller.quit()
                backend.windows[2]=replace(backend.windows[2],**change)
                restarted=Controller(backend,Repository(directory))
                self.assertFalse(restarted._restored_hidden)
                # The ordinary rule: window 1 keeps its own tile.
                self.assertEqual(self.tiles(restarted)[1],(1,True))
                self.assertIn(self.tiles(restarted)[0],[(None,True),(2,True)])
                self.assertEqual(self.tiles(restarted)[0][0],None if 'app_id' in change else 2)

    def test_windows_on_another_display_are_not_brought_back(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,(True,True),displays=2);controller.quit()
            backend.windows[2]=replace(backend.windows[2],display_id='d2')
            self.assertEqual(self.tiles(self.restart(backend,directory)),[(None,True),(1,True)])
            self.assertEqual(backend.windows[2].display_id,'d2')

    def test_windows_of_a_hidden_space_are_put_aside_again(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            controller.switch_space('d1',1)
            controller.profile('d1',1).assignments=[Assignment('chrome',1,True)]
            controller.arrange('d1')
            controller.quit()
            # Quitting gave every window back.
            self.assertEqual({w.state for w in backend.windows.values()},{'normal'})
            restarted=self.restart(backend,directory,start=False)
            self.assertEqual((self.tiles(restarted),self.tiles(restarted,1)),([(2,False),(1,False)],[(1,True)]))
            self.assertEqual(backend.windows[2].state,'normal')
            restarted.start()
            self.assertEqual((backend.windows[1].state,backend.windows[2].state),('normal','minimized'))
            self.assertIn(2,restarted._parked)
            restarted.switch_space('d1',0)
            self.assertEqual({w.state for w in backend.windows.values()},{'normal'})
            self.assertEqual(self.tiles(restarted),[(2,False),(1,False)])
            restarted.quit()

    def test_no_title_is_stored_and_idle_scans_do_not_write(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            backend.windows[1]=replace(backend.windows[1],title='Secret');controller.refresh()
            session=Path(directory)/'session.json'
            self.assertNotIn('Secret',session.read_text())
            self.assertEqual([e['hwnd'] for e in json.loads(session.read_text())['windows']],[2,1])
            session.unlink()
            for _ in range(20): controller.refresh()
            self.assertFalse(session.exists())

    def test_sessions_of_older_versions_and_invalid_entries_are_read(self):
        good=dict(display_id='d1',space=0,index=0,app_id='chrome',hwnd=2,pid=2,process='generation')
        for windows,kept in ((None,[]),('x',[]),([1,{},dict(good,hwnd='2'),dict(good,index=-1),dict(good,space=3),dict(good,process='')],[]),([good],[good])):
            with tempfile.TemporaryDirectory() as directory:
                data={'version':1,'pending':[],'active_spaces':{'d1':1}}
                if windows is not None: data['windows']=windows
                Repository(directory).save_session([],{})
                stored=json.loads((Path(directory)/'session.json').read_text())
                data['version']=stored['version']
                (Path(directory)/'session.json').write_text(json.dumps(data))
                session=Repository(directory).load_session()
                self.assertEqual((session['windows'],session['active_spaces']),(kept,{'d1':1}))

    def test_a_process_without_creation_date_is_not_recorded(self):
        self.assertEqual(Controller._process_key(WindowRef(1,42,'42:1337:5')),'42:1337')
        self.assertEqual(Controller._process_key(WindowRef(1,42,'42::5')),'')


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
            # Another window of the same app tiled in this space: no window for the pin.
            profile.assignments[2]=Assignment('other',None,True)
            self.assertEqual({s['app_id']:s['state'] for s in controller.reserved_slots()}['other'],'NO WINDOW')
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
            self.assertEqual([(s['index'],s['state']) for s in controller.reserved_slots()],[(1,'NO WINDOW')])
            # Its card opens a new window for that tile; the sibling stays in place.
            sibling=assignments[0].window_id
            controller.activate_reserved('d1',0,1)
            assignments=controller.profile('d1',0).assignments
            self.assertEqual(backend.launches,['chrome'])
            backend.add_window('chrome',hwnd=20);controller.refresh()
            assignments=controller.profile('d1',0).assignments
            self.assertEqual(assignments[0].window_id,sibling)
            self.assertNotIn(assignments[1].window_id,(None,sibling))

    def single_instance(self, directory, pinned=True):
        """Two chrome tiles restored with one chrome window open; the launch
        for the second tile opens no new window."""
        display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
        backend=FakeBackend([display],[record(1,'chrome')],[ApplicationRef('chrome','Chrome')])
        backend.auto_launch=False
        controller=Controller(backend,Repository(directory))
        controller.restore_profile(SpaceProfile('d1',0,'2x2',assignments=[
            Assignment('chrome',None,True),Assignment('chrome',None,pinned)]))
        self.assertEqual(backend.launches,['chrome'])
        controller.refresh()
        return backend,controller

    def elapse(self, controller, seconds):
        for pending in controller._pending.values():
            pending['created']-=seconds
            if 'open_since' in pending: pending['open_since']-=seconds
        controller.refresh()

    def test_launch_of_an_already_open_application_expires_to_a_closed_card(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.single_instance(directory)
            self.elapse(controller,29)
            self.assertEqual([(s['index'],s['state']) for s in controller.reserved_slots()],[(1,'OPENING')])
            notified=[];controller.subscribe(lambda:notified.append(1))
            self.elapse(controller,2)
            self.assertEqual((controller._pending,controller._cancelled_launches),({},[]))
            self.assertEqual([(s['index'],s['state']) for s in controller.reserved_slots()],[(1,'NO WINDOW')])
            self.assertTrue(notified)
            self.assertEqual(controller.last_error,'')
            self.assertEqual(Repository(directory).load_session()['pending'],[])
            # The card still opens the application.
            controller.activate_reserved('d1',0,1)
            self.assertEqual(backend.launches,['chrome','chrome'])
            self.assertTrue(controller._pending)

    def test_late_window_fills_the_pinned_tile_of_an_expired_launch(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.single_instance(directory)
            self.elapse(controller,31);self.assertFalse(controller._pending)
            late=backend.add_window('chrome',hwnd=20).ref.hwnd;controller.refresh()
            self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments[:2]],[1,late])
            self.assertNotIn(late,controller._floating)

    def test_expired_launch_releases_an_unpinned_tile(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.single_instance(directory,pinned=False)
            self.elapse(controller,119);self.assertTrue(controller._pending)
            self.elapse(controller,2)
            self.assertFalse(controller._pending)
            self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments if a],[1])
            self.assertEqual(controller.last_error,'')
            late=backend.add_window('chrome',hwnd=20).ref.hwnd;controller.refresh()
            self.assertIn(late,[a.window_id for a in controller.profile('d1',0).assignments if a])

    def test_launch_expires_when_the_open_window_floats(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.single_instance(directory)
            controller._floating.add(1);controller.refresh()
            self.elapse(controller,31)
            self.assertFalse(controller._pending)

    def test_launch_does_not_expire_while_a_window_is_held(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.single_instance(directory,pinned=False)
            controller._interacting.add(1);self.elapse(controller,121)
            self.assertTrue(controller._pending)
            controller._interacting.clear();controller.refresh()
            self.assertFalse(controller._pending)

    def test_launch_without_any_open_window_does_not_expire(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.single_instance(directory)
            backend.remove_window(1);controller.refresh()
            self.elapse(controller,121)
            self.assertTrue(controller._pending)
            self.assertIn('Chrome',controller.last_error)

    def test_each_new_window_restarts_the_wait_of_the_other_launches(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[],[ApplicationRef('chrome','Chrome')]);backend.auto_launch=False
            controller=Controller(backend,Repository(directory))
            controller.restore_profile(SpaceProfile('d1',0,'2x2',assignments=[Assignment('chrome',None,True) for _ in range(3)]))
            self.elapse(controller,60);self.assertEqual(len(controller._pending),3)
            for hwnd,left in ((20,2),(21,1)):
                backend.add_window('chrome',hwnd=hwnd);controller.refresh()
                self.elapse(controller,20)
                self.assertEqual(len(controller._pending),left)
            self.elapse(controller,11);self.assertFalse(controller._pending)


    def test_slow_window_keeps_its_unpinned_place(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1,'chrome'),record(2,'notepad')],[ApplicationRef('chrome','Chrome')])
            backend.auto_launch=False
            controller=Controller(backend,Repository(directory))
            controller.restore_profile(SpaceProfile('d1',0,'2x2',assignments=[
                Assignment('chrome'),Assignment('chrome'),Assignment('notepad')]))
            controller.refresh();self.elapse(controller,60)
            late=backend.add_window('chrome',hwnd=20).ref.hwnd;controller.refresh()
            self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments if a],[1,late,2])

    def test_no_slow_launch_warning_while_the_application_has_a_window(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[],[ApplicationRef('chrome','Chrome')]);backend.auto_launch=False
            controller=Controller(backend,Repository(directory))
            controller.restore_profile(SpaceProfile('d1',0,'2x2',assignments=[Assignment('chrome') for _ in range(3)]))
            self.elapse(controller,100)
            backend.add_window('chrome',hwnd=20);controller.refresh()
            self.elapse(controller,25)
            self.assertEqual((len(controller._pending),controller.last_error),(2,''))
            # Its only window closes: it is waited for again, with the warning.
            backend.remove_window(20);controller.refresh()
            self.assertIn('Chrome',controller.last_error)


class CrossDisplayDropTests(unittest.TestCase):
    def make(self, directory, assignments, preset='auto', tiles=(), windows=()):
        displays=[Display(f'd{i+1}',f'D{i+1}',Rect(1920*i,0,1920,1080),primary=not i) for i in range(2)]
        backend=FakeBackend(displays,[record(1,'app','d2'),*windows],[ApplicationRef('app','App'),ApplicationRef('other','Other')])
        backend.auto_launch=False
        controller=Controller(backend,Repository(directory))
        controller.profiles[('d1',0)]=SpaceProfile('d1',0,preset,list(tiles),assignments)
        controller.start()
        return backend,controller

    def drop(self, backend, controller, index, hwnd=1):
        rect=controller.resolved_rects('d1',0)[index]
        controller._handle_event({'type':'move_start','hwnd':hwnd})
        backend.cursor=(rect.x+rect.width//2,rect.y+rect.height//2)
        backend.windows[hwnd]=replace(backend.windows[hwnd],rect=Rect(rect.x+5,rect.y+5,640,480),display_id='d1')
        controller._handle_event({'type':'move_end','hwnd':hwnd});controller.refresh()

    def tiles(self, controller, display='d1'):
        return [(a.app_id,a.window_id,a.pinned) if a else None for a in controller.profile(display,0).assignments]

    def test_the_empty_pinned_tile_is_filled(self):
        # Aimed at, or elsewhere on its display.
        for aimed in (1,0):
            with self.subTest(aimed=aimed),tempfile.TemporaryDirectory() as directory:
                backend,controller=self.make(directory,[Assignment('note'),Assignment('app',None,True)],windows=[record(2,'note')])
                self.drop(backend,controller,aimed)
                self.assertEqual(self.tiles(controller),[('note',2,False),('app',1,True)])
                self.assertEqual((self.tiles(controller,'d2'),controller.reserved_slots()),([],[]))
                self.assertEqual(backend.windows[1].rect,controller.resolved_rects('d1',0)[1])

    def test_the_pinned_tile_aimed_at_is_chosen(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[Assignment('app',None,True) for _ in range(3)])
            self.drop(backend,controller,1)
            self.assertEqual([t[1] for t in self.tiles(controller)],[None,1,None])

    def test_a_tile_reserved_for_another_app_is_not_taken(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[Assignment('other',None,True)])
            self.drop(backend,controller,0)
            self.assertEqual(self.tiles(controller),[('other',None,True),('app',1,False)])

    def test_the_empty_custom_tile_aimed_at_is_filled(self):
        with tempfile.TemporaryDirectory() as directory:
            thirds=[Tile(str(i),i/3,0,1/3,1) for i in range(3)]
            backend,controller=self.make(directory,[None,None,None],'custom',thirds)
            self.drop(backend,controller,1)
            self.assertEqual(self.tiles(controller),[None,('app',1,False),None])

    def test_the_launch_of_the_filled_tile_is_forgotten(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[Assignment('app',None,True)])
            controller.activate_reserved('d1',0,0);self.assertTrue(controller._pending)
            self.drop(backend,controller,0)
            self.assertFalse(controller._pending)
            late=backend.add_window('app',hwnd=20).ref.hwnd;controller.refresh()
            self.assertEqual([t[1] for t in self.tiles(controller)],[1,late])

    def test_the_drop_is_undone(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[Assignment('app',None,True)])
            self.drop(backend,controller,0)
            controller.undo()
            self.assertEqual(self.tiles(controller),[('app',None,True)])
            self.assertEqual([t[1] for t in self.tiles(controller,'d2')],[1])


class SameDisplayDropTests(unittest.TestCase):
    def make(self, directory, assignments, preset='auto', tiles=()):
        display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
        backend=FakeBackend([display],[record(1,'app'),record(2,'note'),record(3,'note')],
                            [ApplicationRef('app','App'),ApplicationRef('other','Other')])
        backend.auto_launch=False
        controller=Controller(backend,Repository(directory))
        controller.profiles[('d1',0)]=SpaceProfile('d1',0,preset,list(tiles),assignments)
        controller.start()
        return backend,controller

    def drop(self, backend, controller, index, hwnd=1):
        before=backend.windows[hwnd].rect
        backend.cursor=(before.x+before.width//2,before.y+20)
        rect=controller.resolved_rects('d1',0)[index]
        controller._handle_event({'type':'move_start','hwnd':hwnd})
        backend.cursor=(rect.x+rect.width//2,rect.y+rect.height//2)
        backend.windows[hwnd]=replace(backend.windows[hwnd],rect=Rect(rect.x+5,rect.y+5,before.width,before.height))
        controller._handle_event({'type':'move_end','hwnd':hwnd});controller.refresh()

    def tiles(self, controller):
        return [(a.app_id,a.window_id,a.pinned) if a else None for a in controller.profile('d1',0).assignments]

    def test_an_unpinned_window_takes_the_pin_and_leaves_no_hole(self):
        for assignments,aimed,expected in (
                ([Assignment('app',1),Assignment('note',2),Assignment('note',3),Assignment('app',None,True)],3,
                 [('note',2,False),('note',3,False),('app',1,True)]),
                ([Assignment('app',None,True),Assignment('note',2),Assignment('app',1),Assignment('note',3)],0,
                 [('app',1,True),('note',2,False),('note',3,False)])):
            with self.subTest(aimed=aimed),tempfile.TemporaryDirectory() as directory:
                backend,controller=self.make(directory,assignments)
                self.drop(backend,controller,aimed)
                self.assertEqual(self.tiles(controller),expected)
                self.assertFalse(controller.reserved_slots())
                self.assertEqual(len(controller.resolved_rects('d1',0)),3)

    def test_a_custom_layout_keeps_the_tile_left_empty(self):
        with tempfile.TemporaryDirectory() as directory:
            quarters=[Tile(str(i),i/4,0,1/4,1) for i in range(4)]
            backend,controller=self.make(directory,
                [Assignment('app',1),Assignment('note',2),Assignment('note',3),Assignment('app',None,True)],'custom',quarters)
            self.drop(backend,controller,3)
            self.assertEqual(self.tiles(controller),[None,('note',2,False),('note',3,False),('app',1,True)])
            self.assertEqual(backend.windows[1].rect,controller.resolved_rects('d1',0)[3])

    def test_a_pinned_window_or_the_pin_of_another_app_is_swapped(self):
        for first,pin in ((Assignment('app',1,True),Assignment('app',None,True)),(Assignment('app',1),Assignment('other',None,True))):
            with self.subTest(pin=pin.app_id),tempfile.TemporaryDirectory() as directory:
                backend,controller=self.make(directory,[first,Assignment('note',2),Assignment('note',3),pin])
                self.drop(backend,controller,3)
                self.assertEqual(self.tiles(controller),[(pin.app_id,None,True),('note',2,False),('note',3,False),('app',1,first.pinned)])

    def test_the_drop_is_undone(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,
                [Assignment('app',1),Assignment('note',2),Assignment('note',3),Assignment('app',None,True)])
            before=self.tiles(controller)
            self.drop(backend,controller,3);self.assertNotEqual(self.tiles(controller),before)
            controller.undo();self.assertEqual(self.tiles(controller),before)

    def test_a_launch_follows_the_tile_swapped_while_it_opens(self):
        for first,waited in ((Assignment('app',1,True),'app'),(Assignment('app',1),'other')):
            with self.subTest(waited=waited),tempfile.TemporaryDirectory() as directory:
                backend,controller=self.make(directory,[first,Assignment('note',2),Assignment('note',3),Assignment(waited,None,True)])
                controller.activate_reserved('d1',0,3)
                self.drop(backend,controller,3)
                self.assertEqual([k[2] for k in controller._pending],[0])
                self.assertEqual([(s['index'],s['state']) for s in controller.reserved_slots()],[(0,'OPENING')])
                late=backend.add_window(waited,hwnd=20).ref.hwnd;controller.refresh()
                self.assertEqual([t[1] for t in self.tiles(controller)],[late,2,3,1])
                self.assertFalse(controller._pending)

    def test_launches_follow_their_tiles(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[Assignment('app',1),Assignment('note',2),Assignment('note',3),
                Assignment('app',None,True),Assignment('other',None,True)])
            controller.activate_reserved('d1',0,3);controller.activate_reserved('d1',0,4)
            self.assertEqual(sorted(k[2] for k in controller._pending),[3,4])
            self.drop(backend,controller,3)
            self.assertEqual({k[2]:p['app_id'] for k,p in controller._pending.items()},{3:'other'})
            late=backend.add_window('other',hwnd=20).ref.hwnd;controller.refresh()
            self.assertEqual(self.tiles(controller)[2:],[('app',1,True),('other',late,True)])


class ReplacedArrangementLaunchTests(unittest.TestCase):
    """An arrangement replaced while its applications open."""
    def make(self, directory, count=1):
        display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
        backend=FakeBackend([display],[record(1,'note')],[ApplicationRef('app','App'),ApplicationRef('other','Other')])
        backend.auto_launch=False
        controller=Controller(backend,Repository(directory))
        controller.profiles[('d1',0)]=SpaceProfile('d1',0,assignments=[Assignment('note',1)]+[Assignment('app',None,True) for _ in range(count)])
        controller.start()
        for i in range(count): controller.activate_reserved('d1',0,i+1)
        self.assertEqual(backend.launches,['app']*count)
        return backend,controller

    def tiles(self, controller):
        return [(a.app_id,a.window_id) if a else None for a in controller.profile('d1',0).assignments]

    def arrive(self, backend, controller, hwnd=20, app='app'):
        backend.add_window(app,hwnd=hwnd);controller.refresh()
        self.assertNotIn(hwnd,controller._floating)

    def test_apply_keeps_the_launch_of_a_moved_tile(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            controller.swap_draft('d1',0,0,1);controller.apply_drafts()
            self.assertEqual(backend.launches,['app'])
            self.assertEqual([(s['index'],s['state']) for s in controller.reserved_slots()],[(0,'OPENING')])
            self.arrive(backend,controller)
            self.assertEqual((self.tiles(controller),controller._pending),([('app',20),('note',1)],{}))

    def test_apply_keeps_a_window_that_arrived_during_the_edit(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            controller.swap_draft('d1',0,0,1)
            self.arrive(backend,controller)
            controller.apply_drafts()
            self.assertEqual((self.tiles(controller),backend.launches),([('app',20),('note',1)],['app']))
            self.assertFalse(controller._floating)

    def test_a_window_whose_tile_was_cleared_does_not_take_another_tile(self):
        # Tile 2 is empty before the edit, or its window arrives during it.
        for arrived in (False,True):
            with self.subTest(arrived=arrived),tempfile.TemporaryDirectory() as directory:
                display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
                backend=FakeBackend([display],[record(1,'app')],[ApplicationRef('app','App')]);backend.auto_launch=False
                controller=Controller(backend,Repository(directory))
                controller.profiles[('d1',0)]=SpaceProfile('d1',0,assignments=[Assignment('app',1),Assignment('app',None,True)])
                controller.start()
                if arrived: controller.activate_reserved('d1',0,1)
                controller.set_assignment('d1',0,0,None)
                if arrived: self.arrive(backend,controller)
                controller.apply_drafts()
                self.assertEqual(self.tiles(controller),[None,('app',20 if arrived else None)])
                self.assertEqual(controller._floating,{1})

    def opening(self, directory, displays=1):
        """On each display, tile 1 holds a window of the app, tile 2 opens one."""
        shown=[Display(f'd{i+1}',f'D{i+1}',Rect(1920*i,0,1920,1080),primary=not i) for i in range(displays)]
        backend=FakeBackend(shown,[record(i+1,'app',d.id) for i,d in enumerate(shown)],[ApplicationRef('app','App')])
        backend.auto_launch=False
        controller=Controller(backend,Repository(directory))
        for i,d in enumerate(shown):
            controller.profiles[(d.id,0)]=SpaceProfile(d.id,0,assignments=[Assignment('app',i+1),Assignment('app',None,True)])
        controller.start()
        for d in shown: controller.activate_reserved(d.id,0,1)
        return backend,controller

    def test_a_cleared_window_stays_out_when_the_awaited_tile_is_moved(self):
        for clears_first in (True,False):
            with self.subTest(clears_first=clears_first),tempfile.TemporaryDirectory() as directory:
                backend,controller=self.opening(directory)
                if clears_first:
                    controller.set_assignment('d1',0,0,None);controller.swap_draft('d1',0,0,1)
                else:
                    controller.swap_draft('d1',0,0,1);controller.set_assignment('d1',0,1,None)
                self.arrive(backend,controller)
                controller.apply_drafts()
                self.assertEqual([t for t in self.tiles(controller) if t],[('app',20)])
                self.assertEqual((controller._floating,backend.launches),({1},['app']))

    def test_edits_on_two_displays_are_applied_together(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.opening(directory,2)
            for d in ('d1','d2'):
                controller.set_assignment(d,0,0,None);controller.swap_draft(d,0,0,1)
            self.arrive(backend,controller,20);self.arrive(backend,controller,21)
            controller.apply_drafts()
            held=[[a.window_id for a in controller.profile(d,0).assignments if a] for d in ('d1','d2')]
            self.assertEqual((sorted(sum(held,[])),controller._floating),([20,21],{1,2}))
            self.assertEqual([len(h) for h in held],[1,1])

    def test_an_edit_undone_then_made_again_starts_from_the_current_windows(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.opening(directory)
            controller.set_assignment('d1',0,0,None);controller.undo_draft('d1',0)
            self.assertFalse(controller.draft('d1',0).dirty)
            self.arrive(backend,controller)
            # Both windows are in the arrangement the second edit starts from.
            controller.set_assignment('d1',0,1,None)
            controller.apply_drafts()
            self.assertEqual([t for t in self.tiles(controller) if t],[('app',1)])
            self.assertEqual(controller._floating,{20})

    def test_the_window_of_a_cleared_awaited_tile_floats(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,2)
            controller.set_assignment('d1',0,1,None)
            self.arrive(backend,controller)
            controller.apply_drafts()
            self.assertEqual(self.tiles(controller)[1:],[None,('app',None)])
            self.assertEqual(controller._floating,{20})

    def test_apply_launches_only_the_tiles_added(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,2)
            for i in (3,4): controller.set_assignment('d1',0,i,Assignment('app' if i==3 else 'other'))
            controller.apply_drafts()
            self.assertEqual(backend.launches,['app','app','app','other'])
            self.assertEqual(sorted(k[2] for k in controller._pending),[1,2,3,4])

    def test_apply_cancels_the_launch_of_a_cleared_tile(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            controller.set_assignment('d1',0,1,None);controller.apply_drafts()
            self.assertFalse(controller._pending)
            backend.add_window('app',hwnd=20);controller.refresh()
            self.assertIn(20,controller._floating)

    def test_apply_launches_again_an_application_that_never_came(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            for pending in controller._pending.values(): pending['created']-=121
            controller.swap_draft('d1',0,0,1);controller.apply_drafts()
            self.assertEqual(backend.launches,['app','app'])

    def test_apply_of_a_hidden_space_keeps_its_launch(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            controller.switch_space('d1',1)
            controller.swap_draft('d1',0,0,1);controller.apply_drafts([('d1',0)])
            self.assertEqual(([k for k in controller._pending],backend.launches),([('d1',0,0)],['app']))

    def test_opening_tiles_of_a_draft_are_those_that_wait_after_apply(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,2)
            opening=lambda:controller.opening_slots('d1',0,controller.draft('d1',0).profile)
            self.assertEqual(opening(),{1,2})
            controller.swap_draft('d1',0,0,1);self.assertEqual(opening(),{0,2})
            controller.set_assignment('d1',0,2,None);self.assertEqual(opening(),{0})
            controller.set_assignment('d1',0,2,Assignment('other'));self.assertEqual(opening(),{0})
            self.assertEqual(controller.opening_slots('d1',1,controller.draft('d1',0).profile),set())
            controller.apply_drafts()
            self.assertEqual({k[2] for k,p in controller._pending.items() if p['app_id']=='app'},{0})

    def test_the_launch_signature_follows_the_launches(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            self.assertEqual(controller.launch_signature('d1',0),((1,'app'),))
            self.assertEqual(controller.launch_signature('d1',1),())
            for pending in controller._pending.values(): pending['created']-=121
            self.assertEqual(controller.launch_signature('d1',0),())
            for pending in controller._pending.values(): pending['created']+=121
            self.arrive(backend,controller)
            self.assertEqual(controller.launch_signature('d1',0),())

    def test_the_preview_does_not_count_launches_in_progress(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            controller.templates=[LayoutTemplate('t','T','2x2',assignments=[
                Assignment('app'),Assignment('note'),Assignment('other'),Assignment('gone')])]
            counts=lambda space=0:{k:controller.preview_template('t','d1',space)[k] for k in ('reuse','open','unavailable')}
            self.assertEqual(counts(),dict(reuse=1,open=1,unavailable=1))
            # Not for another space, nor for a launch that Restore starts again.
            self.assertEqual(counts(1)['open'],2)
            controller.templates[0].assignments[3]=Assignment('app')
            self.assertEqual(counts()['open'],2)
            for pending in controller._pending.values(): pending['created']-=121
            self.assertEqual(counts()['open'],3)
            for pending in controller._pending.values(): pending['created']+=121
            launched=len(backend.launches);planned=counts()['open']
            controller.restore_template('t','d1',0)
            self.assertEqual(len(backend.launches)-launched,planned)

    def test_a_layout_restored_twice_launches_once(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory)
            layout=SpaceProfile('d1',0,'2x2',assignments=[Assignment('app'),Assignment('note'),Assignment('other')])
            for _ in range(2): controller.restore_profile(layout)
            self.assertEqual(backend.launches,['app','other'])
            self.arrive(backend,controller);self.arrive(backend,controller,21,'other')
            self.assertEqual((self.tiles(controller),controller._pending),([('app',20),('note',1),('other',21)],{}))


class CrossDisplayGuideTests(unittest.TestCase):
    def test_the_guide_shows_what_the_drop_does(self):
        for assignments,aimed,label,shown in (
                ([Assignment('note',2),Assignment('app',None,True)],1,'Move here',1),
                ([Assignment('note',2),Assignment('app',None,True)],0,'Move here',1),
                ([Assignment('note',2)],0,'Insert',None),
                ([Assignment('note',2),Assignment('other',None,True)],1,'Insert',None),
                ([],None,'Insert',None)):
            with self.subTest(label=label,aimed=aimed),tempfile.TemporaryDirectory() as directory:
                displays=[Display(f'd{i+1}',f'D{i+1}',Rect(1920*i,0,1920,1080),primary=not i) for i in range(2)]
                windows=[record(1,'app','d2')]+([record(2,'note')] if assignments else [])
                backend=FakeBackend(displays,windows,[ApplicationRef('app','App')]);backend.auto_launch=False
                controller=Controller(backend,Repository(directory))
                controller.profiles[('d1',0)]=SpaceProfile('d1',0,assignments=deepcopy(assignments))
                controller.start()
                rects=controller.resolved_rects('d1',0)
                aim=rects[aimed] if aimed is not None else Rect(100,100,800,600)
                held=backend.windows[1].rect
                backend.cursor=(held.x+held.width//2,held.y+20)
                controller._handle_event({'type':'move_start','hwnd':1})
                backend.cursor=(aim.x+aim.width//2,aim.y+aim.height//2)
                backend.windows[1]=replace(backend.windows[1],rect=Rect(aim.x+5,aim.y+5,640,480),display_id='d1')
                controller._handle_event({'type':'location','hwnd':1})
                guide=controller.drag_guide
                self.assertEqual((guide[0],guide[2]),('d1',label))
                if shown is not None: self.assertEqual(guide[1],rects[shown])
                elif rects: self.assertEqual((guide[1].x,guide[1].right),(min(r.x for r in rects),max(r.right for r in rects)))
                else: self.assertEqual(guide[1],displays[0].work_area)
                controller._handle_event({'type':'move_end','hwnd':1});controller.refresh()
                held=next(i for i,a in enumerate(controller.profile('d1',0).assignments) if a and a.window_id==1)
                self.assertEqual(held,shown if shown is not None else len(assignments))


class PinnedCardTests(unittest.TestCase):
    def make(self, directory, windows, assignments, displays=1):
        shown=[Display(f'd{i+1}',f'D{i+1}',Rect(1920*i,0,1920,1080),primary=not i) for i in range(displays)]
        backend=FakeBackend(shown,windows,[ApplicationRef('chrome','Chrome'),ApplicationRef('notepad','Notepad')])
        controller=Controller(backend,Repository(directory))
        controller.profiles[('d1',0)]=SpaceProfile('d1',0,assignments=assignments)
        controller.start()
        return backend,controller

    def states(self, controller):
        return [(s['index'],s['state']) for s in controller.reserved_slots() if s['display_id']=='d1']

    def test_a_click_opens_only_its_tile(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[],[Assignment(a,None,True) for a in ('chrome','notepad','chrome')])
            self.assertEqual(backend.launches,[])
            controller.activate_reserved('d1',0,2)
            self.assertEqual(backend.launches,['chrome'])
            controller.refresh()
            assignments=controller.profile('d1',0).assignments
            self.assertEqual([bool(a.window_id) for a in assignments],[False,False,True])
            self.assertEqual(self.states(controller),[(0,'NO WINDOW'),(1,'CLOSED')])

    def test_restoring_a_minimized_or_floating_window_opens_nothing(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[record(1,'chrome')],[Assignment('chrome',None,True),Assignment('notepad',None,True)])
            backend.minimize(1);controller.refresh()
            self.assertEqual(self.states(controller),[(0,'MINIMIZED'),(1,'CLOSED')])
            controller.activate_reserved('d1',0,0);controller.refresh()
            self.assertEqual((backend.launches,backend.windows[1].state),([],'normal'))
            controller.toggle_float(1)
            self.assertEqual(self.states(controller),[(0,'FLOATING'),(1,'CLOSED')])
            controller.activate_reserved('d1',0,0);controller.refresh()
            self.assertEqual(backend.launches,[])
            self.assertEqual(self.states(controller),[(1,'CLOSED')])

    def test_a_click_on_the_card_of_a_cleared_tile_does_nothing(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[],[None,Assignment('chrome',None,True)])
            undo=controller.history.undo_count
            for index in (0,5): controller.activate_reserved('d1',0,index)
            self.assertEqual((backend.launches,controller.history.undo_count),([],undo))

    def test_open_elsewhere_on_another_display_opens_one_window_here(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[record(1,'chrome','d2')],
                [Assignment('chrome',None,True),Assignment('notepad',None,True)],displays=2)
            self.assertEqual(self.states(controller),[(0,'OPEN ELSEWHERE'),(1,'CLOSED')])
            controller.activate_reserved('d1',0,0);controller.refresh()
            self.assertEqual(backend.launches,['chrome'])
            self.assertNotIn(controller.profile('d1',0).assignments[0].window_id,(None,1))
            self.assertEqual(backend.windows[1].display_id,'d2')

    def test_a_window_of_another_space_is_shown_by_one_click_in_the_clicked_tile(self):
        with tempfile.TemporaryDirectory() as directory:
            backend,controller=self.make(directory,[record(1,'chrome')],[])
            controller.profiles[('d1',1)]=SpaceProfile('d1',1)
            controller.active_spaces['d1']=0
            controller.switch_space('d1',1);controller.refresh(auto=False)
            self.assertEqual(backend.windows[1].state,'minimized')
            controller.profile('d1',1).assignments=[Assignment('chrome',None,True),Assignment('chrome',None,True)]
            self.assertEqual(self.states(controller),[(0,'OPEN ELSEWHERE'),(1,'OPEN ELSEWHERE')])
            controller.activate_reserved('d1',1,1);controller.refresh()
            self.assertEqual(backend.launches,[])
            self.assertEqual([a.window_id for a in controller.profile('d1',1).assignments],[None,1])
            self.assertEqual(backend.windows[1].state,'normal')

    def test_restart_serves_the_pins_of_the_active_space_first(self):
        for count,expected in ((0,[None]),(1,[1]),(2,[1,2])):
            with self.subTest(count=count),tempfile.TemporaryDirectory() as directory:
                Repository(directory).save_layouts([],{('d1',0):SpaceProfile('d1',0,assignments=[
                    Assignment('chrome'),Assignment('chrome',None,True)])})
                display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
                backend=FakeBackend([display],[record(i,'chrome') for i in range(1,count+1)],[ApplicationRef('chrome','Chrome')])
                controller=Controller(backend,Repository(directory));controller.start()
                assignments=[a for a in controller.profile('d1',0).assignments if a]
                pin=next(a for a in assignments if a.pinned)
                self.assertEqual(pin.window_id,1 if count else None)
                self.assertEqual(sorted([a.window_id for a in assignments],key=lambda v:v or 0),expected)

    def test_restart_pin_of_a_hidden_space_takes_nothing(self):
        with tempfile.TemporaryDirectory() as directory:
            Repository(directory).save_layouts([],{
                ('d1',0):SpaceProfile('d1',0,assignments=[Assignment('chrome')]),
                ('d1',1):SpaceProfile('d1',1,assignments=[Assignment('chrome',None,True)])})
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1,'chrome')],[ApplicationRef('chrome','Chrome')])
            controller=Controller(backend,Repository(directory));controller.start()
            self.assertEqual(controller.profile('d1',0).assignments[0].window_id,1)
            self.assertIsNone(controller.profile('d1',1).assignments[0].window_id)

    def test_restart_keeps_a_window_for_the_pending_launch_of_a_pin(self):
        with tempfile.TemporaryDirectory() as directory:
            repository=Repository(directory)
            repository.save_layouts([],{('d1',0):SpaceProfile('d1',0,assignments=[
                Assignment('chrome',None,True),Assignment('chrome')])})
            repository.save_session([dict(display_id='d1',space=0,index=0,app_id='chrome',created_utc=time.time())],{'d1':0})
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1,'chrome')],[ApplicationRef('chrome','Chrome')])
            controller=Controller(backend,Repository(directory));controller.start()
            self.assertEqual([a.window_id for a in controller.profile('d1',0).assignments if a],[1])
            self.assertFalse(controller._pending)


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

    def test_every_move_reaches_the_palette_before_the_queue(self):
        with tempfile.TemporaryDirectory() as directory:
            controller,backend=self.make(directory)
            moved=[];controller.subscribe_follow(moved.append)
            for i in range(3): controller._enqueue_event({'type':'location','hwnd':1,'n':i})
            controller._enqueue_event({'type':'title','hwnd':1})
            self.assertEqual(moved,[1,1,1])


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


class HiddenCloseTests(unittest.TestCase):
    def test_a_window_hidden_before_it_closes_shrinks_the_grid(self):
        # Electron apps (Bruno) hide their window, then destroy it.
        from dataclasses import replace
        from smartgrid.core.geometry import effective_layout
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(h,f'app{h}') for h in range(1,16)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            for assignment in controller.profile('d1',0).assignments:
                assignment.pinned=True
            backend.add_window('bruno',hwnd=99);controller.refresh(auto=True)
            self.assertEqual(effective_layout(controller.profile('d1',0),runtime=True)[0],'4x4')
            backend.windows[99]=replace(backend.windows[99],eligible=False);controller.refresh(auto=True)
            backend.remove_window(99);controller.refresh(auto=True)
            self.assertEqual(effective_layout(controller.profile('d1',0),runtime=True)[0],'5x3')
            self.assertEqual(backend.windows[15].rect,controller.resolved_rects('d1',0)[14])


class FreeMoveTests(unittest.TestCase):
    def test_moving_a_floating_window_rescans_only_at_release(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            controller.toggle_float(2)
            scans=[];discover=backend.discover_windows
            backend.discover_windows=lambda: scans.append(1) or discover()
            controller._handle_event({'type':'move_start','hwnd':2})
            controller._last_reconcile=0.0
            controller._handle_event({'type':'location','hwnd':2})
            self.assertFalse(scans)
            controller._handle_event({'type':'move_end','hwnd':2})
            self.assertTrue(scans);self.assertFalse(controller._free_moving)


class NativeMovementTests(unittest.TestCase):
    def test_windows_jump_to_their_tile_even_with_a_retired_animation_setting(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            controller.settings.window_animations=True
            calls=[];place=backend.place
            backend.place=lambda hwnd,rect,**kwargs: calls.append(kwargs) or place(hwnd,rect,**kwargs)
            controller.arrange('d1')
            self.assertTrue(calls)
            self.assertFalse(any(c.get('animate') for c in calls))


class LiveMaximizeTests(unittest.TestCase):
    def test_a_window_maximized_since_the_last_scan_is_not_restored(self):
        from dataclasses import replace
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            backend.windows[1]=replace(backend.windows[1],state='maximized')
            backend.operations.clear()
            controller._reflow_all(include_new=True)
            self.assertEqual(backend.windows[1].state,'maximized')
            self.assertFalse(backend.operations)

    def test_a_window_maximized_during_an_arrangement_is_not_restored(self):
        from dataclasses import replace
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2),record(3)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            place=backend.place_result
            def placing(hwnd,rect,*args,**kwargs):
                # The user maximizes window 3 and minimizes window 2 while 1 is placed.
                if hwnd==1:
                    backend.windows[3]=replace(backend.windows[3],state='maximized')
                    backend.windows[2]=replace(backend.windows[2],state='minimized')
                return place(hwnd,rect,*args,**kwargs)
            backend.place_result=placing
            controller.profile('d1',0).resize_tiles=[]
            for hwnd in (1,2,3): backend.windows[hwnd]=replace(backend.windows[hwnd],rect=Rect(0,0,10,10))
            controller.refresh(auto=False)
            controller._reflow_all(include_new=True)
            self.assertEqual(backend.windows[3].state,'maximized')
            self.assertEqual(backend.windows[2].state,'minimized')


class EarlyCardTests(unittest.TestCase):
    def test_pinned_cards_are_known_before_the_windows_are_placed(self):
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(h,f'app{h}') for h in (1,2,3)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            for assignment in controller.profile('d1',0).assignments:
                assignment.pinned=True
            seen=[];place=backend.place_result
            controller.subscribe_layout(lambda: seen.append([s['state'] for s in controller.reserved_slots()]))
            backend.place_result=lambda *a,**k: (seen.append('place'),place(*a,**k))[1]
            backend.minimize(2);controller.refresh(auto=True)
            self.assertEqual(seen[0],['MINIMIZED'])


class PlacementGhostTests(unittest.TestCase):
    def test_a_dropped_window_glides_from_where_it_was_released(self):
        from dataclasses import replace
        with tempfile.TemporaryDirectory() as directory:
            display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
            backend=FakeBackend([display],[record(1),record(2)],[])
            controller=Controller(backend,Repository(directory));controller.start()
            profile=controller.profile('d1',0)
            tiles={a.window_id:r for a,r in zip(profile.assignments,controller.resolved_rects('d1',0))}
            released=Rect(1200,300,640,480)
            backend.windows[1]=replace(backend.windows[1],rect=released)
            controller.refresh(auto=False)
            guides=[];controller.subscribe_guides(lambda: guides.append(1))
            profile.assignments.reverse();controller._apply_profile(profile)
            starts={event[5]:event[1] for event in controller.motion_events}
            self.assertEqual(starts[1],released)
            self.assertEqual(starts[2],tiles[2])
            self.assertTrue(guides)


class AtomicSaveTests(unittest.TestCase):
    def test_a_briefly_locked_file_is_saved_after_a_retry(self):
        from unittest import mock
        from smartgrid.storage import repository
        with tempfile.TemporaryDirectory() as directory:
            path=Path(directory)/'layouts.json'
            replace=repository.os.replace;calls=[]
            def locked(source,target):
                calls.append(1)
                if len(calls)<3: raise PermissionError(5,'Access is denied')
                replace(source,target)
            with mock.patch.object(repository.os,'replace',locked):
                repository.atomic_json(path,{'a':1})
            self.assertEqual(json.loads(path.read_text()),{'a':1})
            self.assertEqual(len(calls),3)
            self.assertEqual([p.name for p in Path(directory).iterdir()],['layouts.json'])


class SlotGuardTests(unittest.TestCase):
    def make(self, directory):
        from dataclasses import replace
        display=Display('d1','D',Rect(0,0,1920,1080),primary=True)
        backend=FakeBackend([display],[record(1),record(2),record(3)],[])
        controller=Controller(backend,Repository(directory));controller.start()
        tile=controller._tile_rects[1][1]
        grow=lambda: backend.windows.__setitem__(1,replace(backend.windows[1],rect=Rect(tile.x,tile.y,tile.width+93,tile.height)))
        return controller,backend,tile,grow

    def test_a_window_growing_out_of_its_tile_is_clamped_back(self):
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            grow();controller._guard_slots()
            self.assertEqual(backend.windows[1].rect,tile)
            # At most every 0.22 s per window: no fight with a stubborn app.
            grow();controller._guard_slots()
            self.assertNotEqual(backend.windows[1].rect,tile)

    def test_the_guard_leaves_windows_alone_when_it_must(self):
        from dataclasses import replace
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            cases=[('settings','force_resize',False),('interacting',None,None),('floating',None,None),('maximized',None,None)]
            for case,_,_ in cases:
                controller._guard_fixed.clear();grow()
                if case=='settings': controller.settings.force_resize=False
                if case=='interacting': controller._interacting.add(2)
                if case=='floating': controller._floating.add(1)
                if case=='maximized': backend.windows[1]=replace(backend.windows[1],state='maximized')
                controller._guard_slots()
                self.assertNotEqual(backend.windows[1].rect,tile,case)
                controller.settings.force_resize=True;controller._interacting.clear();controller._floating.discard(1)
                backend.windows[1]=replace(backend.windows[1],state='normal')

    def test_small_differences_are_tolerated(self):
        from dataclasses import replace
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            nearly=Rect(tile.x,tile.y,tile.width+6,tile.height)
            backend.windows[1]=replace(backend.windows[1],rect=nearly)
            controller._guard_slots()
            self.assertEqual(backend.windows[1].rect,nearly)

    def test_a_move_event_clamps_at_once_and_a_stubborn_window_is_not_fought(self):
        import time as clock
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            grow();controller._handle_event({'type':'location','hwnd':1})
            self.assertEqual(backend.windows[1].rect,tile)
            # A window that keeps growing back is then corrected at most every 0.22 s.
            now=clock.monotonic()
            controller._guard_fixed[1]=tuple(now-.1*i for i in range(8,0,-1))
            grow();controller._handle_event({'type':'location','hwnd':1})
            self.assertNotEqual(backend.windows[1].rect,tile)

    def test_a_move_is_clamped_before_the_queue_unless_the_runtime_is_busy(self):
        import threading
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            grow();controller._enqueue_event({'type':'location','hwnd':1})
            self.assertEqual(backend.windows[1].rect,tile)
            # An arrangement (another thread holding the runtime) is never interrupted.
            controller._guard_fixed.clear();grow()
            held,release=threading.Event(),threading.Event()
            def arrangement():
                with controller._lock:
                    held.set();release.wait(5)
            worker=threading.Thread(target=arrangement);worker.start();held.wait(5)
            controller._enqueue_event({'type':'location','hwnd':2})
            controller._enqueue_event({'type':'location','hwnd':1})
            self.assertNotEqual(backend.windows[1].rect,tile)
            release.set();worker.join(5)
            self.assertIn(1,controller._pending_locations)

    def test_a_grabbed_window_is_never_clamped_before_its_gesture_is_handled(self):
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            # Windows reports the start of the drag, then moves; the runtime
            # has not handled the start yet.
            controller._enqueue_event({'type':'move_start','hwnd':1})
            grow();controller._enqueue_event({'type':'location','hwnd':1})
            controller._guard_slots()
            self.assertNotEqual(backend.windows[1].rect,tile)
            # Drag and drop and swap still work as before.
            controller._handle_event({'type':'move_start','hwnd':1})
            controller._enqueue_event({'type':'move_end','hwnd':1})
            self.assertTrue(controller._interacting)
            controller._guard_slots()
            self.assertNotEqual(backend.windows[1].rect,tile)

    def test_undo_is_not_reverted_by_the_guard(self):
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            before={h:backend.windows[h].rect for h in (1,2)}
            controller.history.push(controller._snapshot())
            profile=controller.profile('d1',0)
            profile.assignments[0],profile.assignments[1]=profile.assignments[1],profile.assignments[0]
            controller._apply_profile(profile)
            self.assertNotEqual(backend.windows[1].rect,before[1])
            controller.undo();controller._guard_fixed.clear();controller._guard_slots()
            self.assertEqual({h:backend.windows[h].rect for h in (1,2)},before)

    def test_a_window_that_never_stays_in_its_tile_is_released(self):
        import time as clock
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            now=clock.monotonic()
            controller._guard_fixed[1]=tuple(now-2.9+.3*i for i in range(9))
            grow();controller._guard_slots()
            self.assertIn(1,controller._guard_released)
            self.assertNotEqual(backend.windows[1].rect,tile)
            # Placed again by an arrangement: guarded again.
            controller._apply_profile(controller.profile('d1',0))
            self.assertNotIn(1,controller._guard_released)

    def test_a_gesture_without_an_end_does_not_disable_the_guard(self):
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            controller._enqueue_event({'type':'move_start','hwnd':999})
            grow();controller._guard_slots()
            self.assertEqual(backend.windows[1].rect,tile)

    def test_fullscreen_and_changing_displays_are_left_alone(self):
        from dataclasses import replace
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            # A video going fullscreen in a tiled browser.
            backend.windows[1]=replace(backend.windows[1],rect=Rect(0,0,1920,1080))
            controller._guard_slots()
            self.assertEqual(backend.windows[1].rect,Rect(0,0,1920,1080))
            # Resolution changed, not rescanned yet: the tile may be stale.
            controller._guard_fixed.clear();grow()
            backend.displays=[Display('d1','D',Rect(0,0,2560,1440),primary=True)]
            controller._guard_slots()
            self.assertNotEqual(backend.windows[1].rect,tile)

    def test_a_window_moved_under_the_held_mouse_is_not_pulled_back(self):
        from dataclasses import replace
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            # An application dragging its own window (no native move reported).
            backend.mouse_pressed=True
            dragged=Rect(tile.x+300,tile.y+40,tile.width,tile.height)
            backend.windows[1]=replace(backend.windows[1],rect=dragged)
            controller._guard_slots()
            self.assertEqual(backend.windows[1].rect,dragged)
            # Growing in place is clamped at once, even with the button held.
            controller._guard_fixed.clear();grow();controller._guard_slots()
            self.assertEqual(backend.windows[1].rect,tile)

    def test_an_app_growing_now_and_then_stays_guarded(self):
        import time as clock
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            # Ten corrections spread over 9 s: clicks, not a stubborn app.
            now=clock.monotonic()
            controller._guard_fixed[1]=tuple(now-9+i for i in range(9))
            grow();controller._guard_slots()
            self.assertNotIn(1,controller._guard_released)
            self.assertEqual(backend.windows[1].rect,tile)

    def test_a_move_during_the_cooldown_is_checked_at_its_end(self):
        with tempfile.TemporaryDirectory() as directory:
            controller,backend,tile,grow=self.make(directory)
            grow();controller._guard_slots(only=1)
            self.assertEqual(backend.windows[1].rect,tile)
            grow();controller._guard_slots(only=1)
            fixed=controller._guard_fixed[1][-1]
            self.assertAlmostEqual(controller._guard_recheck,fixed+.03,places=3)
