"""Offscreen Qt regressions. These verify widgets, not the Windows compositor."""
import os
os.environ.setdefault('QT_QPA_PLATFORM','offscreen')
import tempfile
import time
import unittest
from pathlib import Path
from PySide6.QtWidgets import QApplication
from smartgrid.__main__ import demo_backend
from smartgrid.core.controller import Controller
from smartgrid.core.models import Assignment
from smartgrid.core.geometry import valid_tiles
from smartgrid.storage.repository import Repository
from smartgrid.ui.bridge import bridge_for
from smartgrid.ui.studio import Studio
from smartgrid.ui.preferences import Preferences
from smartgrid.ui.quick_switcher import QuickSwitcher
from smartgrid.ui.theme import apply_theme


class UITests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app=QApplication.instance() or QApplication([])

    def setUp(self):
        self.temp=tempfile.TemporaryDirectory()
        self.controller=Controller(demo_backend(),Repository(self.temp.name))
        self.controller.refresh_apps()
        apply_theme(self.controller.settings)
        self.widgets=[]

    def tearDown(self):
        bridge_for(self.controller).shutdown()
        for widget in self.widgets:
            widget.hide()
            widget.deleteLater()
        self.app.processEvents()
        self.controller.quit()
        self.temp.cleanup()

    def settle(self):
        deadline=time.monotonic()+5
        while bridge_for(self.controller).pending and time.monotonic()<deadline:
            self.app.processEvents()
            time.sleep(.01)
        self.app.processEvents()
        self.assertEqual(bridge_for(self.controller).pending,0)

    def test_preferences_apply_immediately_reset_and_flush_on_escape(self):
        prefs=Preferences(self.controller);self.widgets.append(prefs)
        prefs.controls['gap'].setValue(24)
        deadline=time.monotonic()+2
        while time.monotonic()<deadline and self.controller.settings.gap!=24:
            self.app.processEvents();time.sleep(.01)
        self.settle();self.assertEqual(self.controller.settings.gap,24)
        prefs.resets['gap'].click();prefs.reject();self.settle()
        self.assertEqual(self.controller.settings.gap,12)
        self.assertEqual(prefs.tabs.tabText(1),'Applications')

    def test_saved_preview_requires_edit_and_save_is_inline(self):
        profile=self.controller.draft('demo-primary',0).profile
        self.controller.save_profile_template('Work',profile)
        studio=Studio(self.controller);self.widgets.append(studio);studio.show();self.settle()
        studio.mode_combo.setCurrentIndex(studio.mode_combo.findData('edit'))
        self.assertFalse(studio.canvas.isEnabled())
        studio.edit_button.click();self.assertTrue(studio.canvas.isEnabled())
        studio.save_template();self.assertTrue(studio.name_row.isVisible())
        self.assertFalse(studio.replace_row.isVisible())
        studio.hide();studio.close()

    def test_saved_list_selection_and_inline_space_sharing(self):
        self.controller.start()
        first=self.controller.save_profile_template('First',self.controller.profile('demo-primary',0))
        second=self.controller.save_profile_template('Second',self.controller.profile('demo-primary',0))
        studio=Studio(self.controller);self.widgets.append(studio);studio.show();self.settle()
        studio.mode_combo.setCurrentIndex(studio.mode_combo.findData('edit'))
        for row in range(studio.saved_list.count()):
            if studio.saved_list.item(row).data(256)==second.id: studio.saved_list.setCurrentRow(row);break
        self.assertEqual(studio.template_id,second.id);self.assertFalse(studio.canvas.isEnabled())
        # A window is shared by assigning it in another space of the same display.
        studio.mode_combo.setCurrentIndex(0);self.settle()
        window_id=self.controller.draft('demo-primary',0).profile.assignments[0].window_id
        studio._space_changed(1);studio.select_tile(0)
        self.assertIs(studio.stack.currentWidget(),studio.library)
        studio.drop_assignment(0,{'app_id':'browser','window_id':window_id});self.settle()
        self.assertIs(studio.stack.currentWidget(),studio.canvas)
        self.assertTrue(any(a and a.window_id==window_id for a in self.controller.draft('demo-primary',0).profile.assignments))
        self.assertTrue(any(a and a.window_id==window_id for a in self.controller.draft('demo-primary',1).profile.assignments))
        self.assertEqual(studio._window_location(self.controller._window(window_id)),'Space 1, Space 2')

    def test_picker_row_click_restores_and_has_no_preview_panel(self):
        self.controller.save_profile_template('Work',self.controller.profile('demo-primary',0))
        picker=QuickSwitcher(self.controller);self.widgets.append(picker);picker.show();self.settle()
        self.assertFalse(picker.preview.parentWidget().isVisible())
        self.assertFalse(picker.restore.isVisible())
        picker.list.itemClicked.emit(picker.list.currentItem());self.settle()
        self.assertFalse(picker.isVisible());self.assertTrue(self.controller.running)

    def test_studio_and_preferences_construct_at_small_and_dense_sizes(self):
        studio=Studio(self.controller); self.widgets.append(studio)
        studio.show(); self.settle()
        for preset in ['2x2','5x5']:
            self.controller.set_preset('demo-primary',0,preset)
            studio.refresh_from_controller()
            studio.resize(650,500); self.app.processEvents()
            self.assertGreater(studio.canvas.width(),100)
        prefs=Preferences(self.controller); self.widgets.append(prefs)
        prefs.show(); self.settle()
        settings=prefs.collect_settings()
        self.assertEqual(settings.gap,self.controller.settings.gap)

    def test_auto_five_windows_convert_to_full_custom_partition(self):
        studio=Studio(self.controller); self.widgets.append(studio)
        profile=self.controller.draft('demo-primary',0).profile
        profile.assignments=[Assignment('editor') for _ in range(5)]
        studio._ensure_custom(profile)
        self.assertTrue(valid_tiles(profile.tiles))
        self.assertEqual(len(profile.tiles),6)

    def test_quick_switcher_effects_and_template_render(self):
        self.controller.set_assignment('demo-primary',0,0,Assignment('editor',2,True))
        profile=self.controller.draft('demo-primary',0).profile
        template=self.controller.save_profile_template('Work',profile)
        picker=QuickSwitcher(self.controller); self.widgets.append(picker)
        picker.show(); self.settle()
        self.assertEqual(sum(bool(picker.list.item(i).data(256)) for i in range(picker.list.count())),1)
        self.assertIn(template.id,picker.previews)

    def test_visual_artifact(self):
        self.controller.start()
        studio=Studio(self.controller); self.widgets.append(studio)
        studio.resize(1120,760); studio.show(); self.settle()
        screenshot=Path(self.temp.name)/'studio.png'
        self.assertTrue(studio.grab().save(str(screenshot)))

    def test_application_lifetime_tray_commands_and_reserved_cards(self):
        from smartgrid.ui.app import DesktopUI
        ui=DesktopUI(self.controller,self.app)
        self.settle()
        self.assertFalse(ui.windows)
        self.assertFalse(self.controller.running)
        self.assertFalse(self.controller.backend.operations)
        ui.open('studio')
        self.widgets.extend(ui.windows.values())
        self.settle()
        self.assertTrue(ui.windows['studio'].isVisible())
        self.assertFalse(self.controller.running)
        ui.add('Test', 'stop').trigger()
        self.settle()
        self.assertFalse(self.controller.running)
        ui.timer.stop()
        ui.focus_frame.close()
        ui.window_actions.close()
        ui.tray.hide()

    def test_screen_mapping_does_not_depend_on_monitor_names(self):
        from smartgrid.ui.app import DesktopUI
        ui=DesktopUI(self.controller,self.app);self.settle()
        display=self.controller.displays[0]
        self.assertIsNotNone(ui._screen(display))
        ui.tray.hide()

    def test_preview_scales_physical_padding_and_complete_auto_grid(self):
        studio=Studio(self.controller); self.widgets.append(studio)
        profile=self.controller.draft('demo-primary',0).profile
        profile.assignments=[Assignment('editor') for _ in range(5)]
        studio.refresh_from_controller()
        self.assertEqual(len(studio.canvas.cards),6)
        studio.canvas.resize(600,400)
        studio.canvas._layout_cards()
        # Gap in the preview is the desktop gap scaled to the preview, never 8 raw Qt pixels.
        r=studio.canvas.rectangles
        self.assertLessEqual(r[1].x-r[0].right,4)

    def test_partial_placement_is_reported_as_an_error(self):
        from smartgrid.core.commands import ApplyResult
        studio=Studio(self.controller); self.widgets.append(studio)
        messages=[]
        studio.show_error=messages.append
        studio._applied(ApplyResult(failures={2:'Placement refused'}))
        self.assertTrue(messages)
        self.assertIn('Partially',messages[0])

    def test_library_skips_hidden_helper_windows_and_preserves_unchanged_items(self):
        from smartgrid.core.models import WindowRecord,WindowRef,Rect
        for hwnd in range(100,1100):
            self.controller.backend.windows[hwnd]=WindowRecord(WindowRef(hwnd,hwnd,str(hwnd)),'helper',
                'System helper',Rect(0,0,1,1),'demo-primary',eligible=False,exclusion_reason='Not visible')
        self.controller.refresh(auto=False)
        studio=Studio(self.controller);self.widgets.append(studio)
        library=studio.library
        self.assertEqual(library.windows.count(),0)
        library.section_buttons[0].setChecked(True)
        self.assertEqual(library.windows.count(),4)
        original=library.windows.item(0)
        library.windows.setCurrentItem(original)
        library.refresh_from_controller()
        self.assertIs(library.windows.item(0),original)
        self.assertIs(library.windows.currentItem(),original)

    def test_library_remembers_each_search_and_does_not_rebuild_hidden_catalogue(self):
        studio=Studio(self.controller);self.widgets.append(studio)
        library=studio.library
        library.search.setText('terminal')
        library.refresh_from_controller()
        library.tabs.setCurrentIndex(1);self.settle()
        self.assertEqual(library.search.text(),'')
        library.search.setText('editor')
        library.refresh_from_controller()
        original=library.apps.item(0)
        self.assertIsNotNone(original)
        library.tabs.setCurrentIndex(0)
        self.assertEqual(library.search.text(),'terminal')
        library.refresh_from_controller()
        self.assertIs(library.apps.item(0),original)
        library.tabs.setCurrentIndex(1)
        self.assertEqual(library.search.text(),'editor')
        self.assertIs(library.apps.item(0),original)

    def test_studio_activation_and_deactivation_restore_original_desktop(self):
        original={h:w.rect for h,w in self.controller.backend.windows.items()}
        studio=Studio(self.controller);self.widgets.append(studio)
        studio.show()
        studio.toggle_button.click()
        self.settle()
        self.assertTrue(self.controller.running)
        self.assertEqual(len(self.controller.last_result.placed),4)
        self.assertEqual(studio.toggle_button.text(),'Arrange windows')
        studio.toggle_button.click()
        self.settle()
        self.assertFalse(self.controller.running)
        self.assertEqual({h:w.rect for h,w in self.controller.backend.windows.items()},original)

    def test_current_preview_matches_adaptive_desktop_and_new_template_stays_independent(self):
        self.controller.profile('demo-primary',0).preset='5x5'
        self.controller.start()
        studio=Studio(self.controller);self.widgets.append(studio)
        studio.show();self.settle()
        self.assertEqual(studio.canvas.effective_preset,'2x2')
        self.assertEqual(len(studio.canvas.cards),4)
        self.assertFalse(studio.preset_combo.isVisible())
        studio.mode_combo.setCurrentIndex(studio.mode_combo.findData('new'))
        studio.drop_assignment(0,{'app_id':'editor'})
        self.settle()
        studio.mode_combo.setCurrentIndex(0)
        studio.mode_combo.setCurrentIndex(studio.mode_combo.findData('new'))
        self.assertEqual(studio.current_profile().assignments[0].app_id,'editor')
        self.assertTrue(studio.preset_strip.isVisible())
        self.assertFalse(studio.preset_combo.isVisible())

    def test_template_capacity_cannot_hide_assignments_and_custom_starts_with_one_tile(self):
        studio=Studio(self.controller);self.widgets.append(studio)
        studio.mode_combo.setCurrentIndex(studio.mode_combo.findData('new'))
        studio.drop_assignment(3,{'app_id':'editor'})
        self.settle()
        studio.preset_combo.setCurrentIndex(studio.preset_combo.findData('1x1'))
        self.assertEqual(studio.current_profile().preset,'2x2')
        self.assertIsNotNone(studio.current_profile().assignments[3])
        studio.clear_tile(3);self.settle()
        studio.preset_combo.setCurrentIndex(studio.preset_combo.findData('custom'))
        self.assertEqual(len(studio.current_profile().tiles),1)
        self.assertTrue(valid_tiles(studio.current_profile().tiles))

    def test_quick_switcher_search_cannot_restore_hidden_selection(self):
        from smartgrid.core.models import SpaceProfile,Tile
        self.controller.save_profile_template('Work',SpaceProfile('demo-primary',0,preset='2x2'))
        self.controller.save_profile_template('Reading',SpaceProfile('demo-primary',0,preset='custom',tiles=[Tile('all',0,0,1,1)]))
        picker=QuickSwitcher(self.controller);self.widgets.append(picker)
        picker.show();self.settle()
        self.assertEqual(picker.list.item(1).data(256),self.controller.templates[1].id)
        picker.search.setText('work')
        self.assertIn('Work',picker.list.currentItem().text())
        picker.search.setText('missing')
        self.assertFalse(picker.restore.isEnabled())
        operations=len(self.controller.backend.operations)
        picker.restore_selected();self.settle()
        self.assertEqual(len(self.controller.backend.operations),operations)
        picker.search.clear()
        self.assertFalse(picker.list.currentItem().isHidden())

    def test_template_restore_uses_open_hide_policy_and_preserves_other_drafts(self):
        self.controller.start()
        self.controller.set_assignment('demo-primary',1,0,Assignment('editor',None,True))
        studio=Studio(self.controller);self.widgets.append(studio)
        studio.mode_combo.setCurrentIndex(studio.mode_combo.findData('new'))
        studio.drop_assignment(0,{'app_id':'editor','window_id':2});self.settle()
        studio.apply();self.settle()
        self.assertEqual(self.controller.last_result.placed,[2])
        self.assertEqual(len(self.controller.last_result.hidden),3)
        self.assertTrue(self.controller.draft('demo-primary',1).dirty)

    def test_arrange_updates_clean_studio_draft_and_does_not_hide_failure(self):
        studio=Studio(self.controller);self.widgets.append(studio)
        picker=QuickSwitcher(self.controller);self.widgets.append(picker)
        self.settle()
        self.controller.backend.failures.add(2)
        picker.show();picker.arrange_open();self.settle()
        self.assertTrue(picker.isVisible())
        self.assertIn('failed',picker.message.text())
        self.assertEqual(self.controller.draft('demo-primary',0).profile.assignments,
                         self.controller.profile('demo-primary',0).assignments)
        self.assertEqual(len(studio.canvas.cards),4)
