import unittest
from dataclasses import asdict

from smartgrid.core.models import Settings
from smartgrid.core.window_actions import (REMOTE_HANDLE_COLOR, action_style, placement, rule_value, set_rule,
    tooltip_placement)

MSTSC = r'exe:C:\Windows\System32\mstsc.exe'


class WindowActionRuleTests(unittest.TestCase):
    def test_remote_viewers_default_to_an_orange_handle_on_the_right(self):
        self.assertEqual(action_style(MSTSC, Settings()), ('right', REMOTE_HANDLE_COLOR))
        self.assertEqual(action_style(r'exe:C:\Program Files\VMware\vmware.exe', Settings()), ('right', REMOTE_HANDLE_COLOR))
        self.assertEqual(action_style(r'exe:C:\Apps\code.exe', Settings()), ('left', None))
        # The Start menu entry "Connexion Bureau à distance" is an AppID, not an executable.
        self.assertEqual(action_style('aumid:microsoft.windows.remotedesktop', Settings()), ('right', REMOTE_HANDLE_COLOR))

    def test_own_rule_wins_over_the_executable_default_and_removal_clears_both(self):
        rules = set_rule(Settings().action_positions, MSTSC, 'top')
        self.assertEqual(rule_value(rules, MSTSC), 'top')
        self.assertNotIn('mstsc.exe', rules)
        self.assertIsNone(rule_value(set_rule(rules, MSTSC, None), MSTSC))
        self.assertEqual(rule_value({'mstsc.exe': 'right', MSTSC: 'bottom'}, MSTSC), 'bottom')

    def test_settings_validate_rules_and_old_files_get_the_defaults(self):
        data = asdict(Settings())
        data.pop('action_positions'); data.pop('action_colors')
        self.assertEqual(Settings.from_dict(data).action_positions['mstsc.exe'], 'right')
        for name, value in (('action_positions', {'a': 'middle'}), ('action_colors', {'a': 'orange'})):
            with self.assertRaises(ValueError):
                Settings.from_dict({**asdict(Settings()), name: value})

    def test_a_default_added_later_reaches_saved_settings_once(self):
        data = asdict(Settings())
        data.pop('action_defaults')
        for rules in (data['action_positions'], data['action_colors']):
            rules.pop('Microsoft.Windows.RemoteDesktop')
        data['action_positions'].pop('vmware.exe')  # removed before the tracking existed: comes back once
        settings = Settings.from_dict(data)
        self.assertEqual(action_style('aumid:microsoft.windows.remotedesktop', settings), ('right', REMOTE_HANDLE_COLOR))
        removed = asdict(settings)
        for rules in (removed['action_positions'], removed['action_colors']):
            rules.pop('Microsoft.Windows.RemoteDesktop')
        again = Settings.from_dict(removed)
        self.assertEqual(action_style('aumid:microsoft.windows.remotedesktop', again), ('left', None))

    def test_handle_sits_against_the_chosen_edge_inside_the_screen(self):
        rect, area = (100, 100, 800, 600), (0, 0, 1920, 1080)
        self.assertEqual(placement('left', rect, area, (28, 40)), (104, 380))
        self.assertEqual(placement('right', rect, area, (28, 40)), (868, 380))
        self.assertEqual(placement('top', rect, area, (40, 28)), (480, 104))
        self.assertEqual(placement('bottom', rect, area, (40, 28)), (480, 668))
        self.assertEqual(placement('right', (1800, 0, 400, 300), area, (28, 40)), (1886, 130))

    def test_tooltip_goes_away_from_the_window_edge(self):
        area = (0, 0, 1920, 1080)
        self.assertEqual(tooltip_placement('left', (100, 100, 36, 36), (80, 20), area), (144, 108))
        self.assertEqual(tooltip_placement('right', (800, 100, 36, 36), (80, 20), area), (712, 108))
        self.assertEqual(tooltip_placement('top', (100, 100, 36, 36), (80, 20), area), (78, 144))
        self.assertEqual(tooltip_placement('bottom', (100, 500, 36, 36), (80, 20), area), (78, 472))


if __name__ == '__main__':
    unittest.main()
