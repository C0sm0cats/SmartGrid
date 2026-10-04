"""Built-in exclusions."""
import unittest
from smartgrid.core.exclusions import excluded_class, excluded_window


class ExclusionTests(unittest.TestCase):
    def test_overlays_launchers_and_call_windows_are_excluded(self):
        for title in ('Spotify Premium', 'Steam', 'VLC media player', 'Incoming call', 'OBS 30.1 - Profile',
                      'Battle.net', 'Picture in picture', 'Search', 'Volume Control'):
            self.assertTrue(excluded_window(title), title)

    def test_ordinary_windows_are_not_caught_by_substrings(self):
        for title in ('Getting started - Microsoft Edge', 'research.docx - Word', 'Steamboat notes - Notepad',
                      'Mozilla Firefox', 'Recall settings', 'Pipeline.py - Visual Studio Code', 'Start menu ideas.txt'):
            self.assertFalse(excluded_window(title), title)

    def test_small_teams_windows_are_notifications(self):
        teams = r'exe:c:\users\me\appdata\local\microsoft\teams\ms-teams.exe'
        self.assertTrue(excluded_window('Microsoft Teams', teams, 400, 150))
        self.assertFalse(excluded_window('Chat | Microsoft Teams', teams, 1200, 800))

    def test_shell_and_dialog_classes(self):
        self.assertTrue(excluded_class('#32770'))
        self.assertTrue(excluded_class('Progman'))
        self.assertFalse(excluded_class('MozillaWindowClass'))


if __name__ == '__main__':
    unittest.main()
