"""Built-in exclusions."""
import unittest
from smartgrid.core.exclusions import TITLE_WORDS, excluded_app, excluded_class, excluded_window, overlay_words


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

    def test_keyword_changes_apply_on_top_of_the_built_in_list(self):
        words = overlay_words(['Acme Tool'], ['call'])
        self.assertTrue(excluded_window('Acme tool - settings', words=words))
        self.assertFalse(excluded_window('Team call', words=words))
        self.assertTrue(excluded_window('Incoming call', words=words))
        self.assertEqual(overlay_words(), TITLE_WORDS)
        self.assertFalse(excluded_window('Spotify', words=()))

    def test_application_names_in_the_list_float_by_default(self):
        for name in ('Spotify', 'Steam', 'VLC media player', 'OBS Studio', 'Discord'):
            self.assertTrue(excluded_app(name), name)
        for name in ('Microsoft Teams', 'Visual Studio Code', 'Google Chrome', 'Notepad'):
            self.assertFalse(excluded_app(name), name)


if __name__ == '__main__':
    unittest.main()
