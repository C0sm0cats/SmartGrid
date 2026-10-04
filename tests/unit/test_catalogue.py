"""Start menu catalogue parsing, without PowerShell or Win32."""
import json
import unittest
from types import SimpleNamespace
from unittest.mock import patch
from smartgrid.platform.windows import apps as catalogue


class CatalogueTests(unittest.TestCase):
    def discover(self, rows, files=()):
        result = catalogue.AppCatalogue(api=None, cache_dir='unused')
        result.icon = lambda target: ''
        result.known_folder = lambda guid: r'C:\Windows\System32'
        output = SimpleNamespace(returncode=0, stdout=json.dumps(rows).encode())
        with patch.object(catalogue.subprocess, 'run', return_value=output), \
             patch.object(catalogue.os.path, 'isfile', side_effect=lambda path: path in files):
            return {app.name: app for app in result.discover()}

    def test_desktop_appids_become_executables_and_duplicates_merge(self):
        apps = self.discover([
            {'name': 'Notepad', 'aumid': r'{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\notepad.exe', 'executable': ''},
            {'name': 'File Explorer', 'aumid': 'Microsoft.Windows.Explorer', 'executable': ''},
            {'name': 'File Explorer', 'aumid': '', 'executable': r'C:\Windows\explorer.exe'},
            {'name': 'Terminal', 'aumid': 'Microsoft.WindowsTerminal_8wekyb3d8bbwe!App', 'executable': ''},
        ], files={r'C:\Windows\System32\notepad.exe'})
        self.assertEqual(apps['Notepad'].id, r'exe:c:\windows\system32\notepad.exe')
        self.assertEqual(apps['File Explorer'].id, r'exe:c:\windows\explorer.exe')
        self.assertEqual(apps['Terminal'].id, 'aumid:microsoft.windowsterminal_8wekyb3d8bbwe!app')
        self.assertTrue(all(app.installed for app in apps.values()))

    def test_uninstallers_documents_and_links_are_skipped(self):
        apps = self.discover([
            {'name': 'Uninstall Foo', 'aumid': '', 'executable': r'C:\Foo\unins000.exe'},
            {'name': 'Foo Help', 'aumid': r'C:\Foo\help.chm', 'executable': ''},
            {'name': 'Foo website', 'aumid': 'https://foo.example', 'executable': ''},
            {'name': 'Foo', 'aumid': '', 'executable': r'C:\Foo\foo.exe'},
        ])
        self.assertEqual(list(apps), ['Foo'])


if __name__ == '__main__':
    unittest.main()
