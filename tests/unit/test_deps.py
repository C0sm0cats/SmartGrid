"""Dependency check used by smartgrid.bat before it starts the application."""
import tempfile
import unittest
from pathlib import Path
from smartgrid import deps


class DependencyCheckTests(unittest.TestCase):
    def check(self, text):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'requirements.txt'
            path.write_text(text, encoding='utf-8')
            return deps.missing(path)

    def test_installed_compatible_requirement_passes(self):
        self.assertEqual(self.check('# comment\n\npip>=1\n'), [])

    def test_missing_or_incompatible_requirements_are_reported(self):
        self.assertEqual(self.check('smartgrid-not-installed>=1\npip<1\n'),
                         ['smartgrid-not-installed>=1', 'pip<1'])

    def test_inapplicable_marker_is_ignored(self):
        self.assertEqual(self.check('smartgrid-not-installed; sys_platform == "nonexistent"\n'), [])


if __name__ == '__main__':
    unittest.main()
