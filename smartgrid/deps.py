"""Dependency check for the Windows launchers: exit 0 when requirements.txt is satisfied.

Run as ``python -m smartgrid.deps``; it never imports the application itself.
"""
import importlib.metadata
from pathlib import Path
import sys

REQUIREMENTS = Path(__file__).resolve().parents[1] / 'requirements.txt'


def missing(path=REQUIREMENTS):
    """Return the requirement lines that are not installed in a compatible version."""
    try:
        from pip._vendor.packaging.requirements import Requirement
    except ImportError:
        Requirement = None
    result = []
    for line in Path(path).read_text(encoding='utf-8').splitlines():
        line = line.split('#', 1)[0].strip()
        if not line:
            continue
        try:
            if Requirement is None:
                name = line.split(';')[0].split('[')[0]
                for separator in '<>=!~ ':
                    name = name.split(separator)[0]
                importlib.metadata.version(name)
                continue
            requirement = Requirement(line)
            if requirement.marker and not requirement.marker.evaluate():
                continue
            version = importlib.metadata.version(requirement.name)
            if requirement.url or not requirement.specifier.contains(version, prereleases=True):
                result.append(line)
        except (importlib.metadata.PackageNotFoundError, ValueError):
            result.append(line)
    return result


def main():
    absent = missing()
    if not absent:
        try:
            import PySide6.QtWidgets  # noqa: F401  (a broken Qt install must reinstall)
            import PIL  # noqa: F401
        except ImportError as error:
            absent = [str(error)]
    for line in absent:
        print(f'Missing or incompatible: {line}', file=sys.stderr)
    return 1 if absent else 0


if __name__ == '__main__':
    raise SystemExit(main())
