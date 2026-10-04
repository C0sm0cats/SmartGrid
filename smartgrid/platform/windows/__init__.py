"""Import-safe Windows services. Native DLLs load only when instantiated."""
from .backend import WindowsBackend, PlacementResult
from .hotkeys import HotkeyManager
from .instance import InstanceGuard

__all__ = ['WindowsBackend', 'PlacementResult', 'HotkeyManager', 'InstanceGuard']
