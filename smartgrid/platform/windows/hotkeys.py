from __future__ import annotations
import logging
from .api import Win32
from .messages import MessageThread

log = logging.getLogger(__name__)
MODIFIERS = {'alt':1, 'ctrl':2, 'control':2, 'shift':4, 'win':8, 'windows':8, 'meta':8}
KEYS = {'left':0x25, 'up':0x26, 'right':0x27, 'down':0x28, 'escape':0x1b, 'esc':0x1b, 'space':0x20, 'enter':0x0d, 'return':0x0d, 'tab':9, 'delete':0x2e, 'backspace':8, 'home':0x24, 'end':0x23, 'pageup':0x21, 'pagedown':0x22}

def parse_shortcut(shortcut):
    parts = [part.strip().lower() for part in shortcut.split('+')]
    if not parts or not parts[-1]:
        raise ValueError('Shortcut needs a key')
    modifiers = 0x4000  # MOD_NOREPEAT
    for part in parts[:-1]:
        if part not in MODIFIERS:
            raise ValueError(f'Unknown modifier: {part}')
        modifiers |= MODIFIERS[part]
    key = parts[-1]
    if key in KEYS:
        return modifiers, KEYS[key]
    if len(key) == 1 and key.isascii() and key.isalnum():
        return modifiers, ord(key.upper())
    if key.startswith('f') and key[1:].isdigit() and 1 <= int(key[1:]) <= 24:
        return modifiers, 0x6f + int(key[1:])
    raise ValueError(f'Unsupported key: {key}')

class HotkeyManager(MessageThread):
    def __init__(self, callback, api=None):
        super().__init__(api or Win32(), 'SmartGrid hotkeys')
        self.callback = callback
        self.registered = {}
        self._cleanup.append(self._unregister_owned)

    def register(self, mapping):
        def owned():
            self._unregister_owned()
            conflicts, seen = {}, set()
            for identifier, (action, shortcut) in enumerate(mapping.items(), 1):
                if not shortcut.strip():
                    continue
                try:
                    combination = parse_shortcut(shortcut)
                    if combination in seen:
                        raise ValueError('Duplicate shortcut')
                    seen.add(combination)
                    if not self.api.RegisterHotKey(None, identifier, *combination):
                        raise ValueError('Shortcut is reserved or already registered')
                    self.registered[identifier] = action
                except ValueError as error:
                    conflicts[action] = str(error)
            return conflicts
        return self.call(owned)

    def _unregister_owned(self):
        for identifier in self.registered:
            self.api.UnregisterHotKey(None, identifier)
        self.registered.clear()

    def unregister(self):
        if self.thread and self.thread.is_alive() and not self._closing:
            self.call(self._unregister_owned)

    def process_message(self, message, wparam, lparam=0):
        if message != 0x312 or wparam not in self.registered:
            return False
        try:
            self.callback(self.registered[wparam])
        except Exception:
            log.exception('Hotkey callback failed')
        return True

    def on_message(self, message):
        self.process_message(message.message, message.wParam, message.lParam)
