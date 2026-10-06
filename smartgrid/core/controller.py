"""One source of truth for drafts, spaces, native placements and their history.

Native calls are serialized. UI commands should execute in the UI bridge's worker;
native event callbacks only enqueue data and never rearrange from the callback.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import asdict, replace
import logging
import queue
import threading
import time
import uuid

from .models import app_display_name, Assignment, Draft, LayoutTemplate, Settings, SpaceProfile, Tile, Rect
from .geometry import capacity, resolve_layout, auto_preset, effective_layout, directional_neighbor, edge_neighbors, swap_neighbor, nearest_slot, valid_tiles, linked_resize
from .reconcile import compact_assignments
from .exclusions import excluded_app, excluded_window, overlay_words
from .history import History
from .commands import ApplyResult
from smartgrid.storage.archive import export_archive, import_archive

log = logging.getLogger(__name__)
# Smallest tile a linked resize leaves when windows are forced into their tiles.
FORCED_MINIMUM = (160, 120)

DEFAULT_HOTKEYS = {
    "toggle": "Ctrl+Alt+T", "arrange": "Ctrl+Alt+R", "studio": "Ctrl+Alt+P",
    "quick_switcher": "Ctrl+Alt+L", "swap": "Ctrl+Alt+S", "float": "Ctrl+Alt+F",
    "undo": "Ctrl+Alt+Z", "redo": "Ctrl+Alt+Y", "space1": "Ctrl+Alt+1", "space2": "Ctrl+Alt+2",
    "space3": "Ctrl+Alt+3", "stop": "Ctrl+Alt+Q",
    # Ctrl+Win+Left/Right switches Windows
    # virtual desktops, so SmartGrid defaults to Ctrl+Alt+arrow.
    "focus_left": "Ctrl+Alt+Left", "focus_right": "Ctrl+Alt+Right",
    "focus_up": "Ctrl+Alt+Up", "focus_down": "Ctrl+Alt+Down",
}
# Earlier defaults, replaced when found unchanged in saved settings.
RETIRED_HOTKEYS = {"focus_left": "Ctrl+Win+Left", "focus_right": "Ctrl+Win+Right",
                   "focus_up": "Ctrl+Win+Up", "focus_down": "Ctrl+Win+Down"}


class Controller:
    def __init__(self, backend, repository):
        self.backend, self.repository = backend, repository
        self.settings = repository.load_settings()
        self._migrate_hotkeys()
        self.templates, self.profiles = repository.load_layouts()
        self.displays = []
        self.windows = []
        self.apps = []
        self.active_spaces = {}
        self.drafts = {}
        self._draft_histories = {}
        self.history = History(10)
        self.running = False
        self.paused = False
        self.last_error = ""
        self.status = "Paused — the desktop is unchanged."
        self.last_result = ApplyResult()
        self.hotkey_conflicts = {}
        self._callbacks = []
        self._lock = threading.RLock()
        self._queue = queue.Queue()
        self._stop_event = threading.Event()
        self._thread = None
        self._hotkeys = None
        self._ui_action = None
        self._originals = {}
        self._parked = set()
        self._manual_minimized = set()
        self._hidden_by_template = set()
        self._floating = set()
        self._float_slots = {}
        self._pending = {}
        self._cancelled_launches = []
        self._late_unmanaged = set()
        self._generation = 0
        self._suppress_events_until = 0.0
        self._interaction_started = {}
        self._placed_at = {}
        self._interacting = set()
        self._move_origins = {}
        # 'move' or 'resize' per window being dragged, and its state at the start.
        self._gesture_kinds = {}
        self._gesture_states = {}
        # Latest pending location event per window: a burst of moves is handled once.
        self._pending_locations = {}
        self._pending_lock = threading.Lock()
        self._guide_callbacks = []
        self._follow_callbacks = []
        self._minimize_snapshots = {}
        self._resize_snapshot = None
        self._swap_snapshot = None
        self._swap_hwnd = None
        self._deleted_template = None
        self._last_border = None
        self._last_tiled_selection = None
        self.preview_rectangles = []
        # UI guides drawn by Qt: drag target with its
        # Swap/Move label, the last space switch for the OSD, swap arrows.
        self.drag_guide = None
        self.space_event = None
        self.motion_events = []
        self.swap_event = None
        self._tile_rects = {}
        self._closed = False
        self._last_reconcile = 0.0
        self._last_preview = 0.0
        self._interaction_minimums = {}
        session=repository.load_session()
        self.active_spaces.update(session.get('active_spaces',{}))
        for entry in session.get('pending',[]):
            key=(entry['display_id'],entry['space'],entry['index'])
            profile=self.profiles.get(key[:2])
            if profile and key[2]<len(profile.assignments) and profile.assignments[key[2]] and profile.assignments[key[2]].app_id==entry['app_id']:
                self._pending[key]={'app_id':entry['app_id'],'generation':self._generation,
                    'created_utc':entry['created_utc'],'created':time.monotonic()-max(0,time.time()-entry['created_utc'])}
        self._session_signature=None
        self._desktop_id=getattr(backend,'desktop_id',None)
        self.refresh(auto=False)
        self._bind_restored_profiles()

    def _bind_restored_profiles(self):
        """Rebind persisted occupants after a restart; forget apps that are closed.

        Profiles are stored without HWNDs. A regular occupant whose application
        is no longer open is not a launch request (only pins and
        real pending launches stay reserved), so it must not become OPENS ON APPLY.
        """
        with self._lock:
            pending={(key[0].split('@desktop:',1)[0],key[1],key[2]) for key in self._pending}
            for (display_id,space),profile in self.profiles.items():
                used={a.window_id for a in profile.assignments if a and a.window_id}
                for index,assignment in enumerate(profile.assignments):
                    if not assignment or assignment.window_id or assignment.pinned:
                        continue
                    if (display_id.split('@desktop:',1)[0],space,index) in pending:
                        continue
                    window=next((w for w in self.windows if w.app_id==assignment.app_id and w.display_id==display_id
                                 and self._eligible(w) and w.ref.hwnd not in used),None)
                    if window:
                        assignment.window_id=window.ref.hwnd;used.add(window.ref.hwnd)
                    else:
                        profile.assignments[index]=None
                if self.settings.compact_close:
                    profile.assignments=compact_assignments(profile.assignments)
                while profile.assignments and profile.assignments[-1] is None:
                    profile.assignments.pop()

    def _migrate_hotkeys(self):
        """Replace retired default shortcuts that the user never changed."""
        from smartgrid.platform.windows.hotkeys import parse_shortcut
        changed = False
        for action, old in RETIRED_HOTKEYS.items():
            value = self.settings.hotkeys.get(action)
            try:
                same = bool(value) and parse_shortcut(value.replace('Meta+', 'Win+')) == parse_shortcut(old)
            except ValueError:
                same = False
            if same:
                self.settings.hotkeys[action] = DEFAULT_HOTKEYS[action]
                changed = True
        if changed and not self.repository.read_only:
            self.repository.save_settings(self.settings)

    def subscribe(self, callback):
        self._callbacks.append(callback)

    def subscribe_guides(self, callback):
        """Called when only the drag guides changed (drop target, resize preview)."""
        self._guide_callbacks.append(callback)

    def _notify_guides(self):
        for callback in list(self._guide_callbacks):
            try:
                callback()
            except Exception:
                log.exception("Guide subscriber failed")

    def subscribe_follow(self, callback):
        """Called with the window of every native move, straight from the
        event thread: the palette follows its edge without waiting for a rescan."""
        self._follow_callbacks.append(callback)

    def _notify_follow(self, hwnd):
        for callback in list(self._follow_callbacks):
            try:
                callback(hwnd)
            except Exception:
                log.exception("Follow subscriber failed")

    def set_ui_action_handler(self, callback):
        self._ui_action = callback

    def _notify(self):
        try: self._save_session()
        except OSError as error:
            self.last_error='Cannot save the runtime session: '+str(error)
            log.exception('Unable to save runtime session')
        # Native arrangements must update a read-only/current Studio view, while
        # retaining every context the user is actively editing.
        for key,draft in list(self.drafts.items()):
            if not draft.dirty and key in self.profiles and draft.profile!=self.profiles[key]:
                draft.profile=deepcopy(self.profiles[key])
        for callback in list(self._callbacks):
            try:
                callback()
            except Exception:
                log.exception("Subscriber failed")

    def _window(self, hwnd):
        return next((w for w in self.windows if w.ref.hwnd == hwnd), None)

    def _display(self, display_id):
        result = next((d for d in self.displays if d.id == display_id), None)
        if result is None:
            raise ValueError("This display is no longer connected.")
        return result

    def _require_desktop(self):
        desktops=getattr(self.backend,'virtual_desktops',None)
        if desktops and desktops.available:
            current=desktops.current(self.backend.foreground())
            if current!=getattr(self.backend,'desktop_id',None) or not current:
                self.refresh(auto=False)
        if not getattr(self.backend,'desktop_known',True):
            raise ValueError('Cannot identify the current Windows desktop; arrangement is suspended.')

    def _ruled_out(self, window):
        """Always floating, or a built-in exclusion whose Always floating box was not unticked."""
        if window.app_id in self.settings.included_apps:
            return False
        if window.app_id in self.settings.excluded_apps:
            return True
        if not self.settings.builtin_exclusions:
            return False
        if window.app_id in self.settings.overlay_apps_added:
            return True
        if getattr(self, '_app_names', (None,))[0] is not self.apps:
            self._app_names = (self.apps, {a.id: a.name for a in self.apps})
        name = self._app_names[1].get(window.app_id) or app_display_name(window.app_id)
        words = overlay_words(self.settings.overlay_words_added, self.settings.overlay_words_removed)
        return bool(excluded_app(name, words) or
                    excluded_window(window.title, window.app_id, window.rect.width, window.rect.height, words))

    def _eligible(self, window):
        if not window.eligible or window.ref.hwnd in self._floating:
            return False
        return not self._ruled_out(window)

    def current_display_id(self):
        foreground = getattr(self.backend, "foreground", lambda: None)()
        window = self._window(foreground)
        if window and window.eligible and any(d.id==window.display_id for d in self.displays):
            return window.display_id
        if not self.displays:
            raise ValueError("No display available.")
        return next((d.id for d in self.displays if d.primary), self.displays[0].id)

    def profile(self, display_id, space):
        if space not in (0, 1, 2):
            raise ValueError("Space must be 1, 2 or 3.")
        key = (display_id, space)
        if key not in self.profiles:
            self.profiles[key] = SpaceProfile(display_id, space,ratio=self.settings.master_ratio)
        return self.profiles[key]

    def draft(self, display_id, space):
        key = (display_id, space)
        if key not in self.drafts:
            self.drafts[key] = Draft(deepcopy(self.profile(display_id, space)))
        return self.drafts[key]

    def _draft_checkpoint(self, display_id, space):
        key = (display_id, space)
        self._draft_histories.setdefault(key, History(50)).push(self.draft(*key).profile)

    def _draft_changed(self, display_id, space):
        draft = self.draft(display_id, space)
        draft.dirty = draft.profile != self.profile(display_id, space)
        self._notify()

    def _save(self):
        self.repository.save_layouts(self.templates, self.profiles)
        self._save_session()

    def _save_session(self):
        pending=[dict(display_id=key[0],space=key[1],index=key[2],app_id=value['app_id'],
                      created_utc=value.get('created_utc',time.time()-(time.monotonic()-value['created'])))
                 for key,value in sorted(self._pending.items())]
        # Store a wall-clock origin once; polling must not rewrite the file.
        for entry in pending:
            self._pending[(entry['display_id'],entry['space'],entry['index'])]['created_utc']=entry['created_utc']
        signature=(repr(pending),tuple(sorted(self.active_spaces.items())))
        if signature!=self._session_signature:
            self.repository.save_session(pending,self.active_spaces)
            self._session_signature=signature

    def refresh_apps(self):
        with self._lock:
            self.apps = self.backend.discover_apps()
            self._notify()
        return self.apps

    def refresh(self, auto=True):
        with self._lock:
            previous = {w.ref.hwnd: w for w in self.windows}
            previous_profiles=deepcopy(self.profiles)
            previous_error=self.last_error
            old_displays = {d.id: (d.work_area, d.dpi) for d in self.displays}
            self.displays = self.backend.discover_displays()
            desktop=getattr(self.backend,'desktop_id',None)
            if desktop!=self._desktop_id:
                self._desktop_id=desktop
                self.history.clear();self._hide_overlay();self._clear_border()
                self._last_tiled_selection=None
                self._interacting.clear();self._move_origins.clear();self._resize_snapshot=None
                if self._swap_snapshot is not None:
                    # Never roll an old-desktop snapshot onto the new desktop.
                    self._swap_snapshot=None;self._swap_hwnd=None
                    if self._hotkeys:
                        self.hotkey_conflicts=self._hotkeys.register(DEFAULT_HOTKEYS | self.settings.hotkeys)
            self.windows = self.backend.discover_windows()
            for window in self.windows:
                window.floating = window.ref.hwnd in self._floating or self._ruled_out(window)
            live = {w.ref.hwnd: w for w in self.windows}
            for display in self.displays:
                base=display.id.split('@desktop:',1)[0]
                if base!=display.id:
                    if base in self.active_spaces:
                        self.active_spaces.setdefault(display.id,self.active_spaces.pop(base))
                    for old_key in [key for key in self.profiles if key[0]==base]:
                        profile=self.profiles.pop(old_key);profile.display_id=display.id
                        self.profiles[(display.id,old_key[1])]=profile
                    for old_key,value in list(self.repository.variants.items()):
                        if value.get('display_id')==base:
                            profile=SpaceProfile.from_dict(value);profile.display_id=display.id
                            self.repository.variants.pop(old_key)
                            self.repository.remember_variant(profile)
                self.active_spaces.setdefault(display.id, 0)
                self.profile(display.id, self.active_spaces[display.id])
            changed = old_displays != {d.id: (d.work_area, d.dpi) for d in self.displays}
            layout_changed=changed
            closed = set(previous) - set(live)
            created = set(live) - set(previous)
            # HWND reuse must never restore a different application/process generation.
            replaced = {h for h in previous.keys() & live.keys() if previous[h].ref != live[h].ref}
            # A window hidden before it is destroyed (Electron apps such as Bruno)
            # is already ineligible when it closes: its tile still counts.
            assigned = {a.window_id for p in self.profiles.values() for a in p.assignments if a and a.window_id}
            for hwnd in closed | replaced:
                self._originals.pop(hwnd, None)
                self._floating.discard(hwnd)
                self._late_unmanaged.discard(hwnd)
                self._parked.discard(hwnd)
                self._hidden_by_template.discard(hwnd)
                self._manual_minimized.discard(hwnd)
                for profile in self.profiles.values():
                    for index, assignment in enumerate(profile.assignments):
                        if assignment and assignment.window_id == hwnd:
                            assignment.window_id = None
                            if not assignment.pinned and self.settings.compact_close:
                                profile.assignments[index] = None
            changed |= bool(closed or created or replaced)
            layout_changed |= any(self._eligible(previous[h]) for h in closed|replaced) or bool((closed|replaced) & assigned) or any(self._eligible(live[h]) for h in created|replaced)
            for hwnd, window in live.items():
                old = previous.get(hwnd)
                if old and window.state != old.state:
                    changed = True
                    layout_changed |= self._eligible(window) or self._eligible(old)
                    if window.state == "minimized" and hwnd not in self._parked and hwnd not in self._hidden_by_template:
                        context = (window.display_id, self.active_spaces.get(window.display_id, 0))
                        self._minimize_snapshots[hwnd] = (context, deepcopy(self.profile(*context)),
                                                        {a.window_id for a in self.profile(*context).assignments if a and a.window_id})
                        self._manual_minimized.add(hwnd)
                    elif window.state == "normal" and hwnd in self._manual_minimized:
                        self._manual_minimized.discard(hwnd)
                        saved = self._minimize_snapshots.pop(hwnd, None)
                        if saved and self.active_spaces.get(saved[0][0], 0) == saved[0][1]:
                            current = self.profile(*saved[0])
                            occupants = {a.window_id for a in current.assignments if a and a.window_id} | {hwnd}
                            if saved[2] == occupants and all(h in live for h in occupants):
                                self.profiles[saved[0]] = saved[1]
                    if window.state == "normal" and hwnd in self._parked:
                        self._parked.discard(hwnd)
                        # Restoring manually through taskbar brings the window into active space.
                        profile = self.profile(window.display_id, self.active_spaces.get(window.display_id, 0))
                        if not any(a and a.window_id == hwnd for a in profile.assignments):
                            profile.assignments.append(Assignment(window.app_id, hwnd))
            self._bind_pending()
            if auto and self.running and not self.paused and layout_changed and not self._interacting and self._swap_snapshot is None:
                self._reflow_all(include_new=True, compact=bool(closed and self.settings.compact_close))
            self._update_border()
            if changed or previous!=live or previous_profiles!=self.profiles or previous_error!=self.last_error:
                self._notify()
        return self.windows

    def start_runtime(self):
        if self._thread and self._thread.is_alive():
            return
        self._stop_event.clear()
        self.backend.start_events(self._enqueue_event)
        self._thread = threading.Thread(target=self._event_loop, name="SmartGrid-controller", daemon=True)
        self._thread.start()
        try:
            from smartgrid.platform.windows.hotkeys import HotkeyManager
            if getattr(self.backend, "is_fake", False):
                self._hotkeys = None
            else:
                self._hotkeys = HotkeyManager(lambda action: self._queue.put({"type": "hotkey", "action": action}))
                self.hotkey_conflicts = self._hotkeys.register(DEFAULT_HOTKEYS | self.settings.hotkeys)
        except (OSError, RuntimeError, ImportError):
            log.exception("Global hotkeys unavailable")
            self.last_error = "Global shortcuts are unavailable; the tray menu remains available."

    def _enqueue_event(self, event):
        """Queue a native event; location events are coalesced per window, so a
        drag never waits behind a backlog of positions already out of date."""
        if event.get("type") in ("location", "location_change") and event.get("hwnd"):
            self._notify_follow(event["hwnd"])
            with self._pending_lock:
                queued = event["hwnd"] in self._pending_locations
                self._pending_locations[event["hwnd"]] = event
            if queued:
                return
        self._queue.put(event)

    def _event_loop(self):
        while not self._stop_event.is_set():
            try:
                event = self._queue.get(timeout=1.0)
            except queue.Empty:
                event = {"type": "reconcile"}
            try:
                if event.get("type") in ("location", "location_change") and event.get("hwnd"):
                    with self._pending_lock:
                        event = self._pending_locations.pop(event["hwnd"], event)
                if event.get("type") == "hotkey":
                    self.handle_action(event["action"])
                else:
                    self._handle_event(event)
            except Exception as error:
                log.exception("Runtime command failed")
                self.last_error = str(error)
                self._notify()

    def _handle_event(self, event):
        kind, hwnd = event.get("type", ""), event.get("hwnd")
        with self._lock:
            if kind in ("move_start", "resize_start", "movesize_start") and hwnd:
                # Start following the gesture at once: no full rescan here.
                # Only an unknown window needs one.
                if self._window(hwnd) is None:
                    self.refresh(auto=False)
                # Moving a floating or unmanaged window is an ordinary move: no
                # drop target, no snapping, and its reserved tile is kept.
                if not self._is_tiled(hwnd):
                    return
                self._interacting.add(hwnd)
                self._interaction_started[hwnd]=time.monotonic()
                self._last_preview=0.0
                self._interaction_minimums.clear()
                w = self._window(hwnd)
                if w:
                    rect = getattr(self.backend, "window_rect", lambda h: None)(hwnd) or w.rect
                    self._move_origins[hwnd] = (rect, w.display_id)
                    # Windows reports every move or resize loop as move_start.
                    kind = "resize" if kind == "resize_start" else \
                        getattr(self.backend, "gesture_kind", lambda h: None)(hwnd)
                    self._gesture_kinds[hwnd] = kind
                    # Its own state for undo; the full snapshot is taken only for a resize.
                    self._gesture_states[hwnd] = (w.ref, self.backend.snapshot(hwnd))
                    log.info("Gesture %s on %s (%s) from %s", kind or "unknown", hwnd, w.app_id, rect)
                return
            if kind in ("move_end", "resize_end", "movesize_end") and hwnd:
                if hwnd not in self._interacting:
                    return
                self._interacting.discard(hwnd)
                original = self._move_origins.pop(hwnd, None)
                gesture = self._gesture_kinds.pop(hwnd, None)
                start_state = self._gesture_states.pop(hwnd, None)
                # A size change made by SmartGrid itself during the gesture (an
                # arrangement running meanwhile) is not a user resize or drop.
                if self._placed_at.get(hwnd,0)>=self._interaction_started.pop(hwnd,time.monotonic()):
                    original=None
                self.refresh(auto=False)
                window = self._window(hwnd)
                if original and window:
                    log.info("Gesture %s on %s ended: %s -> %s", gesture or "unknown", hwnd, original[0], window.rect)
                if self.running and not self.paused and original and original[1] in {d.id for d in self.displays} and window and window.state == "normal":
                    before, display_id = original
                    resized = abs(window.rect.width - before.width) > 8 or abs(window.rect.height - before.height) > 8
                    if window.rect != before or display_id != window.display_id:
                        # Snapshot for undo, taken only now that the gesture
                        # changes something: the current state, with this
                        # window as it was before the gesture.
                        self._resize_snapshot = self._snapshot()
                        if start_state:
                            self._resize_snapshot["windows"][hwnd] = start_state
                    if display_id!=window.display_id:
                        self._native_drop(hwnd,display_id,window.display_id)
                    elif resized and gesture != "move":
                        self._native_resize(hwnd, before, window.rect, self._resize_snapshot)
                    elif window.rect != before or display_id != window.display_id:
                        self._native_drop(hwnd, display_id, window.display_id)
                self._resize_snapshot = None
                self._interaction_minimums.clear()
                self._hide_overlay()
                return
            if kind in ("location", "location_change") and hwnd in self._interacting:
                now=time.monotonic()
                if now-self._last_preview>=.015:
                    self._last_preview=now
                    self._preview_native_drag(hwnd)
                return
            if time.monotonic() < self._suppress_events_until and kind not in ("foreground", "destroyed", "display_change"):
                return
            # While a window is dragged, other windows' moves and title changes
            # wait: a rescan would delay the drag. Lifecycle events still count.
            if self._interacting and kind in ("location", "location_change", "title", "changed"):
                return
            if kind == "foreground":
                self._update_border()
                return
            now = time.monotonic()
            if kind in ("location", "location_change", "title", "changed") and now-self._last_reconcile < self.settings.debounce_ms/1000:
                return
            self._last_reconcile = now
            self.refresh(auto=True)

    def _capture(self, hwnd):
        window = self._window(hwnd)
        if window and hwnd not in self._originals:
            self._originals[hwnd] = (window.ref, self.backend.snapshot(hwnd))

    def _snapshot(self):
        return {"profiles": deepcopy(self.profiles), "spaces": dict(self.active_spaces),
                "parked": set(self._parked), "manual": set(self._manual_minimized),
                "hidden": set(self._hidden_by_template), "floating": set(self._floating),
                "pending": deepcopy(self._pending), "running": self.running, "paused": self.paused,
                "focus": getattr(self.backend, "foreground", lambda: None)(),
                "references": {w.ref.hwnd:w.ref for w in self.windows},
                "windows": {w.ref.hwnd: (w.ref, self.backend.snapshot(w.ref.hwnd)) for w in self.windows if w.eligible}}

    def _restore_snapshot(self, snapshot):
        self._generation += 1
        self.profiles = deepcopy(snapshot["profiles"])
        self.active_spaces = dict(snapshot["spaces"])
        self._parked = set(snapshot["parked"])
        self._manual_minimized = set(snapshot["manual"])
        self._hidden_by_template = set(snapshot["hidden"])
        self._floating = set(snapshot["floating"])
        self._cancel_pending()
        self._pending = deepcopy(snapshot["pending"])
        restored={(p["app_id"],p["created"]) for p in self._pending.values()}
        self._cancelled_launches=[p for p in self._cancelled_launches if (p["app_id"],p["created"]) not in restored]
        self.running, self.paused = snapshot["running"], snapshot["paused"]
        valid_references = {w.ref.hwnd: w.ref for w in self.windows}
        for profile in self.profiles.values():
            for assignment in profile.assignments:
                if assignment and assignment.window_id:
                    expected = snapshot["windows"].get(assignment.window_id)
                    reference=snapshot.get("references",{}).get(assignment.window_id) or (expected[0] if expected else None)
                    if not reference or valid_references.get(assignment.window_id) != reference:
                        assignment.window_id = None
        self._suppress_events_until = time.monotonic() + 0.5
        failures = []
        for hwnd, (reference, state) in snapshot["windows"].items():
            window = self._window(hwnd)
            if window and window.ref == reference and self.backend.alive(hwnd):
                if not self.backend.restore(hwnd, state):
                    failures.append(hwnd)
        if snapshot["focus"] and self.backend.alive(snapshot["focus"]):
            self.backend.focus(snapshot["focus"])
        current_displays={display.id for display in self.displays}
        for key in list(self.drafts):
            if key[0] in current_displays:
                self.drafts.pop(key);self._draft_histories.pop(key,None)
        self._save()
        self.refresh(auto=False)
        self.status = "Arrangement restored." if not failures else "Partially restored: some windows refused placement."
        self._notify()

    def start(self):
        with self._lock:
            self._require_desktop()
            if self.running:
                return self.resume()
            self.refresh(auto=False)
            self.history.push(self._snapshot())
            self.running, self.paused = True, False
            self._bind_pending()
            self._reflow_all(include_new=True, explicit=True)
            self.status = self.last_result.summary
            self._notify()

    def pause(self):
        with self._lock:
            self.paused = True
            self.status = "Paused — layout kept."
            self._hide_overlay()
            self._notify()

    def resume(self):
        with self._lock:
            self._require_desktop()
            if not self.running:
                return self.start()
            self.paused = False
            self._reflow_all(include_new=True, explicit=True)
            self.status = self.last_result.summary
            self._notify()

    def stop(self):
        with self._lock:
            self._generation += 1
            self._cancel_pending()
            self._hide_overlay()
            if self._swap_snapshot is not None:
                self.finish_swap(False)
            self.refresh(auto=False)
            failures = []
            for hwnd, (reference, state) in self._originals.items():
                window = self._window(hwnd)
                if window and window.ref == reference and self.backend.alive(hwnd):
                    if not self.backend.restore(hwnd, state):
                        failures.append(hwnd)
            self._originals = {hwnd: value for hwnd, value in self._originals.items() if hwnd in failures}
            self._parked.clear()
            self._hidden_by_template.clear()
            self._manual_minimized.clear()
            self._floating.clear()
            self.running, self.paused = False, False
            self.history.clear()
            self._clear_border()
            self.status = "Windows restored." if not failures else "Stopped; some windows could not be restored."
            self._save()
            self._notify()
            return not failures

    def quit(self):
        if self._closed:
            return
        self._closed = True
        self._stop_event.set()
        self.stop()
        if self._hotkeys:
            self._hotkeys.unregister()
            self._hotkeys.stop()
        self.backend.stop()
        if self._thread and self._thread is not threading.current_thread():
            self._thread.join(timeout=3)

    def _resolve_assignments(self, profile, include_new=False, allow_cross_display=False, allow_maximized=False):
        used = set()
        assignments = deepcopy(profile.assignments)
        protected = {a.window_id for a in assignments if a and a.window_id
                     and self._window(a.window_id) and self._window(a.window_id).app_id == a.app_id
                     and self._eligible(self._window(a.window_id))}
        for slot_index, assignment in enumerate(assignments):
            if assignment is None:
                continue
            candidate = self._window(assignment.window_id)
            if candidate and candidate.ref.hwnd in self._floating:
                if assignment.pinned:
                    used.add(candidate.ref.hwnd)
                else:
                    assignments[slot_index] = None
                continue
            if candidate and (candidate.app_id != assignment.app_id or candidate.ref.hwnd in used or not self._eligible(candidate)):
                candidate = None
            if candidate is None:
                candidates = [w for w in self.windows if w.app_id == assignment.app_id and self._eligible(w)
                              and w.ref.hwnd not in used and w.ref.hwnd not in protected
                              and (allow_cross_display or w.display_id == profile.display_id)]
                candidate = next((w for w in candidates if w.state == "normal"), next(iter(candidates), None))
            assignment.window_id = candidate.ref.hwnd if candidate else None
            if candidate:
                used.add(candidate.ref.hwnd)
        if include_new:
            inactive = {a.window_id for (d, s), p in self.profiles.items()
                        if d == profile.display_id and s != self.active_spaces.get(d, 0)
                        for a in p.assignments if a and a.window_id}
            for window in self.windows:
                hwnd = window.ref.hwnd
                if window.display_id == profile.display_id and self._eligible(window) and (window.state == "normal" or (allow_maximized and window.state == "maximized")) and hwnd not in used and hwnd not in inactive and hwnd not in self._hidden_by_template:
                    empty = next((i for i, a in enumerate(assignments) if a is None), None)
                    value = Assignment(window.app_id, hwnd)
                    if empty is None:
                        assignments.append(value)
                    else:
                        assignments[empty] = value
                    used.add(hwnd)
        return assignments

    def _profile_rects(self, profile, assignments=None):
        display = self._display(profile.display_id)
        slots = profile.assignments if assignments is None else assignments
        padding = self.settings.margins if self.settings.independent_padding else self.settings.padding
        preset, tiles, total = effective_layout(profile, runtime=True, assignments=slots)
        return resolve_layout(display.work_area, total, preset=preset, gap=self.settings.gap,
                              padding=padding, ratio=profile.ratio, tiles=tiles, adaptive=False)

    def resolved_rects(self, display_id, space, draft=False):
        profile = self.draft(display_id, space).profile if draft else self.profile(display_id, space)
        return self._profile_rects(profile)

    def _apply_profile(self, profile, include_new=False, launch_missing=False, hide_surplus=False, compact=False, explicit=False):
        self._display(profile.display_id)
        if not getattr(self.backend,'desktop_known',True):
            raise ValueError('Cannot identify the current Windows desktop; arrangement is suspended.')
        result = ApplyResult()
        blockers=[w for w in self.windows if w.display_id==profile.display_id and
                  (w.state=="fullscreen" or (not explicit and w.state=="maximized" and self._eligible(w)))]
        if not hide_surplus and blockers:
            if explicit:
                result.failures={w.ref.hwnd:"A fullscreen window occupies this display. Leave fullscreen before arranging." for w in blockers}
            self.last_result=result
            return result
        assignments = self._resolve_assignments(profile, include_new=include_new, allow_cross_display=hide_surplus, allow_maximized=explicit)
        # Custom layouts keep tile positions (no compaction) while their tiles can
        # hold the windows; an overflowing custom layout falls back to Auto, which
        # grows, shrinks and compacts like any classic layout.
        custom_fits = profile.preset == "custom" and sum(a is not None for a in assignments) <= len(profile.tiles)
        if not custom_fits and (compact or (self.settings.compact_minimize and include_new)):
            assignments = compact_assignments([
                None if a and not a.pinned and a.window_id in self._manual_minimized else a
                for a in assignments])
        # Custom definitions retain their exact slot geometry unless they overflow.
        if profile.preset == "custom" and len(assignments) <= len(profile.tiles):
            assignments += [None] * (len(profile.tiles) - len(assignments))
        rectangles = self._profile_rects(profile, assignments)
        collapsed = any(t.width < .03 or t.height < .03 for t in profile.resize_tiles)
        if profile.resize_tiles and profile.preset != 'custom' and (collapsed or (
                not self.settings.force_resize and self._resize_too_small(assignments, rectangles))):
            # A stored linked resize that no longer fits these windows is dropped;
            # the normal adaptive layout is used instead.
            log.info('Dropping resized geometry of %s space %s: below application minimum sizes', profile.display_id, profile.space)
            profile.resize_tiles = []
            rectangles = self._profile_rects(profile, assignments)
        log.info('Arrange %s space %s: %s, %s',profile.display_id,profile.space,
                 effective_layout(profile,runtime=True,assignments=assignments)[0],
                 [(a.window_id if a else None,(r.x,r.y,r.width,r.height)) for a,r in zip(assignments,rectangles)])
        while assignments and assignments[-1] is None and not custom_fits:
            assignments.pop()
        if len(rectangles) < len(assignments):
            raise ValueError("This layout does not have enough tiles.")
        selected = {a.window_id for a in assignments if a and a.window_id}
        if hide_surplus:
            for window in self.windows:
                hwnd = window.ref.hwnd
                if window.display_id == profile.display_id and window.state == "normal" and self._eligible(window) and hwnd not in selected:
                    self._capture(hwnd)
                    if self.backend.minimize(hwnd):
                        self._hidden_by_template.add(hwnd)
                        result.hidden.append(hwnd)
        self._suppress_events_until = time.monotonic() + .6
        for index, assignment in enumerate(assignments):
            if assignment is None:
                continue
            hwnd = assignment.window_id
            if hwnd is None:
                key = (profile.display_id, profile.space, index)
                if launch_missing and key not in self._pending:
                    app = next((a for a in self.apps if a.id == assignment.app_id), None)
                    if app is None:
                        result.unavailable.append(assignment.app_id)
                    else:
                        self._pending[key] = {"app_id": app.id, "generation": self._generation, "created": time.monotonic(), "created_utc":time.time()}
                        self._save()
                        if self.backend.launch(app):
                            result.launched.append(app.id)
                        else:
                            self._pending.pop(key, None)
                            result.unavailable.append(app.id)
                continue
            window = self._window(hwnd)
            if not window or not self.backend.alive(hwnd):
                assignment.window_id = None
                continue
            if window.state == "fullscreen":
                result.failures[hwnd] = "Fullscreen window: placement skipped."
                continue
            if hwnd in self._floating:
                continue
            if hwnd in self._manual_minimized and not hide_surplus:
                continue
            self._capture(hwnd)
            self._floating.discard(hwnd)
            self._parked.discard(hwnd)
            self._hidden_by_template.discard(hwnd)
            if hide_surplus:
                self._manual_minimized.discard(hwnd)
            if window.state in ("minimized", "maximized", "hidden"):
                self.backend.show(hwnd)
            if window.display_id != profile.display_id:
                for (display_id, _), other in self.profiles.items():
                    if display_id != profile.display_id:
                        for j, a in enumerate(other.assignments):
                            if a and a.window_id == hwnd:
                                other.assignments[j] = None
            rectangle = rectangles[index]
            self._placed_at[hwnd] = time.monotonic()
            previous_tile = self._tile_rects.get(hwnd)
            if previous_tile and previous_tile[0] == profile.display_id and self.settings.animations and \
                    any(abs(a - b) > 2 for a, b in zip(asdict(previous_tile[1]).values(), asdict(rectangle).values())):
                # A ghost travels from the old tile to the new one.
                self.motion_events = (self.motion_events + [(profile.display_id, previous_tile[1], rectangle,
                                                              window.app_id, time.monotonic())])[-16:]
            self._tile_rects[hwnd] = (profile.display_id, rectangle)
            if self.backend.place(hwnd, rectangle, animate=self.settings.window_animations and self.settings.animations,
                                  duration=self.settings.visual_duration() / 1000,
                                  fps=self.settings.animation_fps, effect=self.settings.animation_effect,
                                  timeout=self.settings.tile_timeout, retries=self.settings.tile_retries,
                                  force=self.settings.force_resize):
                result.placed.append(hwnd)
            else:
                last=getattr(self.backend,'last_result',None)
                log.warning('Placement of %s refused: %s',hwnd,getattr(last,'reason','') or 'unknown')
                result.failures[hwnd] = "The window refused the requested size or placement."
        profile.assignments = assignments
        self.last_result = result
        if result.unavailable:
            names = [next((a.name for a in self.apps if a.id == app_id), app_display_name(app_id)) for app_id in result.unavailable]
            self.last_error = f"Could not open: {', '.join(names)}. Their tiles remain reserved."
        return result

    def _reflow_all(self, include_new=False, compact=False, explicit=False):
        if not getattr(self.backend,'desktop_known',True):
            self.last_error='Cannot identify the current Windows desktop; arrangement is suspended.'
            return
        combined = ApplyResult()
        for display in self.displays:
            space = self.active_spaces.get(display.id, 0)
            profile = self.profile(display.id, space)
            result = self._apply_profile(profile, include_new=include_new, compact=compact, explicit=explicit)
            for field in ("placed", "hidden", "launched", "unavailable"):
                getattr(combined, field).extend(getattr(result, field))
            combined.failures.update(result.failures)
            self.repository.remember_variant(profile)
            key = (display.id, space)
            if key in self.drafts and not self.drafts[key].dirty:
                self.drafts[key] = Draft(deepcopy(profile))
        self.last_result = combined
        self.status = combined.summary
        self._save()
        self._update_border()

    def arrange(self, display_id=None):
        with self._lock:
            self._require_desktop()
            self.refresh(auto=False)
            self.history.push(self._snapshot())
            self.running, self.paused = True, False
            display_id = display_id or self.current_display_id()
            profile = self.profile(display_id, self.active_spaces.get(display_id, 0))
            profile.resize_tiles = []
            for hwnd in list(self._late_unmanaged):
                window=self._window(hwnd)
                if window and window.display_id==display_id:
                    self._floating.discard(hwnd)
                    self._late_unmanaged.discard(hwnd)
            # Explicit Arrange open windows brings back windows hidden by a template.
            for hwnd in list(self._hidden_by_template):
                window = self._window(hwnd)
                if window and window.display_id == display_id:
                    self.backend.show(hwnd)
                    self._hidden_by_template.discard(hwnd)
            self.refresh(auto=False)
            self._apply_profile(profile, include_new=True, compact=True, explicit=True)
            self.repository.remember_variant(profile)
            self._save()
            self.status = self.last_result.summary
            self._notify()
            return self.last_result

    def set_assignment(self, display_id, space, index, assignment):
        with self._lock:
            if index < 0 or index >= 128:
                raise ValueError("Tile index out of range.")
            self._draft_checkpoint(display_id, space)
            profile = self.draft(display_id, space).profile
            while len(profile.assignments) <= index:
                profile.assignments.append(None)
            if assignment and assignment.window_id:
                window = self._window(assignment.window_id)
                if window is None or not window.eligible:
                    raise ValueError("This window is no longer available.")
                for i, a in enumerate(profile.assignments):
                    if i != index and a and a.window_id == assignment.window_id:
                        profile.assignments[i] = None
            profile.assignments[index] = deepcopy(assignment)
            self._draft_changed(display_id, space)

    def set_preset(self, display_id, space, preset):
        with self._lock:
            self._draft_checkpoint(display_id, space)
            draft = self.draft(display_id, space)
            profile = draft.profile
            if preset != "auto" and preset != "custom" and capacity(preset) < sum(a is not None for a in profile.assignments):
                raise ValueError("This preset does not have enough tiles for the assigned apps.")
            variant = self.repository.get_variant(display_id, space, preset)
            if variant is not None and not draft.dirty:
                draft.profile = variant
            else:
                profile.preset = preset
                profile.resize_tiles = []
                if preset == "custom" and not profile.tiles:
                    profile.tiles = [Tile(uuid.uuid4().hex, 0, 0, 1, 1)]
                elif preset != "custom":
                    profile.tiles = []
            self._draft_changed(display_id, space)

    def set_draft_geometry(self, display_id, space, tiles, assignments=None):
        with self._lock:
            if not valid_tiles(tiles):
                raise ValueError("Tiles must cover the screen without gaps or overlaps.")
            profile = deepcopy(self.draft(display_id, space).profile)
            profile.preset, profile.tiles, profile.resize_tiles = "custom", deepcopy(tiles), []
            if assignments is not None:
                profile.assignments = deepcopy(assignments)
            if len(profile.assignments) > len(tiles):
                if any(a for a in profile.assignments[len(tiles):]):
                    raise ValueError("Merged tile: keep which app?")
                profile.assignments = profile.assignments[:len(tiles)]
            profile.assignments += [None] * (len(tiles) - len(profile.assignments))
            SpaceProfile.from_dict(asdict(profile))
            self._draft_checkpoint(display_id, space)
            self.drafts[(display_id, space)].profile = profile
            self._draft_changed(display_id, space)

    def set_draft_profile(self, display_id, space, profile):
        with self._lock:
            value = deepcopy(profile)
            value.display_id, value.space = display_id, space
            SpaceProfile.from_dict(asdict(value))
            self._profile_rects(value)
            self._draft_checkpoint(display_id, space)
            self.drafts[(display_id, space)].profile = value
            self._draft_changed(display_id, space)

    def reset_draft(self, display_id, space):
        with self._lock:
            self.drafts[(display_id, space)] = Draft(deepcopy(self.profile(display_id, space)))
            self._draft_histories.setdefault((display_id, space), History(50)).clear()
            self._notify()

    def set_shared_spaces(self, display_id, space, index, spaces):
        with self._lock:
            if any(s not in (0, 1, 2) for s in spaces):
                raise ValueError("Invalid shared space.")
            source = self.draft(display_id, space).profile
            if not (0 <= index < len(source.assignments)) or source.assignments[index] is None:
                raise ValueError("Choose an app for this tile first.")
            assignment = source.assignments[index]
            if not assignment.window_id:
                raise ValueError("Sharing needs an open window.")
            updated = {}
            for target_space in (0, 1, 2):
                profile = deepcopy(self.draft(display_id, target_space).profile)
                matches = [i for i, a in enumerate(profile.assignments) if a and a.window_id == assignment.window_id]
                if target_space in set(spaces) | {space} and not matches:
                    size = len(self._profile_rects(profile))
                    profile.assignments += [None] * max(0, size - len(profile.assignments))
                    free = index if index < len(profile.assignments) and profile.assignments[index] is None else next((i for i, a in enumerate(profile.assignments) if a is None), None)
                    if free is None:
                        if profile.preset != "auto":
                            raise ValueError(f"Space {target_space + 1} has no free tile.")
                        profile.assignments.append(deepcopy(assignment))
                    else:
                        profile.assignments[free] = deepcopy(assignment)
                elif target_space not in spaces and target_space != space:
                    for i in matches:
                        profile.assignments[i] = None
                updated[target_space] = profile
            for target_space, profile in updated.items():
                if profile != self.draft(display_id, target_space).profile:
                    self._draft_checkpoint(display_id, target_space)
                    self.drafts[(display_id, target_space)].profile = profile
                    self._draft_changed(display_id, target_space)

    def swap_draft(self, display_id, space, a, b):
        with self._lock:
            profile = self.draft(display_id, space).profile
            count = len(self._profile_rects(profile))
            if not (0 <= a < count and 0 <= b < count):
                raise ValueError("Tile unavailable.")
            self._draft_checkpoint(display_id, space)
            profile.assignments += [None] * max(0, count - len(profile.assignments))
            profile.assignments[a], profile.assignments[b] = profile.assignments[b], profile.assignments[a]
            self._draft_changed(display_id, space)

    def clear_space(self, display_id, space):
        with self._lock:
            self._draft_checkpoint(display_id, space)
            profile = self.draft(display_id, space).profile
            profile.assignments = [None] * len(profile.assignments)
            self._draft_changed(display_id, space)

    def undo_draft(self, display_id, space):
        with self._lock:
            draft = self.draft(display_id, space)
            value = self._draft_histories.setdefault((display_id, space), History(50)).undo(draft.profile)
            if value is not None:
                draft.profile = value
                self._draft_changed(display_id, space)

    def redo_draft(self, display_id, space):
        with self._lock:
            draft = self.draft(display_id, space)
            value = self._draft_histories.setdefault((display_id, space), History(50)).redo(draft.profile)
            if value is not None:
                draft.profile = value
                self._draft_changed(display_id, space)

    def apply_drafts(self, contexts=None):
        with self._lock:
            self._require_desktop()
            self.refresh(auto=False)
            dirty = {key: deepcopy(draft.profile) for key, draft in self.drafts.items()
                     if draft.dirty and (contexts is None or key in contexts)
                     and (contexts is not None or key[0] in {d.id for d in self.displays})}
            if not dirty:
                self.last_result=ApplyResult()
                self.status="Nothing to apply. Use Arrange windows to arrange open windows."
                self._notify()
                return self.last_result
            # Validate all contexts first: invalid draft must not half-apply earlier ones.
            for profile in dirty.values():
                self._profile_rects(profile)
            self.history.push(self._snapshot())
            self._generation += 1
            self.running, self.paused = True, False
            combined = ApplyResult()
            for key, profile in dirty.items():
                old = self.profile(*key)
                removed = {a.window_id for a in old.assignments if a and a.window_id} - {a.window_id for a in profile.assignments if a and a.window_id}
                self._cancel_pending(key)
                self.profiles[key] = profile
                other_memberships = {a.window_id for k, p in self.profiles.items() if k != key for a in p.assignments if a and a.window_id}
                for hwnd in removed - other_memberships:
                    self._floating.add(hwnd)
                    self._parked.discard(hwnd)
                    if self.backend.alive(hwnd):
                        self.backend.show(hwnd)
                if profile.space == self.active_spaces.get(profile.display_id, 0):
                    result = self._apply_profile(profile, launch_missing=True, explicit=True)
                    for field in ("placed", "hidden", "launched", "unavailable"):
                        getattr(combined, field).extend(getattr(result, field))
                    combined.failures.update(result.failures)
                self.repository.remember_variant(profile)
                self.drafts[key] = Draft(deepcopy(profile))
                self._draft_histories.setdefault(key, History(50)).clear()
            self._save()
            self.last_result = combined
            self.status = combined.summary if dirty else "Nothing to apply."
            self._notify()
            return self.last_result

    def switch_space(self, display_id, space):
        with self._lock:
            self._require_desktop()
            self._display(display_id)
            if space not in (0, 1, 2):
                raise ValueError("Invalid space.")
            old_space = self.active_spaces.get(display_id, 0)
            if old_space == space:
                return
            self.refresh(auto=False)
            self.history.push(self._snapshot())
            old, new = self.profile(display_id, old_space), self.profile(display_id, space)
            old_ids = {a.window_id for a in old.assignments if a and a.window_id}
            new.assignments = self._resolve_assignments(new)
            new_ids = {a.window_id for a in new.assignments if a and a.window_id}
            for hwnd in old_ids - new_ids:
                window = self._window(hwnd)
                if window and window.state != "minimized" and hwnd not in self._floating:
                    self._capture(hwnd)
                    if self.backend.minimize(hwnd):
                        self._parked.add(hwnd)
            self.active_spaces[display_id] = space
            for hwnd in new_ids:
                if hwnd in self._parked and hwnd not in self._manual_minimized:
                    self.backend.show(hwnd)
                    self._parked.discard(hwnd)
            self.refresh(auto=False)
            if self.running and not self.paused:
                self._apply_profile(new)
            self._save()
            self.space_event = (display_id, space, time.monotonic())
            index = next((i for i, d in enumerate(self.displays) if d.id == display_id), 0)
            self.status = f"Space {space + 1} · Display {index + 1}"
            self._notify()

    def save_profile_template(self, name, profile, template_id=None):
        with self._lock:
            name = name.strip()
            if not name or len(name) > 60:
                raise ValueError("Enter a layout name.")
            if any(t.name.casefold() == name.casefold() and t.id != template_id for t in self.templates):
                raise ValueError("A layout with this name already exists.")
            template = LayoutTemplate(template_id or uuid.uuid4().hex, name, "custom" if profile.resize_tiles else profile.preset,
                                      deepcopy(profile.resize_tiles or profile.tiles), deepcopy(profile.assignments), profile.ratio)
            for assignment in template.assignments:
                if assignment:
                    assignment.window_id = None
            LayoutTemplate.from_dict(asdict(template))
            self.templates = [t for t in self.templates if t.id != template.id] + [template]
            self._save()
            self.status = "Layout saved. The desktop is unchanged."
            self._notify()
            return template

    def save_template(self, name, display_id, space, template_id=None):
        return self.save_profile_template(name, self.draft(display_id, space).profile, template_id)

    def load_template_draft(self, template_id, display_id, space):
        template = next((t for t in self.templates if t.id == template_id), None)
        if not template:
            raise ValueError("Layout not found.")
        return SpaceProfile(display_id, space, template.preset, deepcopy(template.tiles),
                            deepcopy(template.assignments), template.ratio)

    def preview_template(self, template_id, display_id, space):
        with self._lock:
            profile = self.load_template_draft(template_id, display_id, space)
            assignments = self._resolve_assignments(profile, allow_cross_display=True)
            selected = {a.window_id for a in assignments if a and a.window_id}
            known = {a.id for a in self.apps}
            missing = [a.app_id for a in assignments if a and not a.window_id]
            reuse = len(selected)
            opened = sum(app_id in known for app_id in missing)
            hidden = sum(w.display_id == display_id and w.state == "normal" and self._eligible(w) and w.ref.hwnd not in selected for w in self.windows)
            return {"reuse": reuse, "open": opened, "hide": hidden, "unavailable": len(missing) - opened,
                    "profile": profile, "assignments": assignments}

    def restore_template(self, template_id, display_id, space):
        with self._lock:
            return self.restore_profile(self.load_template_draft(template_id,display_id,space))

    def restore_profile(self, profile):
        """Restore a saved or edited template with the same reuse/open/hide policy."""
        with self._lock:
            self._require_desktop()
            self.refresh(auto=False)
            if not self.apps:
                self.refresh_apps()
            profile=deepcopy(profile)
            display_id,space=profile.display_id,profile.space
            self._profile_rects(profile)
            self.history.push(self._snapshot())
            self._generation += 1
            self._cancel_pending((display_id, space))
            self.running, self.paused = True, False
            self.profiles[(display_id, space)] = profile
            self.active_spaces[display_id] = space
            self._apply_profile(profile, launch_missing=True, hide_surplus=True, compact=True, explicit=True)
            self.repository.remember_variant(profile)
            self.drafts[(display_id, space)] = Draft(deepcopy(profile))
            self._save()
            self.status = self.last_result.summary
            self._notify()
            return self.last_result

    def rename_template(self, template_id, name):
        profile = self.load_template_draft(template_id, self.current_display_id(), 0)
        return self.save_profile_template(name, profile, template_id)

    def delete_template(self, template_id):
        with self._lock:
            template = next((t for t in self.templates if t.id == template_id), None)
            if template:
                self._deleted_template = deepcopy(template)
                self.templates.remove(template)
                self._save()
                self._notify()

    def undo_delete_template(self):
        with self._lock:
            if self._deleted_template:
                template = self._deleted_template
                self._deleted_template = None
                if any(t.name.casefold() == template.name.casefold() for t in self.templates):
                    template.name = template.name[:45] + " (restored)"
                self.templates.append(template)
                self._save()
                self._notify()

    def _cancel_pending(self, context=None):
        for key,pending in list(self._pending.items()):
            if context is None or key[:2]==context:
                self._cancelled_launches.append(dict(pending))
                self._pending.pop(key,None)

    def _bind_pending(self):
        if not self.running or not getattr(self.backend,'desktop_known',True): return
        used = {a.window_id for p in self.profiles.values() for a in p.assignments if a and a.window_id}
        now=time.monotonic()
        self._cancelled_launches=[p for p in self._cancelled_launches if now-p["created"]<=120]
        for pending in list(self._cancelled_launches):
            candidate=next((w for w in self.windows if w.app_id==pending['app_id'] and w.ref.hwnd not in used and w.ref.hwnd not in self._late_unmanaged and self._eligible(w)),None)
            if candidate:
                self._late_unmanaged.add(candidate.ref.hwnd)
                self._floating.add(candidate.ref.hwnd)
                candidate.floating=True
                self._cancelled_launches.remove(pending)
        for key, pending in list(self._pending.items()):
            display_id, space, index = key
            if display_id not in {d.id for d in self.displays}: continue
            if time.monotonic()-pending["created"] > 120 and not pending.get('warned'):
                pending['warned']=True
                self.last_error = "The application has not opened a window yet. Its destination remains reserved."
            profile = self.profile(display_id, space)
            if index >= len(profile.assignments):
                self._pending.pop(key, None)
                continue
            assignment = profile.assignments[index]
            if not assignment or assignment.app_id != pending["app_id"]:
                self._pending.pop(key, None)
                continue
            window = next((w for w in self.windows if w.app_id == pending["app_id"] and self._eligible(w) and w.ref.hwnd not in used), None)
            if window is None:
                continue
            assignment.window_id = window.ref.hwnd
            used.add(window.ref.hwnd)
            self._capture(window.ref.hwnd)
            self._pending.pop(key, None)
            if self.active_spaces.get(display_id, 0) == space and self.running and not self.paused:
                self._apply_profile(profile)
            elif self.running and not self.paused and self.active_spaces.get(display_id, 0) != space and self.backend.minimize(window.ref.hwnd):
                self._parked.add(window.ref.hwnd)
            self._save()

    def focus_direction(self, direction):
        with self._lock:
            self._require_desktop()
            if not self.running or self.paused or self._interacting or self._swap_snapshot:
                return
            hwnd = getattr(self.backend, "foreground", lambda: None)()
            display_id = self.current_display_id()
            profile = self.profile(display_id, self.active_spaces.get(display_id, 0))
            index = next((i for i, a in enumerate(profile.assignments) if a and a.window_id == hwnd), None)
            if index is None:
                return
            rects = self._profile_rects(profile)
            available = {i for i, a in enumerate(profile.assignments) if a and a.window_id and self._window(a.window_id) and self._window(a.window_id).state == "normal" and a.window_id not in self._floating}
            target = directional_neighbor(rects, index, direction, available)
            if target is not None:
                self.backend.focus(profile.assignments[target].window_id)
                self._update_border()

    def toggle_float(self, hwnd=None):
        with self._lock:
            self._require_desktop()
            hwnd = hwnd or getattr(self.backend, "foreground", lambda: None)()
            window = self._window(hwnd)
            if not window or not window.eligible:
                raise ValueError("Select an application window.")
            self.history.push(self._snapshot())
            if hwnd in self._floating:
                self._floating.remove(hwnd)
                self._reflow_all(include_new=True)
            else:
                self._capture(hwnd)
                self._floating.add(hwnd)
                self.backend.border(hwnd, None)
                state = self._originals.get(hwnd)
                if state:
                    # Its size from before tiling, centred on its display.
                    centered = getattr(self.backend, "restore_centered", None)
                    display = next((d for d in self.displays if d.id == window.display_id), None)
                    if centered and display:
                        centered(hwnd, state[1], display.work_area)
                    else:
                        self.backend.restore(hwnd, state[1])
                self._reflow_all(include_new=False)
            self._notify()

    def window_action(self, action, reference):
        """Explicit palette actions target a generation, never a recycled HWND."""
        with self._lock:
            self._require_desktop()
            self.refresh(auto=False)
            window=self._window(reference.hwnd)
            if not self.running or not window or window.ref!=reference or not window.eligible:
                return False
            if action=='float':
                self.toggle_float(reference.hwnd)
                return True
            if action not in ('minimize','maximize','close'):
                raise ValueError('Unknown window action')
            self._capture(reference.hwnd)
            if action=='maximize' and window.state=='maximized':
                success=self.backend.show(reference.hwnd)
            else:
                success=getattr(self.backend,action)(reference.hwnd)
            if success:
                self.refresh(auto=True)
            else:
                self.last_error='Windows refused the window action.'
                self._notify()
            return success

    @property
    def focus_color(self):
        if self.settings.border_custom_color:
            return self.settings.border_color
        return getattr(self.backend,'system_accent_color',lambda:'#3584e4')()

    def include_application(self, app_id):
        with self._lock:
            if app_id not in self.settings.included_apps:
                self.settings.included_apps.append(app_id)
            self.settings.excluded_apps = [a for a in self.settings.excluded_apps if a != app_id]
            self.repository.save_settings(self.settings)
            if self.running and not self.paused:
                self._reflow_all(include_new=True)
            self._notify()

    def undo(self):
        with self._lock:
            self._require_desktop()
            snapshot = self.history.undo(self._snapshot())
            if snapshot is not None:
                self._restore_snapshot(snapshot)

    def redo(self):
        with self._lock:
            self._require_desktop()
            snapshot = self.history.redo(self._snapshot())
            if snapshot is not None:
                self._restore_snapshot(snapshot)

    def begin_swap(self):
        with self._lock:
            self._require_desktop()
            if self._swap_snapshot:
                return self.finish_swap(True)
            if not self.running or self.paused:
                raise ValueError("Turn on Arrange windows before swap mode.")
            self._swap_hwnd = getattr(self.backend, "foreground", lambda: None)()
            if self._window(self._swap_hwnd) is None:
                raise ValueError("Select a tiled window.")
            self._swap_snapshot = self._snapshot()
            if self._hotkeys:
                conflicts = self._hotkeys.register((DEFAULT_HOTKEYS | self.settings.hotkeys) | {
                    "swap_left": "Left", "swap_right": "Right", "swap_up": "Up", "swap_down": "Down",
                    "swap_confirm": "Enter", "swap_cancel": "Escape"})
                if any(k.startswith("swap_") for k in conflicts):
                    self.finish_swap(False)
                    raise ValueError("Swap mode keys are already in use.")
            self.status = "Swap: arrows move, Enter accepts, Escape cancels."
            self._update_border()
            self._notify()

    def _swap_context(self):
        window = self._window(self._swap_hwnd)
        if not window:
            return None
        profile = self.profile(window.display_id, self.active_spaces.get(window.display_id, 0))
        index = next((i for i, a in enumerate(profile.assignments) if a and a.window_id == self._swap_hwnd), None)
        if index is None:
            return None
        # Every tile holding a window, pinned or not; an empty (reserved) tile is not a target.
        available = {i for i, a in enumerate(profile.assignments) if a and a.window_id}
        return window, profile, index, available

    def swap_direction(self, direction):
        with self._lock:
            if self._swap_snapshot is None:
                return
            context = self._swap_context()
            if context is None:
                return self.finish_swap(False) if not self._window(self._swap_hwnd) else None
            window, profile, index, available = context
            # Only a window across that edge, never one further away.
            target = swap_neighbor(self._profile_rects(profile), index, direction, available)
            log.info('Swap %s from tile %s of %s: %s', direction, index, len(profile.assignments),
                     'tile %s' % target if target is not None else 'no neighbour')
            if target is not None:
                self._swap_tiles(window, profile, index, target)

    def _swap_tiles(self, window, profile, index, target):
        rects = self._profile_rects(profile)
        other = profile.assignments[target]
        app = next((a for a in self.apps if other and a.id == other.app_id), None)
        label = ('Swap · ' + (app.name if app else app_display_name(other.app_id))) if other and other.window_id else 'Move here'
        # Source and target guides with a label, 360 ms.
        self.swap_event = (window.display_id, rects[index], rects[target], label, time.monotonic())
        profile.assignments[index], profile.assignments[target] = profile.assignments[target], profile.assignments[index]
        self._apply_profile(profile)
        self.backend.focus(self._swap_hwnd)
        self._notify()

    def finish_swap(self, commit=True):
        with self._lock:
            if self._swap_snapshot is None:
                return
            snapshot = self._swap_snapshot
            self._swap_snapshot, self._swap_hwnd = None, None
            if commit:
                self.history.push(snapshot)
                self._save()
            else:
                self._restore_snapshot(snapshot)
            if self._hotkeys:
                self.hotkey_conflicts = self._hotkeys.register(DEFAULT_HOTKEYS | self.settings.hotkeys)
            self._update_border()
            self._notify()

    def resize_tiles(self, display_id, space, tiles):
        with self._lock:
            if not valid_tiles(tiles):
                raise ValueError("Invalid resized geometry.")
            self.history.push(self._snapshot())
            profile = self.profile(display_id, space)
            profile.resize_tiles = deepcopy(tiles)
            if self.running and not self.paused and self.active_spaces.get(display_id, 0) == space:
                self._apply_profile(profile)
            self._save()
            self._notify()

    def _resize_too_small(self, assignments, rectangles):
        minimum = getattr(self.backend, 'min_size', None)
        for assignment, rectangle in zip(assignments, rectangles):
            if not assignment or not assignment.window_id:
                continue
            width, height = minimum(assignment.window_id) if minimum else (1, 1)
            if rectangle.width < max(width, 40) or rectangle.height < max(height, 40):
                return True
        return False

    def _reachable_minimums(self, rects, minimums):
        """Minimum tile sizes for a linked resize.

        With forced tiling, application minimum sizes do not apply: tiles keep
        a small usable size. Otherwise a tile already below its application
        minimum never shrinks further, instead of refusing the whole gesture."""
        if self.settings.force_resize:
            minimums = [FORCED_MINIMUM] * len(minimums)
        return [(min(m[0], r.width), min(m[1], r.height)) for r, m in zip(rects, minimums)]

    def _native_resize(self, hwnd, before, after, snapshot):
        window = self._window(hwnd)
        if not window:
            return
        profile = self.profile(window.display_id, self.active_spaces.get(window.display_id, 0))
        index = next((i for i, a in enumerate(profile.assignments) if a and a.window_id == hwnd), None)
        if index is None:
            return
        rects = self._profile_rects(profile)
        min_sizes = [getattr(self.backend, "min_size", lambda h: (80, 80))(a.window_id) if a and a.window_id else (3, 3) for a in profile.assignments]
        min_sizes += [(3, 3)] * (len(rects) - len(min_sizes))
        min_sizes = self._reachable_minimums(rects, min_sizes)
        changed = rects
        for edge, delta in (("left", after.x - before.x), ("right", after.right - before.right),
                            ("top", after.y - before.y), ("bottom", after.bottom - before.bottom)):
            if abs(delta) > 5:
                changed = linked_resize(changed, index, edge, delta, min_sizes=min_sizes)
        if snapshot:
            self.history.push(snapshot)
        # Normalize shared virtual edges including gaps against the actual inner bounds.
        gap = self.settings.gap
        display = self._display(window.display_id)
        padding = self.settings.margins if self.settings.independent_padding else dict.fromkeys(("left", "right", "top", "bottom"), self.settings.padding)
        left, top = display.work_area.x + padding["left"], display.work_area.y + padding["top"]
        right, bottom = display.work_area.right - padding["right"], display.work_area.bottom - padding["bottom"]
        width, height = right - left, bottom - top
        tiles = []
        for i, rectangle in enumerate(changed):
            x = rectangle.x - left - (gap // 2 if rectangle.x > left else 0)
            y = rectangle.y - top - (gap // 2 if rectangle.y > top else 0)
            r = rectangle.right - left + (gap - gap // 2 if rectangle.right < right else 0)
            b = rectangle.bottom - top + (gap - gap // 2 if rectangle.bottom < bottom else 0)
            tiles.append(Tile(str(i), x / width, y / height, (r - x) / width, (b - y) / height))
        # Never store a collapsed geometry (a resize computed against stale slots
        # produced a 3 px tile): the Studio's 3 % minimum applies here too.
        if valid_tiles(tiles) and all(t.width >= .03 and t.height >= .03 for t in tiles):
            profile.resize_tiles = tiles
        self._apply_profile(profile)
        self._save()
        self._notify()

    def _native_drop(self, hwnd, old_display, target_display):
        window = self._window(hwnd)
        if not window:
            return
        target = self.profile(target_display, self.active_spaces.get(target_display, 0))
        source = self.profile(old_display, self.active_spaces.get(old_display, 0))
        source_index = next((i for i, a in enumerate(source.assignments) if a and a.window_id == hwnd), None)
        if source_index is None:
            return
        self.history.push(self._resize_snapshot or self._snapshot())
        if old_display != target_display:
            value = source.assignments[source_index]
            for (display_id, _), profile in self.profiles.items():
                if display_id == old_display:
                    profile.assignments = [None if a and a.window_id == hwnd else a for a in profile.assignments]
            target.assignments.append(value)
            self._apply_profile(source, compact=True)
            self._apply_profile(target, include_new=True, compact=True)
        else:
            rects = self._profile_rects(target)
            x, y = window.rect.x + window.rect.width / 2, window.rect.y + window.rect.height / 2
            index = nearest_slot(rects, x, y)
            if index is not None and index != source_index:
                target.assignments += [None] * max(0, len(rects) - len(target.assignments))
                target.assignments[source_index], target.assignments[index] = target.assignments[index], target.assignments[source_index]
            self._apply_profile(target)
        self._hide_overlay()
        self._save()
        self._notify()

    def _is_tiled(self, hwnd):
        """A window placed in a tile of the active space of its display."""
        window = self._window(hwnd)
        if not window or not self.running or self.paused or hwnd in self._floating or not self._eligible(window):
            return False
        profile = self.profile(window.display_id, self.active_spaces.get(window.display_id, 0))
        return any(a and a.window_id == hwnd for a in profile.assignments)

    def _preview_native_drag(self, hwnd):
        if not self.running or self.paused:
            return
        window = self._window(hwnd)
        if not window:
            return
        rect=self.backend.window_rect(hwnd)
        if not rect:
            return
        x,y=rect.x+rect.width/2,rect.y+rect.height/2
        display=next((d for d in self.displays if d.work_area.x<=x<d.work_area.right and d.work_area.y<=y<d.work_area.bottom),None)
        window=replace(window,rect=rect,display_id=display.id if display else window.display_id)
        self.windows=[window if w.ref==window.ref else w for w in self.windows]
        profile = self.profile(window.display_id, self.active_spaces.get(window.display_id, 0))
        rects = self._profile_rects(profile)
        original = self._move_origins.get(hwnd)
        resizing=self._gesture_kinds.get(hwnd)!="move"
        if resizing and original and original[1]==window.display_id and (abs(window.rect.width-original[0].width)>8 or abs(window.rect.height-original[0].height)>8):
            index=next((i for i,a in enumerate(profile.assignments) if a and a.window_id==hwnd),None)
            if index is not None:
                before,after=original[0],window.rect
                minimums=[]
                for assignment in profile.assignments:
                    if assignment and assignment.window_id:
                        handle=assignment.window_id
                        if handle not in self._interaction_minimums:
                            self._interaction_minimums[handle]=self.backend.min_size(handle)
                        minimums.append(self._interaction_minimums[handle])
                    else:
                        minimums.append((3,3))
                minimums += [(3,3)]*(len(rects)-len(minimums))
                minimums=self._reachable_minimums(rects,minimums)
                for edge,delta in (("left",after.x-before.x),("right",after.right-before.right),
                                   ("top",after.y-before.y),("bottom",after.bottom-before.bottom)):
                    if abs(delta)>5:
                        rects=linked_resize(rects,index,edge,delta,min_sizes=minimums)
                self.preview_rectangles=[(window.display_id,r) for r in rects]
                self._notify_guides()
                return
        self.preview_rectangles=[]
        index = nearest_slot(rects, window.rect.x + window.rect.width / 2, window.rect.y + window.rect.height / 2)
        if index is None:
            guide = None
        else:
            target = profile.assignments[index] if index < len(profile.assignments) else None
            if target and target.window_id == hwnd:
                label = None
            elif target and target.window_id:
                app = next((a for a in self.apps if a.id == target.app_id), None)
                label = 'Swap · ' + (app.name if app else app_display_name(target.app_id))
            else:
                label = 'Move here'
            guide = (window.display_id, rects[index], label, window.app_id)
        if guide != self.drag_guide:
            self.drag_guide = guide
            self._notify_guides()

    def _hide_overlay(self):
        changed=bool(self.preview_rectangles) or self.drag_guide is not None
        self.preview_rectangles=[]
        self.drag_guide=None
        if changed:
            self._notify_guides()

    def swap_hints(self):
        """(display_id, window rect, [(direction, target, x, y, primary)]) for the window in swap mode.

        One arrow per neighbouring window, centred on the edge they share, so
        every possible swap is shown, including several neighbours on one side.
        primary marks the neighbour the arrow key swaps with.
        """
        if self._swap_snapshot is None or not self._swap_hwnd:
            return None
        context = self._swap_context()
        if context is None or context[0].state != 'normal':
            return None
        window, profile, index, available = context
        rects = self._profile_rects(profile)
        source = rects[index]
        arrows = []
        for direction in ('left', 'right', 'up', 'down'):
            keyed = swap_neighbor(rects, index, direction, available)
            for target, centre in edge_neighbors(rects, index, direction, available):
                if direction in ('left', 'right'):
                    x, y = (source.x if direction == 'left' else source.right), centre
                else:
                    x, y = centre, (source.y if direction == 'up' else source.bottom)
                arrows.append((direction, target, x, y, target == keyed))
        return window.display_id, source, arrows

    def reserved_slots(self):
        """Empty pinned tiles and their state: MINIMIZED, FLOATING,
        OPENING, OPEN ELSEWHERE or CLOSED. Hidden while a fullscreen window
        occupies the display."""
        result = []
        if not self.running or self.paused:
            return result
        for display in self.displays:
            if any(w.display_id == display.id and w.state == 'fullscreen' for w in self.windows):
                continue
            space = self.active_spaces.get(display.id, 0)
            profile = self.profile(display.id, space)
            rectangles = self._profile_rects(profile)
            for index, assignment in enumerate(profile.assignments):
                if not assignment or not assignment.pinned or index >= len(rectangles):
                    continue
                window = self._window(assignment.window_id)
                if window is not None and window.state != 'minimized' and assignment.window_id not in self._floating:
                    continue
                if window is not None:
                    state = 'FLOATING' if assignment.window_id in self._floating else 'MINIMIZED'
                elif (display.id, space, index) in self._pending:
                    state = 'OPENING'
                elif self._open_elsewhere(assignment.app_id, display.id, space):
                    state = 'OPEN ELSEWHERE'
                else:
                    state = 'CLOSED'
                result.append({"display_id": display.id, "space": space, "index": index, "rect": rectangles[index],
                               "app_id": assignment.app_id, "state": state})
        return result

    def _open_elsewhere(self, app_id, display_id, space):
        """A window of the app on another display or in another space of this one.
        Other windows of the same app tiled in this space do not count: the
        pinned app itself is closed, and clicking its card opens it."""
        here = {a.window_id for a in self.profile(display_id, space).assignments if a and a.window_id}
        other_spaces = {a.window_id for s in range(3) if s != space
                        for a in self.profile(display_id, s).assignments if a and a.window_id}
        return any(w.app_id == app_id and self._eligible(w) and w.ref.hwnd not in here and
                   (w.display_id != display_id or w.ref.hwnd in other_spaces) for w in self.windows)

    def activate_reserved(self, display_id, space, index):
        with self._lock:
            self._require_desktop()
            self.history.push(self._snapshot())
            profile = self.profile(display_id, space)
            assignment = profile.assignments[index]
            if assignment and assignment.window_id:
                self._floating.discard(assignment.window_id)
                self._manual_minimized.discard(assignment.window_id)
                self.backend.show(assignment.window_id)
            if not self.apps:
                self.refresh_apps()
            self._apply_profile(profile, launch_missing=True, hide_surplus=False)
            self._save()
            self._notify()

    def _clear_border(self):
        if self._last_border:
            self.backend.border(self._last_border, None)
        self._last_border = None

    def _update_border(self):
        if not self.running or self.paused or not getattr(self.backend,'desktop_known',True):
            return self._clear_border()
        focused = self._swap_hwnd or getattr(self.backend, "foreground", lambda: None)()
        # Opening the tray menu activates the taskbar and then SmartGrid's own
        # menu; the outline stays on the window that was focused.
        if focused and not self._swap_hwnd and getattr(self.backend, "is_shell_surface", lambda h: False)(focused):
            focused = self._last_border
        tiled = {a.window_id for d in self.displays for a in self.profile(d.id, self.active_spaces.get(d.id, 0)).assignments if a and a.window_id and a.window_id not in self._floating}
        if focused in tiled:
            self._last_tiled_selection = focused
        else:
            focused = None
        # Swap mode always outlines its window: the swap arrows are drawn on it.
        if not self._swap_snapshot and (not self.settings.border_width or not self.settings.active_border):
            focused = None
        color="#ff9b91" if self._swap_snapshot else self.focus_color
        if self._last_border != focused or getattr(self,'_last_border_color',None)!=color:
            self._clear_border()
            if focused:
                # The Qt outline (thickness, halo) replaces the 1 px DWM border when a UI is present.
                if getattr(self,'native_border',True):
                    self.backend.border(focused, color)
                self._last_border = focused
                self._last_border_color=color

    def update_settings(self, settings):
        with self._lock:
            settings = Settings.from_dict(asdict(settings))
            newly_excluded=(set(settings.excluded_apps)-set(settings.included_apps))-(set(self.settings.excluded_apps)-set(self.settings.included_apps))
            restore_failures=[]
            for window in self.windows:
                if window.app_id in newly_excluded:
                    saved=self._originals.get(window.ref.hwnd)
                    if saved and saved[0]==window.ref and not self.backend.restore(window.ref.hwnd,saved[1]):
                        restore_failures.append(window.ref.hwnd)
            if settings.master_ratio!=self.settings.master_ratio:
                for profile in self.profiles.values():
                    profile.ratio=settings.master_ratio
            self.settings = settings
            self.repository.save_settings(settings)
            if self._hotkeys:
                self.hotkey_conflicts = self._hotkeys.register(DEFAULT_HOTKEYS | settings.hotkeys)
            self._clear_border()
            if self.running and not self.paused:
                self.refresh(auto=False)
                self._reflow_all()
            if restore_failures:
                self.last_error='Windows refused to restore floating applications: '+', '.join(map(str,restore_failures))
            self._notify()

    def export_archive(self, path):
        export_archive(path, self.templates, self.profiles, self.repository.variants)

    def import_archive(self, path):
        with self._lock:
            result = import_archive(path, self.templates, self.profiles, self.repository.variants)
            self.templates, self.profiles = result["templates"], result["profiles"]
            self.repository.variants = result["variants"]
            self._save()
            self.status = f"Imported {result['added']} layouts and {result['profiles_added']} profiles"
            self._notify()
            return result

    def handle_action(self, action):
        if action in ("studio", "quick_switcher", "preferences", "quit"):
            if self._ui_action:
                self._ui_action(action)
            elif action == "quit":
                self.quit()
            return
        if action == "toggle":
            return self.stop() if self.running and not self.paused else self.resume()
        if action == "pause":
            return self.pause()
        if action == "arrange":
            return self.arrange()
        if action == "swap":
            return self.begin_swap()
        if action == "float":
            return self.toggle_float()
        if action == "undo":
            return self.undo()
        if action == "redo":
            return self.redo()
        if action == "stop":
            return self.stop()
        if action == "recycle_bin":
            import os
            if os.name == "nt":
                os.startfile("shell:RecycleBinFolder")
            return
        if action.startswith("focus_"):
            return self.focus_direction(action[6:])
        if action.startswith("space") and action[-1:] in "123":
            self._require_desktop()
            return self.switch_space(self.current_display_id(), int(action[-1]) - 1)
        if action.startswith("swap_"):
            if action == "swap_confirm":
                return self.finish_swap(True)
            if action == "swap_cancel":
                return self.finish_swap(False)
            return self.swap_direction(action[5:])
