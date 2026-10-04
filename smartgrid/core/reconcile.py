"""Deterministic lifecycle reconciliation, stable app matching and reservations."""
from __future__ import annotations

from copy import deepcopy
from typing import Iterable, Mapping

from .models import Assignment, SpaceProfile, WindowRecord


def compact_assignments(assignments: Iterable[Assignment | None]) -> list[Assignment | None]:
    """Pack movable occupants around immovable pinned and pending reservations.

    List capacity and pinned indices never change. Pending assignments reserve
    their destinations independently of pin state; there is no invisible loss
    when several windows belonging to the same application are assigned.
    """
    assignments = deepcopy(list(assignments))
    fixed = {i for i, a in enumerate(assignments) if a is not None and (a.pinned or a.window_id is None)}
    movable = [a for i, a in enumerate(assignments) if a is not None and i not in fixed]
    result = [a if i in fixed else None for i, a in enumerate(assignments)]
    free = [i for i in range(len(result)) if i not in fixed]
    for i, assignment in zip(free, movable):
        result[i] = assignment
    return result


def reconcile_assignments(assignments: Iterable[Assignment | None], windows: Iterable[WindowRecord],
                          compact: bool = True, include_unassigned: bool = False,
                          preserve_pending: bool = True) -> list[Assignment | None]:
    """Retain matching live occupants, then fill reservations by stable HWND.

    Preserve valid explicit bindings before matching any app-only reservation.
    This prevents a newly appeared window from stealing a sibling's slot. Dead
    pins always become app reservations. With compaction disabled dead regular
    occupants become holes; their numeric slot does not move. Pending non-pins
    can be removed explicitly via preserve_pending=False, e.g. cancellation.
    Minimized records remain live: caller controls minimize compaction policy.
    """
    result = deepcopy(list(assignments))
    candidates = sorted((w for w in windows if w.eligible and not w.floating),
                        key=lambda w: (w.app_id, w.ref.hwnd, w.ref.pid, w.ref.generation))
    by_hwnd = {w.ref.hwnd: w for w in candidates}
    used = set()
    needs_matching = []
    for i, a in enumerate(result):
        if a is None:
            continue
        live = by_hwnd.get(a.window_id)
        if live is not None and live.app_id == a.app_id and live.ref.hwnd not in used:
            used.add(live.ref.hwnd)
            continue
        originally_pending = a.window_id is None
        a.window_id = None
        needs_matching.append((i, originally_pending))
    for i, originally_pending in needs_matching:
        a = result[i]
        match = next((w for w in candidates if w.app_id == a.app_id and w.ref.hwnd not in used), None)
        if match is not None:
            a.window_id = match.ref.hwnd
            used.add(match.ref.hwnd)
        elif not (a.pinned or (originally_pending and preserve_pending)):
            result[i] = None
    if compact:
        result = compact_assignments(result)
    if include_unassigned:
        for w in candidates:
            if w.ref.hwnd in used:
                continue
            assignment = Assignment(w.app_id, w.ref.hwnd)
            try:
                slot = result.index(None)
            except ValueError:
                result.append(assignment)
            else:
                result[slot] = assignment
            used.add(w.ref.hwnd)
    return result


def reconcile_profile(profile: SpaceProfile, windows: Iterable[WindowRecord], compact: bool = True,
                      include_unassigned: bool = False, preserve_pending: bool = True) -> SpaceProfile:
    result = deepcopy(profile)
    result.assignments = reconcile_assignments(profile.assignments,
        (w for w in windows if w.display_id == profile.display_id), compact,
        include_unassigned, preserve_pending)
    return result


def memberships(profiles: Mapping[object, SpaceProfile] | Iterable[SpaceProfile]) -> dict[int, set[tuple[str, int]]]:
    """Represent shared membership without duplicating or moving assignments."""
    result = {}
    values = profiles.values() if isinstance(profiles, Mapping) else profiles
    for profile in values:
        for assignment in profile.assignments:
            if assignment is not None and assignment.window_id is not None:
                result.setdefault(assignment.window_id, set()).add((profile.display_id, profile.space))
    return result
