"""Bounded, isolated snapshots for application and per-context draft history."""
from copy import deepcopy


class History:
    def __init__(self, limit: int = 10):
        if not isinstance(limit, int) or isinstance(limit, bool) or limit < 1:
            raise ValueError('History limit must be a positive integer')
        self.limit = limit
        self._undo = []
        self._redo = []

    @property
    def can_undo(self) -> bool:
        return bool(self._undo)

    @property
    def can_redo(self) -> bool:
        return bool(self._redo)

    @property
    def undo_count(self) -> int:
        return len(self._undo)

    @property
    def redo_count(self) -> int:
        return len(self._redo)

    def push(self, snapshot):
        self._undo.append(deepcopy(snapshot))
        del self._undo[:-self.limit]
        self._redo.clear()

    def undo(self, current):
        if not self._undo:
            return None
        self._redo.append(deepcopy(current))
        return deepcopy(self._undo.pop())

    def redo(self, current):
        if not self._redo:
            return None
        self._undo.append(deepcopy(current))
        del self._undo[:-self.limit]
        return deepcopy(self._redo.pop())

    def clear(self):
        self._undo.clear()
        self._redo.clear()
