import unittest
from smartgrid.core.history import History
from smartgrid.core.models import Assignment, SpaceProfile


class HistoryTests(unittest.TestCase):
    def test_undo_redo_isolates_mutations_and_new_edits_clear_redo(self):
        history = History(2)
        first = SpaceProfile('display',0,assignments=[Assignment('app',1)])
        history.push(first)
        first.assignments[0].window_id = 2
        restored = history.undo(first)
        self.assertEqual(restored.assignments[0].window_id,1)
        restored.assignments[0].window_id = 9
        self.assertEqual(history.redo(restored).assignments[0].window_id,2)
        self.assertEqual(history.undo(first).assignments[0].window_id,9)
        history.push({'fresh':True})
        self.assertFalse(history.can_redo)
        self.assertEqual(history.undo({}),{'fresh':True})

    def test_limit_empty_and_clear(self):
        history = History(2)
        self.assertIsNone(history.undo(0))
        self.assertIsNone(history.redo(0))
        for i in range(3):
            history.push(i)
        self.assertEqual(history.undo(3),2)
        self.assertEqual(history.undo(2),1)
        self.assertIsNone(history.undo(1))
        self.assertTrue(history.can_redo)
        history.clear()
        self.assertFalse(history.can_redo)
        self.assertFalse(history.can_undo)
        for limit in (0,-1,True,1.5):
            with self.assertRaises(ValueError):
                History(limit)
