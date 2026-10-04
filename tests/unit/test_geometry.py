import math
import random
import unittest

from smartgrid.core.geometry import (auto_preset, capacity, directional_neighbor,
    layout_tiles, linked_resize, merge_tiles, move_edge, nearest_slot, presets,
    resolve_layout, set_tile_geometry, split_tile, valid_tiles)
from smartgrid.core.models import Rect, Tile


class GeometryTests(unittest.TestCase):
    def assert_nonoverlap(self, rects):
        for i, a in enumerate(rects):
            for b in rects[i+1:]:
                self.assertFalse(min(a.right, b.right) > max(a.x, b.x) and
                                 min(a.bottom, b.bottom) > max(a.y, b.y))

    def test_auto_has_no_ceiling_and_preserves_legacy_choices(self):
        self.assertEqual([auto_preset(n) for n in (1,2,3,4,6,9,12,15)],
                         ['1x1', '2x1', 'master_stack', '2x2', '3x2', '3x3', '4x3', '5x3'])
        for area in (Rect(-1920, 30, 1920, 1050), Rect(0,0,1080,1920), Rect(0,0,3440,1440)):
            for count in (0,1,2,3,4,6,9,12,15,16,20,25,30,42,100,173):
                rects = resolve_layout(area, count)
                self.assertEqual(len(rects), count)
                for r in rects:
                    self.assertGreater(r.width, 0)
                    self.assertGreater(r.height, 0)
                    self.assertGreaterEqual(r.x, area.x+8)
                    self.assertGreaterEqual(r.y, area.y+8)
                    self.assertLessEqual(r.right, area.right-8)
                    self.assertLessEqual(r.bottom, area.bottom-8)
                self.assert_nonoverlap(rects)

    def test_explicit_capacity_and_partial_presets(self):
        self.assertIn(('5x5', '5 × 5'), presets())
        self.assertEqual(capacity('5x5'), 25)
        for preset in ('full', 'side_by_side', 'master_stack', '5x5'):
            count = capacity(preset)
            self.assertEqual(len(resolve_layout(Rect(0,0,2000,1400), count, preset)), count)
        single = resolve_layout(Rect(0,0,1000,1000), 1, '2x2')
        self.assertEqual(single, [Rect(8,8,488,488)])
        with self.assertRaisesRegex(ValueError, 'capacity'):
            resolve_layout(Rect(0,0,1000,1000), 5, '2x2')

    def test_padding_and_rounding_no_pixel_leak(self):
        rects = resolve_layout(Rect(-500,-200,1001,803), 25, '5x5', gap=7,
            padding={'top':5,'right':11,'bottom':13,'left':17})
        self.assertEqual(min(r.x for r in rects), -483)
        self.assertEqual(max(r.right for r in rects), 490)
        self.assertEqual(min(r.y for r in rects), -195)
        self.assertEqual(max(r.bottom for r in rects), 590)
        self.assert_nonoverlap(rects)
        for row in range(5):
            for col in range(4):
                self.assertEqual(rects[row*5+col+1].x-rects[row*5+col].right, 7)
        for gap,pad in ((20,100),(500,0),(0,160)):
            with self.assertRaises(ValueError):
                resolve_layout(Rect(0,0,320,240),15,'5x3', gap=gap,padding=pad)
        with self.assertRaises(ValueError):
            resolve_layout(Rect(0,0,100,100),2,gap=-1)

    def test_custom_validation_detects_holes_overlap_nan_and_duplicate_ids(self):
        base = [Tile('a',0,0,.5,1), Tile('b',.5,0,.5,1)]
        self.assertTrue(valid_tiles(base))
        for bad in ([], [base[0]], [base[0], Tile('b',.4,0,.5,1)],
                    [base[0], Tile('a',.5,0,.5,1)], [Tile('a',math.nan,0,1,1)],
                    [Tile('a',0,0,1,math.inf)], [Tile('a',0,0,-1,1)]):
            self.assertFalse(valid_tiles(bad))
        self.assertTrue(valid_tiles(layout_tiles(30, '6x5')))
        self.assertEqual(len(resolve_layout(Rect(0,0,1920,1080),30,'custom',tiles=layout_tiles(30,'6x5'))),30)
        with self.assertRaises(ValueError):
            resolve_layout(Rect(0,0,1000,1000),3,'custom',tiles=base)

    def test_repeated_splits_merges_preserve_partition_and_ids(self):
        tiles = [Tile('first',0,0,1,1)]
        rng = random.Random(772)
        for _ in range(29):
            i = rng.randrange(len(tiles))
            tiles = split_tile(tiles,i,'vertical' if tiles[i].width >= tiles[i].height else 'horizontal')
            self.assertTrue(valid_tiles(tiles))
        self.assertEqual(len(tiles),30)
        self.assertEqual(len(set(t.id for t in tiles)),30)
        tiles = split_tile([Tile('first',0,0,1,1)],0,'vertical')
        tiles = merge_tiles(tiles,0,1)
        self.assertEqual(tiles,[Tile('first',0,0,1,1)])
        master = layout_tiles(3,'master_stack')
        with self.assertRaisesRegex(ValueError,'complete edge'):
            merge_tiles(master,0,1)
        self.assertEqual(merge_tiles(master,1,2),[master[0],Tile('stack-1',.6,0,.4,1)])

    def test_shared_t_junction_edge_propagates_to_both_neighbors(self):
        base = layout_tiles(3,'master_stack')
        moved = move_edge(base,0,'right',.15)
        self.assertTrue(valid_tiles(moved))
        self.assertAlmostEqual(moved[0].width,.75)
        self.assertAlmostEqual(moved[1].x,.75)
        self.assertAlmostEqual(moved[2].x,.75)
        self.assertEqual(base, layout_tiles(3,'master_stack'))
        clamped = move_edge(base,0,'right',9)
        self.assertTrue(valid_tiles(clamped))
        self.assertAlmostEqual(clamped[1].width,1e-5)
        self.assertEqual(move_edge(base,0,'left',.1),base)

    def test_disconnected_segments_and_arbitrary_nonslicing_partition(self):
        base = layout_tiles(4,'2x2')
        moved = move_edge(base,0,'right',.1)
        self.assertTrue(valid_tiles(moved))
        self.assertAlmostEqual(moved[0].width,.6)
        self.assertEqual(moved[2],base[2])
        pinwheel = [Tile('a',0,0,.7,.3), Tile('b',.7,0,.3,.7),
                    Tile('c',.3,.7,.7,.3), Tile('d',0,.3,.3,.7), Tile('e',.3,.3,.4,.4)]
        self.assertTrue(valid_tiles(pinwheel))
        for i, edge in ((0,'bottom'),(1,'left'),(2,'top'),(3,'right')):
            self.assertTrue(valid_tiles(move_edge(pinwheel,i,edge,.04)))

    def test_geometry_fields_preserve_partition_and_reject_outer_conflicts(self):
        base = layout_tiles(3,'master_stack')
        result = set_tile_geometry(base,0,width=.7)
        self.assertTrue(valid_tiles(result))
        self.assertAlmostEqual(result[1].width,.3)
        with self.assertRaises(ValueError):
            set_tile_geometry(base,0,x=.1)
        with self.assertRaises(ValueError):
            set_tile_geometry(base,0,width=-.1)

    def test_focus_and_hit_test_use_effective_rectangles(self):
        rects = [Rect(-100,0,100,200),Rect(10,0,100,100),Rect(10,110,100,90)]
        self.assertEqual(directional_neighbor(rects,0,'right'),1)
        self.assertEqual(directional_neighbor(rects,0,'right',available={1}),1)
        self.assertEqual(directional_neighbor(rects,1,'down'),2)
        self.assertEqual(directional_neighbor(rects,2,'left'),0)
        self.assertIsNone(directional_neighbor(rects,0,'left'))
        self.assertEqual(nearest_slot(rects, 5,150),0)
        self.assertEqual(nearest_slot(rects, 25,180),2)
        self.assertEqual(nearest_slot(rects, 110,200),2)
        self.assertIsNone(nearest_slot([],0,0))

    def test_linked_resize_cascades_minima_and_clamps_outer_bounds(self):
        rects = [Rect(0,0,300,500),Rect(310,0,300,500),Rect(620,0,300,500)]
        resized = linked_resize(rects,0,'right',250,[(100,100)]*3)
        self.assertEqual(resized,[Rect(0,0,550,500),Rect(560,0,100,500),Rect(670,0,250,500)])
        maximum = linked_resize(rects,0,'right',10000,[(100,100)]*3)
        self.assertEqual(maximum,[Rect(0,0,700,500),Rect(710,0,100,500),Rect(820,0,100,500)])
        self.assertEqual(linked_resize(rects,1,'left',-1000,[(100,100)]*3)[0].width,100)
        self.assert_nonoverlap(maximum)
        self.assertEqual(linked_resize(rects,0,'left',100),rects)
        with self.assertRaisesRegex(ValueError,'minimum sizes'):
            linked_resize(rects,0,'right',10,[(301,100)]*3)

    def test_linked_t_junction_corners_and_unequal_minimums(self):
        rects = resolve_layout(Rect(0,0,1000,800),3,'master_stack',gap=8,padding=0)
        resized = linked_resize(rects,0,'right',100,[(10,10),(150,10),(200,10)])
        self.assertEqual(resized[1].x,resized[2].x)
        self.assertEqual(resized[0].right+8,resized[1].x)
        self.assert_nonoverlap(resized)
        grid = resolve_layout(Rect(0,0,800,800),4,'2x2',gap=8,padding=0)
        corner = linked_resize(grid,0,'bottom_right',(35,40))
        self.assertEqual(corner[0].width,grid[0].width+35)
        self.assertEqual(corner[0].height,grid[0].height+40)
        self.assert_nonoverlap(corner)


if __name__ == '__main__':
    unittest.main()


class ResizedLayoutCountTests(unittest.TestCase):
    def test_linked_resize_applies_only_to_its_window_count(self):
        from smartgrid.core.models import Assignment,SpaceProfile,Tile
        from smartgrid.core.geometry import effective_layout
        tiles=[Tile('a',0,0,.55,.5),Tile('b',.55,0,.45,.5),Tile('c',0,.5,.55,.5),Tile('d',.55,.5,.45,.5)]
        profile=SpaceProfile('d',0,assignments=[Assignment(str(i),i+1) for i in range(4)],resize_tiles=tiles)
        self.assertEqual(effective_layout(profile,runtime=True)[0],'custom')
        profile.assignments=profile.assignments[:3]
        self.assertEqual(effective_layout(profile,runtime=True)[0],'master_stack')
