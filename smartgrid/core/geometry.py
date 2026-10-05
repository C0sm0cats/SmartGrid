"""One pure geometry engine for previews, placement, hit testing and editing.

Rectangles use physical integer coordinates. Custom tiles describe a complete
partition of the normalized work area; gaps are applied only at resolution.
"""
from __future__ import annotations

from copy import deepcopy
from dataclasses import replace
import math
import re
from typing import Iterable, Mapping

from .models import Rect, Tile

EPS = 1e-9
MIN_TILE_SIZE = 1e-5
_ALIASES = {'full': '1x1', 'side_by_side': '2x1', 'master': 'master_stack'}


def presets() -> list[tuple[str, str]]:
    return [('auto', 'Auto'), ('1x1', 'Full'), ('2x1', 'Split'),
            ('master_stack', 'Focus')] + [
        (f'{cols}x{rows}', f'{cols} × {rows}')
        for cols, rows in ((2, 2), (3, 2), (3, 3), (4, 3), (5, 3), (4, 4), (5, 4), (5, 5))
    ] + [('custom', 'Custom')]


def preset_name(preset: str) -> str:
    """Display name for a layout identifier."""
    preset = _ALIASES.get(preset, preset)
    return dict(presets()).get(preset) or preset.replace('x', ' × ')


def auto_preset(count: int) -> str:
    _count(count)
    if count <= 1:
        return '1x1'
    if count == 2:
        return '2x1'
    if count == 3:
        return 'master_stack'
    for size, preset in ((4, '2x2'), (6, '3x2'), (9, '3x3'), (12, '4x3'),
                         (15, '5x3'), (16, '4x4'), (20, '5x4'), (25, '5x5')):
        if count <= size:
            return preset
    cols = math.ceil(math.sqrt(count))
    return f'{cols}x{math.ceil(count / cols)}'


def effective_layout(profile, runtime=False, assignments=None):
    """One layout decision for the desktop and its Studio preview.

    Classic templates keep their editable capacity, but their live arrangement
    grows and shrinks. Interior holes and pins still reserve their positions.
    """
    slots = profile.assignments if assignments is None else assignments
    requested = len(slots)
    if runtime:
        while requested and slots[requested - 1] is None:
            requested -= 1
    # A native linked resize applies to the window count it was made for.
    # With fewer or more windows
    # (minimize, close, new window) the normal layout is used and the resized
    # geometry returns when the count matches again.
    resized = bool(profile.resize_tiles) and requested == len(profile.resize_tiles)
    tiles = profile.resize_tiles if resized else profile.tiles
    preset = 'custom' if resized else profile.preset
    if preset == 'custom' and requested <= len(tiles):
        return preset, tiles, len(tiles)
    if runtime or preset == 'auto' or requested > capacity(preset, tiles):
        preset, tiles = auto_preset(requested), []
        return preset, tiles, capacity(preset) if requested or not runtime else 0
    return preset, tiles, max(requested, capacity(preset, tiles))


def _count(count):
    if not isinstance(count, int) or isinstance(count, bool) or count < 0:
        raise ValueError('Window count must be a nonnegative integer')


def _grid(preset: str) -> tuple[int, int]:
    match = re.fullmatch(r'([1-9][0-9]*)x([1-9][0-9]*)', _ALIASES.get(preset, preset))
    if not match:
        raise ValueError(f'Unknown layout preset: {preset}')
    return tuple(map(int, match.groups()))


def capacity(preset: str, tiles=None) -> int:
    preset = _ALIASES.get(preset, preset)
    if preset == 'auto':
        return 0  # Dynamic capacity, determined from the number of assignments.
    if preset == 'custom':
        if not valid_tiles(tiles or []):
            raise ValueError('Custom tiles must cover the work area without overlaps')
        return len(tiles)
    if preset == 'master_stack':
        return 3
    cols, rows = _grid(preset)
    return cols * rows


def layout_tiles(count: int, preset='auto', ratio=.6, tiles=None, adaptive=False) -> list[Tile]:
    """Return normalized tiles, including full preset capacity on request."""
    _count(count)
    if count == 0:
        return []
    preset = _ALIASES.get(preset, preset)
    if preset == 'auto':
        preset = auto_preset(count)
    if preset == 'custom':
        if not valid_tiles(tiles or []):
            raise ValueError('Custom tiles must cover the work area without overlaps')
        if count > len(tiles):
            raise ValueError('Custom layout capacity exceeded; choose Auto for overflow')
        return deepcopy(list(tiles))[:count]
    if preset == 'master_stack':
        if not isinstance(ratio, (float, int)) or not math.isfinite(ratio) or not 0 < ratio < 1:
            raise ValueError('Master ratio must be between zero and one')
        if count > 3:
            raise ValueError('Master layout capacity exceeded')
        return [Tile('master', 0, 0, ratio, 1), Tile('stack-1', ratio, 0, 1-ratio, .5),
                Tile('stack-2', ratio, .5, 1-ratio, .5)][:count]
    cols, rows = _grid(preset)
    if count > cols * rows:
        raise ValueError('Layout capacity exceeded; choose Auto for overflow')
    result = []
    for row in range(rows):
        row_count = min(cols, count - row * cols)
        if row_count <= 0:
            break
        effective_cols = row_count if adaptive else cols
        for col in range(row_count):
            result.append(Tile(f'{row}:{col}', col/effective_cols, row/rows,
                               1/effective_cols, 1/rows))
    return result


def resolve_layout(area: Rect, count: int, preset='auto', gap=8, padding=8,
                   ratio=.6, tiles=None, adaptive=False) -> list[Rect]:
    _count(count)
    if count == 0:
        return []
    if not isinstance(area, Rect) or any(type(v) is not int for v in (area.x, area.y, area.width, area.height)):
        raise ValueError('Work area must use integer physical coordinates')
    if not isinstance(gap, int) or isinstance(gap, bool) or gap < 0:
        raise ValueError('Gap must be a nonnegative integer')
    margins = dict.fromkeys(('top', 'right', 'bottom', 'left'), padding)
    if isinstance(padding, Mapping):
        if set(padding) != set(margins):
            raise ValueError('Padding requires top, right, bottom and left')
        margins = dict(padding)
    if any(not isinstance(v, int) or isinstance(v, bool) or v < 0 for v in margins.values()):
        raise ValueError('Padding must contain nonnegative integers')
    inner = Rect(area.x + margins['left'], area.y + margins['top'],
                 area.width - margins['left'] - margins['right'],
                 area.height - margins['top'] - margins['bottom'])
    if inner.width <= 0 or inner.height <= 0:
        raise ValueError('Padding leaves no usable work area')
    normalized = layout_tiles(count, preset, ratio, tiles, adaptive)
    result = []
    for t in normalized:
        x = inner.x + round(t.x * inner.width) + (gap//2 if t.x > EPS else 0)
        y = inner.y + round(t.y * inner.height) + (gap//2 if t.y > EPS else 0)
        right = inner.x + round(t.right * inner.width) - (gap-gap//2 if t.right < 1-EPS else 0)
        bottom = inner.y + round(t.bottom * inner.height) - (gap-gap//2 if t.bottom < 1-EPS else 0)
        if right <= x or bottom <= y:
            raise ValueError('Work area too small for layout, gaps and padding')
        result.append(Rect(x, y, right-x, bottom-y))
    return result


def valid_tiles(tiles: Iterable[Tile]) -> bool:
    """Bounded, non-overlapping rectangles must sum to the complete unit area."""
    try:
        tiles = list(tiles)
        if not tiles or len({t.id for t in tiles}) != len(tiles):
            return False
        for t in tiles:
            if not isinstance(t.id, str) or not t.id:
                return False
            if not all(isinstance(v, (int, float)) and not isinstance(v, bool) and math.isfinite(v)
                       for v in (t.x, t.y, t.width, t.height)):
                return False
            if (t.x < -EPS or t.y < -EPS or t.width < MIN_TILE_SIZE-EPS or
                    t.height < MIN_TILE_SIZE-EPS or t.right > 1+EPS or t.bottom > 1+EPS):
                return False
        for i, a in enumerate(tiles):
            for b in tiles[i+1:]:
                if min(a.right, b.right)-max(a.x, b.x) > EPS and min(a.bottom, b.bottom)-max(a.y, b.y) > EPS:
                    return False
        return abs(math.fsum(t.width*t.height for t in tiles)-1) <= EPS
    except (AttributeError, TypeError, ValueError):
        return False


def _validated(tiles, index):
    if not valid_tiles(tiles):
        raise ValueError('Tiles must form a complete non-overlapping partition')
    if not isinstance(index, int) or not 0 <= index < len(tiles):
        raise IndexError('Tile index out of range')


def _fresh_id(tiles, base):
    existing = {t.id for t in tiles}
    index = 1
    while f'{base}-{index}' in existing:
        index += 1
    return f'{base}-{index}'


def split_tile(tiles, index, axis) -> list[Tile]:
    _validated(tiles, index)
    if axis not in ('horizontal', 'vertical'):
        raise ValueError('Split axis must be horizontal or vertical')
    result = deepcopy(list(tiles))
    tile = result[index]
    new_id = _fresh_id(result, tile.id)
    if axis == 'vertical':
        first = replace(tile, width=tile.width/2)
        second = replace(tile, id=new_id, x=tile.x+tile.width/2, width=tile.width/2)
    else:
        first = replace(tile, height=tile.height/2)
        second = replace(tile, id=new_id, y=tile.y+tile.height/2, height=tile.height/2)
    result[index:index+1] = [first, second]
    if not valid_tiles(result):
        raise ValueError('Split would create tiles below the minimum size')
    return result


def merge_tiles(tiles, index, other) -> list[Tile]:
    _validated(tiles, index)
    _validated(tiles, other)
    if index == other:
        raise ValueError('Choose two different tiles to merge')
    a, b = tiles[index], tiles[other]
    horizontal = abs(a.y-b.y) < EPS and abs(a.height-b.height) < EPS and (
        abs(a.right-b.x) < EPS or abs(b.right-a.x) < EPS)
    vertical = abs(a.x-b.x) < EPS and abs(a.width-b.width) < EPS and (
        abs(a.bottom-b.y) < EPS or abs(b.bottom-a.y) < EPS)
    if not (horizontal or vertical):
        raise ValueError('Merged tiles must share a complete edge and form a rectangle')
    merged = Tile(a.id, min(a.x, b.x), min(a.y, b.y),
                  max(a.right, b.right)-min(a.x, b.x), max(a.bottom, b.bottom)-min(a.y, b.y))
    return [deepcopy(merged if i == index else t) for i, t in enumerate(tiles) if i != other]


def _overlap(a, b):
    return min(a[1], b[1]) - max(a[0], b[0]) > EPS


def move_edge(tiles, index, edge, delta) -> list[Tile]:
    """Move a connected shared edge, preserving T-junctions and full coverage.

    Detached segments on the same axis remain independently editable. The
    requested movement is clamped at neighboring minimum sizes and outer edges.
    """
    _validated(tiles, index)
    if edge not in ('left', 'right', 'top', 'bottom'):
        raise ValueError('Unknown tile edge')
    if not isinstance(delta, (int, float)) or isinstance(delta, bool) or not math.isfinite(delta):
        raise ValueError('Edge movement must be finite')
    vertical = edge in ('left', 'right')
    low, high = ('left', 'right') if vertical else ('top', 'bottom')
    def coord(t, side):
        return (t.x if side == 'left' else t.right if side == 'right' else
                t.y if side == 'top' else t.bottom)
    def interval(t):
        return (t.y, t.bottom) if vertical else (t.x, t.right)
    line = coord(tiles[index], edge)
    if line < EPS or line > 1-EPS:
        return deepcopy(list(tiles))
    candidates = [(i, side) for i, t in enumerate(tiles) for side in (low, high)
                  if abs(coord(t, side)-line) < EPS]
    selected = {(index, edge)}
    changed = True
    while changed:
        changed = False
        for node in candidates:
            if node not in selected and any(_overlap(interval(tiles[node[0]]), interval(tiles[n[0]]))
                                            for n in selected):
                selected.add(node)
                changed = True
    lower, upper = -math.inf, math.inf
    for i, side in selected:
        t = tiles[i]
        if side == low:
            upper = min(upper, coord(t, high)-line-MIN_TILE_SIZE)
        else:
            lower = max(lower, coord(t, low)-line+MIN_TILE_SIZE)
    delta = max(lower, min(upper, delta))
    result = deepcopy(list(tiles))
    for i, side in selected:
        t = result[i]
        if side == 'left':
            t.x += delta
            t.width -= delta
        elif side == 'right':
            t.width += delta
        elif side == 'top':
            t.y += delta
            t.height -= delta
        else:
            t.height += delta
    if not valid_tiles(result):
        raise ValueError('Edge movement cannot preserve complete coverage')
    return result


def set_tile_geometry(tiles, index, **kwargs) -> list[Tile]:
    _validated(tiles, index)
    if set(kwargs) - {'x', 'y', 'width', 'height'}:
        raise ValueError('Geometry fields are x, y, width and height')
    original = tiles[index]
    target = replace(original, **kwargs)
    if (not all(isinstance(v, (float, int)) and math.isfinite(v)
                for v in (target.x, target.y, target.width, target.height)) or
            target.width < MIN_TILE_SIZE or target.height < MIN_TILE_SIZE or
            target.x < 0 or target.y < 0 or target.right > 1+EPS or target.bottom > 1+EPS):
        raise ValueError('Requested geometry is outside the normalized work area')
    result = deepcopy(list(tiles))
    # Expanding first avoids a transient collapsed tile when translating.
    operations = [('left', target.x-original.x), ('right', target.right-original.right),
                  ('top', target.y-original.y), ('bottom', target.bottom-original.bottom)]
    operations.sort(key=lambda op: ((op[1] > 0) if op[0] in ('left', 'top') else (op[1] < 0)))
    for edge, _ in operations:
        current = result[index]
        value = {'left': target.x, 'right': target.right, 'top': target.y, 'bottom': target.bottom}[edge]
        present = {'left': current.x, 'right': current.right, 'top': current.y, 'bottom': current.bottom}[edge]
        result = move_edge(result, index, edge, value-present)
    if any(abs(getattr(result[index], f)-getattr(target, f)) > EPS for f in ('x', 'y', 'width', 'height')):
        raise ValueError('Requested geometry conflicts with work area bounds or neighboring tiles')
    return result


def nearest_slot(rects, x, y) -> int | None:
    if not rects:
        return None
    for i, rect in enumerate(rects):
        if rect.contains(x, y):
            return i
    return min(range(len(rects)), key=lambda i: (
        max(rects[i].x-x, 0, x-rects[i].right)**2 +
        max(rects[i].y-y, 0, y-rects[i].bottom)**2, i))


def directional_neighbor(rects, index, direction, available=None) -> int | None:
    if direction not in ('left', 'right', 'up', 'down', 'top', 'bottom'):
        raise ValueError('Direction must be left, right, up or down')
    if not 0 <= index < len(rects):
        return None
    allowed = set(range(len(rects))) if available is None else set(available)
    source = rects[index]
    cx, cy = source.x+source.width/2, source.y+source.height/2
    vertical = direction in ('up', 'down', 'top', 'bottom')
    sign = -1 if direction in ('left', 'up', 'top') else 1
    candidates = []
    for i, r in enumerate(rects):
        if i == index or i not in allowed:
            continue
        rx, ry = r.x+r.width/2, r.y+r.height/2
        primary = sign * ((ry-cy) if vertical else (rx-cx))
        if primary <= EPS:
            continue
        source_span = (source.x, source.right) if vertical else (source.y, source.bottom)
        target_span = (r.x, r.right) if vertical else (r.y, r.bottom)
        aligned = _overlap(source_span, target_span)
        orthogonal = abs(rx-cx) if vertical else abs(ry-cy)
        edge_distance = max(0, (r.y-source.bottom if sign > 0 else source.y-r.bottom) if vertical
                            else (r.x-source.right if sign > 0 else source.x-r.right))
        candidates.append(((not aligned, edge_distance, orthogonal, primary, i), i))
    return min(candidates)[1] if candidates else None


EDGE_SLACK = 2


def edge_neighbors(rects, index, direction, available=None) -> list[tuple[int, float]]:
    """Every tile directly across one edge of a tile, with the centre of the shared span.

    A tile can have several neighbours on one side (the large tile of Focus,
    custom layouts, uneven grids). Only tiles that share part of the edge and
    have nothing in between are returned, as (index, centre along the edge).
    """
    if direction not in ('left', 'right', 'up', 'down'):
        raise ValueError('Direction must be left, right, up or down')
    if not 0 <= index < len(rects):
        return []
    allowed = set(range(len(rects))) if available is None else set(available)
    source = rects[index]
    horizontal = direction in ('left', 'right')

    def distance(r):
        return {'right': r.x - source.right, 'left': source.x - r.right,
                'down': r.y - source.bottom, 'up': source.y - r.bottom}[direction]

    def span(r):
        return (max(source.y, r.y), min(source.bottom, r.bottom)) if horizontal else \
               (max(source.x, r.x), min(source.right, r.right))

    # A few pixels of tolerance: rounded fractional tiles can touch or overlap by 1 px.
    facing = [(i, r) for i, r in enumerate(rects) if i != index and distance(r) >= -EDGE_SLACK
              and span(r)[1] - span(r)[0] > EDGE_SLACK]
    result = []
    for i, r in facing:
        if i not in allowed:
            continue
        low, high = span(r)
        # Blocked when another facing tile sits closer over the same stretch.
        if any(j != i and distance(o) < distance(r) - EDGE_SLACK and min(high, span(o)[1]) - max(low, span(o)[0]) > EDGE_SLACK
               for j, o in facing):
            continue
        result.append((i, (low + high) / 2))
    return sorted(result, key=lambda item: item[1])


def swap_neighbor(rects, index, direction, available=None) -> int | None:
    """The tile an arrow key swaps with: a tile across that edge, never one beyond.

    With several neighbours on one side, the one sharing the longest stretch
    of the edge, then the one closest to the centre of the edge.
    """
    neighbors = edge_neighbors(rects, index, direction, available)
    if not neighbors:
        return None
    source = rects[index]
    horizontal = direction in ('left', 'right')
    middle = source.y + source.height / 2 if horizontal else source.x + source.width / 2

    def shared(i):
        r = rects[i]
        return (min(source.bottom, r.bottom) - max(source.y, r.y)) if horizontal else \
               (min(source.right, r.right) - max(source.x, r.x))
    return min(neighbors, key=lambda item: (-shared(item[0]), abs(item[1] - middle), item[0]))[0]


def linked_resize(rects, index, edge, delta, min_sizes=None) -> list[Rect]:
    """Resize shared dividers with bounded cascading to adjacent minimum sizes.

    Pixel gaps stay constant. Moving a divider can push the next divider when a
    neighbor reaches its minimum; outer work-area edges stay fixed. At the last
    constraint, movement is clamped. Corners accept (dx, dy) or a scalar delta.
    """
    rects = list(rects)
    if not isinstance(index, int) or not 0 <= index < len(rects):
        raise IndexError('Rectangle index out of range')
    corner = edge.replace('_', '').replace('-', '')
    aliases = {'topleft': ('left', 'top'), 'topright': ('right', 'top'),
               'bottomleft': ('left', 'bottom'), 'bottomright': ('right', 'bottom')}
    if corner in aliases:
        dx, dy = delta if isinstance(delta, (tuple, list)) and len(delta) == 2 else (delta, delta)
        first, second = aliases[corner]
        return linked_resize(linked_resize(rects, index, first, dx, min_sizes), index, second, dy, min_sizes)
    if edge not in ('left', 'right', 'top', 'bottom'):
        raise ValueError('Unknown resize edge')
    if not isinstance(delta, (int, float)) or isinstance(delta, bool) or not math.isfinite(delta):
        raise ValueError('Resize movement must be finite')
    if any(r.width < 1 or r.height < 1 for r in rects):
        raise ValueError('Resize requires positive rectangles')
    minimums = [(1, 1)] * len(rects)
    if min_sizes is not None:
        if isinstance(min_sizes, Mapping):
            minimums = [min_sizes.get(i, (1, 1)) for i in range(len(rects))]
        else:
            if len(min_sizes) != len(rects):
                raise ValueError('A minimum size is required for each rectangle')
            minimums = list(min_sizes)
    for r, m in zip(rects, minimums):
        if (not isinstance(m, (tuple, list)) or len(m) != 2 or
                any(not isinstance(v, (int, float)) or isinstance(v, bool) or not math.isfinite(v) or v <= 0 for v in m)):
            raise ValueError('Minimum sizes must be positive (width, height) pairs')
        if r.width < math.ceil(m[0]) or r.height < math.ceil(m[1]):
            raise ValueError('Existing layout cannot satisfy application minimum sizes')
    vertical = edge in ('left', 'right')
    def interval(r):
        return (r.y, r.bottom) if vertical else (r.x, r.right)
    starts = [r.x if vertical else r.y for r in rects]
    ends = [r.right if vertical else r.bottom for r in rects]
    neighbors = [(i, j, starts[j]-ends[i]) for i, a in enumerate(rects) for j, b in enumerate(rects)
                 if i != j and _overlap(interval(a), interval(b)) and starts[j] >= ends[i]]
    gap = min((distance for _, _, distance in neighbors), default=0)
    parent = list(range(2*len(rects)))
    def find(a):
        while parent[a] != a:
            parent[a] = parent[parent[a]]
            a = parent[a]
        return a
    def union(a, b):
        parent[find(b)] = find(a)
    for i, j, distance in neighbors:
        if distance == gap:
            union(2*i+1, 2*j)
    groups = [(find(2*i), find(2*i+1)) for i in range(len(rects))]
    fixed = {g for i, pair in enumerate(groups) for g, position in zip(pair, (starts[i], ends[i]))
             if position == min(starts) or position == max(ends)}
    target = groups[index][0 if edge in ('left', 'top') else 1]
    if target in fixed or round(delta) == 0:
        return list(rects)
    constraints = [(left, right, ends[i]-starts[i]-math.ceil(minimums[i][0 if vertical else 1]))
                   for i, (left, right) in enumerate(groups)]
    all_groups = set(g for pair in groups for g in pair)
    positive = delta > 0
    def displacements(distance):
        values = dict.fromkeys(all_groups, 0)
        values[target] = distance if positive else -distance
        # Difference constraints with nonnegative slack converge in at most V
        # passes. This propagates only the displacement required at a minimum.
        for _ in range(len(all_groups)+1):
            changed = False
            for left, right, slack in constraints:
                if values[left]-values[right] <= slack:
                    continue
                node = right if positive else left
                required = values[left]-slack if positive else values[right]+slack
                if node in fixed or node == target:
                    return None
                values[node] = required
                changed = True
            if not changed:
                return values
        return None
    requested = abs(round(delta))
    values = displacements(requested)
    if values is None:
        lo, hi = 0, requested
        while lo < hi:
            mid = (lo+hi+1)//2
            if displacements(mid) is None:
                hi = mid-1
            else:
                lo = mid
        values = displacements(lo)
    result = []
    for i, r in enumerate(rects):
        left, right = groups[i]
        start, end = starts[i]+values[left], ends[i]+values[right]
        result.append(Rect(start, r.y, end-start, r.height) if vertical else
                      Rect(r.x, start, r.width, end-start))
    for i, a in enumerate(result):
        for b in result[i+1:]:
            if _overlap((a.x, a.right), (b.x, b.right)) and _overlap((a.y, a.bottom), (b.y, b.bottom)):
                raise ValueError('Resize cannot satisfy the layout without overlap')
    return result
