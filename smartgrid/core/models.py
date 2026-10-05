"""Serializable application state. This module has no UI or platform imports."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field, fields
from copy import deepcopy
from typing import Any, Mapping, TypeVar
import math
import re


class Serializable:
    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Mapping[str, Any]):
        return model_from_dict(cls, data)


@dataclass(frozen=True)
class Rect(Serializable):
    x: int
    y: int
    width: int
    height: int

    @property
    def right(self) -> int:
        return self.x + self.width

    @property
    def bottom(self) -> int:
        return self.y + self.height

    def contains(self, x: float, y: float) -> bool:
        return self.x <= x < self.right and self.y <= y < self.bottom


@dataclass
class Tile(Serializable):
    id: str
    x: float
    y: float
    width: float
    height: float

    @property
    def right(self) -> float:
        return self.x + self.width

    @property
    def bottom(self) -> float:
        return self.y + self.height


@dataclass
class Display(Serializable):
    id: str
    name: str
    work_area: Rect
    dpi: int = 96
    primary: bool = False


@dataclass
class ApplicationRef(Serializable):
    id: str
    name: str
    executable: str = ''
    aumid: str = ''
    icon_path: str = ''
    installed: bool = False


@dataclass(frozen=True)
class WindowRef(Serializable):
    hwnd: int
    pid: int = 0
    generation: str = ''


@dataclass
class WindowRecord(Serializable):
    ref: WindowRef
    app_id: str
    title: str
    rect: Rect
    display_id: str
    state: str = 'normal'
    eligible: bool = True
    exclusion_reason: str = ''
    floating: bool = False
    original: dict | None = None


@dataclass
class Assignment(Serializable):
    app_id: str
    window_id: int | None = None
    pinned: bool = False


@dataclass
class LayoutTemplate(Serializable):
    id: str
    name: str
    preset: str = '2x2'
    tiles: list[Tile] = field(default_factory=list)
    assignments: list[Assignment | None] = field(default_factory=list)
    ratio: float = .6


@dataclass
class SpaceProfile(Serializable):
    display_id: str
    space: int
    preset: str = 'auto'
    tiles: list[Tile] = field(default_factory=list)
    assignments: list[Assignment | None] = field(default_factory=list)
    ratio: float = .6
    resize_tiles: list[Tile] = field(default_factory=list)


@dataclass
class Draft(Serializable):
    profile: SpaceProfile
    dirty: bool = False


@dataclass
class Settings(Serializable):
    gap: int = 12
    padding: int = 12
    independent_padding: bool = False
    margins: dict[str, int] = field(default_factory=lambda: dict(top=12, right=12, bottom=12, left=12))
    master_ratio: float = .6
    theme: str = 'system'
    accent: str = '#8ce8c3'
    animations: bool = True
    window_animations: bool = False
    animation_duration: int = 140
    animation_effect: str = 'crit_damped'
    animation_fps: int = 60
    sounds: bool = False  # no longer offered; kept so that existing settings files still load
    compact_minimize: bool = True
    compact_close: bool = True
    # No longer offered; kept so that existing settings files still load.
    start_active: bool = False
    show_last_selection: bool = False
    debounce_ms: int = 100
    tile_timeout: float = 2.0
    tile_retries: int = 3
    excluded_apps: list[str] = field(default_factory=list)
    included_apps: list[str] = field(default_factory=list)
    hotkeys: dict[str, str] = field(default_factory=dict)
    border_width: int = 2
    border_style: str = 'outline'
    active_border: bool = True
    border_custom_color: bool = False
    border_color: str = '#3584e4'
    animation_speed: str = 'normal'
    animation_curve: str = 'ease-out'
    force_resize: bool = True
    builtin_exclusions: bool = True
    # Changes to the built-in overlay keywords; the list itself stays in code.
    overlay_words_added: list[str] = field(default_factory=list)
    overlay_words_removed: list[str] = field(default_factory=list)
    # Applications added to the overlay list (app IDs).
    overlay_apps_added: list[str] = field(default_factory=list)

    def visual_duration(self, base=140):
        duration={'fast':90,'normal':140,'slow':220}.get(self.animation_speed,self.animation_duration)
        return round(base*duration/140) if self.animations else 0


def app_display_name(app_id: str) -> str:
    """Readable fallback for an application missing from the catalogue."""
    kind, _, value = app_id.partition(':')
    if kind == 'exe' and value:
        name = re.split(r'[\\/]', value)[-1]
        return name.rsplit('.', 1)[0] if name.lower().endswith(('.exe', '.com')) else name
    if kind == 'aumid' and value:
        package = value.split('!')[0].split('_')[0]
        return package.split('.')[-1] or value
    return app_id


T = TypeVar('T')


def model_from_dict(cls: type[T], data: Mapping[str, Any]) -> T:
    """Decode nested dataclasses; reject unknown fields rather than lose archive data."""
    if not isinstance(data, Mapping):
        raise ValueError(f'{cls.__name__} must be an object')
    names = {f.name for f in fields(cls)}
    unknown = set(data) - names
    if unknown:
        raise ValueError(f'Unknown {cls.__name__} fields: {", ".join(sorted(unknown))}')
    values = deepcopy(dict(data))
    nested = {
        Display: {'work_area': Rect},
        WindowRecord: {'ref': WindowRef, 'rect': Rect},
        Draft: {'profile': SpaceProfile},
    }
    for name, kind in nested.get(cls, {}).items():
        if name in values and not isinstance(values[name], kind):
            values[name] = kind.from_dict(values[name])
    if cls in (LayoutTemplate, SpaceProfile):
        for name in ('tiles', 'resize_tiles'):
            if name in values:
                if not isinstance(values[name], list):
                    raise ValueError(f'{name} must be a list')
                values[name] = [v if isinstance(v, Tile) else Tile.from_dict(v) for v in values[name]]
        if 'assignments' in values:
            if not isinstance(values['assignments'], list):
                raise ValueError('assignments must be a list')
            values['assignments'] = [v if v is None or isinstance(v, Assignment) else Assignment.from_dict(v)
                                     for v in values['assignments']]
    try:
        result = cls(**values)
        _validate_model(result)
        return result
    except (TypeError, ValueError) as exc:
        raise ValueError(f'Invalid {cls.__name__}: {exc}') from exc


def settings_from_dict(data: Mapping[str, Any]) -> Settings:
    return Settings.from_dict(data)


def profile_from_dict(data: Mapping[str, Any]) -> SpaceProfile:
    return SpaceProfile.from_dict(data)


def template_from_dict(data: Mapping[str, Any]) -> LayoutTemplate:
    return LayoutTemplate.from_dict(data)


def _validate_model(value):
    """Strict boundary validation for persisted state and untrusted archives."""
    def number(name, lo, hi, integer=False):
        v = getattr(value, name)
        kind = type(v) is int if integer else type(v) in (int, float)
        if not kind or not lo <= v <= hi or not math.isfinite(v):
            raise ValueError(f'{name} must be {"an integer" if integer else "a number"} between {lo} and {hi}')
    def string(name, empty=False, maximum=4096):
        v = getattr(value, name)
        if not isinstance(v, str) or len(v) > maximum or (not empty and not v.strip()):
            raise ValueError(f'{name} must be a {"possibly empty " if empty else "nonempty "}string')
    def boolean(name):
        if type(getattr(value, name)) is not bool:
            raise ValueError(f'{name} must be a boolean')
    if isinstance(value, Rect):
        for name in ('x', 'y'):
            number(name, -(2**31), 2**31-1, True)
        # Empty native rectangles are representable; the geometry resolver and
        # backend placement enforce strictly positive placement dimensions.
        for name in ('width', 'height'):
            number(name, 0, 2**31-1, True)
    elif isinstance(value, Tile):
        string('id', maximum=1024)
        for name in ('x', 'y'):
            number(name, 0, 1)
        for name in ('width', 'height'):
            number(name, 1e-5, 1)
        if value.right > 1+1e-9 or value.bottom > 1+1e-9:
            raise ValueError('Tile bounds extend outside the normalized work area')
    elif isinstance(value, Display):
        string('id', maximum=1024)
        string('name')
        number('dpi', 24, 1536, True)
        boolean('primary')
        _validate_model(value.work_area)
    elif isinstance(value, ApplicationRef):
        string('id', maximum=1024)
        string('name')
        for name in ('executable', 'aumid', 'icon_path'):
            string(name, empty=True, maximum=32768)
    elif isinstance(value, WindowRef):
        number('hwnd', 1, 2**64-1, True)
        number('pid', 0, 2**32-1, True)
        string('generation', empty=True)
    elif isinstance(value, WindowRecord):
        string('app_id', maximum=1024)
        string('title', empty=True, maximum=32768)
        string('display_id', maximum=1024)
        string('state')
        string('exclusion_reason', empty=True)
        boolean('eligible')
        boolean('floating')
        _validate_model(value.ref)
        _validate_model(value.rect)
        if value.original is not None and not isinstance(value.original, dict):
            raise ValueError('original must be an object or null')
    elif isinstance(value, Assignment):
        string('app_id', maximum=1024)
        if value.window_id is not None:
            number('window_id', 1, 2**64-1, True)
        boolean('pinned')
    elif isinstance(value, (LayoutTemplate, SpaceProfile)):
        if isinstance(value, LayoutTemplate):
            string('id', maximum=1024)
            string('name')
        else:
            string('display_id', maximum=1024)
            number('space', 0, 2, True)
        string('preset', maximum=64)
        number('ratio', 1e-5, 1-1e-5)
        from .geometry import capacity, valid_tiles
        for name in ('tiles', 'resize_tiles'):
            if hasattr(value, name) and getattr(value, name) and not valid_tiles(getattr(value, name)):
                raise ValueError(f'{name} must completely cover the work area without overlaps')
        for assignment in value.assignments:
            if assignment is not None:
                _validate_model(assignment)
        for tile in value.tiles + getattr(value, 'resize_tiles', []):
            _validate_model(tile)
        cap = capacity(value.preset, value.tiles)
        # Auto assignments are unbounded, but explicit layouts cannot discard
        # occupied/reserved assignments outside their capacity.
        if isinstance(value, LayoutTemplate) and cap and len(value.assignments) > cap and any(value.assignments[cap:]):
            raise ValueError('Assignments exceed the layout capacity')
    elif isinstance(value, Draft):
        boolean('dirty')
        _validate_model(value.profile)
    elif isinstance(value, Settings):
        for name in ('gap', 'padding'):
            number(name, 0, 10000, True)
        for name in ('independent_padding', 'animations', 'window_animations', 'sounds',
                     'compact_minimize', 'compact_close', 'start_active', 'show_last_selection','active_border','border_custom_color','force_resize','builtin_exclusions'):
            boolean(name)
        for name, lo, hi in (('animation_duration', 0, 10000), ('animation_fps', 1, 240),
                             ('debounce_ms', 0, 10000), ('tile_retries', 0, 20), ('border_width', 0, 32)):
            number(name, lo, hi, True)
        number('tile_timeout', .01, 120)
        number('master_ratio', .25, .75)
        if value.theme not in ('system', 'light', 'dark'):
            raise ValueError('Unknown theme')
        for color in (value.accent,value.border_color):
            if not isinstance(color, str) or not re.fullmatch(r'#[0-9a-fA-F]{6}', color):
                raise ValueError('Color must be a #RRGGBB color')
        if value.animation_speed not in ('fast','normal','slow','custom') or value.animation_curve not in ('ease-out','linear','ease-in-out'):
            raise ValueError('Unknown visual animation setting')
        if value.animation_effect not in ('crit_damped', 'spring', 'curved', 'curve', 'linear'):
            raise ValueError('Unknown animation effect')
        if value.border_style not in ('outline', 'solid', 'glow'):
            raise ValueError('Unknown border style')
        if not isinstance(value.margins, dict) or set(value.margins) != {'top', 'right', 'bottom', 'left'}:
            raise ValueError('Margins require top, right, bottom and left')
        if any(type(v) is not int or not 0 <= v <= 10000 for v in value.margins.values()):
            raise ValueError('Margins must contain nonnegative integer distances')
        for name in ('excluded_apps', 'included_apps', 'overlay_apps_added'):
            items = getattr(value, name)
            if not isinstance(items, list) or any(not isinstance(v, str) or not v.strip() or len(v) > 1024 for v in items):
                raise ValueError(f'{name} must be a list of nonempty app identities')
        for name in ('overlay_words_added', 'overlay_words_removed'):
            items = getattr(value, name)
            if not isinstance(items, list) or any(not isinstance(v, str) or not v.strip() or len(v) > 100 for v in items):
                raise ValueError(f'{name} must be a list of nonempty words')
        if not isinstance(value.hotkeys, dict) or any(not isinstance(k, str) or not k or not isinstance(v, str)
                                                    for k, v in value.hotkeys.items()):
            raise ValueError('Hotkeys must map action strings to shortcut strings')
