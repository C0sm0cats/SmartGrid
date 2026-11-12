"""Per-application placement and color of the window action handle and palette.

A rule key is either a full app ID (``exe:C:\\...\\mstsc.exe``, ``aumid:...``)
or a bare name: an executable (``mstsc.exe``) that matches that program
wherever it is installed, or a Start menu AppID (``Microsoft.Windows.RemoteDesktop``,
the localized "Remote Desktop Connection" entry). The shipped defaults use bare names.
"""
import re

ACTION_POSITIONS = ('left', 'right', 'top', 'bottom')
REMOTE_HANDLE_COLOR = '#ed5b00'

# Remote desktop and virtual machine viewers: their handle moves to the right
# and turns orange, so it never covers the handle of a session shown inside.
REMOTE_VIEWERS = ('mstsc.exe', 'msrdc.exe', 'msrdcw.exe', 'vmconnect.exe', 'vmware.exe', 'vmplayer.exe',
                  'VirtualBoxVM.exe', 'AnyDesk.exe', 'TeamViewer.exe', 'rustdesk.exe', 'vncviewer.exe',
                  'tvnviewer.exe', 'mRemoteNG.exe', 'RDCMan.exe', 'nxplayer.exe', 'x2goclient.exe',
                  'Moonlight.exe', 'parsecd.exe', 'remote-viewer.exe', 'Microsoft.Windows.RemoteDesktop')


def default_positions():
    return {name: 'right' for name in REMOTE_VIEWERS}


def default_colors():
    return {name: REMOTE_HANDLE_COLOR for name in REMOTE_VIEWERS}


def _bare_name(app_id):
    """Executable name of an exe ID, or the AppID of an aumid ID."""
    kind, _, value = app_id.partition(':')
    if kind == 'exe' and value:
        return re.split(r'[\\/]', value)[-1].casefold()
    return value.casefold() if kind == 'aumid' else ''


def rule_matches(key, app_id):
    """True when the rule key applies to this application."""
    if not key or not app_id:
        return False
    if key.casefold() == app_id.casefold():
        return True
    return ':' not in key and key.casefold() == _bare_name(app_id)


def rule_value(rules, app_id):
    """The value for this application: its own ID first, then its executable name."""
    if not isinstance(rules, dict):
        return None
    matches = [(key, value) for key, value in rules.items() if rule_matches(key, app_id)]
    exact = [value for key, value in matches if ':' in key]
    return (exact or [value for _, value in matches] or [None])[0]


def set_rule(rules, app_id, value):
    """A copy of the rules where this application has only `value` (None: no rule)."""
    result = {key: v for key, v in rules.items() if not rule_matches(key, app_id)}
    if value is not None:
        result[app_id] = value
    return result


def matching_rules(rules, app_id):
    return {key: value for key, value in rules.items() if rule_matches(key, app_id)}


def valid_color(value):
    return value.lower() if isinstance(value, str) and re.fullmatch(r'#[0-9a-fA-F]{6}', value) else None


def action_style(app_id, settings):
    """(position, color or None) of the handle for a window of this application."""
    position = rule_value(getattr(settings, 'action_positions', {}), app_id)
    color = valid_color(rule_value(getattr(settings, 'action_colors', {}), app_id))
    return (position if position in ACTION_POSITIONS else 'left'), color


def horizontal(position):
    return position in ('top', 'bottom')


def placement(position, rect, area, size):
    """Top-left corner of a control of `size` (w, h) inside `rect`, against
    the chosen edge, centered along it and kept within the work `area`.
    Rectangles are (x, y, width, height)."""
    x0, y0, w0, h0 = rect
    ax, ay, aw, ah = area
    width, height = size
    x = x0 + w0 - width - 4 if position == 'right' else x0 + 4 if position == 'left' else x0 + (w0 - width) // 2
    y = y0 + h0 - height - 4 if position == 'bottom' else y0 + 4 if position == 'top' else y0 + (h0 - height) // 2
    clamp = lambda value, low, high: min(max(value, low), max(low, high))
    return clamp(x, ax + 6, ax + aw - width - 6), clamp(y, ay + 8, ay + ah - height - 8)


def tooltip_placement(position, button, tooltip, area):
    """Tooltip beside an action button, on the side away from the window edge."""
    bx, by, bw, bh = button
    tw, th = tooltip
    ax, ay, aw, ah = area
    if horizontal(position):
        x = bx + (bw - tw) // 2
        below, above = by + bh + 8, by - th - 8
        y = (below if below + th <= ay + ah - 6 else above) if position == 'top' else (above if above >= ay + 6 else below)
    else:
        y = by + (bh - th) // 2
        right, left = bx + bw + 8, bx - tw - 8
        x = (left if left >= ax + 6 else right) if position == 'right' else (right if right + tw <= ax + aw - 6 else left)
    clamp = lambda value, low, high: min(max(value, low), max(low, high))
    return clamp(x, ax + 6, ax + aw - tw - 6), clamp(y, ay + 6, ay + ah - th - 6)


def lighten(color, amount=.45):
    """Toward white, for a border that stays visible on a dark color."""
    channels = [int(color[i:i + 2], 16) for i in (1, 3, 5)]
    return '#' + ''.join(f'{round(c + (255 - c) * amount):02x}' for c in channels)
