"""Built-in exclusions: windows kept out of the grid by default.

Applications that are usually small, transient or overlays (media players,
game launchers, streaming and monitoring tools, other window managers, call
windows) float unless their Always floating box is unticked. Titles are matched on whole words;
the most generic words only when they make up the whole title, so ordinary
windows such as "Getting started" or "research.docx" are not caught.
"""
import ntpath
import re
from functools import lru_cache

# Matched as whole words anywhere in the title.
TITLE_WORDS = (
    'zscaler', 'spotify', 'discord', 'steam', 'incoming call', 'call', 'meeting', 'join',
    'obs', 'streamlabs', 'twitch studio', 'nvidia overlay', 'geforce experience',
    'shadowplay', 'radeon software', 'amd relive', 'rainmeter', 'wallpaper engine',
    'lively wallpaper', 'msi afterburner', 'rtss', 'rivatuner', 'hwinfo', 'hwmonitor',
    'displayfusion', 'actual window', 'aquasnap', 'powertoys', 'fancyzones',
    'picture in picture', 'picture-in-picture', 'miniplayer', 'mini player', 'youtube music',
    'vlc media player', 'media player classic', 'battle.net', 'epic games',
    'gog galaxy', 'uplay', 'ubisoft connect', 'ea app', 'game bar', 'xbox',
    'volume control', 'realtek audio console',
)
# Too generic to match inside a title: excluded only when they are the whole title.
WHOLE_TITLES = (
    'origin', 'pip', 'notification', 'toast', 'popup', 'tooltip', 'splash', 'alert', 'flyout',
    'brightness', 'program manager', 'start', 'cortana', 'search',
)
# Window classes that are never application windows.
CLASSES = frozenset(name.casefold() for name in (
    'Chrome_RenderWidgetHostHWND', 'OperationStatusWindow', 'Windows.UI.Core.CoreWindow',
    'ForegroundStaging', 'WorkerW', 'Progman', 'Shell_TrayWnd', 'Shell_SecondaryTrayWnd',
    'RealTimeDisplay', 'Credential Dialog Xaml Host', 'MultitaskingViewFrame', 'TaskSwitcherWnd',
    'XamlExplorerHostIslandWindow', '#32770', 'Windows.UI.Popupwindowclass', 'PopupHostWindow',
    'Microsoft.UI.Content.PopupWindowSiteBridge', 'NotepadShellExperienceHost', 'TRectangleCapture',
))
TEAMS_TOAST = (520, 300)



def normalize_word(word):
    return ' '.join((word or '').casefold().split())


def overlay_words(added=(), removed=()):
    """The keyword list in use: the built-in words, minus removed, plus added.

    Only the user's changes are saved, so words added to the built-in list in
    a later version still apply.
    """
    removed = {normalize_word(w) for w in removed}
    words = [w for w in TITLE_WORDS if w not in removed] + [normalize_word(w) for w in added]
    return tuple(dict.fromkeys(w for w in words if w))


@lru_cache(maxsize=8)
def _pattern(words):
    if not words:
        return None
    return re.compile(r'(?<![\w.])(' + '|'.join(re.escape(w) for w in sorted(words, key=len, reverse=True)) + r')(?![\w])')


def _keyword(text, words):
    pattern = _pattern(tuple(words))
    match = pattern.search(text) if pattern else None
    return match.group(1) if match else ''


def excluded_class(class_name):
    return (class_name or '').casefold() in CLASSES


def excluded_app(name, words=TITLE_WORDS):
    """The built-in exclusion an application name matches, or ''.

    Applications in the list float by default; Preferences shows them with
    Always floating already ticked.
    """
    text = ' '.join((name or '').casefold().split())
    if not text:
        return ''
    if text in WHOLE_TITLES:
        return text
    return _keyword(text, words)


def excluded_window(title, app_id='', width=0, height=0, words=TITLE_WORDS):
    """Return the reason a window is excluded by default, or ''."""
    executable = ntpath.basename(app_id.partition(':')[2]) if app_id.startswith('exe:') else ''
    if executable == 'ms-teams.exe' and (height < 200 or (width < TEAMS_TOAST[0] and height < TEAMS_TOAST[1])):
        return 'Microsoft Teams notification'
    text = ' '.join((title or '').casefold().split())
    if not text:
        return ''
    if text in WHOLE_TITLES:
        return f'Built-in exclusion: {text}'
    word = _keyword(text, words)
    return f'Built-in exclusion: {word}' if word else ''
