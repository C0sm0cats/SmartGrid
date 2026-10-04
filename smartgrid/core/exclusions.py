"""Built-in exclusions: windows kept out of the grid by default.

Applications that are usually small, transient or overlays (media players,
game launchers, streaming and monitoring tools, other window managers, call
windows) float unless explicitly included. Titles are matched on whole words;
the most generic words only when they make up the whole title, so ordinary
windows such as "Getting started" or "research.docx" are not caught.
"""
import ntpath
import re

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

_WORDS = re.compile(r'(?<![\w.])(' + '|'.join(re.escape(w) for w in sorted(TITLE_WORDS, key=len, reverse=True)) + r')(?![\w])')


def excluded_class(class_name):
    return (class_name or '').casefold() in CLASSES


def excluded_window(title, app_id='', width=0, height=0):
    """Return the reason a window is excluded by default, or ''."""
    executable = ntpath.basename(app_id.partition(':')[2]) if app_id.startswith('exe:') else ''
    if executable == 'ms-teams.exe' and (height < 200 or (width < TEAMS_TOAST[0] and height < TEAMS_TOAST[1])):
        return 'Microsoft Teams notification'
    text = ' '.join((title or '').casefold().split())
    if not text:
        return ''
    if text in WHOLE_TITLES:
        return f'Built-in exclusion: {text}'
    match = _WORDS.search(text)
    return f'Built-in exclusion: {match.group(1)}' if match else ''
