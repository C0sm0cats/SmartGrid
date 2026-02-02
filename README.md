# SmartGrid

[![Release](https://img.shields.io/github/v/release/C0sm0cats/SmartGrid)](https://github.com/C0sm0cats/SmartGrid/releases/latest)
[![Downloads](https://img.shields.io/github/downloads/C0sm0cats/SmartGrid/total)](https://github.com/C0sm0cats/SmartGrid/releases)
[![License](https://img.shields.io/github/license/C0sm0cats/SmartGrid)](LICENSE)
[![Platform](https://img.shields.io/badge/platform-Windows%2010%2F11-blue)](https://github.com/C0sm0cats/SmartGrid)

> Dynamic tiling window manager for Windows — **pure Python** (Win32 + DWM)

SmartGrid gives you instant tiling, drag & drop snapping, swap mode, and **workspaces per monitor** — with a system tray UI and global hotkeys.

## Features

- **Dynamic layouts** (1 → full, 2 → split, 3 → master/stack, 4+ → grid up to 5×3)
- **Maximize-safe tiling:** maximizing a tiled window won’t reshuffle other windows; restore returns to the original slot
- **Drag & drop snap:** drag a tiled window by the title bar, preview appears, drop to snap (supports cross-monitor)
- **Swap Mode:** red border + arrow keys to swap with adjacent windows
- **Floating windows toggle:** keep specific windows out of the grid (video/chat/reference)
- **Workspaces per monitor:** 3 workspaces per screen, instant switching, layout remembered
- **System tray menu:** toggle tiling, retile, swap mode, move workspace, settings (gap/padding), hotkeys, quit
- **Active border:** green border follows the active tiled window

## Hotkeys

| Shortcut             | Action                                                                      |
|:---------------------|-----------------------------------------------------------------------------|
| `Ctrl + Alt + T`     | Toggle tiling (on/off)                                                      |
| `Ctrl + Alt + R`     | Force re-tile all windows now                                               |
| `Ctrl + Alt + S`     | Enter Swap Mode (red border + arrows)                                       |
| `Ctrl + Alt + M`     | Move current workspace to next monitor                                      |
| `Ctrl + Alt + F`     | Toggle Floating Selected Window                                             |
| `Ctrl + Alt + 1/2/3` | Switch to workspace 1/2/3 (current monitor)                                 |
| `Ctrl + Alt + Q`     | Quit SmartGrid                                                              |

## Install & Run

### Option A — Download the latest release

https://github.com/C0sm0cats/SmartGrid/releases/latest

### Option B — Run from source

Requirements:
- Windows 10 / 11 (64-bit)
- Python 3.9+
- Dependencies: `pywin32`, `pystray`, `pillow` (PIL)

```bash
git clone https://github.com/C0sm0cats/SmartGrid.git
cd SmartGrid
python -m pip install --upgrade pip
python -m pip install pywin32 pystray pillow
python smartgrid.py
```

Press `Ctrl + Alt + T` to enable tiling.

## Usage

- Launch SmartGrid → nothing moves (you see the welcome message)
- Press `Ctrl+Alt+T` → instant tiling + auto-retile activated
- From now on: restore a window, minimize one, open whatever you want → layout updates **automatically**
- Press `Ctrl+Alt+T` again → free mode (move windows manually)
- Press `Ctrl+Alt+T` again → everything snaps back into perfect order

## Workspaces (per monitor)

SmartGrid gives you **3 independent workspaces per monitor** — like having multiple virtual desktops, but better.

**How it works:**
1. Tile your windows on workspace 1 (default)
2. Press `Ctrl+Alt+2` → workspace 1 windows **hide instantly** (no minimize animation)
3. Tile different windows on workspace 2
4. Press `Ctrl+Alt+1` → back to your first context, **pixel-perfect**

**Example workflow:**
```
Monitor 1, Workspace 1: [Browser, VSCode, Terminal]  ← Dev environment
Monitor 1, Workspace 2: [Spotify, Discord, OBS]      ← Entertainment/Streaming
Monitor 1, Workspace 3: [Email, Slack, Calendar]     ← Communication

→ Switch contexts instantly without cluttering your taskbar!
```

## Drag & Drop Snap

1. You have 6 windows tiled
2. You grab **one** by the title bar
3. **A blue preview rectangle appears** showing exactly where it will snap
4. You **drop**
→ **BAM**. It snaps perfectly to the previewed position.
→ If you dropped on another window → they **swap instantly**
→ If you dropped in empty space → it **moves** there
→ Works **across monitors**
→ Works **across workspaces**
→ No keys. No thinking. Pure flow.

## Multi-Monitor Workflow

**Move current workspace to another screen:**
1. Tile your windows on monitor 1, workspace 1
2. Press `Ctrl + Alt + M` → workspace 1 jumps to monitor 2, perfectly resized
3. Press again → continues cycling (monitor 3, back to 1...)

**Important:** Only the **current workspace** moves. Other workspaces stay on their monitors.

**What happens if target workspace has windows?**  
→ They **merge** and re-tile together (like i3/Sway behavior)

## Notes / Troubleshooting

- **Maximize behavior:** while a window is maximized, SmartGrid intentionally avoids background reshuffles so other windows don’t move.
- **Hotkeys don’t work:** another tool may be using the same shortcuts (PowerToys/FancyZones, DisplayFusion, etc.).
- **Some windows don’t tile:** SmartGrid filters overlays/toasts/taskbar/etc. You can tune the rules in `is_useful_window()` in `smartgrid.py`.
- **Border colors:** DWM border coloring works best on Windows 11; on some Windows 10 builds it may be ignored.

## Contributing

Ideas, issues and PRs are welcome. For bug reports, please include:
- Windows version (10/11 + build)
- Monitor setup (count + resolution + scaling)
- App names involved (and whether they were maximized/minimized/restored)

## Why This Script Exists

Many great tiling solutions exist for Windows, but a surprising number of modern applications resist standard window-management APIs. SmartGrid forces every window into perfect obedience using raw Win32 + DWM tricks.
Plus, it adds **workspace management** that most Windows tiling tools don't have and a **system tray icon** with a **context menu** for quick access to all major features.

## Author

Made with passion and pure determination by [@C0sm0cats](https://github.com/C0sm0cats)

---

**SmartGrid — Because sometimes you just want your windows to line up perfectly.**

Press `Ctrl + Alt + T` and feel the difference.
