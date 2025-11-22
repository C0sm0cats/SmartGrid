# SmartGrid 
## Powerful Pure-Python Dynamic Tiling Window Manager for Windows

**SmartGrid** is not just a tiler.
It’s a **real, living, breathing dynamic tiling window manager** for Windows 10 & 11 — written in pure Python.

No config. No bullshit. Just press one key and live in perfect harmony.

## What It Actually Does Now

- **True dynamic tiling** — open, close, minimize, restore → layout adapts **instantly**  
  1 window → full screen  
  2 → perfect side-by-side  
  3 → master + stack  
  4+ → intelligent grid (up to 5×3)
- **Workspaces per monitor** — 3 independent workspaces on each screen  
  → Switch instantly with `Ctrl+Alt+1/2/3`  
  → Each workspace remembers its layout **perfectly** (position + grid coords)  
  → Hidden windows restore automatically (even from taskbar)  
  → Smooth transitions with **zero flickering**
- **Drag & Drop Snap** — Grab any tiled window by the title bar → drop it anywhere → **it snaps perfectly**  
  → Works **across monitors**  
  → Automatically **swaps** if target cell is occupied  
  → **Zero hotkeys needed** — pure mouse bliss
- **SWAP Mode** — `Ctrl+Alt+S` → red border → arrow keys → **direct swap** with adjacent windows
  → Navigate with ← → ↑ ↓  
  → The red window **follows your movements**  
  → Press Enter or `Ctrl+Alt+S` to exit
- **Green border** that **always** follows the active window
- **Workspace-aware monitor cycling** — `Ctrl+Alt+M` → current workspace jumps to next monitor  
  → Other workspaces stay intact  
  → Merges smoothly if target workspace has windows
- Works with **everything**: Electron, UWP, WPF, acrylic, custom-drawn, stubborn apps — **all obey**

## Hotkeys

| Shortcut         | Action                                                                    |
|:-----------------|---------------------------------------------------------------------------|
| `Ctrl + Alt + T` | Toggle persistent tiling mode (on/off)                                    |
| `Ctrl + Alt + R` | Force re-tile all visible windows now                                     |
| `Ctrl + Alt + M` | Move current workspace to next monitor                                    |
| `Ctrl + Alt + S` | Enter SWAP MODE (red border + arrow keys) to exchange window positions   |
|                  | ↳ Use ← → ↑ ↓ to navigate, Enter or Ctrl+Alt+S to exit                    |
| `Ctrl + Alt + 1` | Switch to workspace 1 (current monitor)                                   |
| `Ctrl + Alt + 2` | Switch to workspace 2 (current monitor)                                   |
| `Ctrl + Alt + 3` | Switch to workspace 3 (current monitor)                                   |
| `Ctrl + Alt + Q` | Quit SmartGrid                                                            |

> **Pro tip**: After the first `Ctrl+Alt+T`, you’ll almost never touch `R` again.

## Behavior

- Launch SmartGrid → nothing moves (you see the welcome message)
- Press `Ctrl+Alt+T` → **BAM**. Instant perfect tiling + auto-retile activated
- From now on: restore a window, minimize one, open Firefox → layout updates **automatically**
- Press `Ctrl+Alt+T` again → free mode (move windows manually)
- Press `Ctrl+Alt+T` again → everything snaps back into perfect order

## Workspace System

SmartGrid gives you **3 independent workspaces per monitor** — like having multiple virtual desktops, but better.

**How it works:**
1. Tile your windows on workspace 1 (default)
2. Press `Ctrl+Alt+2` → workspace 1 windows **hide instantly** (no minimize animation)
3. Tile different windows on workspace 2
4. Press `Ctrl+Alt+1` → back to your first context, **pixel-perfect**

**Smart features:**
- ✅ Restores **exact positions** (no layout recalculation)
- ✅ Brings back **minimized windows** from taskbar
- ✅ Shows **hidden windows** automatically
- ✅ Skips **dead windows** (closed apps)
- ✅ **Smooth transitions** (no flickering)
- ✅ Works **independently** on each monitor

**Example workflow:**
```
Monitor 1, Workspace 1: [Browser, VSCode, Terminal]  ← Dev environment
Monitor 1, Workspace 2: [Spotify, Discord, OBS]      ← Entertainment/Streaming
Monitor 1, Workspace 3: [Email, Slack, Calendar]     ← Communication

→ Switch contexts instantly without cluttering your taskbar!
```

## Drag & Drop Snap Feature

This is the one that makes people go **"wait… how?!"**

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


## Requirements

- Windows 10 / 11 (64-bit)
- Python 3.9+
- `pip install pywin32`

## Install & Run

```bash
git clone https://github.com/yourusername/smartgrid.git
cd smartgrid
pip install pywin32
python smartgrid.py
```

Press `Ctrl + Alt + T` → enjoy instant, perfect tiling.

## Why This Script Exists

Many great tiling solutions exist for Windows, but a surprising number of modern applications resist standard window-management APIs. SmartGrid forces every window into perfect obedience using raw Win32 + DWM tricks.
Plus, it adds **workspace management** that most Windows tiling tools don't have.

## Author

Made with passion and pure determination by [@C0sm0cats](https://github.com/C0sm0cats)

---

**SmartGrid — Because sometimes you just want your windows to line up perfectly.**

Press `Ctrl + Alt + T` and feel the difference.