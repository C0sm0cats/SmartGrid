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
- **Drag & Drop Snap** — Grab any tiled window by the title bar → drop it anywhere → **it snaps perfectly**  
  → Works **across monitors**  
  → Automatically **swaps** if target cell is occupied  
  → **Zero hotkeys needed** — pure mouse bliss
- **SWAP Mode** — `Ctrl+Alt+S` → red border → arrow keys → press Enter → instant position exchange
- **Green border** that **always** follows the active window
- **One-key multi-monitor cycling** — `Ctrl+Alt+M` → entire layout jumps to next monitor, perfectly resized
- Works with **everything**: Electron, UWP, WPF, acrylic, custom-drawn, stubborn apps — **all obey**

## Hotkeys

| Shortcut         | Action                                                                    |
|:-----------------|---------------------------------------------------------------------------|
| `Ctrl + Alt + T` | Toggle persistent tiling mode (on/off)                                    |
| `Ctrl + Alt + R` | Force re-tile all visible windows now                                     |
| `Ctrl + Alt + M` | Move all tiled windows to next monitor                                    |
| `Ctrl + Alt + S` | Enter SWAP MODE (red border + arrow keys) to exchange window positions   |
|                  | ↳ Use ← → ↑ ↓ to navigate, Enter or Ctrl+Alt+S to exit                    |
| `Ctrl + Alt + Q` | Quit SmartGrid                                                            |

> **Pro tip**: After the first `Ctrl+Alt+T`, you’ll almost never touch `R` again.

## Behavior

- Launch SmartGrid → nothing moves (you see the welcome message)
- Press `Ctrl+Alt+T` → **BAM**. Instant perfect tiling + auto-retile activated
- From now on: restore a window, minimize one, open Firefox → layout updates **automatically**
- Press `Ctrl+Alt+T` again → free mode (move windows manually)
- Press `Ctrl+Alt+T` again → everything snaps back into perfect order

## Drag & Drop Snap Feature

This is the one that makes people go **"wait… how?!"**

1. You have 6 windows tiled  
2. You grab **one** by the title bar  
3. You drag it over another monitor, or over an empty cell, or over another window  
4. You **drop**  
→ **BAM**. It snaps perfectly.  
→ If you dropped on another window → they **swap instantly**  
→ If you dropped in empty space → it **moves** there  
→ Works **across monitors**  
→ No keys. No thinking. Pure flow.

## Multi-Monitor Workflow

1. Tile your windows on monitor 1
2. Press `Ctrl + Alt + M` → everything jumps to monitor 2, perfectly resized
3. Press again → back to monitor 1 (or to monitor 3, 4… fully cyclic)

No manual dragging. No resizing. Just one key.

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

## Author

Made with passion and pure determination by [@C0sm0cats](https://github.com/C0sm0cats)

---

**SmartGrid — Because sometimes you just want your windows to line up perfectly.**

Press `Ctrl + Alt + T` and feel the difference.