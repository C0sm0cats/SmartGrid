# SmartGrid 
## Powerful Pure-Python Dynamic Tiling Window Manager for Windows

**SmartGrid** is not just a tiler.
It’s a **real, living, breathing dynamic tiling window manager** for Windows 10 & 11 — written in pure Python.

No config. No bullshit. Just press one key and live in perfect harmony.

## What It Actually Does Now

- **True dynamic tiling**: every time you restore, minimize, open or close a window → **the layout instantly adapts** (1 → full, 2 → side-by-side, 3 → master+stack, 4 → 2×2, up to 5×3 grid)
- **Zero manual re-tile needed** — it just *knows*
- Perfect green border that **always** follows the active window (even after 100 minimizes)
- Full multi-monitor cycling: `Ctrl+Alt+M` moves your entire living grid to the next monitor
- Works with the most stubborn apps (Electron, UWP, WPF, acrylic, custom frames, etc.)
- Under 500 lines of clean, readable, battle-tested Python

## Hotkeys

| Shortcut            | Action                                              |
|---------------------|-----------------------------------------------------|
| `Ctrl + Alt + T`    | Toggle persistent tiling mode                       |
| `Ctrl + Alt + R`    | One-shot re-tile of all visible windows             |
| `Ctrl + Alt + M`    | Cycle all tiled windows to the next monitor         |
| `Ctrl + Alt + Q`    | Quit SmartGrid                                      |

> **Pro tip**: After the first `Ctrl+Alt+T`, you’ll almost never touch `R` again.

## Behavior

- Launch SmartGrid → nothing moves (you see the welcome message)
- Press `Ctrl+Alt+T` → **BAM**. Instant perfect tiling + auto-retile activated
- From now on: restore a window, minimize one, open Firefox → layout updates **automatically**
- Press `Ctrl+Alt+T` again → free mode (move windows manually)
- Press `Ctrl+Alt+T` again → everything snaps back into perfect order

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

Many great tiling solutions exist for Windows, but a surprising number of modern applications resist standard window-management APIs. SmartGrid uses low-level Win32 calls, DWM border compensation, and aggressive repositioning to make every visible window obey the grid — reliably and instantly.

All in under 500 lines of clean, readable Python.


## Author

Made with passion and pure determination by [@C0sm0cats](https://github.com/C0sm0cats)

---

**SmartGrid — Because sometimes you just want your windows to line up perfectly.**

Press `Ctrl + Alt + T` and feel the difference.