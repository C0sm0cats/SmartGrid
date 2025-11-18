# SmartGrid — Powerful & Reliable Window Tiling for Windows

**SmartGrid** is a lightweight, aggressive, pure-Python window tiler that **just works** on Windows 10 & 11 — even with the most stubborn modern applications.

No config files. No admin rights. Just press a key and get a perfect grid.

## What It Does

- Instantly tiles all visible windows into clean 3×3, 4×3 or 5×3 grids (up to 15 windows)  
- Uses constant gaps and edge padding for a polished look  
- Fully supports multi-monitor setups:  
  → `Ctrl+Alt+M` moves **your entire grid** to the next monitor and automatically resizes everything  
- Works reliably with apps that have custom frames, invisible borders, acrylic effects, GPU rendering, etc.  
- Zero perceptible lag on hotkeys

## Hotkeys

| Shortcut            | Action                                              |
|---------------------|-----------------------------------------------------|
| `Ctrl + Alt + T`    | Toggle persistent tiling mode                       |
| `Ctrl + Alt + R`    | One-shot re-tile of all visible windows             |
| `Ctrl + Alt + M`    | Cycle all tiled windows to the next monitor         |
| `Ctrl + Alt + Q`    | Quit SmartGrid                                      |

## Requirements

- Windows 10 or Windows 11 (64-bit)
- Python 3.9 or newer
- One tiny dependency:

## Quick Start

```bash
git clone https://github.com/yourusername/smartgrid.git
cd smartgrid
pip install pywin32
python smartgrid.py
```

Press `Ctrl + Alt + T` → enjoy instant, perfect tiling.

## Multi-Monitor Workflow

1. Tile your windows on monitor 1
2. Press `Ctrl + Alt + M` → everything jumps to monitor 2, perfectly resized
3. Press again → back to monitor 1 (or to monitor 3, 4… fully cyclic)

No manual dragging. No resizing. Just one key.

## Why This Script Exists

Many great tiling solutions exist for Windows, but a surprising number of modern applications resist standard window-management APIs. SmartGrid uses low-level Win32 calls, DWM border compensation, and aggressive repositioning to make every visible window obey the grid — reliably and instantly.

All in under 500 lines of clean, readable Python.

## Windows Only

Built from the ground up for the Windows Desktop Window Manager. Proudly Windows-only.


## Author

Made with passion and pure determination by [@C0sm0cats](https://github.com/C0sm0cats)

---

**SmartGrid — Because sometimes you just want your windows to line up perfectly.**

Press `Ctrl + Alt + T` and feel the difference.