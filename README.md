# Window Layout

Fast Windows session-state management for TSD workflows. Captures window positions, process metadata, and geometry to restore your workspace with a single command. Optionally relaunches closed apps and captures/restores Microsoft Edge tabs per window.

## Features

- Save window layout to JSON (positions, sizes, minimized/maximized state)
- Restore layout using `SetWindowPlacement` for reliable max/min/normal handling
- Relaunch missing apps from saved executable paths
- Optional Edge tab capture and restore via remote debugging (per-window mapping)
- Verbose mode (`-v`) on save and restore for live per-window feedback
- Hotkey listener for instant layout switching without leaving your keyboard
- Interactive setup wizard for first-time capture
- Lightweight PySide6 GUI with speed-menu buttons and layout editor

## Requirements

- Windows 10 / 11
- Python 3.9+
- CLI packages: `psutil`, `pywin32`
- GUI packages (optional): `PySide6`

```bash
pip install "window-layout[cli]"
pip install "window-layout[gui]"
```

## Offline Install

Build a shareable offline bundle on an internet-connected machine:

```bash
python scripts/build_offline_bundle.py --python-versions 3.13 3.12 3.11
```

This produces `offline_bundle.zip` containing wheels, a built package, and installer scripts. Copy the zip to the offline machine, extract it, then run one of:

| Shell | Command |
|---|---|
| CMD | `scripts\\offline-install.cmd` |
| PowerShell | `./scripts/offline-install.ps1` |
| Zsh / Git Bash | `./scripts/offline-install.sh` |

Or manually:
```bash
python -m pip install --no-index --find-links wheels --find-links dist window-layout
```

---

## Quick Start

```bash
# Capture current desktop
python window_layout.py save layout.json

# Restore (move existing windows back only)
python window_layout.py restore layout.json

# Restore + launch missing apps
python window_layout.py restore layout.json --launch-missing

# Restore + Edge tabs (per-window)
python window_layout.py restore layout.json --edge-tabs

# Restore + Edge tabs (close flagged windows first)
python window_layout.py restore layout.json --edge-tabs --destructive

# Restore with verbose output to see every window as it moves
python window_layout.py restore layout.json --verbose

# Restore with diagnostics (shows match scores per window)
python window_layout.py restore layout.json --diagnostics
```

---

## How Save Works

`save` enumerates all top-level windows that appear in Alt-Tab / the taskbar. For each window it records:

- Window title, class name, process name, and full executable path
- `normal_rect` -- the restored position from `GetWindowPlacement` (correct even when minimized or maximized)
- `show_cmd` -- whether the window was normal, maximized, or minimized at save time
- A `launch` spec (exe + args + cwd) used to relaunch closed apps on restore

Filtered automatically (never saved):
- System UWP hosts: `TextInputHost.exe`, `ApplicationFrameHost.exe`, `StartMenuExperienceHost.exe`, `SearchHost.exe`
- Tool windows, owned pop-ups, windows smaller than 120x80px (excluding minimized windows)
- Windows with no title

```bash
python window_layout.py save layout.json

# Also capture Edge tabs (requires Edge debug session -- see Edge Tabs below)
python window_layout.py save layout.json --edge-tabs

# See each window as it is captured
python window_layout.py save layout.json --verbose
```

---

## How Restore Works

Restore matches each saved window against currently running windows using a scoring heuristic, then uses `SetWindowPlacement` to atomically apply the saved position and state.

Restore modes:
- no args: move existing windows only
- `--launch-missing`: relaunch missing apps before positioning
- `--edge-tabs`: restore Edge tabs per window and launch missing Edge windows (uses `about:blank` if no tabs are saved)
- `--destructive`: only affects windows that have `windows[*].destructive = true` in the layout. When set, matching windows with that flag are closed and relaunched clean.

Match scoring (max 175 points):

| Signal | Points | Notes |
|---|---|---|
| Exe path match | 50 | Most stable identifier across restarts |
| Process name match | 25 | |
| Window class match | 15 | |
| Title exact match | 40 | Not used for Edge (tab title changes) |
| Title partial match | 15 | |
| Geometry close (<= 40px) | 30 | Primary Edge differentiator |
| Geometry nearby (<= 120px) | 15 | |

Matching is pre-filtered by process: only windows running the same executable are ever compared against each other, preventing cross-process false matches.

A window must score >= 40 to be matched. Windows below that threshold are skipped unless you enable relaunch.

```bash
# Basic: move existing windows only
python window_layout.py restore layout.json

# Launch missing apps before positioning
python window_layout.py restore layout.json --launch-missing

# Restore Edge tabs per window
python window_layout.py restore layout.json --edge-tabs

# Print per-window match scores
python window_layout.py restore layout.json --diagnostics

# Show top 5 match candidates per window
python window_layout.py restore layout.json --diagnostics --diagnostics-top 5

# Verbose: print each window as it is positioned
python window_layout.py restore layout.json --verbose
```

---

## Edge Tabs (Optional)

Edge tab capture requires Edge to be running with remote debugging enabled. This is a separate Edge profile from your normal one -- you set up your workflow tabs inside the debug session.

Step-by-step first-time setup:

```bash
# 1. Launch a debug-enabled Edge session
python window_layout.py edge-debug --port 9222

# 2. Open the tabs and windows you want to capture inside that Edge window

# 3. Save with tab capture
python window_layout.py save layout.json --edge-tabs

# 4. Restore including tabs
python window_layout.py restore layout.json --edge-tabs
```

Additional Edge commands:

```bash
# Launch Edge debug session with a custom profile directory
python window_layout.py edge-debug --port 9222 --profile-dir C:\\edge-work-profile

# Capture tabs into an already-saved layout (without re-saving window positions)
python window_layout.py edge-capture layout.json --port 9222

# Manually set URLs for an Edge window (no debug session required)
python window_layout.py edge-urls layout.json https://example.com https://docs.example.com

# Append URLs without replacing existing ones
python window_layout.py edge-urls layout.json https://newsite.com --append

# Manually reassign which tabs belong to which Edge window
python window_layout.py edit layout.json
```

How tab-to-window assignment works:

When you have multiple Edge windows, tabs are matched to windows using the CDP `windowId` field if available. If `windowId` is missing, the tool falls back to title token overlap, then round-robin. When only one Edge window is saved, tabs are opened into the existing Edge window (`--new-tab`) rather than forcing a new window.

Notes:
- Internal pages (`edge://`, `chrome://`) are skipped automatically
- Tabs are stored on each window entry as `windows[*].edge_tabs`
- `--edge-tabs` restores tabs per window and avoids re-launching tabs if Edge is already in the correct position
- Legacy layouts using `browser_tabs`, `edge_sessions`, or `open_urls.edge` are auto-migrated on load

---

## First-Time Setup Wizard

Guided capture for new users:

```bash
python window_layout.py wizard
```

The wizard optionally launches an Edge debug session, waits for you to set up tabs, then saves everything in one flow. Default output is `layouts/layout.json`.

---

## Hotkeys

Configure global hotkeys in `config.json` to trigger save/restore without leaving your keyboard:

```json
{
  "hotkeys": [
    {
      "keys": "Ctrl+Alt+S",
      "action": "save",
      "args": ["layouts/daily.json"]
    },
    {
      "keys": "Ctrl+Alt+R",
      "action": "restore",
      "args": ["layouts/daily.json", "--edge-tabs"]
    }
  ]
}
```

```bash
python window_layout.py hotkeys
```

Supported modifier keys: `Ctrl`, `Alt`, `Shift`, `Win`. Supported keys: letters, numbers, F1-F24, and common specials (Tab, Enter, Esc, Space, Delete, Home, End, arrow keys).

---

## GUI (PySide6)

```bash
python gui_app.py
# or if installed as package:
window-layout-gui
```

The GUI renders a grid of speed-menu buttons from `config.json`. The editor has two columns: layouts found in `layouts_root` on the left, and your current speed menu on the right.

```json
{
  "layouts_root": "C:\\Users\\Jared\\layouts",
  "speed_menu": {
    "buttons": [
      {
        "label": "Daily",
        "emoji": ":rocket:",
        "layout": "daily.json",
        "args": ["--launch-missing"]
      },
      {
        "label": "Focus",
        "emoji": ":brain:",
        "layout": "focus.json",
        "args": ["--edge-tabs"]
      }
    ]
  }
}
```

- `layouts_root` sets the folder scanned for layout files (default: `layouts/`)
- Relative `layout` values are resolved from `layouts_root` first
- `args` is passed directly to `restore` -- use `--launch-missing`, `--edge-tabs`, `--destructive`, or no args
- GUI tabs: Settings, Speed Menu, Speed Menu Editor, Layout Editor

---

## Dev Scripts

```powershell
# PowerShell
./scripts/dev.ps1 save
./scripts/dev.ps1 restore
./scripts/dev.ps1 edge-debug
./scripts/dev.ps1 edge-save
./scripts/dev.ps1 edge-restore
```

```bash
# Git Bash / Zsh
./scripts/dev.sh save
./scripts/dev.sh restore
./scripts/dev.sh edge-debug
./scripts/dev.sh edge-save
./scripts/dev.sh edge-restore
```

```bat
:: CMD
scripts\\dev.cmd save
scripts\\dev.cmd edge-save
scripts\\dev.cmd edge-restore
```

---

## Diagnostics & Troubleshooting

Run the testbench for a full desktop audit:

```bash
python testbench.py            # guided walkthrough
python testbench.py --snapshot # one-shot dump, no prompts
python testbench.py --edge-only
python testbench.py --compare before.json after.json
```

The testbench shows: all enumerated windows with filter reasons, running processes, Edge debug port status, match simulation, and a list of known bottlenecks.

Common issues:

| Symptom | Cause | Fix |
|---|---|---|
| Window not captured | Process is in the block-list (UWP/system) | Expected -- these can't be moved |
| Window skipped on restore | Title changed since save | Use `--diagnostics` to see score; consider `edge-urls` or `edit` to update |
| Edge tabs not captured | Debug port closed | Run `edge-debug` first, then re-save |
| Maximized window restores to wrong size | Old script used `MoveWindow` | Updated script uses `SetWindowPlacement` -- pull latest |
| Minimized window wasn't saved | Old size filter excluded ~0px iconic windows | Fixed in current version |
| Window restored to wrong monitor | Multi-monitor DPI mismatch | Ensure script runs with DPI awareness (default on Windows 10+) |
| Elevated app won't move | Process is running as Administrator | Run `window_layout.py` as Administrator too |

---

## Known Limitations

- Elevated/protected processes cannot be repositioned by a non-elevated script
- Edge tab capture requires a dedicated debug-profile Edge session
- Some UWP Store apps are "cloaked" and enumerate but cannot be reliably moved
- Window titles that change between save and restore reduce match confidence (use `--diagnostics` to inspect)

---

## Roadmap

| Phase | Status | Description |
|---|---|---|
| CLI core | Complete | Save, restore, edge tabs, hotkeys, wizard |
| Testbench | Complete | Verbose diagnostic tool for debugging layouts |
| GUI Phase 1 | Started | Session manager in `gui_app.py` |
| GUI Phase 2 | Planned | Visual Edge tab assignment editor |
| GUI Phase 3 | Planned | Scheduled capture, profile presets, diagnostics log view |

Architecture note: CLI logic remains the source of truth. The GUI calls into the same Python functions rather than reimplementing them.
