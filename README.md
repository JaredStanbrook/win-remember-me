# Window Layout CLI

Fast Windows session-state management for daily TSD workflows. This tool captures window positions, PIDs, and metadata to restore your workspace quickly. It can also relaunch closed apps and (optionally) capture/reopen Microsoft Edge tabs.

## Features
- Save window layout to JSON.
- Restore layout with window positioning (including snapped layouts).
- Relaunch missing apps using absolute executable paths.
- Optional Edge tab capture/restore via remote debugging.

## Requirements
- Windows
- Python 3.9+
- Packages: `psutil`, `pywin32`

Install:
```bash
pip install psutil pywin32
```

## Offline Install (v0.1)
On an online machine:
```bash
python -m pip download -r requirements.txt -d wheels
python -m pip wheel . -w dist
```

Copy `dist/` and `wheels/` to the offline machine, then run:
```bash
python -m pip install --no-index --find-links wheels --find-links dist window-layout-cli
```

This installs the `window-layout` CLI entry point.

## Quick Start
Save layout:
```bash
python window_layout.py save layout.json
```

Restore layout:
```bash
python window_layout.py restore layout.json
```

Restore and relaunch missing apps:
```bash
python window_layout.py restore layout.json --launch-missing
```

## Dev Commands
PowerShell:
```powershell
./scripts/dev.ps1 save
./scripts/dev.ps1 restore
./scripts/dev.ps1 edge-debug
./scripts/dev.ps1 edge-save
./scripts/dev.ps1 edge-restore
```

Git Bash:
```bash
./scripts/dev.sh save
./scripts/dev.sh restore
./scripts/dev.sh edge-debug
./scripts/dev.sh edge-save
./scripts/dev.sh edge-restore
```

CMD:
```bat
scripts\dev.cmd save
scripts\dev.cmd edge-save
scripts\dev.cmd edge-restore
```

## First-Time Setup Wizard
For TSD staff who want a quick guided capture:
```bash
python window_layout.py wizard
```
The wizard can launch an Edge debug session, capture tabs, and save a layout in one flow.

## Edge Tabs (Optional)
To capture tabs, Edge must be launched with remote debugging. The tool can start a debug instance:

```bash
python window_layout.py edge-debug
```

Then save tabs:
```bash
python window_layout.py save layout.json --edge-tabs
```

Restore tabs:
```bash
python window_layout.py restore layout.json --restore-edge-tabs
```

Notes:
- Internal pages like `edge://settings` are skipped.
- If no tabs are captured, `--restore-edge-tabs` still reopens Edge (if it was open at save time) and restores its window position.

## Common Options
- `--dry-run`: show matches without moving windows.
- `--min-score`: adjust matching sensitivity (default: 40).
- `--launch-wait`: seconds to wait after relaunching (default: 6).

Examples:
```bash
python window_layout.py restore layout.json --dry-run
python window_layout.py restore layout.json --launch-missing --launch-wait 8
```

## How It Works
- Each saved window includes metadata (title, class, process, exe, rects).
- Restore uses a scoring heuristic to match current windows.
- Explorer windows store their folder path and relaunch into that folder if closed.

## Limitations
- Some windows (elevated/UWP/protected) cannot be moved reliably.
- Tab capture depends on Edge’s remote debugging endpoint.
- Window titles changing between save/restore can reduce match accuracy.
