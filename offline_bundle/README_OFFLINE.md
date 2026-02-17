# Offline Bundle

This bundle was generated for offline installation of `window-layout-cli`.

## Included folders
- `wheels/` (downloaded dependency wheels)
- `dist/` (project wheel builds)
- `scripts/` install helpers

## Install options
- CMD: `scripts\offline-install.cmd`
- PowerShell: `./scripts/offline-install.ps1`
- Zsh: `./scripts/offline-install.zsh`

These wrappers call the underlying `install_offline.*` scripts and install with:
`python -m pip install --no-index --find-links wheels --find-links dist window-layout-cli`

## Build metadata
Requested Python versions: 3.13, 3.12
Resolved builders: C:\Users\Jared\python-apps\.venv\Scripts\python.exe
