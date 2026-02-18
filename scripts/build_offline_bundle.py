#!/usr/bin/env python3
"""Build offline bundle assets for multiple Python versions (multi-version safe)."""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List, Optional
from zipfile import ZIP_DEFLATED, ZipFile

REPO_ROOT = Path(__file__).resolve().parents[1]


def _run(cmd: List[str], cwd: Path) -> None:
    print("+", " ".join(cmd))
    subprocess.run(cmd, cwd=str(cwd), check=True)


def _find_python_commands(versions: List[str]) -> List[List[str]]:
    commands: List[List[str]] = []
    if sys.platform.startswith("win"):
        # Windows launcher: py -3.13 etc
        for v in versions:
            commands.append(["py", f"-{v}"])
    else:
        # *nix: python3.13 etc
        for v in versions:
            commands.append([f"python{v}"])
        commands.append(["python3"])
        commands.append(["python"])
    # Also include current interpreter if it matches requested versions
    current_version = f"{sys.version_info.major}.{sys.version_info.minor}"
    if current_version in versions:
        commands.append([sys.executable])
    return commands


def _available_python_commands(commands: List[List[str]]) -> List[List[str]]:
    available: List[List[str]] = []
    seen = set()
    for cmd in commands:
        key = tuple(cmd)
        if key in seen:
            continue
        seen.add(key)
        try:
            subprocess.run(
                [*cmd, "--version"],
                check=True,
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
            available.append(cmd)
        except Exception:
            continue
    return available


def _detect_py_tag(py_cmd: List[str]) -> str:
    """
    Return a stable tag like 'py311' or 'py313' by querying that interpreter.
    This avoids collisions and keeps bundle layout predictable.
    """
    p = subprocess.run(
        [*py_cmd, "-c", "import sys; print(f'{sys.version_info.major}{sys.version_info.minor}')"],
        check=True,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
    )
    mm = p.stdout.strip()
    if not mm.isdigit() or len(mm) < 2:
        raise RuntimeError(f"Failed to detect python tag for: {' '.join(py_cmd)}")
    return f"py{mm}"


def _zip_dir(source_dir: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zf:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(source_dir))


def _write_bundle_install_scripts(bundle_scripts_dir: Path) -> None:
    """
    Write self-contained offline install scripts that:
    - Detect the running Python version (py311/py312/py313)
    - Install from matching wheels/dist folders
    """

    install_cmd = r"""@echo off
setlocal enabledelayedexpansion

REM Usage:
REM   offline-install.cmd [package-name]
REM Default package: window-layout
set PACKAGE=%~1
if "%PACKAGE%"=="" set PACKAGE=window-layout

REM Detect py tag from current python on PATH
for /f "usebackq delims=" %%i in (`python -c "import sys; print(f'py{sys.version_info.major}{sys.version_info.minor}')"`) do set PYTAG=%%i

set SCRIPT_DIR=%~dp0
set ROOT_DIR=%SCRIPT_DIR%..

set WHEELS_DIR=%ROOT_DIR%\wheels\%PYTAG%
set DIST_DIR=%ROOT_DIR%\dist\%PYTAG%

if not exist "%WHEELS_DIR%" (
  echo ERROR: Missing wheels folder for %PYTAG%: "%WHEELS_DIR%"
  echo Available versions:
  dir /b "%ROOT_DIR%\wheels" 2>nul
  exit /b 1
)

if not exist "%DIST_DIR%" (
  echo ERROR: Missing dist folder for %PYTAG%: "%DIST_DIR%"
  echo Available versions:
  dir /b "%ROOT_DIR%\dist" 2>nul
  exit /b 1
)

echo Using %PYTAG%
echo Installing %PACKAGE% offline...
python -m pip install --no-index --find-links "%WHEELS_DIR%" --find-links "%DIST_DIR%" "%PACKAGE%"
endlocal
"""

    install_ps1 = r"""param(
  [string]$Package = "window-layout"
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RootDir = (Resolve-Path (Join-Path $ScriptDir "..")).Path

$PyTag = & python -c "import sys; print(f'py{sys.version_info.major}{sys.version_info.minor}')"
$PyTag = $PyTag.Trim()

$WheelsDir = Join-Path $RootDir ("wheels\" + $PyTag)
$DistDir   = Join-Path $RootDir ("dist\" + $PyTag)

if (!(Test-Path $WheelsDir)) {
  Write-Error "Missing wheels folder for $PyTag: $WheelsDir"
  Write-Host "Available versions:"
  Get-ChildItem (Join-Path $RootDir "wheels") -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.Name }
  exit 1
}

if (!(Test-Path $DistDir)) {
  Write-Error "Missing dist folder for $PyTag: $DistDir"
  Write-Host "Available versions:"
  Get-ChildItem (Join-Path $RootDir "dist") -Directory -ErrorAction SilentlyContinue | ForEach-Object { $_.Name }
  exit 1
}

Write-Host "Using $PyTag"
Write-Host "Installing $Package offline..."
python -m pip install --no-index --find-links $WheelsDir --find-links $DistDir $Package
"""

    install_sh = r"""#!/usr/bin/env sh
set -eu

# Usage:
#   ./offline-install.sh [package-name]
# Default package: window-layout
PACKAGE="${1:-window-layout}"

SCRIPT_DIR="$(CDPATH= cd -- "$(dirname -- "$0")" && pwd)"
ROOT_DIR="$(cd "$SCRIPT_DIR/.." && pwd)"

PYTAG="$(python -c "import sys; print(f'py{sys.version_info.major}{sys.version_info.minor}')")"
WHEELS_DIR="$ROOT_DIR/wheels/$PYTAG"
DIST_DIR="$ROOT_DIR/dist/$PYTAG"

if [ ! -d "$WHEELS_DIR" ]; then
  echo "ERROR: Missing wheels folder for $PYTAG: $WHEELS_DIR" >&2
  echo "Available versions:" >&2
  ls -1 "$ROOT_DIR/wheels" 2>/dev/null || true
  exit 1
fi

if [ ! -d "$DIST_DIR" ]; then
  echo "ERROR: Missing dist folder for $PYTAG: $DIST_DIR" >&2
  echo "Available versions:" >&2
  ls -1 "$ROOT_DIR/dist" 2>/dev/null || true
  exit 1
fi

echo "Using $PYTAG"
echo "Installing $PACKAGE offline..."
python -m pip install --no-index --find-links "$WHEELS_DIR" --find-links "$DIST_DIR" "$PACKAGE"
"""

    (bundle_scripts_dir / "offline-install.cmd").write_text(install_cmd, encoding="utf-8")
    (bundle_scripts_dir / "offline-install.ps1").write_text(install_ps1, encoding="utf-8")
    sh_path = bundle_scripts_dir / "offline-install.sh"
    sh_path.write_text(install_sh, encoding="utf-8")
    try:
        sh_path.chmod(0o755)
    except Exception:
        pass


def build_bundle(versions: List[str], require_all: bool, extras: List[str]) -> None:
    wheels_dir = REPO_ROOT / "wheels"
    dist_dir = REPO_ROOT / "dist"
    bundle_dir = REPO_ROOT / "offline_bundle"
    bundle_scripts_dir = bundle_dir / "scripts"

    # Clean
    for target in [wheels_dir, dist_dir, bundle_dir]:
        if target.exists():
            shutil.rmtree(target)

    wheels_dir.mkdir(parents=True, exist_ok=True)
    dist_dir.mkdir(parents=True, exist_ok=True)
    bundle_scripts_dir.mkdir(parents=True, exist_ok=True)

    requested = _find_python_commands(versions)
    available = _available_python_commands(requested)

    if not available:
        raise RuntimeError("No requested Python runtimes found.")

    if require_all and len(available) < len({tuple(c) for c in requested}):
        raise RuntimeError("Not all requested Python runtimes were found.")

    resolved_tags: List[str] = []

    # Build per Python version into wheels/<tag> and dist/<tag>
    for py_cmd in available:
        tag = _detect_py_tag(py_cmd)
        resolved_tags.append(tag)
        print(f"Building artifacts with {' '.join(py_cmd)} (tag={tag})")

        (wheels_dir / tag).mkdir(parents=True, exist_ok=True)
        (dist_dir / tag).mkdir(parents=True, exist_ok=True)

        _run([*py_cmd, "-m", "pip", "download", "-r", "requirements.txt", "-d", str(wheels_dir / tag)], REPO_ROOT)
        if extras:
            extra_spec = ",".join(extras)
            _run([*py_cmd, "-m", "pip", "download", f".[{extra_spec}]", "-d", str(wheels_dir / tag)], REPO_ROOT)
        _run([*py_cmd, "-m", "pip", "wheel", ".", "-w", str(dist_dir / tag), "--no-deps"], REPO_ROOT)

    # Copy to offline_bundle preserving per-version structure
    (bundle_dir / "wheels").mkdir(parents=True, exist_ok=True)
    (bundle_dir / "dist").mkdir(parents=True, exist_ok=True)

    shutil.copytree(wheels_dir, bundle_dir / "wheels", dirs_exist_ok=True)
    shutil.copytree(dist_dir, bundle_dir / "dist", dirs_exist_ok=True)

    _write_bundle_install_scripts(bundle_scripts_dir)

    readme = f"""# Offline Bundle (Multi-Python)

This bundle was generated for offline installation of `window-layout` for multiple Python versions.

## Layout
- `wheels/<pyTAG>/` downloaded dependency wheels per Python version
- `dist/<pyTAG>/` project wheel per Python version
- `scripts/` offline install helpers

`pyTAG` format: `py311`, `py312`, `py313`, etc.

## Install (auto-selects correct version folder)
### Windows (CMD)
`scripts\\offline-install.cmd`

### Windows (PowerShell)
`./scripts/offline-install.ps1`

### macOS/Linux
`./scripts/offline-install.sh`

## Manual install (example for py313)
CLI only:
`python -m pip install --no-index --find-links wheels/py313 --find-links dist/py313 "window-layout[cli]"`

GUI:
`python -m pip install --no-index --find-links wheels/py313 --find-links dist/py313 "window-layout[gui]"`

## Extras
Extras bundled: {', '.join(extras) if extras else 'none'}
If you need GUI on the offline machine, build with `--extras gui` and install with `window-layout[gui]`.

## Build metadata
Requested Python versions: {', '.join(versions)}
Requested extras: {', '.join(extras) if extras else 'none'}
Resolved builders/tags: {', '.join(sorted(set(resolved_tags)))}
"""
    (bundle_dir / "README_OFFLINE.md").write_text(readme, encoding="utf-8")

    _zip_dir(bundle_dir, REPO_ROOT / "offline_bundle.zip")
    print("Created offline_bundle.zip")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build offline_bundle.zip for multiple Python runtimes.")
    parser.add_argument(
        "--python-versions",
        nargs="+",
        default=["3.13", "3.12", "3.11"],
        help="Python versions to attempt (default: 3.13 3.12 3.11)",
    )
    parser.add_argument(
        "--require-all",
        action="store_true",
        help="Fail if any requested Python runtime is not available.",
    )
    parser.add_argument(
        "--extras",
        nargs="*",
        default=[],
        help="Optional extras to include (e.g. gui).",
    )
    args = parser.parse_args()
    build_bundle(args.python_versions, args.require_all, args.extras)


if __name__ == "__main__":
    main()
