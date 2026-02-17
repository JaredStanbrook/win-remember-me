#!/usr/bin/env python3
"""Build offline bundle assets for multiple Python versions."""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path
from typing import List
from zipfile import ZIP_DEFLATED, ZipFile


REPO_ROOT = Path(__file__).resolve().parents[1]


def _run(cmd: List[str], cwd: Path) -> None:
    print("+", " ".join(cmd))
    subprocess.run(cmd, cwd=str(cwd), check=True)


def _find_python_commands(versions: List[str]) -> List[List[str]]:
    commands: List[List[str]] = []
    if sys.platform.startswith("win"):
        for version in versions:
            commands.append(["py", f"-{version}"])
    else:
        for version in versions:
            commands.append([f"python{version}"])
        commands.append(["python3"])
        commands.append(["python"])
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
            subprocess.run([*cmd, "--version"], check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
            available.append(cmd)
        except Exception:
            continue
    return available


def _python_tag(py_cmd: List[str]) -> str:
    safe = "_".join(py_cmd).replace("-", "")
    return safe


def _write_offline_wrapper_scripts(bundle_scripts_dir: Path) -> None:
    cmd_wrapper = """@echo off
setlocal
set SCRIPT_DIR=%~dp0
call "%SCRIPT_DIR%install_offline.cmd" %*
endlocal
"""
    ps_wrapper = """param(
    [string]$Package = "window-layout-cli",
    [string]$DistDir = "dist",
    [string]$WheelsDir = "wheels"
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$ScriptDir\\install_offline.ps1" -Package $Package -DistDir $DistDir -WheelsDir $WheelsDir
"""
    zsh_wrapper = """#!/usr/bin/env zsh
set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
"$SCRIPT_DIR/install_offline.sh" "$@"
"""

    (bundle_scripts_dir / "offline-install.cmd").write_text(cmd_wrapper, encoding="utf-8")
    (bundle_scripts_dir / "offline-install.ps1").write_text(ps_wrapper, encoding="utf-8")
    (bundle_scripts_dir / "offline-install.zsh").write_text(zsh_wrapper, encoding="utf-8")


def _copy_install_scripts(bundle_scripts_dir: Path) -> None:
    for script_name in ["install_offline.cmd", "install_offline.ps1", "install_offline.sh"]:
        shutil.copy2(REPO_ROOT / "scripts" / script_name, bundle_scripts_dir / script_name)


def _zip_dir(source_dir: Path, zip_path: Path) -> None:
    if zip_path.exists():
        zip_path.unlink()
    with ZipFile(zip_path, "w", compression=ZIP_DEFLATED) as zf:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                zf.write(path, path.relative_to(source_dir))


def build_bundle(versions: List[str], require_all: bool) -> None:
    wheels_dir = REPO_ROOT / "wheels"
    dist_dir = REPO_ROOT / "dist"
    bundle_dir = REPO_ROOT / "offline_bundle"
    bundle_scripts_dir = bundle_dir / "scripts"

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

    for py_cmd in available:
        tag = _python_tag(py_cmd)
        print(f"Building artifacts with {' '.join(py_cmd)} (tag={tag})")
        _run([*py_cmd, "-m", "pip", "download", "-r", "requirements.txt", "-d", str(wheels_dir / tag)], REPO_ROOT)
        _run([*py_cmd, "-m", "pip", "wheel", ".", "-w", str(dist_dir / tag), "--no-deps"], REPO_ROOT)

    for src_base, dst_base in [(wheels_dir, bundle_dir / "wheels"), (dist_dir, bundle_dir / "dist")]:
        dst_base.mkdir(parents=True, exist_ok=True)
        for path in src_base.rglob("*"):
            if path.is_file():
                shutil.copy2(path, dst_base / path.name)

    _copy_install_scripts(bundle_scripts_dir)
    _write_offline_wrapper_scripts(bundle_scripts_dir)

    readme = f"""# Offline Bundle

This bundle was generated for offline installation of `window-layout-cli`.

## Included folders
- `wheels/` (downloaded dependency wheels)
- `dist/` (project wheel builds)
- `scripts/` install helpers

## Install options
- CMD: `scripts\\offline-install.cmd`
- PowerShell: `./scripts/offline-install.ps1`
- Zsh: `./scripts/offline-install.zsh`

These wrappers call the underlying `install_offline.*` scripts and install with:
`python -m pip install --no-index --find-links wheels --find-links dist window-layout-cli`

## Build metadata
Requested Python versions: {', '.join(versions)}
Resolved builders: {', '.join(' '.join(c) for c in available)}
"""
    (bundle_dir / "README_OFFLINE.md").write_text(readme, encoding="utf-8")

    _zip_dir(bundle_dir, REPO_ROOT / "offline_bundle.zip")
    print("Created offline_bundle.zip")


def main() -> None:
    parser = argparse.ArgumentParser(description="Build offline_bundle.zip for common Python runtimes.")
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
    args = parser.parse_args()
    build_bundle(args.python_versions, args.require_all)


if __name__ == "__main__":
    main()
