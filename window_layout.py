import argparse
import json
import os
import subprocess
import time
import urllib.request
from dataclasses import dataclass, asdict
from typing import Dict, List, Optional, Tuple

import psutil
import win32con
import win32api
import win32gui
import win32process
import win32com.client


@dataclass
class WindowEntry:
    title: str
    class_name: str
    pid: int
    process_name: str
    exe: str
    is_visible: bool
    is_minimized: bool
    is_maximized: bool
    rect: Tuple[int, int, int, int]  # (left, top, right, bottom)
    normal_rect: Tuple[int, int, int, int]  # from GetWindowPlacement (restored size/pos)
    show_cmd: int  # SW_SHOWNORMAL / SW_SHOWMAXIMIZED / SW_SHOWMINIMIZED etc.


def _safe_get_text(hwnd: int) -> str:
    try:
        return win32gui.GetWindowText(hwnd) or ""
    except Exception:
        return ""


def _safe_get_class(hwnd: int) -> str:
    try:
        return win32gui.GetClassName(hwnd) or ""
    except Exception:
        return ""


def _get_pid(hwnd: int) -> int:
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid or 0)
    except Exception:
        return 0


def _proc_info(pid: int) -> Tuple[str, str]:
    if not pid:
        return ("", "")
    try:
        p = psutil.Process(pid)
        name = p.name()
        exe = p.exe() if p.exe() else ""
        return (name or "", exe or "")
    except Exception:
        return ("", "")


def _window_rect(hwnd: int) -> Tuple[int, int, int, int]:
    try:
        return tuple(win32gui.GetWindowRect(hwnd))
    except Exception:
        return (0, 0, 0, 0)


def _window_placement(hwnd: int) -> Tuple[int, Tuple[int, int, int, int]]:
    """
    Returns (showCmd, normalPositionRect)
    normalPositionRect is a RECT tuple (left, top, right, bottom)
    """
    try:
        placement = win32gui.GetWindowPlacement(hwnd)
        # placement = (flags, showCmd, ptMin, ptMax, rcNormalPosition)
        show_cmd = int(placement[1])
        rc_normal = tuple(placement[4])
        return show_cmd, rc_normal
    except Exception:
        return (win32con.SW_SHOWNORMAL, (0, 0, 0, 0))


def _is_interesting_window(hwnd: int) -> bool:
    """
    Keep this conservative to avoid capturing tool windows, hidden owners, etc.
    """
    if not win32gui.IsWindow(hwnd):
        return False

    # top-level only (has no parent)
    if win32gui.GetParent(hwnd):
        return False

    # ignore invisible
    if not win32gui.IsWindowVisible(hwnd):
        return False

    # ignore tiny/empty title windows (still keep some with empty titles? up to you)
    title = _safe_get_text(hwnd).strip()
    if len(title) == 0:
        return False

    # ignore cloaked windows (some UWP) – best-effort: if it errors, ignore check
    # (There isn't a simple pywin32 call; leaving out to keep dependencies minimal.)
    return True


def capture_windows() -> List[Dict]:
    entries: List[Tuple[WindowEntry, int]] = []
    explorer_paths = _explorer_window_paths()

    def enum_cb(hwnd, _):
        if not _is_interesting_window(hwnd):
            return

        title = _safe_get_text(hwnd).strip()
        class_name = _safe_get_class(hwnd).strip()
        pid = _get_pid(hwnd)
        proc_name, exe = _proc_info(pid)

        rect = _window_rect(hwnd)
        show_cmd, normal_rect = _window_placement(hwnd)

        entry = WindowEntry(
            title=title,
            class_name=class_name,
            pid=pid,
            process_name=proc_name,
            exe=exe,
            is_visible=bool(win32gui.IsWindowVisible(hwnd)),
            is_minimized=bool(win32gui.IsIconic(hwnd)),
            is_maximized=bool(show_cmd == win32con.SW_SHOWMAXIMIZED),
            rect=rect,
            normal_rect=normal_rect,
            show_cmd=show_cmd,
        )
        entries.append((entry, hwnd))

    win32gui.EnumWindows(enum_cb, None)

    # Sort for stable output
    entries.sort(key=lambda e: (e[0].process_name.lower(), e[0].class_name.lower(), e[0].title.lower()))
    payload: List[Dict] = []
    for e, hwnd in entries:
        data = asdict(e)
        if e.exe:
            launch_args: List[str] = []
            if e.process_name.lower() == "explorer.exe":
                path = explorer_paths.get(hwnd)
                if path:
                    launch_args = [path]
            data["launch"] = {"exe": e.exe, "args": launch_args, "cwd": ""}
        payload.append(data)
    return payload


def _explorer_window_paths() -> Dict[int, str]:
    paths: Dict[int, str] = {}
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        windows = shell.Windows()
        for window in windows:
            try:
                hwnd = int(getattr(window, "HWND", 0) or 0)
                location = str(getattr(window, "LocationURL", "") or "").strip()
                if location.startswith("file:///"):
                    path = location.replace("file:///", "").replace("/", "\\")
                    if hwnd and path:
                        paths[hwnd] = path
            except Exception:
                continue
    except Exception:
        return paths
    return paths


def _fetch_edge_tabs(debug_port: int = 9222) -> List[Dict]:
    url = f"http://127.0.0.1:{int(debug_port)}/json/list"
    try:
        with urllib.request.urlopen(url, timeout=1.5) as response:
            raw = response.read().decode("utf-8", errors="replace")
        payload = json.loads(raw)
    except Exception:
        return []

    tabs: List[Dict] = []
    for item in payload:
        if item.get("type") != "page":
            continue
        tab_url = str(item.get("url") or "").strip()
        if not tab_url or tab_url.startswith("edge://") or tab_url.startswith("chrome://"):
            continue
        tabs.append({
            "title": str(item.get("title") or "").strip(),
            "url": tab_url,
        })
    return tabs


def save_layout(path: str, capture_edge_tabs: bool = False, edge_debug_port: int = 9222) -> None:
    data = {
        "schema": "window-layout.v1",
        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
        "windows": capture_windows(),
    }
    if capture_edge_tabs:
        tabs = _fetch_edge_tabs(edge_debug_port)
        data["browser_tabs"] = {
            "edge": {
                "debug_port": int(edge_debug_port),
                "captured_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                "tabs": tabs,
                "note": "Requires Edge started with --remote-debugging-port"
            }
        }
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    tab_count = 0
    if capture_edge_tabs:
        tab_count = len(data.get("browser_tabs", {}).get("edge", {}).get("tabs", []))
    print(f"Saved {len(data['windows'])} windows, {tab_count} Edge tabs -> {path}")
    if capture_edge_tabs and tab_count == 0:
        print("Note: no Edge tabs captured. Start Edge with --remote-debugging-port and retry.")


def _score_match(candidate: Dict, target: Dict) -> int:
    """
    Higher is better. We try to avoid relying on HWND (not stable).
    """
    score = 0

    # Strong match: exe path if present
    if target.get("exe") and candidate.get("exe") and target["exe"].lower() == candidate["exe"].lower():
        score += 50

    # Next: process name
    if target.get("process_name") and candidate.get("process_name") and \
       target["process_name"].lower() == candidate["process_name"].lower():
        score += 25

    # Class name
    if target.get("class_name") and candidate.get("class_name") and \
       target["class_name"].lower() == candidate["class_name"].lower():
        score += 15

    # Title: exact then partial
    t_title = (target.get("title") or "").strip().lower()
    c_title = (candidate.get("title") or "").strip().lower()
    if t_title and c_title:
        if t_title == c_title:
            score += 30
        elif t_title in c_title or c_title in t_title:
            score += 10

    return score


def _current_windows_with_hwnds() -> List[Dict]:
    results: List[Dict] = []

    def enum_cb(hwnd, _):
        if not win32gui.IsWindow(hwnd):
            return
        if win32gui.GetParent(hwnd):
            return
        if not win32gui.IsWindowVisible(hwnd):
            return

        title = _safe_get_text(hwnd).strip()
        class_name = _safe_get_class(hwnd).strip()
        pid = _get_pid(hwnd)
        proc_name, exe = _proc_info(pid)
        show_cmd, normal_rect = _window_placement(hwnd)
        rect = _window_rect(hwnd)

        results.append({
            "hwnd": hwnd,
            "title": title,
            "class_name": class_name,
            "pid": pid,
            "process_name": proc_name,
            "exe": exe,
            "show_cmd": show_cmd,
            "normal_rect": normal_rect,
            "rect": rect
        })

    win32gui.EnumWindows(enum_cb, None)
    return results


def _get_launch_spec(target: Dict) -> Optional[Tuple[str, List[str], str]]:
    launch = target.get("launch")
    exe = ""
    args: List[str] = []
    cwd = ""

    if isinstance(launch, dict):
        exe = str(launch.get("exe") or "").strip()
        raw_args = launch.get("args") or []
        if isinstance(raw_args, str):
            args = [raw_args]
        elif isinstance(raw_args, list):
            args = [str(a) for a in raw_args if str(a).strip()]
        cwd = str(launch.get("cwd") or "").strip()

    if not exe:
        exe = str(target.get("exe") or "").strip()

    if not exe:
        return None

    if os.path.basename(exe).lower() == "applicationframehost.exe":
        return None

    return exe, args, cwd


def _launch_target(target: Dict, dry_run: bool = False) -> bool:
    spec = _get_launch_spec(target)
    if not spec:
        return False

    exe, args, cwd = spec
    if not os.path.exists(exe):
        return False

    if dry_run:
        print(f"[DRY] Launch -> {exe} {' '.join(args)}")
        return True

    try:
        subprocess.Popen([exe, *args], cwd=cwd or None)
        return True
    except Exception:
        return False


def _best_match(target: Dict, current: List[Dict], used_hwnds: set, min_score: int) -> Tuple[Optional[Dict], int]:
    best = None
    best_score = -1
    for c in current:
        if c["hwnd"] in used_hwnds:
            continue
        score = _score_match(c, target)
        if score > best_score:
            best_score = score
            best = c
    if not best or best_score < min_score:
        return None, best_score
    return best, best_score


def _edge_exe_from_targets(targets: List[Dict]) -> Optional[str]:
    for t in targets:
        if str(t.get("process_name") or "").lower() == "msedge.exe":
            exe = str(t.get("exe") or "").strip()
            if exe:
                return exe
    return None


def _launch_edge_tabs(exe: str, tabs: List[Dict], dry_run: bool = False) -> int:
    urls = [t.get("url") for t in tabs if str(t.get("url") or "").strip()]
    if not urls:
        return 0

    if not os.path.exists(exe):
        return 0

    launched = 0
    chunk_size = 10
    for idx in range(0, len(urls), chunk_size):
        chunk = urls[idx:idx + chunk_size]
        args = ["--new-window", *chunk]
        if dry_run:
            print(f"[DRY] Launch Edge tabs -> {exe} {' '.join(args)}")
            launched += len(chunk)
            continue
        try:
            subprocess.Popen([exe, *args])
            launched += len(chunk)
        except Exception:
            break
    return launched


def _find_edge_exe() -> Optional[str]:
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for path in candidates:
        if os.path.exists(path):
            return path
    return None


def launch_edge_debug(port: int = 9222, profile_dir: Optional[str] = None, dry_run: bool = False) -> bool:
    exe = _find_edge_exe()
    if not exe:
        print("Edge not found. Install Edge or provide a custom path.")
        return False

    if not profile_dir:
        profile_dir = os.path.join(os.environ.get("TEMP", r"C:\Temp"), "edge-debug")

    args = [
        "--remote-debugging-port=" + str(int(port)),
        "--user-data-dir=" + profile_dir,
    ]

    if dry_run:
        print(f"[DRY] Launch Edge debug -> {exe} {' '.join(args)}")
        return True

    try:
        subprocess.Popen([exe, *args])
        return True
    except Exception:
        return False


def _prompt(text: str, default: Optional[str] = None) -> str:
    suffix = f" [{default}]" if default else ""
    value = input(f"{text}{suffix}: ").strip()
    return value or (default or "")


def _prompt_yes_no(text: str, default: bool = False) -> bool:
    default_str = "Y/n" if default else "y/N"
    value = input(f"{text} ({default_str}): ").strip().lower()
    if not value:
        return default
    return value in ("y", "yes")


def run_setup_wizard() -> None:
    print("TSD Workspace Setup Wizard")
    print("This will capture your current window layout for fast restores.")
    out_path = _prompt("Output layout path", "layout.json")

    capture_edge = _prompt_yes_no("Capture Edge tabs (requires Edge debug)", default=False)
    edge_port = 9222
    if capture_edge:
        port_raw = _prompt("Edge debug port", "9222")
        try:
            edge_port = int(port_raw)
        except ValueError:
            edge_port = 9222

        if _prompt_yes_no("Launch Edge debug instance now", default=True):
            ok = launch_edge_debug(port=edge_port)
            if ok:
                print("Edge debug session launched.")
                if not _prompt_yes_no("Ready to capture now", default=True):
                    input("Press Enter to capture when ready...")
            else:
                print("Failed to launch Edge debug session. Tabs may not capture.")

    save_layout(out_path, capture_edge_tabs=capture_edge, edge_debug_port=edge_port)
    print("Wizard complete.")


def _apply_window_position(hwnd: int, entry: Dict) -> bool:
    """
    Restore window to normal, move/resize to saved normal_rect, then
    re-apply minimized/maximized state.
    """
    try:
        # Bring out of minimized/maximized so MoveWindow works reliably
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.02)

        left, top, right, bottom = entry["normal_rect"]
        if int(entry.get("show_cmd", win32con.SW_SHOWNORMAL)) == win32con.SW_SHOWNORMAL:
            rect = entry.get("rect")
            if isinstance(rect, (list, tuple)) and len(rect) == 4:
                r_left, r_top, r_right, r_bottom = rect
                if abs((r_right - r_left) - (right - left)) > 50 or abs((r_bottom - r_top) - (bottom - top)) > 50:
                    left, top, right, bottom = rect
        width = max(50, right - left)
        height = max(50, bottom - top)

        left, top, width, height = _clamp_to_visible_bounds(left, top, width, height)

        # Move & resize
        win32gui.MoveWindow(hwnd, int(left), int(top), int(width), int(height), True)

        # Re-apply show state (best effort)
        desired = int(entry.get("show_cmd", win32con.SW_SHOWNORMAL))
        if desired == win32con.SW_SHOWMAXIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWMAXIMIZED)
        elif desired == win32con.SW_SHOWMINIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWMINIMIZED)
        else:
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNORMAL)

        return True
    except Exception:
        return False


def _clamp_to_visible_bounds(left: int, top: int, width: int, height: int) -> Tuple[int, int, int, int]:
    try:
        monitors = win32api.EnumDisplayMonitors()
        bounds = []
        for monitor in monitors:
            info = win32api.GetMonitorInfo(monitor[0])
            bounds.append(info.get("Monitor", info.get("Work")))

        if not bounds:
            return left, top, width, height

        # Only clamp if fully outside all monitor bounds.
        rect = (left, top, left + width, top + height)
        if not any(_rects_intersect(rect, wa) for wa in bounds):
            primary = bounds[0]
            max_left = primary[2] - width
            max_top = primary[3] - height
            left = min(max(left, primary[0]), max_left)
            top = min(max(top, primary[1]), max_top)
            return left, top, width, height

        return left, top, width, height
    except Exception:
        return left, top, width, height


def _rects_intersect(a: Tuple[int, int, int, int], b: Tuple[int, int, int, int]) -> bool:
    return not (a[2] <= b[0] or a[0] >= b[2] or a[3] <= b[1] or a[1] >= b[3])


def restore_layout(
    path: str,
    min_score: int = 40,
    dry_run: bool = False,
    launch_missing: bool = False,
    launch_wait: float = 6.0,
    restore_edge_tabs: bool = False,
) -> None:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if data.get("schema") != "window-layout.v1":
        raise ValueError("Unsupported JSON schema (expected window-layout.v1)")

    targets = data.get("windows", [])
    current = _current_windows_with_hwnds()

    used_hwnds = set()
    applied = 0
    skipped = 0
    missing: List[Dict] = []
    edge_tabs_launched = 0
    edge_tabs_present = bool(data.get("browser_tabs", {}).get("edge", {}).get("tabs", []))

    for t in targets:
        best, best_score = _best_match(t, current, used_hwnds, min_score)
        if not best:
            missing.append(t)
            continue

        used_hwnds.add(best["hwnd"])

        if dry_run:
            print(f"[DRY] Match score={best_score:3d} | {t['process_name']} | {t['title']}  ->  hwnd={best['hwnd']}")
            applied += 1
            continue

        ok = _apply_window_position(best["hwnd"], t)
        if ok:
            applied += 1
        else:
            skipped += 1

    launched = 0
    if launch_missing and missing:
        for t in missing:
            if edge_tabs_present and restore_edge_tabs and str(t.get("process_name") or "").lower() == "msedge.exe":
                continue
            if _launch_target(t, dry_run=dry_run):
                launched += 1

        if launched and not dry_run:
            time.sleep(max(0.5, float(launch_wait)))

        if launched:
            current = _current_windows_with_hwnds()
            remaining: List[Dict] = []
            for t in missing:
                best, best_score = _best_match(t, current, used_hwnds, min_score)
                if not best:
                    remaining.append(t)
                    continue

                used_hwnds.add(best["hwnd"])

                if dry_run:
                    print(f"[DRY] Match score={best_score:3d} | {t['process_name']} | {t['title']}  ->  hwnd={best['hwnd']}")
                    applied += 1
                    continue

                ok = _apply_window_position(best["hwnd"], t)
                if ok:
                    applied += 1
                else:
                    skipped += 1

            missing = remaining

    if restore_edge_tabs and missing:
        edge_missing = [t for t in missing if str(t.get("process_name") or "").lower() == "msedge.exe"]
        if edge_missing:
            edge_tabs = data.get("browser_tabs", {}).get("edge", {}).get("tabs", [])
            edge_exe = _edge_exe_from_targets(targets)
            if edge_tabs and edge_exe:
                edge_tabs_launched = _launch_edge_tabs(edge_exe, edge_tabs, dry_run=dry_run)
            else:
                for t in edge_missing:
                    if _launch_target(t, dry_run=dry_run):
                        launched += 1

            if (edge_tabs_launched or launched) and not dry_run:
                time.sleep(max(0.5, float(launch_wait)))

            current = _current_windows_with_hwnds()
            remaining: List[Dict] = []
            for t in missing:
                best, best_score = _best_match(t, current, used_hwnds, min_score)
                if not best:
                    remaining.append(t)
                    continue

                used_hwnds.add(best["hwnd"])

                if dry_run:
                    print(f"[DRY] Match score={best_score:3d} | {t['process_name']} | {t['title']}  ->  hwnd={best['hwnd']}")
                    applied += 1
                    continue

                ok = _apply_window_position(best["hwnd"], t)
                if ok:
                    applied += 1
                else:
                    skipped += 1

            missing = remaining

    skipped += len(missing)

    if restore_edge_tabs and not edge_tabs_launched:
        edge_tabs = data.get("browser_tabs", {}).get("edge", {}).get("tabs", [])
        edge_exe = _edge_exe_from_targets(targets)
        if edge_exe and edge_tabs:
            edge_tabs_launched = _launch_edge_tabs(edge_exe, edge_tabs, dry_run=dry_run)

    print(
        f"Restore complete. Applied={applied}, Skipped={skipped}, "
        f"TotalTargets={len(targets)}, Launched={launched}, EdgeTabs={edge_tabs_launched}"
    )
    if skipped:
        print("Note: windows may be skipped if titles changed, apps closed, elevated windows, or restricted window types.")


def main():
    parser = argparse.ArgumentParser(description="Save/restore window layout to/from JSON (Windows).")
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_save = sub.add_parser("save", help="Capture current window positions -> JSON")
    p_save.add_argument("json_path", help="Output JSON path")
    p_save.add_argument("--edge-tabs", action="store_true", help="Capture Edge tabs via remote debugging")
    p_save.add_argument("--edge-debug-port", type=int, default=9222, help="Edge remote debugging port (default: 9222)")

    p_edge = sub.add_parser("edge-debug", help="Launch Edge with remote debugging enabled")
    p_edge.add_argument("--port", type=int, default=9222, help="Remote debugging port (default: 9222)")
    p_edge.add_argument("--profile-dir", help="User data dir for debug session (default: %TEMP%\\edge-debug)")
    p_edge.add_argument("--dry-run", action="store_true", help="Only show launch command")

    sub.add_parser("wizard", help="Interactive first-time setup wizard")

    p_restore = sub.add_parser("restore", help="Restore window positions from JSON")
    p_restore.add_argument("json_path", help="Input JSON path")
    p_restore.add_argument("--min-score", type=int, default=40, help="Matching threshold (default: 40)")
    p_restore.add_argument("--dry-run", action="store_true", help="Only show matches, do not move windows")
    p_restore.add_argument("--launch-missing", action="store_true", help="Launch apps for missing windows before restore")
    p_restore.add_argument("--launch-wait", type=float, default=6.0, help="Seconds to wait after launch (default: 6)")
    p_restore.add_argument("--restore-edge-tabs", action="store_true", help="Reopen Edge tabs captured during save")

    p_help = sub.add_parser("help", help="Show quick usage")
    p_help.add_argument("--full", action="store_true", help="Show full argparse help")

    args = parser.parse_args()

    if args.cmd == "save":
        save_layout(
            args.json_path,
            capture_edge_tabs=args.edge_tabs,
            edge_debug_port=args.edge_debug_port,
        )
    elif args.cmd == "edge-debug":
        ok = launch_edge_debug(port=args.port, profile_dir=args.profile_dir, dry_run=args.dry_run)
        if ok:
            print("Edge debug session launched.")
        else:
            print("Failed to launch Edge debug session.")
    elif args.cmd == "wizard":
        run_setup_wizard()
    elif args.cmd == "restore":
        restore_layout(
            args.json_path,
            min_score=args.min_score,
            dry_run=args.dry_run,
            launch_missing=args.launch_missing,
            launch_wait=args.launch_wait,
            restore_edge_tabs=args.restore_edge_tabs,
        )
    elif args.cmd == "help":
        if args.full:
            parser.print_help()
            return
        print("Quick Help")
        print("  save:    python window_layout.py save layout.json")
        print("  restore: python window_layout.py restore layout.json")
        print("  restore (launch missing): python window_layout.py restore layout.json --launch-missing")
        print("  edge debug: python window_layout.py edge-debug")
        print("  edge tabs:  python window_layout.py save layout.json --edge-tabs")
        print("  restore tabs: python window_layout.py restore layout.json --restore-edge-tabs")
        print("  wizard: python window_layout.py wizard")


if __name__ == "__main__":
    main()
