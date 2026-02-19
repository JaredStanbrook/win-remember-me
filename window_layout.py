import argparse
import hashlib
import json
import os
import socket
import subprocess
import sys
import time
import urllib.request
import uuid
from dataclasses import dataclass, asdict
from typing import Dict, Iterable, List, Optional, Tuple

import psutil
import win32con
import win32api
import win32gui
import win32process
import win32com.client

SCHEMA_V1 = "window-layout.v1"
SCHEMA_V2 = "window-layout.v2"
CONFIG_PATH = "config.json"


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


def _now() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")


def _generate_window_id() -> str:
    return str(uuid.uuid4())


def _ensure_window_ids(windows: List[Dict]) -> None:
    for window in windows:
        if not str(window.get("window_id") or "").strip():
            window["window_id"] = _generate_window_id()


def _is_schema_v2(data: Dict) -> bool:
    return str(data.get("schema") or "").strip().lower() == SCHEMA_V2


def _normalize_edge_tabs(tabs: Iterable[Dict]) -> List[Dict]:
    normalized: List[Dict] = []
    for tab in tabs:
        url = str(tab.get("url") or "").strip()
        if not url:
            continue
        normalized.append({
            "title": str(tab.get("title") or "").strip(),
            "url": url,
        })
    return normalized


def _coerce_url_list(raw: Iterable) -> List[Dict]:
    urls: List[Dict] = []
    for item in raw:
        if isinstance(item, dict):
            url = str(item.get("url") or "").strip()
            if url:
                urls.append({"title": str(item.get("title") or "").strip(), "url": url})
        else:
            url = str(item or "").strip()
            if url:
                urls.append({"title": "", "url": url})
    return urls


def _load_config(path: str = CONFIG_PATH) -> Dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}
    return data if isinstance(data, dict) else {}


def _parse_hotkey_keys(keys: str) -> Optional[Tuple[int, int]]:
    if not keys:
        return None
    parts = [p.strip() for p in keys.replace("-", "+").split("+") if p.strip()]
    if not parts:
        return None

    modifiers = 0
    key_part = ""
    for part in parts:
        lowered = part.lower()
        if lowered in ("ctrl", "control"):
            modifiers |= win32con.MOD_CONTROL
        elif lowered in ("alt",):
            modifiers |= win32con.MOD_ALT
        elif lowered in ("shift",):
            modifiers |= win32con.MOD_SHIFT
        elif lowered in ("win", "windows", "meta"):
            modifiers |= win32con.MOD_WIN
        else:
            key_part = part

    if not key_part:
        return None

    key_upper = key_part.upper()
    if len(key_upper) == 1 and key_upper.isalnum():
        vk = ord(key_upper)
        return modifiers, vk

    if key_upper.startswith("F") and key_upper[1:].isdigit():
        fn = int(key_upper[1:])
        if 1 <= fn <= 24:
            return modifiers, win32con.VK_F1 + (fn - 1)

    vk_map = {
        "TAB": win32con.VK_TAB,
        "ENTER": win32con.VK_RETURN,
        "RETURN": win32con.VK_RETURN,
        "ESC": win32con.VK_ESCAPE,
        "ESCAPE": win32con.VK_ESCAPE,
        "SPACE": win32con.VK_SPACE,
        "BACKSPACE": win32con.VK_BACK,
        "DELETE": win32con.VK_DELETE,
        "HOME": win32con.VK_HOME,
        "END": win32con.VK_END,
        "PGUP": win32con.VK_PRIOR,
        "PAGEUP": win32con.VK_PRIOR,
        "PGDN": win32con.VK_NEXT,
        "PAGEDOWN": win32con.VK_NEXT,
        "LEFT": win32con.VK_LEFT,
        "RIGHT": win32con.VK_RIGHT,
        "UP": win32con.VK_UP,
        "DOWN": win32con.VK_DOWN,
    }
    if key_upper in vk_map:
        return modifiers, vk_map[key_upper]
    return None


def _load_hotkeys(path: str = CONFIG_PATH) -> List[Dict]:
    data = _load_config(path)
    raw = data.get("hotkeys") or []
    if not isinstance(raw, list):
        return []
    entries: List[Dict] = []
    for item in raw:
        if not isinstance(item, dict):
            continue
        keys = str(item.get("keys") or "").strip()
        action = str(item.get("action") or "").strip()
        args = item.get("args") or []
        if isinstance(args, str):
            args = [args]
        if not isinstance(args, list):
            args = []
        args = [str(a) for a in args if str(a).strip()]
        if not keys or not action:
            continue
        entries.append({"keys": keys, "action": action, "args": args})
    return entries


def _run_hotkey_action(action: str, args: List[str]) -> None:
    cmd = [sys.executable, os.path.abspath(__file__), action, *args]
    try:
        subprocess.Popen(cmd)
    except Exception:
        pass


def run_hotkey_listener(config_path: str = CONFIG_PATH) -> None:
    hotkeys = _load_hotkeys(config_path)
    if not hotkeys:
        print("No hotkeys configured.")
        return

    try:
        win32gui.PeekMessage(None, 0, 0, win32con.PM_NOREMOVE)
    except Exception:
        pass

    registered: Dict[int, Dict] = {}
    next_id = 1
    for entry in hotkeys:
        parsed = _parse_hotkey_keys(entry["keys"])
        if not parsed:
            print(f"Skipping invalid hotkey: {entry['keys']}")
            continue
        modifiers, vk = parsed
        try:
            win32gui.RegisterHotKey(None, next_id, modifiers, vk)
            registered[next_id] = entry
            print(f"Registered {entry['keys']} -> {entry['action']} {' '.join(entry['args'])}")
            next_id += 1
        except Exception as exc:
            print(f"Failed to register hotkey: {entry['keys']} ({exc})")

    if not registered:
        print("No hotkeys registered.")
        return

    try:
        while True:
            msg = win32gui.GetMessage(None, 0, 0)
            if not msg:
                continue
            payload = msg
            if isinstance(msg, (list, tuple)) and len(msg) == 2:
                payload = msg[1]
            message = None
            wparam = None
            if isinstance(payload, (list, tuple)):
                if len(payload) == 6:
                    _hwnd, message, wparam, _lparam, _time, _pt = payload
                elif len(payload) == 3:
                    message, wparam, _lparam = payload
                elif len(payload) == 2:
                    message, wparam = payload
            if message == win32con.WM_QUIT:
                break
            if message == win32con.WM_HOTKEY:
                entry = registered.get(wparam)
                if entry:
                    _run_hotkey_action(entry["action"], entry.get("args", []))
    except KeyboardInterrupt:
        pass
    finally:
        for hotkey_id in registered:
            try:
                win32api.UnregisterHotKey(None, hotkey_id)
            except Exception:
                continue


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


def _normalize_exe_path(path: str) -> str:
    raw = str(path or "").strip()
    if not raw:
        return ""
    normalized = os.path.normpath(raw).replace("/", "\\")
    return normalized.lower()


def _quantize_dimension(value: int, bucket_size: int = 100) -> int:
    if value <= 0:
        return 0
    return int(round(value / float(bucket_size)) * bucket_size)


def _window_size_bucket(rect: Tuple[int, int, int, int]) -> str:
    width = max(0, int(rect[2]) - int(rect[0]))
    height = max(0, int(rect[3]) - int(rect[1]))
    return f"{_quantize_dimension(width)}x{_quantize_dimension(height)}"


def _monitor_bucket(rect: Tuple[int, int, int, int]) -> str:
    left = int(rect[0])
    top = int(rect[1])
    return f"{left // 400}:{top // 300}"


def _build_match_fingerprint(window: Dict) -> str:
    raw_rect = window.get("rect") or (0, 0, 0, 0)
    rect = tuple(raw_rect) if isinstance(raw_rect, (list, tuple)) else (0, 0, 0, 0)
    if len(rect) != 4:
        rect = (0, 0, 0, 0)
    payload = "|".join([
        _normalize_exe_path(window.get("exe") or ""),
        str(window.get("class_name") or "").strip().lower(),
        str(window.get("process_name") or "").strip().lower(),
        _window_size_bucket(rect),
        _monitor_bucket(rect),
    ])
    if not payload.replace("|", ""):
        return ""
    return hashlib.sha1(payload.encode("utf-8")).hexdigest()


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

    # keep taskbar-style app windows and skip helper/tool windows.
    if not _is_taskbar_window(hwnd):
        return False

    rect = _window_rect(hwnd)
    if (rect[2] - rect[0]) < 120 or (rect[3] - rect[1]) < 80:
        return False

    # ignore cloaked windows (some UWP) – best-effort: if it errors, ignore check
    # (There isn't a simple pywin32 call; leaving out to keep dependencies minimal.)
    return True


def _is_taskbar_window(hwnd: int) -> bool:
    """
    Heuristic: include windows users typically interact with from taskbar/Alt-Tab.
    """
    try:
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        owner = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
    except Exception:
        return True

    if ex_style & win32con.WS_EX_TOOLWINDOW:
        return False

    if owner and not (ex_style & win32con.WS_EX_APPWINDOW):
        return False

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
        data["window_id"] = _generate_window_id()
        data["match_fingerprint"] = _build_match_fingerprint(data)
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
            "window_id": item.get("windowId"),
            "target_id": str(item.get("id") or "").strip(),
        })
    return tabs


def _normalize_edge_window_title(title: str) -> str:
    normalized = (title or "").replace("Microsoft\u200b Edge", "Microsoft Edge").strip()
    lowered = normalized.lower()
    marker = " - microsoft edge"
    if marker in lowered:
        idx = lowered.rfind(marker)
        normalized = normalized[:idx]
    return normalized.strip().lower()


def _assign_edge_tabs_to_windows(windows: List[Dict], tabs: List[Dict], target_windows: Optional[List[Dict]] = None) -> None:
    if target_windows is not None:
        edge_windows = target_windows
    else:
        edge_windows = [w for w in windows if str(w.get("process_name") or "").lower() == "msedge.exe"]
    for window in edge_windows:
        window["edge_tabs"] = []

    if not edge_windows:
        return

    windows_by_title: Dict[str, List[int]] = {}
    for idx, window in enumerate(edge_windows):
        normalized_title = _normalize_edge_window_title(str(window.get("title") or ""))
        if normalized_title:
            windows_by_title.setdefault(normalized_title, []).append(idx)

    assigned_windows = set()
    ungrouped_tabs: List[Dict] = []
    tabs_by_window_id: Dict[int, List[Dict]] = {}
    for tab in tabs:
        window_id = tab.get("window_id")
        if isinstance(window_id, int):
            tabs_by_window_id.setdefault(window_id, []).append(tab)
        else:
            ungrouped_tabs.append(tab)

    for grouped_tabs in tabs_by_window_id.values():
        match_idx: Optional[int] = None
        for tab in grouped_tabs:
            tab_title = _normalize_edge_window_title(str(tab.get("title") or ""))
            candidates = windows_by_title.get(tab_title, [])
            for idx in candidates:
                if idx not in assigned_windows:
                    match_idx = idx
                    break
            if match_idx is not None:
                break

        if match_idx is None:
            for idx in range(len(edge_windows)):
                if idx not in assigned_windows:
                    match_idx = idx
                    break

        if match_idx is None:
            continue

        assigned_windows.add(match_idx)
        for tab in grouped_tabs:
            edge_windows[match_idx]["edge_tabs"].append({
                "title": tab.get("title", ""),
                "url": tab.get("url", ""),
            })

    if ungrouped_tabs:
        idx = 0
        for tab in ungrouped_tabs:
            edge_windows[idx % len(edge_windows)]["edge_tabs"].append({
                "title": tab.get("title", ""),
                "url": tab.get("url", ""),
            })
            idx += 1


def _assign_tabs_to_edge_windows_stably(edge_windows: List[Dict], tabs: List[Dict]) -> None:
    if not edge_windows or not tabs:
        return

    windows_by_id: Dict[str, Dict] = {}
    windows_by_title: Dict[str, List[Dict]] = {}
    for window in edge_windows:
        window_id = str(window.get("window_id") or "").strip()
        if window_id:
            windows_by_id[window_id] = window
        title = _normalize_edge_window_title(str(window.get("title") or ""))
        if title:
            windows_by_title.setdefault(title, []).append(window)

    rr_index = 0
    for tab in tabs:
        assigned_window: Optional[Dict] = None
        tab_window_id = str(tab.get("window_id") or "").strip()
        if tab_window_id:
            assigned_window = windows_by_id.get(tab_window_id)

        if assigned_window is None:
            tab_title = _normalize_edge_window_title(str(tab.get("title") or ""))
            candidates = windows_by_title.get(tab_title) or []
            if candidates:
                assigned_window = candidates[0]

        if assigned_window is None:
            assigned_window = edge_windows[rr_index % len(edge_windows)]
            rr_index += 1

        assigned_window.setdefault("edge_tabs", []).append({
            "title": str(tab.get("title") or "").strip(),
            "url": str(tab.get("url") or "").strip(),
        })


def _migrate_legacy_edge_tab_storage(data: Dict) -> Dict:
    migrated = dict(data)
    windows = [dict(w) for w in migrated.get("windows", [])]
    _ensure_window_ids(windows)

    edge_windows = [w for w in windows if str(w.get("process_name") or "").lower() == "msedge.exe"]
    for window in edge_windows:
        window["edge_tabs"] = _normalize_edge_tabs(window.get("edge_tabs") or [])

    sessions = migrated.get("edge_sessions") if isinstance(migrated.get("edge_sessions"), list) else []
    for session in sessions:
        if not isinstance(session, dict):
            continue
        tabs = _normalize_edge_tabs(session.get("tabs") or [])
        if not tabs:
            continue
        window_ids = [str(wid).strip() for wid in (session.get("window_ids") or []) if str(wid).strip()]
        target_windows = [w for w in edge_windows if str(w.get("window_id") or "") in window_ids]
        if not target_windows:
            target_windows = edge_windows
        _assign_tabs_to_edge_windows_stably(target_windows, tabs)

    browser_edge = migrated.get("browser_tabs", {}).get("edge", {})
    if isinstance(browser_edge, dict):
        tabs = _normalize_edge_tabs(browser_edge.get("tabs") or [])
        _assign_tabs_to_edge_windows_stably(edge_windows, tabs)

    open_urls = migrated.get("open_urls") if isinstance(migrated.get("open_urls"), dict) else {}
    _assign_tabs_to_edge_windows_stably(edge_windows, _coerce_url_list(open_urls.get("edge") or []))

    migrated["windows"] = windows
    migrated.pop("browser_tabs", None)
    migrated.pop("edge_sessions", None)
    migrated.pop("open_urls", None)
    return migrated


def _migrate_v1_to_v2(data: Dict) -> Dict:
    upgraded = _migrate_legacy_edge_tab_storage(data)
    upgraded["schema"] = SCHEMA_V2
    return upgraded


def _ensure_v2_layout(data: Dict) -> Dict:
    if _is_schema_v2(data):
        data = _migrate_legacy_edge_tab_storage(data)
        _ensure_window_ids(data.get("windows", []))
        return data
    return _migrate_v1_to_v2(data)


def _collect_edge_tabs(data: Dict) -> List[Dict]:
    per_window: List[Dict] = []
    for window in data.get("windows", []):
        if str(window.get("process_name") or "").lower() != "msedge.exe":
            continue
        per_window.extend(_normalize_edge_tabs(window.get("edge_tabs") or []))
    return per_window


def _load_existing_metadata(path: str) -> Dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}

    preserved: Dict = {}
    for key in ("speed_menu", "custom_layout_folders", "layouts_root"):
        if key in data:
            preserved[key] = data[key]
    return preserved


def save_layout(
    path: str,
    capture_edge_tabs: bool = False,
    edge_debug_port: int = 9222,
    schema_version: str = SCHEMA_V2,
    edge_profile_dir: Optional[str] = None,
) -> None:
    windows = capture_windows()
    _ensure_window_ids(windows)
    preserved = _load_existing_metadata(path)
    data = {
        "schema": schema_version if schema_version in (SCHEMA_V1, SCHEMA_V2) else SCHEMA_V2,
        "created_at": _now(),
        "windows": windows,
    }
    if preserved:
        data.update(preserved)

    if capture_edge_tabs:
        tabs = _fetch_edge_tabs(edge_debug_port)
        _assign_edge_tabs_to_windows(windows, tabs)
        for window in windows:
            if str(window.get("process_name") or "").lower() != "msedge.exe":
                continue
            if window.get("edge_tabs"):
                window["edge"] = {"session_port": int(edge_debug_port)}
    elif data["schema"] == SCHEMA_V2:
        data = _ensure_v2_layout(data)

    data.pop("browser_tabs", None)
    data.pop("edge_sessions", None)
    data.pop("open_urls", None)

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    tab_count = len(_collect_edge_tabs(data)) if capture_edge_tabs else 0
    print(f"Saved {len(data['windows'])} windows, {tab_count} Edge tabs -> {path}")
    if capture_edge_tabs and tab_count == 0:
        print("Note: no Edge tabs captured. Start Edge with --remote-debugging-port and retry.")


def edge_capture(path: str, edge_debug_port: int = 9222, edge_profile_dir: Optional[str] = None) -> None:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Layout not found: {path}")

    if not _is_debug_endpoint_alive(edge_debug_port):
        print(f"No Edge debug endpoint detected on port {edge_debug_port}.")
        print("Start Edge with --remote-debugging-port and retry.")
        return

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    data = _ensure_v2_layout(data)
    tabs = _fetch_edge_tabs(edge_debug_port)
    if not tabs:
        print("No Edge tabs captured. Start Edge with --remote-debugging-port and retry.")
        return

    windows = data.get("windows", [])
    target_windows = []
    for window in windows:
        edge = window.get("edge") or {}
        try:
            port = int(edge.get("session_port") or 0)
        except (TypeError, ValueError):
            port = 0
        if port == int(edge_debug_port):
            target_windows.append(window)
    if not target_windows:
        target_windows = [w for w in windows if str(w.get("process_name") or "").lower() == "msedge.exe"]

    _assign_edge_tabs_to_windows(windows, tabs, target_windows=target_windows)
    _ensure_window_ids(windows)
    for window in target_windows:
        if window.get("edge_tabs"):
            window["edge"] = {"session_port": int(edge_debug_port)}

    data.pop("browser_tabs", None)
    data.pop("edge_sessions", None)
    data.pop("open_urls", None)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Captured {len(tabs)} Edge tabs into layout windows -> {path}")


def set_edge_open_urls(path: str, urls: List[str], append: bool = False, clear: bool = False) -> None:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Layout not found: {path}")

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    data = _ensure_v2_layout(data)
    edge_windows = [w for w in data.get("windows", []) if str(w.get("process_name") or "").lower() == "msedge.exe"]
    if not edge_windows:
        raise ValueError("No Edge windows found in layout")

    current = _normalize_edge_tabs(edge_windows[0].get("edge_tabs") or [])
    if clear:
        current = []
    if urls:
        new_urls = _coerce_url_list(urls)
        if append:
            current.extend(new_urls)
        else:
            current = new_urls

    edge_windows[0]["edge_tabs"] = _normalize_edge_tabs(current)
    data.pop("browser_tabs", None)
    data.pop("edge_sessions", None)
    data.pop("open_urls", None)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"Saved {len(current)} Edge tabs to first Edge window -> {path}")


def _score_match(candidate: Dict, target: Dict) -> int:
    """
    Higher is better. We try to avoid relying on HWND (not stable).
    """
    score = 0

    # Strong match: exe path if present
    if target.get("exe") and candidate.get("exe") and _normalize_exe_path(target["exe"]) == _normalize_exe_path(candidate["exe"]):
        score += 50

    target_fp = str(target.get("match_fingerprint") or "").strip().lower()
    candidate_fp = str(candidate.get("match_fingerprint") or "").strip().lower()
    if target_fp and candidate_fp and target_fp == candidate_fp:
        score += 90
        if str(target.get("process_name") or "").strip().lower() == "msedge.exe":
            score += 40

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
        results[-1]["match_fingerprint"] = _build_match_fingerprint(results[-1])

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


def _launch_edge_tabs(exe: str, tabs: List[Dict], dry_run: bool = False, base_args: Optional[List[str]] = None) -> int:
    urls = [t.get("url") for t in tabs if str(t.get("url") or "").strip()]
    if not urls:
        return 0

    if not os.path.exists(exe):
        return 0

    launched = 0
    chunk_size = 10
    for idx in range(0, len(urls), chunk_size):
        chunk = urls[idx:idx + chunk_size]
        args = [*base_args] if base_args else []
        args.extend(["--new-window", *chunk])
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


def _launch_edge_tabs_existing(exe: str, tabs: List[Dict], dry_run: bool = False, base_args: Optional[List[str]] = None) -> int:
    urls = [t.get("url") for t in tabs if str(t.get("url") or "").strip()]
    if not urls:
        return 0

    if not os.path.exists(exe):
        return 0

    launched = 0
    chunk_size = 10
    for idx in range(0, len(urls), chunk_size):
        chunk = urls[idx:idx + chunk_size]
        args = [*base_args] if base_args else []
        args.extend(["--new-tab", *chunk])
        if dry_run:
            print(f"[DRY] Launch Edge tabs (existing) -> {exe} {' '.join(args)}")
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


def _is_port_in_use(port: int) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", int(port)), timeout=0.4):
            return True
    except OSError:
        return False


def _is_debug_endpoint_alive(port: int) -> bool:
    url = f"http://127.0.0.1:{int(port)}/json/version"
    try:
        with urllib.request.urlopen(url, timeout=1.0) as response:
            _ = response.read()
        return True
    except Exception:
        return False


def launch_edge_debug(port: int = 9222, profile_dir: Optional[str] = None, dry_run: bool = False) -> bool:
    exe = _find_edge_exe()
    if not exe:
        print("Edge not found. Install Edge or provide a custom path.")
        return False

    if _is_port_in_use(port):
        print(f"Port {port} is already in use. Choose another port or close the existing debug session.")
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
    default_root = os.path.abspath("layouts")
    out_default = os.path.join(default_root, "layout.json")
    out_path = _prompt("Output layout path", out_default)

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


def run_edit_wizard(path: str) -> None:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    schema = str(data.get("schema") or SCHEMA_V1).strip()
    if schema not in (SCHEMA_V1, SCHEMA_V2):
        raise ValueError("Unsupported JSON schema (expected window-layout.v1 or window-layout.v2)")

    if schema == SCHEMA_V1:
        print("Upgrading layout to window-layout.v2 for canonical Edge tab mapping.")
    data = _ensure_v2_layout(data)

    windows = data.get("windows", [])
    edge_windows = [w for w in windows if str(w.get("process_name") or "").lower() == "msedge.exe"]

    if not edge_windows:
        print("No Edge windows found in layout.")
        return

    tabs = _collect_edge_tabs(data)
    if not tabs:
        print("No captured Edge tabs found in layout.")
        return

    print("Edge tab assignment editor")
    print("Select tab indices for each Edge window (comma-separated). Leave blank to keep current mapping.")
    for idx, tab in enumerate(tabs, start=1):
        print(f"  [{idx}] {tab.get('title', '')} -> {tab.get('url', '')}")

    used = set()
    for window in edge_windows:
        title = window.get("title", "(untitled)")
        existing = _normalize_edge_tabs(window.get("edge_tabs") or [])
        current_indices = []
        for tab_idx, tab in enumerate(tabs, start=1):
            if any((tab.get("url") == cur.get("url") and tab.get("title") == cur.get("title")) for cur in existing):
                current_indices.append(str(tab_idx))
        default = ",".join(current_indices)
        selection = _prompt(f"Window: {title} [{window.get('window_id')}]", default)
        if not selection.strip():
            continue

        chosen = []
        for token in selection.split(","):
            token = token.strip()
            if not token.isdigit():
                continue
            tab_idx = int(token)
            if 1 <= tab_idx <= len(tabs):
                chosen.append(tab_idx - 1)

        window["edge_tabs"] = [tabs[i] for i in chosen]
        used.update(chosen)

    unassigned = [tabs[i] for i in range(len(tabs)) if i not in used]
    if unassigned:
        print(f"Unassigned tabs remaining: {len(unassigned)}")

    data.pop("browser_tabs", None)
    data.pop("edge_sessions", None)
    data.pop("open_urls", None)

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Updated Edge tab assignments in {path}")


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


def _is_close_rect(a: Tuple[int, int, int, int], b: Tuple[int, int, int, int], threshold: int) -> bool:
    return (
        abs(a[0] - b[0]) <= threshold and
        abs(a[1] - b[1]) <= threshold and
        abs(a[2] - b[2]) <= threshold and
        abs(a[3] - b[3]) <= threshold
    )


def _edge_size_mismatch(current_rect: Tuple[int, int, int, int], target_rect: Tuple[int, int, int, int], threshold: int = 40) -> bool:
    current_width = current_rect[2] - current_rect[0]
    current_height = current_rect[3] - current_rect[1]
    target_width = target_rect[2] - target_rect[0]
    target_height = target_rect[3] - target_rect[1]
    return abs(current_width - target_width) > threshold or abs(current_height - target_height) > threshold


def _stabilize_edge_window_sizes(applied_edge_matches: List[Tuple[int, Dict]], retries: int = 3, delay_s: float = 0.35) -> int:
    if not applied_edge_matches:
        return 0

    fixed = 0
    for _ in range(max(1, int(retries))):
        resized = False
        time.sleep(max(0.0, float(delay_s)))
        for hwnd, target in applied_edge_matches:
            try:
                current_rect = tuple(win32gui.GetWindowRect(hwnd))
            except Exception:
                continue

            target_rect = tuple(target.get("rect") or target.get("normal_rect") or current_rect)
            if not _edge_size_mismatch(current_rect, target_rect):
                continue
            if _apply_window_position(hwnd, target):
                fixed += 1
                resized = True
        if not resized:
            break
    return fixed


def restore_layout(path: str, mode: str = "basic") -> None:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    schema = str(data.get("schema") or SCHEMA_V1).strip()
    if schema not in (SCHEMA_V1, SCHEMA_V2):
        raise ValueError("Unsupported JSON schema (expected window-layout.v1 or window-layout.v2)")
    if schema == SCHEMA_V2:
        data = _ensure_v2_layout(data)

    normalized_mode = str(mode or "basic").strip().lower()
    if normalized_mode not in ("basic", "smart"):
        raise ValueError("Unsupported restore mode (expected 'basic' or 'smart')")

    restore_edge_tabs = normalized_mode == "smart"
    launch_missing = True
    min_score = 40
    launch_wait = 6.0

    targets = data.get("windows", [])
    current = _current_windows_with_hwnds()

    used_hwnds = set()
    applied = 0
    skipped = 0
    missing: List[Dict] = []
    edge_tabs_launched = 0
    edge_windows_to_restore: List[Dict] = []
    applied_edge_matches: List[Tuple[int, Dict]] = []
    if restore_edge_tabs:
        edge_windows_to_restore = [
            window for window in targets
            if str(window.get("process_name") or "").lower() == "msedge.exe" and _normalize_edge_tabs(window.get("edge_tabs") or [])
        ]

    for t in targets:
        best, _best_score = _best_match(t, current, used_hwnds, min_score)
        if not best:
            missing.append(t)
            continue

        used_hwnds.add(best["hwnd"])
        ok = _apply_window_position(best["hwnd"], t)
        if ok:
            applied += 1
            if str(t.get("process_name") or "").lower() == "msedge.exe":
                applied_edge_matches.append((best["hwnd"], t))
        else:
            skipped += 1

    launched = 0
    if launch_missing and missing:
        for t in missing:
            if restore_edge_tabs and str(t.get("process_name") or "").lower() == "msedge.exe":
                continue
            if _launch_target(t, dry_run=False):
                launched += 1

        if launched:
            time.sleep(max(0.5, float(launch_wait)))
            current = _current_windows_with_hwnds()
            remaining: List[Dict] = []
            for t in missing:
                best, _best_score = _best_match(t, current, used_hwnds, min_score)
                if not best:
                    remaining.append(t)
                    continue

                used_hwnds.add(best["hwnd"])
                ok = _apply_window_position(best["hwnd"], t)
                if ok:
                    applied += 1
                    if str(t.get("process_name") or "").lower() == "msedge.exe":
                        applied_edge_matches.append((best["hwnd"], t))
                else:
                    skipped += 1

            missing = remaining

    skipped += len(missing)

    edge_size_reapplied = _stabilize_edge_window_sizes(applied_edge_matches)

    if restore_edge_tabs and edge_windows_to_restore:
        edge_exe = _edge_exe_from_targets(targets) or _find_edge_exe()
        if edge_exe:
            for window in edge_windows_to_restore:
                tabs = _normalize_edge_tabs(window.get("edge_tabs") or [])
                if not tabs:
                    continue
                base_args: List[str] = []
                edge_meta = window.get("edge") if isinstance(window.get("edge"), dict) else {}
                profile_dir = str(edge_meta.get("profile_dir") or "").strip()
                if profile_dir:
                    base_args.append("--user-data-dir=" + profile_dir)
                edge_tabs_launched += _launch_edge_tabs(
                    edge_exe,
                    tabs,
                    dry_run=False,
                    base_args=base_args,
                )

    print(
        f"Restore complete ({normalized_mode}). Applied={applied}, Skipped={skipped}, "
        f"TotalTargets={len(targets)}, Launched={launched}, EdgeTabs={edge_tabs_launched}, "
        f"EdgeSizeFixes={edge_size_reapplied}"
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
    p_save.add_argument("--edge-profile-dir", help="Edge user data dir for debug session (optional)")
    p_save.add_argument(
        "--schema",
        choices=["v1", "v2"],
        default="v2",
        help="Layout schema version to write (default: v2)",
    )

    p_edge = sub.add_parser("edge-debug", help="Launch Edge with remote debugging enabled")
    p_edge.add_argument("--port", type=int, default=9222, help="Remote debugging port (default: 9222)")
    p_edge.add_argument("--profile-dir", help="User data dir for debug session (default: %TEMP%\\edge-debug)")
    p_edge.add_argument("--dry-run", action="store_true", help="Only show launch command")

    p_edge_capture = sub.add_parser("edge-capture", help="Capture Edge tabs into an existing layout")
    p_edge_capture.add_argument("json_path", help="Layout JSON path")
    p_edge_capture.add_argument("--port", type=int, default=9222, help="Remote debugging port (default: 9222)")
    p_edge_capture.add_argument("--profile-dir", help="Edge user data dir for debug session (optional)")

    p_edge_urls = sub.add_parser("edge-urls", help="Set simple Edge URLs (no debug required)")
    p_edge_urls.add_argument("json_path", help="Layout JSON path")
    p_edge_urls.add_argument("urls", nargs="*", help="URLs to open (space-separated)")
    p_edge_urls.add_argument("--append", action="store_true", help="Append to existing open URLs")
    p_edge_urls.add_argument("--clear", action="store_true", help="Clear existing open URLs")

    p_hotkeys = sub.add_parser("hotkeys", help="Run global hotkey listener (uses config.json)")
    p_hotkeys.add_argument("--config", default=CONFIG_PATH, help="Config path (default: config.json)")

    sub.add_parser("wizard", help="Interactive first-time setup wizard")
    p_edit = sub.add_parser("edit", help="Interactive layout editor")
    p_edit.add_argument("json_path", help="Layout JSON path")

    p_restore = sub.add_parser("restore", help="Restore window positions from JSON")
    p_restore.add_argument("json_path", help="Input JSON path")
    p_restore.add_argument(
        "--mode",
        choices=["basic", "smart"],
        default="basic",
        help="Restore mode: basic (move + launch missing) or smart (basic + Edge tab restore)",
    )

    p_help = sub.add_parser("help", help="Show quick usage")
    p_help.add_argument("--full", action="store_true", help="Show full argparse help")

    args = parser.parse_args()

    if args.cmd == "save":
        save_layout(
            args.json_path,
            capture_edge_tabs=args.edge_tabs,
            edge_debug_port=args.edge_debug_port,
            schema_version=SCHEMA_V1 if args.schema == "v1" else SCHEMA_V2,
            edge_profile_dir=args.edge_profile_dir,
        )
    elif args.cmd == "edge-debug":
        ok = launch_edge_debug(port=args.port, profile_dir=args.profile_dir, dry_run=args.dry_run)
        if ok:
            print("Edge debug session launched.")
        else:
            print("Failed to launch Edge debug session.")
    elif args.cmd == "edge-capture":
        edge_capture(args.json_path, edge_debug_port=args.port, edge_profile_dir=args.profile_dir)
    elif args.cmd == "edge-urls":
        set_edge_open_urls(args.json_path, args.urls, append=args.append, clear=args.clear)
    elif args.cmd == "hotkeys":
        run_hotkey_listener(config_path=args.config)
    elif args.cmd == "wizard":
        run_setup_wizard()
    elif args.cmd == "edit":
        run_edit_wizard(args.json_path)
    elif args.cmd == "restore":
        restore_layout(args.json_path, mode=args.mode)
    elif args.cmd == "help":
        if args.full:
            parser.print_help()
            return
        print("Quick Help")
        print("  save:    python window_layout.py save layout.json")
        print("  restore: python window_layout.py restore layout.json")
        print("  restore basic: python window_layout.py restore layout.json --mode basic")
        print("  edge debug: python window_layout.py edge-debug")
        print("  edge tabs:  python window_layout.py save layout.json --edge-tabs")
        print("  edge capture: python window_layout.py edge-capture layout.json --port 9222")
        print("  edge urls: python window_layout.py edge-urls layout.json https://example.com")
        print("  hotkeys: python window_layout.py hotkeys")
        print("  restore smart: python window_layout.py restore layout.json --mode smart")
        print("  wizard: python window_layout.py wizard")


if __name__ == "__main__":
    main()
