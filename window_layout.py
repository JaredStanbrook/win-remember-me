import argparse
import json
import os
import socket
import subprocess
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

    registered: Dict[int, Dict] = {}
    next_id = 1
    for entry in hotkeys:
        parsed = _parse_hotkey_keys(entry["keys"])
        if not parsed:
            print(f"Skipping invalid hotkey: {entry['keys']}")
            continue
        modifiers, vk = parsed
        try:
            win32api.RegisterHotKey(None, next_id, modifiers, vk)
            registered[next_id] = entry
            print(f"Registered {entry['keys']} -> {entry['action']} {' '.join(entry['args'])}")
            next_id += 1
        except Exception:
            print(f"Failed to register hotkey: {entry['keys']}")

    if not registered:
        print("No hotkeys registered.")
        return

    try:
        while True:
            msg = win32gui.GetMessage(None, 0, 0)
            if msg and msg[1] == win32con.WM_HOTKEY:
                hotkey_id = msg[2]
                entry = registered.get(hotkey_id)
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


def _ensure_open_urls_block(data: Dict) -> None:
    if not isinstance(data.get("open_urls"), dict):
        data["open_urls"] = {}
    open_urls = data["open_urls"]
    if not isinstance(open_urls.get("edge"), list):
        open_urls["edge"] = []


def _migrate_v1_to_v2(data: Dict) -> Dict:
    upgraded = dict(data)
    upgraded["schema"] = SCHEMA_V2
    windows = [dict(w) for w in data.get("windows", [])]
    _ensure_window_ids(windows)

    edge_sessions: List[Dict] = []
    browser_edge = data.get("browser_tabs", {}).get("edge", {})
    if isinstance(browser_edge, dict) and browser_edge.get("tabs"):
        try:
            port = int(browser_edge.get("debug_port") or 0)
        except (TypeError, ValueError):
            port = 0
        tabs = _normalize_edge_tabs(browser_edge.get("tabs") or [])
        window_ids: List[str] = []
        for window in windows:
            if str(window.get("process_name") or "").lower() != "msedge.exe":
                continue
            if window.get("edge_tabs"):
                window_ids.append(str(window.get("window_id")))
                window["edge"] = {"session_port": port} if port else {}
        if tabs:
            edge_sessions.append({
                "port": port,
                "profile_dir": str(browser_edge.get("profile_dir") or ""),
                "captured_at": str(browser_edge.get("captured_at") or data.get("created_at") or _now()),
                "edge_pid": browser_edge.get("edge_pid"),
                "window_ids": window_ids,
                "tabs": tabs,
            })

    upgraded["windows"] = windows
    if edge_sessions:
        upgraded["edge_sessions"] = edge_sessions
    else:
        upgraded.setdefault("edge_sessions", [])

    upgraded.pop("browser_tabs", None)
    _ensure_open_urls_block(upgraded)
    return upgraded


def _ensure_v2_layout(data: Dict) -> Dict:
    if _is_schema_v2(data):
        data.setdefault("edge_sessions", [])
        _ensure_open_urls_block(data)
        _ensure_window_ids(data.get("windows", []))
        return data
    return _migrate_v1_to_v2(data)


def _edge_sessions_from_layout(data: Dict) -> List[Dict]:
    sessions = data.get("edge_sessions") or []
    return [s for s in sessions if isinstance(s, dict)]


def _collect_edge_tabs_by_session(data: Dict) -> List[Dict]:
    sessions = _edge_sessions_from_layout(data)
    if sessions:
        windows = data.get("windows", [])
        by_port: Dict[int, List[Dict]] = {}
        for window in windows:
            edge = window.get("edge") or {}
            try:
                port = int(edge.get("session_port") or 0)
            except (TypeError, ValueError):
                port = 0
            if port <= 0:
                continue
            tabs = window.get("edge_tabs") or []
            if tabs:
                by_port.setdefault(port, []).extend(_normalize_edge_tabs(tabs))
        collected: List[Dict] = []
        for session in sessions:
            try:
                port = int(session.get("port") or 0)
            except (TypeError, ValueError):
                port = 0
            tabs = by_port.get(port) or _normalize_edge_tabs(session.get("tabs") or [])
            if tabs:
                collected.append({
                    "port": port,
                    "profile_dir": str(session.get("profile_dir") or ""),
                    "tabs": tabs,
                })
        return collected

    open_urls = data.get("open_urls") or {}
    edge_urls = _coerce_url_list(open_urls.get("edge") or [])
    if edge_urls:
        return [{
            "port": 0,
            "profile_dir": "",
            "tabs": edge_urls,
        }]
    return []


def _collect_edge_tabs(data: Dict) -> List[Dict]:
    per_window: List[Dict] = []
    for window in data.get("windows", []):
        if str(window.get("process_name") or "").lower() != "msedge.exe":
            continue
        tabs = window.get("edge_tabs") or []
        for tab in tabs:
            url = str(tab.get("url") or "").strip()
            if url:
                per_window.append({"title": str(tab.get("title") or "").strip(), "url": url})
    if per_window:
        return per_window
    return data.get("browser_tabs", {}).get("edge", {}).get("tabs", [])


def _load_existing_metadata(path: str) -> Dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return {}

    preserved: Dict = {}
    for key in ("speed_menu", "custom_layout_folders", "layouts_root", "open_urls", "edge_sessions", "browser_tabs"):
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
        if data["schema"] == SCHEMA_V1:
            data["browser_tabs"] = {
                "edge": {
                    "debug_port": int(edge_debug_port),
                    "captured_at": _now(),
                    "tabs": _normalize_edge_tabs(tabs),
                    "note": "Requires Edge started with --remote-debugging-port"
                }
            }
        else:
            data = _ensure_v2_layout(data)
            window_ids: List[str] = []
            for window in windows:
                if str(window.get("process_name") or "").lower() != "msedge.exe":
                    continue
                if window.get("edge_tabs"):
                    window_ids.append(str(window.get("window_id")))
                    window["edge"] = {"session_port": int(edge_debug_port)}
            session = {
                "port": int(edge_debug_port),
                "profile_dir": str(edge_profile_dir or ""),
                "captured_at": _now(),
                "edge_pid": None,
                "window_ids": window_ids,
                "tabs": _normalize_edge_tabs(tabs),
            }
            data["edge_sessions"] = [session]
            _ensure_open_urls_block(data)
    elif data["schema"] == SCHEMA_V2:
        data = _ensure_v2_layout(data)
    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    tab_count = 0
    if capture_edge_tabs:
        if data.get("schema") == SCHEMA_V2:
            sessions = data.get("edge_sessions") or []
            if sessions:
                tab_count = len(sessions[0].get("tabs") or [])
        else:
            tab_count = len(data.get("browser_tabs", {}).get("edge", {}).get("tabs", []))
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

    window_ids: List[str] = []
    for window in windows:
        if str(window.get("process_name") or "").lower() != "msedge.exe":
            continue
        if window.get("edge_tabs"):
            window_ids.append(str(window.get("window_id")))
            window["edge"] = {"session_port": int(edge_debug_port)}

    sessions = data.get("edge_sessions") or []
    updated = False
    for session in sessions:
        try:
            port = int(session.get("port") or 0)
        except (TypeError, ValueError):
            port = 0
        if port == int(edge_debug_port):
            session["tabs"] = _normalize_edge_tabs(tabs)
            session["profile_dir"] = str(edge_profile_dir or session.get("profile_dir") or "")
            session["captured_at"] = _now()
            session["window_ids"] = window_ids
            updated = True
            break

    if not updated:
        sessions.append({
            "port": int(edge_debug_port),
            "profile_dir": str(edge_profile_dir or ""),
            "captured_at": _now(),
            "edge_pid": None,
            "window_ids": window_ids,
            "tabs": _normalize_edge_tabs(tabs),
        })
    data["edge_sessions"] = sessions
    _ensure_open_urls_block(data)

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Captured {len(tabs)} Edge tabs into session {edge_debug_port} -> {path}")


def set_edge_open_urls(path: str, urls: List[str], append: bool = False, clear: bool = False) -> None:
    if not os.path.exists(path):
        raise FileNotFoundError(f"Layout not found: {path}")

    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    data = _ensure_v2_layout(data)
    _ensure_open_urls_block(data)

    current = _coerce_url_list(data.get("open_urls", {}).get("edge") or [])
    if clear:
        current = []
    if urls:
        new_urls = _coerce_url_list(urls)
        if append:
            current.extend(new_urls)
        else:
            current = new_urls

    data["open_urls"]["edge"] = current

    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    print(f"Saved {len(current)} Edge open URLs -> {path}")


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
        print("Upgrading layout to window-layout.v2 for Edge session mapping.")
        data = _ensure_v2_layout(data)

    windows = data.get("windows", [])
    edge_windows = [w for w in windows if str(w.get("process_name") or "").lower() == "msedge.exe"]
    sessions = _edge_sessions_from_layout(data)

    if not edge_windows:
        print("No Edge windows found in layout.")
        return

    if not sessions:
        print("No Edge sessions found. Capture with edge-capture first.")
        return

    if len(sessions) == 1:
        session = sessions[0]
    else:
        print("Available Edge sessions:")
        for s in sessions:
            print(f"  Port {s.get('port')} | profile: {s.get('profile_dir') or '(default)'}")
        raw = _prompt("Select session port", str(sessions[0].get("port") or ""))
        try:
            port = int(raw)
        except ValueError:
            port = int(sessions[0].get("port") or 0)
        session = next((s for s in sessions if int(s.get("port") or 0) == port), sessions[0])

    tabs = _normalize_edge_tabs(session.get("tabs") or [])
    if not tabs:
        print("No captured Edge tabs found in this session.")
        return

    session_port = int(session.get("port") or 0)
    print("Edge tab assignment editor")
    print("Select tab indices for each Edge window (comma-separated). Leave blank to keep current mapping.")
    for idx, tab in enumerate(tabs, start=1):
        print(f"  [{idx}] {tab.get('title', '')} -> {tab.get('url', '')}")

    used = set()
    for window in edge_windows:
        title = window.get("title", "(untitled)")
        existing = window.get("edge_tabs") or []
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
        if chosen:
            window["edge"] = {"session_port": session_port} if session_port else {}
        used.update(chosen)

    window_ids = []
    for window in edge_windows:
        edge = window.get("edge") or {}
        try:
            port = int(edge.get("session_port") or 0)
        except (TypeError, ValueError):
            port = 0
        if port == session_port and window.get("edge_tabs"):
            window_ids.append(str(window.get("window_id")))

    session["window_ids"] = window_ids

    unassigned = [tabs[i] for i in range(len(tabs)) if i not in used]
    if unassigned:
        print(f"Unassigned tabs remaining: {len(unassigned)}")

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


def restore_layout(
    path: str,
    min_score: int = 40,
    dry_run: bool = False,
    launch_missing: bool = False,
    launch_wait: float = 6.0,
    restore_edge_tabs: bool = False,
    smart_restore: bool = False,
    smart_threshold: int = 20,
) -> None:
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    schema = str(data.get("schema") or SCHEMA_V1).strip()
    if schema not in (SCHEMA_V1, SCHEMA_V2):
        raise ValueError("Unsupported JSON schema (expected window-layout.v1 or window-layout.v2)")
    if schema == SCHEMA_V2:
        data = _ensure_v2_layout(data)

    targets = data.get("windows", [])
    current = _current_windows_with_hwnds()

    used_hwnds = set()
    applied = 0
    skipped = 0
    missing: List[Dict] = []
    edge_tabs_launched = 0
    edge_tabs_present = False
    edge_sessions_to_restore: List[Dict] = []
    if restore_edge_tabs:
        if schema == SCHEMA_V2:
            edge_sessions_to_restore = _collect_edge_tabs_by_session(data)
            edge_tabs_present = any(s.get("tabs") for s in edge_sessions_to_restore)
        else:
            edge_tabs_present = bool(data.get("browser_tabs", {}).get("edge", {}).get("tabs", []))
    edge_existing_in_place = False
    edge_any_running = False

    smart_threshold = max(0, int(smart_threshold))
    if smart_restore:
        # Use a fresh snapshot for position checks.
        current = _current_windows_with_hwnds()
        edge_any_running = any(
            str(c.get("process_name") or "").lower() == "msedge.exe"
            for c in current
        )
        if restore_edge_tabs and edge_tabs_present:
            for t in targets:
                if str(t.get("process_name") or "").lower() != "msedge.exe":
                    continue
                t_rect = tuple(t.get("rect") or (0, 0, 0, 0))
                for c in current:
                    if str(c.get("process_name") or "").lower() != "msedge.exe":
                        continue
                    c_rect = tuple(c.get("rect") or (0, 0, 0, 0))
                    if _is_close_rect(c_rect, t_rect, smart_threshold):
                        edge_existing_in_place = True
                        break
                if edge_existing_in_place:
                    break

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

        if smart_restore:
            current_rect = tuple(best.get("rect") or (0, 0, 0, 0))
            target_rect = tuple(t.get("rect") or (0, 0, 0, 0))
            if _is_close_rect(current_rect, target_rect, smart_threshold):
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

    skipped += len(missing)

    if restore_edge_tabs and edge_tabs_present:
        if not (smart_restore and edge_any_running and not edge_existing_in_place):
            edge_exe = _edge_exe_from_targets(targets) or _find_edge_exe()
            if edge_exe:
                use_existing = smart_restore and edge_existing_in_place
                if schema == SCHEMA_V2:
                    for session in edge_sessions_to_restore:
                        tabs = session.get("tabs") or []
                        if not tabs:
                            continue
                        base_args: List[str] = []
                        profile_dir = str(session.get("profile_dir") or "").strip()
                        if profile_dir:
                            base_args.append("--user-data-dir=" + profile_dir)
                        if use_existing:
                            edge_tabs_launched += _launch_edge_tabs_existing(
                                edge_exe,
                                tabs,
                                dry_run=dry_run,
                                base_args=base_args,
                            )
                        else:
                            edge_tabs_launched += _launch_edge_tabs(
                                edge_exe,
                                tabs,
                                dry_run=dry_run,
                                base_args=base_args,
                            )
                else:
                    edge_tabs = _collect_edge_tabs(data)
                    if edge_tabs:
                        if use_existing:
                            edge_tabs_launched = _launch_edge_tabs_existing(edge_exe, edge_tabs, dry_run=dry_run)
                        else:
                            edge_tabs_launched = _launch_edge_tabs(edge_exe, edge_tabs, dry_run=dry_run)
        else:
            print("Smart restore: Edge is running but not in place; skipping tab restore.")

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
    p_restore.add_argument("--min-score", type=int, default=40, help="Matching threshold (default: 40)")
    p_restore.add_argument("--dry-run", action="store_true", help="Only show matches, do not move windows")
    p_restore.add_argument("--launch-missing", action="store_true", help="Launch apps for missing windows before restore")
    p_restore.add_argument("--launch-wait", type=float, default=6.0, help="Seconds to wait after launch (default: 6)")
    p_restore.add_argument("--restore-edge-tabs", action="store_true", help="Reopen Edge tabs captured during save")
    p_restore.add_argument("--smart", action="store_true", help="Only move windows that are not already in place")
    p_restore.add_argument("--smart-threshold", type=int, default=20, help="Pixel threshold for smart restore (default: 20)")

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
        restore_layout(
            args.json_path,
            min_score=args.min_score,
            dry_run=args.dry_run,
            launch_missing=args.launch_missing,
            launch_wait=args.launch_wait,
            restore_edge_tabs=args.restore_edge_tabs,
            smart_restore=args.smart,
            smart_threshold=args.smart_threshold,
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
        print("  edge capture: python window_layout.py edge-capture layout.json --port 9222")
        print("  edge urls: python window_layout.py edge-urls layout.json https://example.com")
        print("  hotkeys: python window_layout.py hotkeys")
        print("  restore tabs: python window_layout.py restore layout.json --restore-edge-tabs")
        print("  wizard: python window_layout.py wizard")


if __name__ == "__main__":
    main()
