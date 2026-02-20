"""
window_layout.py  –  Save & restore Windows desktop layouts
============================================================

Schema: "window-layout"  (single version, no back-compat cruft)

Key behaviours
  · Capture filters UWP hosts (ApplicationFrameWindow, TextInputHost,
    RtkUWP, SystemSettings, etc.) — they can't be moved or relaunched.
  · Minimised windows ARE captured (size filter bypassed for iconic windows).
  · Z-order captured via EnumWindows (which yields windows in z-order) at
    after placement by calling SetWindowPos bottom-up so the front-most
    window ends up genuinely on top.
  · SetWindowPlacement for atomic position + state restore (max/min/normal).
  · Matching pre-filtered by exe; geometry (+30 pts ≤40px) is the primary
    tiebreaker for identical-title windows (e.g. two Explorer windows,
    same folder).
  · Edge profile (--user-data-dir, --profile-directory) auto-detected from
    the running process cmdline at capture; stored per window.
  · Edge tab restore targets the exact Edge profile session — works with
    multiple concurrent Edge sessions (normal + debug, work + personal).
  · CDP windowId -1/0 treated as "no valid ID"; falls back to token-overlap
    title matching then round-robin for tab->window assignment.
"""

import argparse
import json
import os
import re
import socket
import subprocess
import sys
import time
import urllib.request
import uuid
from typing import Dict, Iterable, List, Optional, Set, Tuple

import psutil
import win32api
import win32con
import win32gui
import win32process
import win32com.client

# ══════════════════════════════════════════════════════════════════════════
#  Constants
# ══════════════════════════════════════════════════════════════════════════
SCHEMA      = "window-layout"
CONFIG_PATH = "config.json"

_EDGE_EXE_CANDIDATES = [
    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
]

# Processes that appear as visible top-level windows but cannot be
# meaningfully captured, repositioned, or relaunched.
_BLOCKED_PROC: Set[str] = {
    "textinputhost.exe",          # Windows Input Experience
    "applicationframehost.exe",   # UWP shell host
    "shellhost.exe",
    "startmenuexperiencehost.exe",
    "searchhost.exe",
    "searchapp.exe",
    "lockapp.exe",
    "systemsettings.exe",         # Settings UWP
    "dwm.exe",
    "fontdrvhost.exe",
    "rtkuwp.exe",                 # Realtek Audio Console UWP
}

# Window classes that are always noise.
_BLOCKED_CLASS: Set[str] = {
    "windows.ui.core.corewindow",  # UWP content host
    "applicationframewindow",      # UWP shell chrome
    "progman",                     # desktop
    "workerw",                     # desktop icon layer
}


# ══════════════════════════════════════════════════════════════════════════
#  Tiny helpers
# ══════════════════════════════════════════════════════════════════════════
def _now() -> str:
    return time.strftime("%Y-%m-%d %H:%M:%S")

def _uuid() -> str:
    return str(uuid.uuid4())

def _ensure_ids(windows: List[Dict]) -> None:
    for w in windows:
        if not str(w.get("window_id") or "").strip():
            w["window_id"] = _uuid()

def _safe_text(hwnd: int) -> str:
    try:    return win32gui.GetWindowText(hwnd) or ""
    except: return ""

def _safe_class(hwnd: int) -> str:
    try:    return win32gui.GetClassName(hwnd) or ""
    except: return ""

def _get_pid(hwnd: int) -> int:
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        return int(pid or 0)
    except: return 0

def _proc_info(pid: int) -> Tuple[str, str]:
    """Returns (process_name, exe_path)."""
    if not pid: return "", ""
    try:
        p = psutil.Process(pid)
        return (p.name() or ""), (p.exe() or "")
    except: return "", ""

def _proc_cmdline(pid: int) -> List[str]:
    try:    return psutil.Process(pid).cmdline()
    except: return []

def _window_rect(hwnd: int) -> Tuple[int, int, int, int]:
    try:    return tuple(win32gui.GetWindowRect(hwnd))
    except: return (0, 0, 0, 0)

def _window_placement(hwnd: int) -> Tuple[int, Tuple[int, int, int, int]]:
    """Returns (showCmd, normalPositionRect).
    normalPositionRect is the RESTORED size/position regardless of
    whether the window is currently minimised or maximised."""
    try:
        pl = win32gui.GetWindowPlacement(hwnd)
        return int(pl[1]), tuple(pl[4])
    except:
        return win32con.SW_SHOWNORMAL, (0, 0, 0, 0)


def _is_snapped(show_cmd: int, live_rect: tuple, normal_rect: tuple,
                threshold: int = 10) -> bool:
    """
    Return True if a window is in a snapped (Aero Snap) position.

    A snapped window has show_cmd == SW_SHOWNORMAL but its live rect
    differs from GetWindowPlacement's normalPosition — because snap moves
    the window without updating normalPosition.  We detect this by
    comparing the two rects.  Minimised and maximised windows are not
    considered snapped (they have their own show_cmd values).
    """
    if show_cmd != win32con.SW_SHOWNORMAL:
        return False
    if len(live_rect) < 4 or len(normal_rect) < 4:
        return False
    return not all(abs(live_rect[i] - normal_rect[i]) <= threshold
                   for i in range(4))

def _capture_z_order() -> Dict[int, int]:
    """
    Build {hwnd: z_index} for all top-level windows, where 0 = topmost.

    EnumWindows is documented to enumerate top-level windows in Z-order
    from front to back, so the first HWND it yields is the frontmost
    visible window.  This is simpler and more correct than GetTopWindow /
    GetNextWindow, which traverse the internal window manager list that
    includes thousands of invisible system windows whose handles do NOT
    match what EnumWindows returns for the same visible windows.
    """
    z: Dict[int, int] = {}
    idx = 0

    def _cb(hwnd, _):
        nonlocal idx
        z[hwnd] = idx
        idx += 1

    try:
        win32gui.EnumWindows(_cb, None)
    except Exception:
        pass
    return z


def _restore_z_order(hwnd_z: List[Tuple[int, int]], verbose: bool = False) -> int:
    """
    Apply z-order to a list of (hwnd, z_index) pairs.

    Strategy: sort by z_index descending (background-most first) and call
    SetWindowPos with HWND_TOP for each.  Because each call brings a window
    to the top, processing back-to-front ends with the z_index=0 window
    genuinely on top.

    SWP flags: NOMOVE | NOSIZE | NOACTIVATE so we only touch z-order.
    Returns the number of windows successfully re-ordered.
    """
    if not hwnd_z:
        return 0
    SWP_FLAGS = (win32con.SWP_NOMOVE | win32con.SWP_NOSIZE |
                 win32con.SWP_NOACTIVATE)
    ordered = sorted(hwnd_z, key=lambda x: x[1], reverse=True)
    success = 0
    for hwnd, z_idx in ordered:
        try:
            win32gui.SetWindowPos(
                hwnd, win32con.HWND_TOP, 0, 0, 0, 0, SWP_FLAGS
            )
            if verbose:
                print(f"  Z-ORDER  hwnd={hex(hwnd)}  z_index={z_idx}")
            success += 1
        except Exception:
            pass
    return success


def _is_close_rect(a, b, thr: int) -> bool:
    return all(abs(a[i] - b[i]) <= thr for i in range(4))

def _rects_intersect(a, b) -> bool:
    return not (a[2] <= b[0] or a[0] >= b[2] or a[3] <= b[1] or a[1] >= b[3])

def _find_edge_exe() -> Optional[str]:
    for p in _EDGE_EXE_CANDIDATES:
        if os.path.exists(p):
            return p
    return None

def _default_edge_udd() -> str:
    return os.path.join(
        os.environ.get("LOCALAPPDATA", ""),
        "Microsoft", "Edge", "User Data",
    )


# ══════════════════════════════════════════════════════════════════════════
#  Edge profile detection (from running process cmdline)
# ══════════════════════════════════════════════════════════════════════════
def _edge_profile_from_pid(pid: int) -> Dict[str, str]:
    """
    Parse --user-data-dir and --profile-directory from a running Edge PID.
    If the process is a standard install (no --user-data-dir flag), the
    default User Data path and 'Default' profile are used.
    """
    result = {"user_data_dir": "", "profile_directory": ""}
    for arg in _proc_cmdline(pid):
        if arg.startswith("--user-data-dir="):
            result["user_data_dir"] = arg.split("=", 1)[1]
        elif arg.startswith("--profile-directory="):
            result["profile_directory"] = arg.split("=", 1)[1]
    # If still no profile_directory, try to derive it from the udd path tail
    if not result["profile_directory"] and result["user_data_dir"]:
        tail = os.path.basename(result["user_data_dir"].rstrip("\\/"))
        if re.match(r"^(Default|Profile \d+)$", tail, re.IGNORECASE):
            result["profile_directory"] = tail
    # Ensure sensible defaults for normal (non-debug) Edge
    if not result["user_data_dir"]:
        result["user_data_dir"] = _default_edge_udd()
    if not result["profile_directory"]:
        result["profile_directory"] = "Default"
    return result


# ══════════════════════════════════════════════════════════════════════════
#  Window filter
# ══════════════════════════════════════════════════════════════════════════
def _is_interesting(hwnd: int) -> bool:
    """True for top-level user-facing windows we can meaningfully save."""
    if not win32gui.IsWindow(hwnd):         return False
    if win32gui.GetParent(hwnd):            return False
    if not win32gui.IsWindowVisible(hwnd):  return False

    title = _safe_text(hwnd).strip()
    if not title:                           return False

    cls = _safe_class(hwnd).strip().lower()
    if cls in _BLOCKED_CLASS:              return False

    try:
        ex_style = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
        owner    = win32gui.GetWindow(hwnd, win32con.GW_OWNER)
    except Exception:
        ex_style, owner = 0, 0

    if (ex_style & win32con.WS_EX_TOOLWINDOW) and not (ex_style & win32con.WS_EX_APPWINDOW):
        return False
    if owner and not (ex_style & win32con.WS_EX_APPWINDOW):
        return False

    pid = _get_pid(hwnd)
    proc, _ = _proc_info(pid)
    if proc.lower() in _BLOCKED_PROC:
        return False

    # Allow minimised windows through even though their live rect is ~0×0
    if not win32gui.IsIconic(hwnd):
        r = _window_rect(hwnd)
        if (r[2] - r[0]) < 120 or (r[3] - r[1]) < 80:
            return False

    return True


# ══════════════════════════════════════════════════════════════════════════
#  Explorer folder paths
# ══════════════════════════════════════════════════════════════════════════
def _explorer_folder_paths() -> Dict[int, str]:
    paths: Dict[int, str] = {}
    try:
        shell = win32com.client.Dispatch("Shell.Application")
        for win in shell.Windows():
            try:
                hwnd = int(getattr(win, "HWND", 0) or 0)
                loc  = str(getattr(win, "LocationURL", "") or "").strip()
                if loc.startswith("file:///"):
                    path = loc.replace("file:///", "").replace("/", "\\")
                    if hwnd and path:
                        paths[hwnd] = path
            except Exception:
                continue
    except Exception:
        pass
    return paths


# ══════════════════════════════════════════════════════════════════════════
#  Capture
# ══════════════════════════════════════════════════════════════════════════
def capture_windows(verbose: bool = False) -> List[Dict]:
    exp_paths = _explorer_folder_paths()

    # Snapshot z-order BEFORE enumeration so every window gets a stable index.
    # Windows not found in the z-order map (shouldn't happen for top-level
    # windows, but guard anyway) get a large sentinel value so they sort last.
    z_order_map = _capture_z_order()
    Z_SENTINEL  = 9999

    entries: List[Tuple[int, Dict]] = []   # (hwnd, data)

    def _cb(hwnd, _):
        if not _is_interesting(hwnd):
            return
        title = _safe_text(hwnd).strip()
        cls   = _safe_class(hwnd).strip()
        pid   = _get_pid(hwnd)
        proc, exe = _proc_info(pid)
        rect  = _window_rect(hwnd)
        show_cmd, normal_rect = _window_placement(hwnd)
        is_min    = show_cmd == win32con.SW_SHOWMINIMIZED
        is_max    = show_cmd == win32con.SW_SHOWMAXIMIZED
        is_snapped = _is_snapped(show_cmd, rect, normal_rect)
        z_idx     = z_order_map.get(hwnd, Z_SENTINEL)

        # For snapped windows, rect is the actual snapped position.
        # normal_rect is the pre-snap position Windows remembers internally —
        # not useful for restore since we want the snapped layout back.
        # We store both so the JSON is informative, but restore_rect is what
        # _apply_placement will use.
        restore_rect = list(rect) if is_snapped else list(normal_rect)

        data: Dict = {
            "window_id":    _uuid(),
            "title":        title,
            "class_name":   cls,
            "pid":          pid,
            "process_name": proc,
            "exe":          exe,
            "is_visible":   True,
            "is_minimized": is_min,
            "is_maximized": is_max,
            "is_snapped":   is_snapped,
            "rect":         list(rect),
            "normal_rect":  list(normal_rect),
            "restore_rect": restore_rect,
            "show_cmd":     show_cmd,
            "z_order":      z_idx,   # 0 = topmost/frontmost
        }

        # Launch spec
        launch_args: List[str] = []
        if proc.lower() == "explorer.exe":
            fp = exp_paths.get(hwnd)
            if fp:
                launch_args = [fp]
        if exe:
            data["launch"] = {"exe": exe, "args": launch_args, "cwd": ""}

        # Edge profile meta — captured from cmdline, no debug port needed
        if proc.lower() == "msedge.exe":
            profile = _edge_profile_from_pid(pid)
            data["edge"] = {
                "user_data_dir":     profile["user_data_dir"],
                "profile_directory": profile["profile_directory"],
                "cdp_window_id":     0,   # filled in if --edge-tabs used
                "debug_port":        0,
            }
            data["edge_tabs"] = []

        if verbose:
            state = "MIN" if is_min else ("MAX" if is_max else ("SNP" if is_snapped else "NRM"))
            print(f"  CAPTURE [z={z_idx:03d}][{state}] {proc}  \"{title[:60]}\"  "
                  f"rect={normal_rect}")

        entries.append((hwnd, data))

    win32gui.EnumWindows(_cb, None)

    # Sort by z_order so the saved list is front-to-back (index 0 = topmost).
    # This makes the JSON human-readable and restore ordering natural.
    entries.sort(key=lambda e: e[1]["z_order"])
    return [d for _, d in entries]


# ══════════════════════════════════════════════════════════════════════════
#  CDP tab fetching
# ══════════════════════════════════════════════════════════════════════════
def _port_open(port: int) -> bool:
    try:
        with socket.create_connection(("127.0.0.1", port), timeout=0.5):
            return True
    except OSError:
        return False

def _cdp_alive(port: int) -> bool:
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/json/version", timeout=1.0
        ) as r:
            r.read()
        return True
    except Exception:
        return False

def _fetch_cdp_tabs(port: int) -> List[Dict]:
    """Fetch page tabs from an Edge CDP debug endpoint."""
    try:
        with urllib.request.urlopen(
            f"http://127.0.0.1:{port}/json/list", timeout=2.0
        ) as r:
            items = json.loads(r.read().decode("utf-8", errors="replace"))
    except Exception:
        return []
    tabs = []
    for item in items:
        if item.get("type") != "page":
            continue
        url = str(item.get("url") or "").strip()
        if not url or url.startswith(("edge://", "chrome://")):
            continue
        # CDP windowId is only useful when it's a real positive integer.
        # Edge sometimes returns -1 or 0 meaning "unknown window" — treat
        # those as None so we fall back to title-based matching.
        raw_wid = item.get("windowId")
        wid = raw_wid if (isinstance(raw_wid, int) and raw_wid > 0) else None
        tabs.append({
            "title":     str(item.get("title") or "").strip(),
            "url":       url,
            "window_id": wid,
        })
    return tabs

def _cdp_open_tab_in_window(port: int, cdp_window_id: int, url: str) -> bool:
    """
    Use CDP Target.createTarget to open a URL in a specific Edge window.
    Only works when a debug session is live on the given port.
    Returns True on success.
    """
    try:
        import json as _json
        payload = _json.dumps({
            "id": 1, "method": "Target.createTarget",
            "params": {"url": url, "windowId": cdp_window_id},
        }).encode()
        req = urllib.request.Request(
            f"http://127.0.0.1:{port}/json/new",
            data=payload,
            headers={"Content-Type": "application/json"},
            method="PUT",
        )
        with urllib.request.urlopen(req, timeout=2.0) as r:
            r.read()
        return True
    except Exception:
        return False


def _normalize_tabs(tabs: Iterable[Dict]) -> List[Dict]:
    out = []
    for t in tabs:
        url = str(t.get("url") or "").strip()
        if url:
            out.append({"title": str(t.get("title") or "").strip(), "url": url})
    return out

def _strip_edge_suffix(title: str) -> str:
    """Remove ' - Microsoft Edge' / ' - Work - Microsoft Edge' suffixes."""
    s   = (title or "").replace("Microsoft\u200b Edge", "Microsoft Edge").strip()
    low = s.lower()
    for m in (" - work - microsoft edge", " - personal - microsoft edge",
              " - microsoft edge"):
        if m in low:
            s   = s[:low.rfind(m)]
            low = s.lower()
    return s.strip().lower()


# ══════════════════════════════════════════════════════════════════════════
#  Assign CDP tabs to saved Edge windows
# ══════════════════════════════════════════════════════════════════════════
def _assign_tabs(edge_windows: List[Dict], tabs: List[Dict]) -> None:
    """
    Distribute CDP-fetched tabs across saved Edge window entries.

    Priority order:
      1. CDP windowId (positive int) matches a stored cdp_window_id hint
      2. The tab's active-tab title token-overlaps with a window title
      3. Round-robin across unassigned windows

    When Edge doesn't return valid windowIds (returns -1 or None for all tabs
    — common in some debug-session configs) every tab lands in no_wid and we
    rely on title matching then round-robin.  With a single Edge window this
    is trivial.  With multiple windows the caller should use `edit` to
    manually reassign if automatic title matching is wrong.
    """
    if not edge_windows or not tabs:
        return
    for w in edge_windows:
        w["edge_tabs"] = []

    by_wid: Dict[int, List[Dict]] = {}
    no_wid: List[Dict] = []
    for tab in tabs:
        wid = tab.get("window_id")
        if isinstance(wid, int) and wid > 0:
            by_wid.setdefault(wid, []).append(tab)
        else:
            no_wid.append(tab)

    # Build lookup structures
    hint_map:  Dict[int, Dict]       = {}
    title_map: Dict[str, List[Dict]] = {}
    for w in edge_windows:
        hint = int((w.get("edge") or {}).get("cdp_window_id", 0) or 0)
        if hint > 0:
            hint_map[hint] = w
        nt = _strip_edge_suffix(str(w.get("title") or ""))
        if nt:
            title_map.setdefault(nt, []).append(w)

    assigned: Set[int] = set()

    def _best_title_match(tab_title: str) -> Optional[Dict]:
        """Find the unassigned window whose title best overlaps with tab_title."""
        tt = _strip_edge_suffix(tab_title)
        tt_toks = set(tt.split()) if tt else set()
        best_w, best_score = None, 0
        for w in edge_windows:
            if id(w) in assigned:
                continue
            wt = _strip_edge_suffix(str(w.get("title") or ""))
            wt_toks = set(wt.split()) if wt else set()
            if not tt_toks or not wt_toks:
                continue
            overlap = len(tt_toks & wt_toks) / max(len(tt_toks | wt_toks), 1)
            if overlap > best_score:
                best_score, best_w = overlap, w
        # Only use title match if it's a reasonably strong overlap
        return best_w if best_score >= 0.3 else None

    def _pick(candidate_tabs: List[Dict], cdp_wid: Optional[int]) -> Optional[Dict]:
        # 1. cdp_window_id hint
        if cdp_wid and cdp_wid in hint_map:
            w = hint_map[cdp_wid]
            if id(w) not in assigned:
                return w
        # 2. Title overlap — use the first tab in the group (usually active tab)
        for tab in candidate_tabs:
            w = _best_title_match(str(tab.get("title") or ""))
            if w is not None:
                return w
        # 3. First unassigned window
        for w in edge_windows:
            if id(w) not in assigned:
                return w
        return None

    # Process tabs with known window IDs first
    for cdp_wid, group in sorted(by_wid.items()):
        w = _pick(group, cdp_wid)
        if w is None:
            continue
        assigned.add(id(w))
        if isinstance(w.get("edge"), dict):
            w["edge"]["cdp_window_id"] = cdp_wid
        w["edge_tabs"].extend(
            {"title": t.get("title", ""), "url": t.get("url", "")}
            for t in group
        )

    # Process tabs with no/invalid window ID
    # Try title matching first, then fall back to round-robin
    rr = 0
    for tab in no_wid:
        w = _best_title_match(str(tab.get("title") or ""))
        if w is None:
            w = edge_windows[rr % len(edge_windows)]
            rr += 1
        w["edge_tabs"].append({
            "title": tab.get("title", ""),
            "url":   tab.get("url", ""),
        })


# ══════════════════════════════════════════════════════════════════════════
#  Save
# ══════════════════════════════════════════════════════════════════════════
def _load_preserved(path: str) -> Dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
    except Exception:
        return {}
    return {k: d[k] for k in ("speed_menu", "custom_layout_folders", "layouts_root")
            if k in d}


def save_layout(
    path: str,
    capture_edge_tabs: bool = False,
    edge_debug_port: int = 9222,
    verbose: bool = False,
) -> None:
    windows  = capture_windows(verbose=verbose)
    _ensure_ids(windows)
    preserved = _load_preserved(path)

    data: Dict = {"schema": SCHEMA, "created_at": _now(), "windows": windows}
    data.update(preserved)

    if capture_edge_tabs:
        if _cdp_alive(edge_debug_port):
            tabs = _fetch_cdp_tabs(edge_debug_port)
            edge_wins = [w for w in windows
                         if str(w.get("process_name") or "").lower() == "msedge.exe"]
            _assign_tabs(edge_wins, tabs)
            for w in edge_wins:
                if w.get("edge_tabs") and isinstance(w.get("edge"), dict):
                    w["edge"]["debug_port"] = edge_debug_port
        else:
            print(f"  Warning: no Edge debug endpoint on port {edge_debug_port}. "
                  f"Run 'edge-debug --port {edge_debug_port}' first.")

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)

    tab_count = sum(len(w.get("edge_tabs") or []) for w in windows)
    print(f"Saved {len(windows)} windows, {tab_count} Edge tabs -> {path}")


# ══════════════════════════════════════════════════════════════════════════
#  Scoring / matching
# ══════════════════════════════════════════════════════════════════════════
def _score(candidate: Dict, target: Dict) -> Tuple[int, Dict]:
    """
    Score a live candidate against a saved target.  Max = 165 pts.

    exe path        +50  (always true when pre-filtered by exe)
    process name    +25
    class name      +15
    title exact     +40  (0 for Edge – active-tab title changes constantly)
    title partial   +15
    geometry ≤ 40px +30  ← PRIMARY tiebreaker for identical-title windows
    geometry ≤120px +15
    """
    comps: Dict = {
        "exe": 0, "process": 0, "class": 0, "title": 0, "geometry": 0,
        "edge_title_deweighted": 0,
    }

    t_exe  = str(target.get("exe") or "").lower()
    c_exe  = str(candidate.get("exe") or "").lower()
    if t_exe and c_exe and t_exe == c_exe:
        comps["exe"] = 50

    t_proc = str(target.get("process_name") or "").lower()
    c_proc = str(candidate.get("process_name") or "").lower()
    if t_proc and c_proc and t_proc == c_proc:
        comps["process"] = 25

    t_cls = str(target.get("class_name") or "").lower()
    c_cls = str(candidate.get("class_name") or "").lower()
    if t_cls and c_cls and t_cls == c_cls:
        comps["class"] = 15

    t_title = str(target.get("title") or "").lower()
    c_title = str(candidate.get("title") or "").lower()
    if t_proc == "msedge.exe":
        comps["edge_title_deweighted"] = 1
        t_toks = set(t_title.split())
        c_toks = set(c_title.split())
        overlap = len(t_toks & c_toks) / max(len(t_toks | c_toks), 1)
        if overlap >= 0.4:
            comps["title"] = 8   # nudge only
    else:
        if t_title and c_title:
            if t_title == c_title:
                comps["title"] = 40
            elif t_title in c_title or c_title in t_title:
                comps["title"] = 15

    t_rect = tuple(target.get("normal_rect") or target.get("rect") or ())
    c_rect = tuple(candidate.get("normal_rect") or candidate.get("rect") or ())
    if len(t_rect) == 4 and len(c_rect) == 4:
        if _is_close_rect(c_rect, t_rect, 40):
            comps["geometry"] = 30
        elif _is_close_rect(c_rect, t_rect, 120):
            comps["geometry"] = 15

    comps["total"] = sum(v for k, v in comps.items()
                         if k not in ("total", "edge_title_deweighted"))
    return comps["total"], comps


def _rank_candidates(target: Dict, current: List[Dict], used: Set) -> List[Dict]:
    """
    Pre-filter to same exe, then score and rank.
    Two windows with the same title (e.g. two Explorer windows in the same
    folder) will be separated purely by geometry — the one closer in position
    gets the higher score.
    """
    t_exe  = str(target.get("exe") or "").lower()
    t_proc = str(target.get("process_name") or "").lower()
    ranked = []
    for c in current:
        if c.get("hwnd") in used:
            continue
        c_exe  = str(c.get("exe") or "").lower()
        c_proc = str(c.get("process_name") or "").lower()
        if t_exe and c_exe:
            if t_exe != c_exe:   continue
        elif t_proc and c_proc:
            if t_proc != c_proc: continue
        score, comps = _score(c, target)
        ranked.append({"candidate": c, "score": score, "components": comps})
    ranked.sort(key=lambda x: x["score"], reverse=True)
    return ranked


def _print_diag(target: Dict, ranked: List[Dict], top_n: int = 3) -> None:
    print(f"[DIAG] Target: proc={target.get('process_name')}  "
          f"title={target.get('title')}")
    if not ranked:
        print("[DIAG]   No candidates")
        return
    for i, item in enumerate(ranked[:max(1, top_n)], 1):
        c  = item["candidate"]
        co = item["components"]
        note = " edge_deweighted" if co.get("edge_title_deweighted") else ""
        print(
            f"[DIAG]   #{i} hwnd={hex(c.get('hwnd', 0))} score={item['score']} "
            f"(exe={co.get('exe',0)} proc={co.get('process',0)} "
            f"cls={co.get('class',0)} title={co.get('title',0)} "
            f"geo={co.get('geometry',0)}){note}"
        )


# ══════════════════════════════════════════════════════════════════════════
#  Live window snapshot (for restore matching)
# ══════════════════════════════════════════════════════════════════════════
def _live_windows() -> List[Dict]:
    """Enumerate current top-level windows for restore matching."""
    z_order_map = _capture_z_order()
    results = []

    def _cb(hwnd, _):
        if not win32gui.IsWindow(hwnd):        return
        if win32gui.GetParent(hwnd):           return
        if not win32gui.IsWindowVisible(hwnd): return
        title = _safe_text(hwnd).strip()
        if not title:                          return
        cls = _safe_class(hwnd).strip()
        if cls.lower() in _BLOCKED_CLASS:      return
        pid = _get_pid(hwnd)
        proc, exe = _proc_info(pid)
        if proc.lower() in _BLOCKED_PROC:      return
        show_cmd, normal_rect = _window_placement(hwnd)
        results.append({
            "hwnd":         hwnd,
            "title":        title,
            "class_name":   cls,
            "pid":          pid,
            "process_name": proc,
            "exe":          exe,
            "show_cmd":     show_cmd,
            "normal_rect":  normal_rect,
            "rect":         _window_rect(hwnd),
            "z_order":      z_order_map.get(hwnd, 9999),
        })

    win32gui.EnumWindows(_cb, None)
    return results


# ══════════════════════════════════════════════════════════════════════════
#  Apply saved position
# ══════════════════════════════════════════════════════════════════════════
def _clamp(left: int, top: int, w: int, h: int) -> Tuple[int, int, int, int]:
    try:
        bounds = [
            win32api.GetMonitorInfo(m[0]).get("Monitor")
            for m in win32api.EnumDisplayMonitors()
        ]
        bounds = [b for b in bounds if b]
        if bounds and not any(_rects_intersect((left, top, left+w, top+h), b)
                              for b in bounds):
            prim = bounds[0]
            left = min(max(left, prim[0]), prim[2] - w)
            top  = min(max(top,  prim[1]), prim[3] - h)
    except Exception:
        pass
    return left, top, w, h


def _apply_placement(hwnd: int, entry: Dict, verbose: bool = False) -> bool:
    """
    Apply saved position + show state to a window.

    Snapped windows (is_snapped=True) are handled differently from normal,
    minimised, and maximised windows:

    Normal / min / max:
        SetWindowPlacement atomically sets the normalPosition AND the show
        state in one call.  This is correct for these states because Windows
        uses normalPosition internally to know where to restore to.

    Snapped:
        The window is SW_SHOWNORMAL but its position was set by Aero Snap,
        which bypasses normalPosition entirely.  SetWindowPlacement would
        write restore_rect into normalPosition and then restore from there —
        which just puts it back to the snapped rect visually, but corrupts
        normalPosition so double-clicking the title bar teleports the window
        somewhere unexpected.  Instead we:
          1. Restore the window to normal show state (in case it drifted).
          2. MoveWindow to the exact snapped rect — this is what Snap does.
          3. Leave normalPosition alone (the user can un-snap normally).
    """
    desired    = int(entry.get("show_cmd") or win32con.SW_SHOWNORMAL)
    is_snapped = bool(entry.get("is_snapped"))

    # Prefer restore_rect (set at capture time to the correct snapped or
    # normal rect).  Fall back to rect then normal_rect for old layouts.
    rr = (entry.get("restore_rect")
          or entry.get("rect")
          or entry.get("normal_rect")
          or [])
    if len(rr) < 4:
        return False

    left, top, right, bottom = rr
    w = max(80, right - left)
    h = max(60, bottom - top)
    left, top, w, h = _clamp(left, top, w, h)
    right, bottom = left + w, top + h

    try:
        if is_snapped:
            # Use MoveWindow — mirrors what Aero Snap does internally.
            win32gui.ShowWindow(hwnd, win32con.SW_SHOWNOACTIVATE)
            win32gui.MoveWindow(hwnd, int(left), int(top), int(w), int(h), True)
            if verbose:
                print(f"  RESTORE [SNP] hwnd={hex(hwnd)} "
                      f"-> ({left},{top},{right},{bottom})")
        else:
            # SetWindowPlacement handles max/min/normal atomically.
            cur = win32gui.GetWindowPlacement(hwnd)
            win32gui.SetWindowPlacement(
                hwnd,
                (cur[0], desired, cur[2], cur[3],
                 (int(left), int(top), int(right), int(bottom)))
            )
            if verbose:
                state = {win32con.SW_SHOWMAXIMIZED: "MAX",
                         win32con.SW_SHOWMINIMIZED: "MIN"}.get(desired, "NRM")
                print(f"  RESTORE [{state}] hwnd={hex(hwnd)} "
                      f"-> ({left},{top},{right},{bottom})")
        return True
    except Exception as exc:
        try:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            time.sleep(0.03)
            win32gui.MoveWindow(hwnd, int(left), int(top), int(w), int(h), True)
            win32gui.ShowWindow(hwnd, desired)
            if verbose:
                print(f"  RESTORE [fallback] hwnd={hex(hwnd)} exc={exc}")
            return True
        except Exception:
            return False


# ══════════════════════════════════════════════════════════════════════════
#  Edge tab restore helpers
# ══════════════════════════════════════════════════════════════════════════
def _open_tabs_via_foreground(
    exe: str,
    hwnd: int,
    tabs: List[Dict],
    user_data_dir: str,
    profile_directory: str,
    focus_delay: float = 0.25,
    tab_delay: float = 0.4,
    verbose: bool = False,
) -> int:
    """
    Bring an Edge window to the foreground then open tabs via --new-tab.

    Edge routes --new-tab into whichever window currently has focus, so
    SetForegroundWindow before the subprocess call lets us deliver tabs
    to a specific window without needing a CDP debug session.

    focus_delay:  seconds to wait after SetForegroundWindow before
                  launching the tab command (gives Windows time to
                  complete the focus transfer).
    tab_delay:    seconds to wait after each batch so Edge finishes
                  routing the tabs before we steal focus away for
                  the next window.
    """
    urls = [t["url"] for t in tabs if str(t.get("url") or "").strip()]
    if not urls or not os.path.exists(exe):
        return 0

    try:
        win32gui.SetForegroundWindow(hwnd)
        # If the window is minimised, restore it first so Edge actually
        # considers it the active window for tab routing.
        placement = win32gui.GetWindowPlacement(hwnd)
        if placement[1] == win32con.SW_SHOWMINIMIZED:
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(focus_delay)
    except Exception as exc:
        if verbose:
            print(f"  [warn] SetForegroundWindow hwnd={hex(hwnd)} failed: {exc}")
        # Continue anyway — --new-tab may still land correctly

    base = []
    if user_data_dir:
        base.append(f"--user-data-dir={user_data_dir}")
    if profile_directory:
        base.append(f"--profile-directory={profile_directory}")

    launched = 0
    for i in range(0, len(urls), 10):
        chunk = urls[i : i + 10]
        if verbose:
            print(
                f"  EDGE TABS [fg hwnd={hex(hwnd)}] --new-tab x{len(chunk)}  "
                f"profile={profile_directory!r}"
            )
        try:
            subprocess.Popen([exe, *base, "--new-tab", *chunk])
            launched += len(chunk)
        except Exception as exc:
            print(f"  Warning: could not open Edge tabs: {exc}")
            break
    if launched:
        time.sleep(tab_delay)
    return launched


def _interactive_tab_assignment(
    edge_exe: str,
    group_targets: List[Dict],
    live_edge_hwnds: List[Tuple[int, str]],
    user_data_dir: str,
    profile_directory: str,
    verbose: bool = False,
) -> int:
    """
    Interactive wizard: show the user the live Edge windows and the saved
    tab groups, let them assign each group to a window, then use
    foreground-shifting to deliver tabs precisely.

    group_targets:    saved Edge window entries that have tabs.
    live_edge_hwnds:  [(hwnd, title)] of currently open Edge windows,
                      in z-order (frontmost first).

    Returns total number of tabs opened.
    """
    if not group_targets or not live_edge_hwnds:
        return 0

    print()
    print("  Edge Tab Assignment")
    print("  ─" * 38)
    print("  Live Edge windows (currently open):")
    for i, (hwnd, title) in enumerate(live_edge_hwnds):
        print(f"    [{i}] {title[:70]}")

    total = 0
    assigned_hwnds: Set[int] = set()

    for gi, t in enumerate(group_targets):
        tabs = _normalize_tabs(t.get("edge_tabs") or [])
        if not tabs:
            continue
        saved_title = str(t.get("title") or "(untitled)")
        print()
        print(f"  Tab group {gi + 1}/{len(group_targets)}: '{saved_title[:60]}'")
        for ti, tab in enumerate(tabs):
            print(f"    {ti + 1}. {tab.get('url', '')}")

        # Suggest the best matching live window by title overlap
        best_i = 0
        best_score = -1.0
        st = _strip_edge_suffix(saved_title).split()
        for i, (hwnd, title) in enumerate(live_edge_hwnds):
            lt = _strip_edge_suffix(title).split()
            union = set(st) | set(lt)
            if union:
                score = len(set(st) & set(lt)) / len(union)
                if score > best_score:
                    best_score, best_i = score, i

        sel = input(
            f"  Open in which window? [0-{len(live_edge_hwnds)-1}] "
            f"(default {best_i}, s=skip): "
        ).strip().lower()

        if sel == "s":
            print("  Skipped.")
            continue
        if not sel:
            sel = str(best_i)
        if not sel.isdigit() or not (0 <= int(sel) < len(live_edge_hwnds)):
            print(f"  Invalid choice, skipping.")
            continue

        target_hwnd, target_title = live_edge_hwnds[int(sel)]
        n = _open_tabs_via_foreground(
            exe=edge_exe, hwnd=target_hwnd, tabs=tabs,
            user_data_dir=user_data_dir, profile_directory=profile_directory,
            verbose=verbose,
        )
        total += n
        assigned_hwnds.add(target_hwnd)
        print(f"  ✓ Opened {n} tab(s) in '{target_title[:50]}'")

    return total



def _open_tabs_in_profile(
    exe: str,
    tabs: List[Dict],
    user_data_dir: str,
    profile_directory: str,
    new_window: bool = False,
    verbose: bool = False,
) -> int:
    """
    Open URLs in a specific Edge profile session.

    --user-data-dir and --profile-directory target the exact profile.
    If Edge is already running with that profile, it reuses the existing
    window/session and opens new tabs inside it.
    If not running, Edge starts a new window for that profile.

    new_window=True forces '--new-window' (used when we couldn't match an
    existing window to move).  new_window=False uses '--new-tab' to open
    into the running session.
    """
    urls = [t["url"] for t in tabs if str(t.get("url") or "").strip()]
    if not urls or not os.path.exists(exe):
        return 0

    base = []
    if user_data_dir:
        base.append(f"--user-data-dir={user_data_dir}")
    if profile_directory:
        base.append(f"--profile-directory={profile_directory}")

    flag = "--new-window" if new_window else "--new-tab"
    launched = 0
    for i in range(0, len(urls), 10):
        chunk = urls[i:i + 10]
        cmd   = [exe, *base, flag, *chunk]
        if verbose:
            print(f"  EDGE TABS {flag}  profile={profile_directory!r}  "
                  f"udd={user_data_dir!r}  urls={len(chunk)}")
        try:
            subprocess.Popen(cmd)
            launched += len(chunk)
        except Exception as exc:
            print(f"  Warning: could not open Edge tabs: {exc}")
            break
    return launched


def _edge_size_mismatch(a, b, thr: int = 40) -> bool:
    return (abs((a[2]-a[0]) - (b[2]-b[0])) > thr or
            abs((a[3]-a[1]) - (b[3]-b[1])) > thr)


def _stabilize_edge(applied: List[Tuple[int, Dict]],
                    retries: int = 3, delay: float = 0.35) -> int:
    """Re-apply position if Edge resizes itself after tab loading."""
    fixed = 0
    for _ in range(retries):
        any_fixed = False
        time.sleep(delay)
        for hwnd, target in applied:
            try:
                cur = tuple(win32gui.GetWindowRect(hwnd))
            except Exception:
                continue
            tgt = tuple(target.get("normal_rect") or target.get("rect") or cur)
            if _edge_size_mismatch(cur, tgt):
                if _apply_placement(hwnd, target):
                    fixed += 1
                    any_fixed = True
        if not any_fixed:
            break
    return fixed


# ══════════════════════════════════════════════════════════════════════════
#  Edge missing-window restore helper
# ══════════════════════════════════════════════════════════════════════════
def _launch_and_position_edge_window(
    exe: str,
    target: Dict,
    used: Set,
    user_data_dir: str,
    profile_directory: str,
    wait: float = 4.0,
    verbose: bool = False,
) -> Optional[int]:
    """
    Open a new Edge window for a saved target that is no longer running,
    then position it to match the saved geometry.

    Strategy:
      1. Open Edge with --new-window and the FIRST tab URL as an anchor
         (so the window has a meaningful title and size from the start).
      2. Wait for a new Edge HWND to appear that isn't already in `used`.
      3. Apply the saved placement (position, size, snap/max state).
      4. Return the new HWND so the caller can foreground-shift remaining tabs.

    Returns the new HWND on success, or None if the window couldn't be found.
    """
    tabs = _normalize_tabs(target.get("edge_tabs") or [])
    urls = [t["url"] for t in tabs if str(t.get("url") or "").strip()]

    base = []
    if user_data_dir:
        base.append(f"--user-data-dir={user_data_dir}")
    if profile_directory:
        base.append(f"--profile-directory={profile_directory}")

    # Open with just the first URL so we can find the window by title/position
    first_url = urls[0] if urls else "about:blank"
    try:
        subprocess.Popen([exe, *base, "--new-window", first_url])
    except Exception as exc:
        if verbose:
            print(f"  [warn] Could not launch Edge window: {exc}")
        return None

    # Poll for a new Edge HWND that isn't in `used`
    deadline = time.time() + wait
    new_hwnd: Optional[int] = None
    while time.time() < deadline and new_hwnd is None:
        time.sleep(0.4)
        for w in _live_windows():
            if (str(w.get("process_name") or "").lower() == "msedge.exe"
                    and w["hwnd"] not in used):
                new_hwnd = w["hwnd"]
                break

    if new_hwnd is None:
        if verbose:
            print(f"  [warn] New Edge window did not appear within {wait}s")
        return None

    used.add(new_hwnd)

    # Position the new window to match the saved geometry
    ok = _apply_placement(new_hwnd, target, verbose=verbose)
    if not ok and verbose:
        print(f"  [warn] Could not position new Edge hwnd={hex(new_hwnd)}")

    if verbose:
        title = str(target.get("title") or "")
        print(f"  LAUNCHED Edge hwnd={hex(new_hwnd)} for '{title[:50]}'")

    return new_hwnd


# ══════════════════════════════════════════════════════════════════════════
#  Launch helpers
# ══════════════════════════════════════════════════════════════════════════
def _get_launch_spec(target: Dict) -> Optional[Tuple[str, List[str], str]]:
    launch = target.get("launch")
    exe    = ""
    args: List[str] = []
    cwd    = ""
    if isinstance(launch, dict):
        exe  = str(launch.get("exe") or "").strip()
        raw  = launch.get("args") or []
        args = [str(a) for a in (raw if isinstance(raw, list) else [raw])
                if str(a).strip()]
        cwd  = str(launch.get("cwd") or "").strip()
    if not exe:
        exe = str(target.get("exe") or "").strip()
    if not exe:
        return None
    if os.path.basename(exe).lower() in _BLOCKED_PROC:
        return None
    return exe, args, cwd


def _launch_target(target: Dict) -> bool:
    spec = _get_launch_spec(target)
    if not spec:
        return False
    exe, args, cwd = spec
    if not os.path.exists(exe):
        return False
    try:
        subprocess.Popen([exe, *args], cwd=cwd or None)
        return True
    except Exception:
        return False


# ══════════════════════════════════════════════════════════════════════════
#  Restore
# ══════════════════════════════════════════════════════════════════════════
def restore_layout(
    path: str,
    launch_missing: bool = False,
    edge_tabs: bool = False,
    destructive: bool = False,
    diagnostics: bool = False,
    diagnostics_top_n: int = 3,
    verbose: bool = False,
) -> None:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)

    if str(data.get("schema") or "").strip() != SCHEMA:
        raise ValueError(
            f"Unrecognised layout schema {data.get('schema')!r}. "
            f"Expected {SCHEMA!r}. Re-save your layout with the current tool."
        )

    all_targets = data.get("windows", [])
    targets = [
        t for t in all_targets
        if str(t.get("process_name") or "").lower() not in _BLOCKED_PROC
        and str(t.get("class_name") or "").lower() not in _BLOCKED_CLASS
    ]
    if verbose and len(targets) < len(all_targets):
        n = len(all_targets) - len(targets)
        print(f"  [info] Skipped {n} blocked-process targets from saved layout")

    min_score   = 40
    launch_wait = 6.0

    current       = _live_windows()
    used: Set     = set()
    applied       = 0
    skipped       = 0
    closed        = 0
    missing: List[Dict]                   = []
    force_launch: List[Dict]              = []
    edge_applied: List[Tuple[int, Dict]]  = []
    # Track (hwnd, saved_z_order) for every successfully placed window
    placed_z: List[Tuple[int, int]]       = []

    # ── Match and position ────────────────────────────────────────────────
    for t in targets:
        destructive_target = destructive and bool(t.get("destructive"))
        ranked = _rank_candidates(t, current, used)
        if diagnostics:
            _print_diag(t, ranked, top_n=diagnostics_top_n)
        if not ranked or ranked[0]["score"] < min_score:
            missing.append(t)
            if destructive_target:
                force_launch.append(t)
            continue
        best = ranked[0]["candidate"]
        used.add(best["hwnd"])
        if destructive_target:
            try:
                win32gui.PostMessage(best["hwnd"], win32con.WM_CLOSE, 0, 0)
                closed += 1
                if verbose:
                    print(f"  CLOSE   hwnd={hex(best['hwnd'])} (destructive)")
            except Exception:
                if verbose:
                    print(f"  [warn] Failed to close hwnd={hex(best['hwnd'])}")
            missing.append(t)
            force_launch.append(t)
        else:
            ok = _apply_placement(best["hwnd"], t, verbose=verbose)
            if ok:
                applied += 1
                saved_z = int(t.get("z_order") if t.get("z_order") is not None else 9999)
                placed_z.append((best["hwnd"], saved_z))
                if str(t.get("process_name") or "").lower() == "msedge.exe":
                    edge_applied.append((best["hwnd"], t))
            else:
                skipped += 1

    # ── Launch missing apps ───────────────────────────────────────────────
    launched = 0
    force_ids = {id(t) for t in force_launch}
    if (launch_missing or force_ids) and missing:
        if closed:
            time.sleep(0.3)
        to_launch = []
        seen_ids = set()
        for t in missing:
            key = id(t)
            if key in seen_ids:
                continue
            seen_ids.add(key)
            if (launch_missing or key in force_ids):
                # Edge windows with tabs handled by edge-tabs flow below
                if edge_tabs and str(t.get("process_name") or "").lower() == "msedge.exe":
                    tabs = _normalize_tabs(t.get("edge_tabs") or [])
                    if tabs:
                        continue
                to_launch.append(t)
        for t in to_launch:
            if _launch_target(t):
                launched += 1

        if launched:
            time.sleep(max(0.5, launch_wait))
            current2       = _live_windows()
            still_miss: List[Dict] = []
            for t in missing:
                ranked2 = _rank_candidates(t, current2, used)
                if not ranked2 or ranked2[0]["score"] < min_score:
                    still_miss.append(t)
                    continue
                best2 = ranked2[0]["candidate"]
                used.add(best2["hwnd"])
                ok = _apply_placement(best2["hwnd"], t, verbose=verbose)
                if ok:
                    applied += 1
                    saved_z = int(t.get("z_order") if t.get("z_order") is not None else 9999)
                    placed_z.append((best2["hwnd"], saved_z))
                    if str(t.get("process_name") or "").lower() == "msedge.exe":
                        edge_applied.append((best2["hwnd"], t))
                else:
                    skipped += 1
            missing = still_miss

    skipped += len(missing)

    # ── Edge size stabilisation ───────────────────────────────────────────
    edge_fixes = _stabilize_edge(edge_applied)

    # ── Z-order restore ───────────────────────────────────────────────────
    # Apply after all geometry is settled so nothing gets buried by a later
    # SetWindowPlacement call.  Apply bottom-up so the z_index=0 window
    # ends up genuinely on top.
    z_fixed = _restore_z_order(placed_z, verbose=verbose)
    if verbose:
        print(f"  Z-order: {z_fixed}/{len(placed_z)} windows re-ordered")

    # ── Smart Edge tab restore ──────────────────────────────────────────────
    # Real-world Edge sessions: multiple windows typically share the same
    # user_data_dir + profile_directory ("Default") — they are different
    # windows of one session, not different profiles.
    #
    # Priority order per profile group:
    #   1. CDP (debug port live) -> Target.createTarget per window  [precise]
    #   2. Foreground-shift      -> SetForegroundWindow + --new-tab [good]
    #   3. Interactive wizard    -> user picks which window per group [manual]
    #   4. Best-effort           -> single --new-tab batch           [fallback]
    edge_tabs_opened = 0
    if edge_tabs:
        edge_exe = _find_edge_exe()
        edge_targets = [
            t for t in targets
            if str(t.get("process_name") or "").lower() == "msedge.exe"
        ]
        if edge_exe and edge_targets:
            default_udd    = _default_edge_udd()
            placed_targets = {id(et) for _, et in edge_applied}
            # hwnd_map: target id -> hwnd (for foreground-shift)
            hwnd_map: Dict[int, int] = {id(t): h for h, t in edge_applied}

            # Group by (udd, profile) — the real session identity
            groups: Dict[Tuple[str, str], List[Dict]] = {}
            for t in edge_targets:
                meta = t.get("edge") if isinstance(t.get("edge"), dict) else {}
                udd  = str(meta.get("user_data_dir") or default_udd).strip()
                prof = str(meta.get("profile_directory") or "Default").strip()
                groups.setdefault((udd, prof), []).append(t)

            for (udd, prof), group_targets in groups.items():

                # ── 1. CDP path ───────────────────────────────────────────
                cdp_port = 0
                for t in group_targets:
                    meta = t.get("edge") if isinstance(t.get("edge"), dict) else {}
                    p = int(meta.get("debug_port") or 0)
                    if p and _cdp_alive(p):
                        cdp_port = p
                        break

                if cdp_port:
                    for t in group_targets:
                        tabs = _normalize_tabs(t.get("edge_tabs") or [])
                        if not tabs:
                            continue
                        meta = t.get("edge") if isinstance(t.get("edge"), dict) else {}
                        wid  = int(meta.get("cdp_window_id") or 0)
                        for tab in tabs:
                            url = str(tab.get("url") or "").strip()
                            if not url:
                                continue
                            if wid and _cdp_open_tab_in_window(cdp_port, wid, url):
                                edge_tabs_opened += 1
                                if verbose:
                                    print(f"  EDGE TAB [CDP wid={wid}] {url}")
                            else:
                                n = _open_tabs_in_profile(
                                    exe=edge_exe, tabs=[tab],
                                    user_data_dir=udd, profile_directory=prof,
                                    new_window=False, verbose=verbose,
                                )
                                edge_tabs_opened += n
                    continue

                # ── 2. Foreground-shift path ──────────────────────────────
                # For each saved window that was matched to a live HWND,
                # bring that HWND to the front then open its tabs via
                # --new-tab. Edge routes tabs into the focused window.
                undelivered: List[Dict] = []
                for t in group_targets:
                    tabs = _normalize_tabs(t.get("edge_tabs") or [])
                    hwnd = hwnd_map.get(id(t))
                    if hwnd:
                        if tabs:
                            n = _open_tabs_via_foreground(
                                exe=edge_exe, hwnd=hwnd, tabs=tabs,
                                user_data_dir=udd, profile_directory=prof,
                                verbose=verbose,
                            )
                            edge_tabs_opened += n
                    else:
                        # Window was missing at restore time; collect for
                        # interactive or fallback handling below.
                        undelivered.append(t)

                if not undelivered:
                    continue

                # ── 3. Launch each missing window individually, position it,
                #       then load its tabs via foreground-shift ──────────────
                # This is the correct fix for the "all tabs in one window"
                # problem: each saved Edge window that is missing gets its own
                # --new-window call, we wait for it to appear, position it to
                # the saved geometry, then foreground-shift to load its tabs.
                # The remaining_undelivered list catches windows that still
                # failed to appear after the wait (very slow machines, etc.)
                remaining_undelivered: List[Dict] = []
                for t in undelivered:
                    tabs = _normalize_tabs(t.get("edge_tabs") or [])
                    new_hwnd = _launch_and_position_edge_window(
                        exe=edge_exe, target=t, used=used,
                        user_data_dir=udd, profile_directory=prof,
                        verbose=verbose,
                    )
                    if new_hwnd:
                        # Skip the first URL — it was used as the anchor when
                        # opening the window, it's already in the new tab.
                        if tabs:
                            remaining_tabs = tabs[1:]
                            if remaining_tabs:
                                _open_tabs_via_foreground(
                                    exe=edge_exe, hwnd=new_hwnd,
                                    tabs=remaining_tabs,
                                    user_data_dir=udd, profile_directory=prof,
                                    verbose=verbose,
                                )
                            edge_tabs_opened += len(tabs)
                        saved_z = int(t.get("z_order") if t.get("z_order") is not None else 9999)
                        placed_z.append((new_hwnd, saved_z))
                    else:
                        remaining_undelivered.append(t)

                if not remaining_undelivered:
                    continue

                # ── 4. Interactive wizard for windows that still didn't appear
                live_hwnds: List[Tuple[int, str]] = [
                    (w["hwnd"], w["title"])
                    for w in _live_windows()
                    if str(w.get("process_name") or "").lower() == "msedge.exe"
                ]
                if live_hwnds:
                    try:
                        n = _interactive_tab_assignment(
                            edge_exe=edge_exe,
                            group_targets=remaining_undelivered,
                            live_edge_hwnds=live_hwnds,
                            user_data_dir=udd,
                            profile_directory=prof,
                            verbose=verbose,
                        )
                        edge_tabs_opened += n
                        continue
                    except EOFError:
                        pass  # Non-interactive (pipe/script), fall through

                # ── 5. Last-resort fallback ───────────────────────────────
                all_tabs: List[Dict] = []
                for t in remaining_undelivered:
                    all_tabs.extend(_normalize_tabs(t.get("edge_tabs") or []))
                if all_tabs:
                    if verbose:
                        print(
                            f"  [fallback] {len(all_tabs)} tab(s) could not be "
                            f"placed into individual windows, opening as group"
                        )
                    n = _open_tabs_in_profile(
                        exe=edge_exe, tabs=all_tabs,
                        user_data_dir=udd, profile_directory=prof,
                        new_window=True, verbose=verbose,
                    )
                    edge_tabs_opened += n


    # ── Summary ───────────────────────────────────────────────────────────
    mode_bits = []
    if launch_missing:
        mode_bits.append("launch-missing")
    if edge_tabs:
        mode_bits.append("edge-tabs")
    mode_label = "+".join(mode_bits) if mode_bits else "basic"
    print(
        f"Restore complete ({mode_label}). "
        f"Applied={applied}, Skipped={skipped}, Total={len(targets)}, "
        f"Closed={closed}, Launched={launched}, ZOrder={z_fixed}, "
        f"EdgeTabs={edge_tabs_opened}, EdgeSizeFixes={edge_fixes}"
    )
    if skipped:
        print("  Skipped: title changed, window elevated, or system-managed.")


# ══════════════════════════════════════════════════════════════════════════
#  Edge debug session launcher
# ══════════════════════════════════════════════════════════════════════════
def launch_edge_debug(port: int = 9222, profile_dir: Optional[str] = None,
                      dry_run: bool = False) -> bool:
    exe = _find_edge_exe()
    if not exe:
        print("Edge executable not found.")
        return False
    if _port_open(port):
        print(f"Port {port} already in use.")
        return False
    if not profile_dir:
        profile_dir = os.path.join(
            os.environ.get("TEMP", r"C:\Temp"), "edge-debug"
        )
    args = [f"--remote-debugging-port={port}",
            f"--user-data-dir={profile_dir}"]
    if dry_run:
        print(f"[DRY] {exe} {' '.join(args)}")
        return True
    try:
        subprocess.Popen([exe, *args])
        return True
    except Exception:
        return False


def edge_capture(path: str, edge_debug_port: int = 9222) -> None:
    """Capture live tabs from a CDP debug session into an existing layout."""
    if not os.path.exists(path):
        raise FileNotFoundError(f"Layout not found: {path}")
    if not _cdp_alive(edge_debug_port):
        print(f"No Edge debug endpoint on port {edge_debug_port}.")
        print("Start one with:  python window_layout.py edge-debug")
        return
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    windows = data.get("windows", [])
    _ensure_ids(windows)
    tabs = _fetch_cdp_tabs(edge_debug_port)
    if not tabs:
        print("No tabs returned from CDP.")
        return
    edge_wins = [w for w in windows
                 if str(w.get("process_name") or "").lower() == "msedge.exe"]
    _assign_tabs(edge_wins, tabs)
    for w in edge_wins:
        if w.get("edge_tabs") and isinstance(w.get("edge"), dict):
            w["edge"]["debug_port"] = edge_debug_port
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Captured {len(tabs)} tabs into {len(edge_wins)} Edge windows -> {path}")


def set_edge_urls(
    path: str,
    urls: List[str],
    append: bool = False,
    clear: bool = False,
    window_id: Optional[str] = None,
) -> None:
    """
    Set or modify the saved Edge tabs for a specific Edge window entry.

    window_id: the UUID stored on the window entry (windows[*].window_id).
               If None, raises ValueError — callers must always specify which
               window to target.  Use list_edge_windows() or the CLI wizard
               to discover valid IDs.
    """
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    edge_wins = [w for w in data.get("windows", [])
                 if str(w.get("process_name") or "").lower() == "msedge.exe"]
    if not edge_wins:
        raise ValueError("No Edge windows in layout")
    if not window_id:
        raise ValueError(
            "window_id is required. Run the CLI wizard or call "
            "list_edge_windows() to find the correct window_id."
        )
    target = next(
        (w for w in edge_wins if str(w.get("window_id") or "") == window_id),
        None,
    )
    if target is None:
        known = [str(w.get("window_id") or "") for w in edge_wins]
        raise KeyError(
            f"No Edge window with window_id={window_id!r}. "
            f"Known Edge window IDs: {known}"
        )
    current = _normalize_tabs(target.get("edge_tabs") or [])
    if clear:
        current = []
    new_tabs = [{"title": "", "url": u} for u in urls if u.strip()]
    target["edge_tabs"] = _normalize_tabs(
        current + new_tabs if append else new_tabs
    )
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    label = str(target.get("title") or window_id)
    tabs_set = len(target["edge_tabs"])
    print(f"Set {tabs_set} URLs on '{label[:60]}' -> {path}")


def list_edge_windows(path: str) -> List[Dict]:
    """
    Return a list of Edge window entries from a layout, each containing
    window_id, title, profile_directory, and saved_tab_count.
    Useful for GUIs and callers that need to pick a window before calling
    set_edge_urls().
    """
    if not os.path.exists(path):
        raise FileNotFoundError(path)
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    result = []
    for w in data.get("windows", []):
        if str(w.get("process_name") or "").lower() != "msedge.exe":
            continue
        result.append({
            "window_id":         str(w.get("window_id") or ""),
            "title":             str(w.get("title") or ""),
            "profile_directory": str((w.get("edge") or {}).get("profile_directory") or "Default"),
            "user_data_dir":     str((w.get("edge") or {}).get("user_data_dir") or ""),
            "saved_tab_count":   len(w.get("edge_tabs") or []),
            "z_order":           int(w.get("z_order") if w.get("z_order") is not None else 9999),
        })
    return result


# ══════════════════════════════════════════════════════════════════════════
#  Hotkey listener
# ══════════════════════════════════════════════════════════════════════════
def _load_config(path: str = CONFIG_PATH) -> Dict:
    if not os.path.exists(path):
        return {}
    try:
        with open(path, encoding="utf-8") as f:
            d = json.load(f)
        return d if isinstance(d, dict) else {}
    except Exception:
        return {}


def _parse_hotkey(keys: str) -> Optional[Tuple[int, int]]:
    parts = [p.strip() for p in keys.replace("-", "+").split("+") if p.strip()]
    mod, key = 0, ""
    for p in parts:
        lo = p.lower()
        if   lo in ("ctrl", "control"):         mod |= win32con.MOD_CONTROL
        elif lo == "alt":                        mod |= win32con.MOD_ALT
        elif lo == "shift":                      mod |= win32con.MOD_SHIFT
        elif lo in ("win", "windows", "meta"):  mod |= win32con.MOD_WIN
        else:                                    key = p
    if not key: return None
    ku = key.upper()
    if len(ku) == 1 and ku.isalnum(): return mod, ord(ku)
    if ku.startswith("F") and ku[1:].isdigit() and 1 <= int(ku[1:]) <= 24:
        return mod, win32con.VK_F1 + int(ku[1:]) - 1
    vk = {"TAB": win32con.VK_TAB, "ENTER": win32con.VK_RETURN,
          "ESC": win32con.VK_ESCAPE, "SPACE": win32con.VK_SPACE,
          "DELETE": win32con.VK_DELETE, "HOME": win32con.VK_HOME,
          "END": win32con.VK_END, "LEFT": win32con.VK_LEFT,
          "RIGHT": win32con.VK_RIGHT, "UP": win32con.VK_UP,
          "DOWN": win32con.VK_DOWN}
    return (mod, vk[ku]) if ku in vk else None


def run_hotkey_listener(config_path: str = CONFIG_PATH) -> None:
    cfg     = _load_config(config_path)
    hotkeys = [h for h in (cfg.get("hotkeys") or [])
               if isinstance(h, dict) and h.get("keys") and h.get("action")]
    if not hotkeys:
        print(f"No hotkeys configured in {config_path}")
        return
    try:
        win32gui.PeekMessage(None, 0, 0, win32con.PM_NOREMOVE)
    except Exception:
        pass
    registered: Dict[int, Dict] = {}
    for i, h in enumerate(hotkeys, 1):
        parsed = _parse_hotkey(str(h["keys"]))
        if not parsed:
            print(f"  Skip invalid hotkey: {h['keys']}")
            continue
        try:
            win32gui.RegisterHotKey(None, i, *parsed)
            registered[i] = h
            print(f"  Registered {h['keys']} -> {h['action']}")
        except Exception as exc:
            print(f"  Failed {h['keys']}: {exc}")
    if not registered:
        return
    try:
        while True:
            msg = win32gui.GetMessage(None, 0, 0)
            payload = msg[1] if isinstance(msg, (list, tuple)) and len(msg) == 2 else msg
            message = wparam = None
            if isinstance(payload, (list, tuple)):
                if   len(payload) == 6: _, message, wparam, *_ = payload
                elif len(payload) == 3: message, wparam, _     = payload
                elif len(payload) == 2: message, wparam        = payload
            if message == win32con.WM_QUIT:
                break
            if message == win32con.WM_HOTKEY:
                h = registered.get(wparam)
                if h:
                    raw_args = h.get("args") or []
                    if isinstance(raw_args, str): raw_args = [raw_args]
                    subprocess.Popen(
                        [sys.executable, os.path.abspath(__file__),
                         str(h["action"]), *[str(a) for a in raw_args]]
                    )
    except KeyboardInterrupt:
        pass
    finally:
        for i in registered:
            try: win32api.UnregisterHotKey(None, i)
            except Exception: pass


# ══════════════════════════════════════════════════════════════════════════
#  Wizards
# ══════════════════════════════════════════════════════════════════════════
def _prompt(text: str, default: str = "") -> str:
    val = input(f"{text} [{default}]: " if default else f"{text}: ").strip()
    return val or default

def _yn(text: str, default: bool = False) -> bool:
    val = input(f"{text} ({'Y/n' if default else 'y/N'}): ").strip().lower()
    return default if not val else val in ("y", "yes")

def run_wizard() -> None:
    print("=== TSD Workspace Setup Wizard ===\n")
    path     = _prompt("Output layout path",
                        os.path.abspath("layouts/layout.json"))
    cap_tabs = _yn("Capture Edge tabs via debug session?", default=False)
    port     = 9222
    if cap_tabs:
        port = int(_prompt("Edge debug port", "9222") or "9222")
        if _yn("Launch Edge debug session now?", default=True):
            if launch_edge_debug(port=port):
                print("Edge debug launched. Set up your tabs, then return.")
                input("Press Enter when ready to capture…")
            else:
                print("Could not launch debug Edge. Continuing without tab capture.")
    save_layout(path, capture_edge_tabs=cap_tabs, edge_debug_port=port,
                verbose=True)
    print(f"\nDone. Restore with:\n"
          f"  python window_layout.py restore {path} --edge-tabs --diagnostics")

def run_edit_wizard(path: str) -> None:
    with open(path, encoding="utf-8") as f:
        data = json.load(f)
    wins      = data.get("windows", [])
    _ensure_ids(wins)
    edge_wins = [w for w in wins
                 if str(w.get("process_name") or "").lower() == "msedge.exe"]
    if not edge_wins:
        print("No Edge windows in layout.")
        return
    all_tabs: List[Dict] = []
    for w in edge_wins:
        all_tabs.extend(_normalize_tabs(w.get("edge_tabs") or []))
    if not all_tabs:
        print("No saved Edge tabs found.")
        return
    print("Saved tabs:")
    for i, t in enumerate(all_tabs, 1):
        print(f"  [{i}] {t.get('title','')} -> {t.get('url','')}")
    used: Set[int] = set()
    for w in edge_wins:
        title    = w.get("title", "(untitled)")
        existing = _normalize_tabs(w.get("edge_tabs") or [])
        cur = [str(i+1) for i, t in enumerate(all_tabs)
               if any(t.get("url") == e.get("url") for e in existing)]
        sel = _prompt(f"Window \"{title}\" tab indices (comma-separated)",
                      ",".join(cur))
        chosen = [int(x)-1 for x in sel.split(",")
                  if x.strip().isdigit() and 0 < int(x) <= len(all_tabs)]
        w["edge_tabs"] = [all_tabs[i] for i in chosen]
        used.update(chosen)
    unassigned = [t for i, t in enumerate(all_tabs) if i not in used]
    if unassigned:
        print(f"  {len(unassigned)} tabs unassigned.")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    print(f"Saved -> {path}")


# ══════════════════════════════════════════════════════════════════════════
#  CLI entry point
# ══════════════════════════════════════════════════════════════════════════
def main() -> None:
    p = argparse.ArgumentParser(
        description="Save/restore Windows desktop layouts."
    )
    s = p.add_subparsers(dest="cmd", required=True)

    sp = s.add_parser("save")
    sp.add_argument("json_path")
    sp.add_argument("--edge-tabs",       action="store_true")
    sp.add_argument("--edge-debug-port", type=int, default=9222)
    sp.add_argument("--verbose", "-v",   action="store_true")

    sp = s.add_parser("restore")
    sp.add_argument("json_path")
    sp.add_argument("--launch-missing",  action="store_true",
                    help="Launch apps that are missing before positioning")
    sp.add_argument("--edge-tabs",       action="store_true",
                    help="Restore saved Edge tabs/windows")
    sp.add_argument("--destructive",     action="store_true",
                    help="Close matching windows and relaunch them clean")
    sp.add_argument("--diagnostics",     action="store_true")
    sp.add_argument("--diagnostics-top", type=int, default=3)
    sp.add_argument("--verbose", "-v",   action="store_true")

    sp = s.add_parser("edge-debug")
    sp.add_argument("--port",       type=int, default=9222)
    sp.add_argument("--profile-dir")
    sp.add_argument("--dry-run",    action="store_true")

    sp = s.add_parser("edge-capture")
    sp.add_argument("json_path")
    sp.add_argument("--port", type=int, default=9222)

    sp = s.add_parser("edge-urls",
        help="Set saved URLs for an Edge window (wizard if no URLs given)")
    sp.add_argument("json_path")
    sp.add_argument("urls", nargs="*",
        help="URLs to set (omit to run interactive wizard)")
    sp.add_argument("--window-id", "-w", default=None,
        dest="window_id",
        help="window_id UUID of the Edge window to target (wizard picks if omitted)")
    sp.add_argument("--append", action="store_true",
        help="Append URLs instead of replacing")
    sp.add_argument("--clear",  action="store_true",
        help="Clear existing URLs before applying")

    sp = s.add_parser("hotkeys")
    sp.add_argument("--config", default=CONFIG_PATH)

    s.add_parser("wizard")
    sp = s.add_parser("edit")
    sp.add_argument("json_path")

    sp = s.add_parser("help")
    sp.add_argument("--full", action="store_true")

    args = p.parse_args()

    if args.cmd == "save":
        save_layout(args.json_path,
                    capture_edge_tabs=args.edge_tabs,
                    edge_debug_port=args.edge_debug_port,
                    verbose=args.verbose)

    elif args.cmd == "restore":
        restore_layout(args.json_path,
                       launch_missing=args.launch_missing,
                       edge_tabs=args.edge_tabs,
                       destructive=args.destructive,
                       diagnostics=args.diagnostics,
                       diagnostics_top_n=args.diagnostics_top,
                       verbose=args.verbose)

    elif args.cmd == "edge-debug":
        ok = launch_edge_debug(port=args.port, profile_dir=args.profile_dir,
                               dry_run=args.dry_run)
        print("Edge debug launched." if ok else "Failed.")

    elif args.cmd == "edge-capture":
        edge_capture(args.json_path, edge_debug_port=args.port)

    elif args.cmd == "edge-urls":
        # ── Wizard mode: no URLs (and no explicit --window-id) provided ───────
        if not args.urls and args.window_id is None:
            _wins = list_edge_windows(args.json_path)
            if not _wins:
                print("No Edge windows found in layout.")
            else:
                print(f"Edge windows in {args.json_path}:")
                for _w in _wins:
                    print(f"  id={_w['window_id']}")
                    print(f"     title={_w['title'][:60] or '(untitled)'}")
                    print(f"     profile={_w['profile_directory']}  saved_tabs={_w['saved_tab_count']}")
                _chosen_id = _prompt("window_id to edit").strip()
                if not _chosen_id:
                    print("No window_id entered — aborting.")
                else:
                    print("Enter URLs one per line. Blank line to finish:")
                    _urls: List[str] = []
                    while True:
                        _u = input("  URL: ").strip()
                        if not _u:
                            break
                        _urls.append(_u)
                    _mode = _prompt("Mode — (r)eplace / (a)ppend / (c)lear", "r").lower()
                    set_edge_urls(
                        args.json_path, _urls,
                        append=_mode.startswith("a"),
                        clear=_mode.startswith("c"),
                        window_id=_chosen_id,
                    )
        else:
            # ── Non-interactive: --window-id and URLs provided ────────────────
            set_edge_urls(
                args.json_path, args.urls,
                append=args.append,
                clear=args.clear,
                window_id=args.window_id,
            )

    elif args.cmd == "hotkeys":
        run_hotkey_listener(config_path=args.config)

    elif args.cmd == "wizard":
        run_wizard()

    elif args.cmd == "edit":
        run_edit_wizard(args.json_path)

    elif args.cmd == "help":
        if args.full:
            p.print_help()
        else:
            print("""
Quick reference
  save:         python window_layout.py save layout.json [-v]
  save+tabs:    python window_layout.py save layout.json --edge-tabs
  restore:      python window_layout.py restore layout.json [--launch-missing] [--edge-tabs] --diagnostics [-v]
  edge-debug:   python window_layout.py edge-debug --port 9222
  edge-capture: python window_layout.py edge-capture layout.json
  edge-urls:    python window_layout.py edge-urls layout.json https://example.com
  hotkeys:      python window_layout.py hotkeys
  wizard:       python window_layout.py wizard
""")


if __name__ == "__main__":
    main()
