"""
=============================================================================
  WINDOW LAYOUT TESTBENCH  –  verbose diagnostic & guided first-run CLI
=============================================================================

PURPOSE
  Deep-dive into everything win32 + psutil can reveal about your running
  desktop so you can understand WHY window_layout.py behaves the way it does
  and what data it collects, matches on, and restores.

USAGE
  python testbench_window_layout.py              # full guided walkthrough
  python testbench_window_layout.py --snapshot   # one-shot dump, no prompts
  python testbench_window_layout.py --edge-only  # only dump Edge/tab info
  python testbench_window_layout.py --compare before.json after.json

REQUIREMENTS
  pip install pywin32 psutil
  (same deps as window_layout.py)

=============================================================================
"""

import argparse
import ctypes
import json
import os
import socket
import sys
import time
import urllib.request
import urllib.error
import traceback
from datetime import datetime
from typing import Dict, List, Optional, Tuple

# ── graceful import ────────────────────────────────────────────────────────
try:
    import psutil
    import win32api
    import win32con
    import win32gui
    import win32process
    import win32com.client
    _WIN32_OK = True
except ImportError as _ie:
    _WIN32_OK = False
    _WIN32_MISSING = str(_ie)

# ══════════════════════════════════════════════════════════════════════════
#  ANSI colour helpers (works on modern Windows terminals / VS Code)
# ══════════════════════════════════════════════════════════════════════════
def _ansi(code: str) -> str:
    return f"\033[{code}m"

R  = _ansi("91")   # red
G  = _ansi("92")   # green
Y  = _ansi("93")   # yellow
B  = _ansi("94")   # blue
M  = _ansi("95")   # magenta
C  = _ansi("96")   # cyan
W  = _ansi("97")   # white
DIM = _ansi("2")
RST = _ansi("0")

def banner(text: str, char: str = "═", colour: str = C) -> None:
    width = 80
    line  = char * width
    print(f"\n{colour}{line}")
    print(f"  {W}{text}{colour}")
    print(f"{line}{RST}")

def section(text: str) -> None:
    print(f"\n{Y}▶ {W}{text}{RST}")

def ok(text: str)   -> None: print(f"  {G}✓ {RST}{text}")
def warn(text: str) -> None: print(f"  {Y}⚠ {RST}{text}")
def err(text: str)  -> None: print(f"  {R}✗ {RST}{text}")
def info(text: str) -> None: print(f"  {C}· {RST}{text}")
def kv(key: str, val) -> None:
    print(f"    {DIM}{key:<28}{RST} {val}")

def pause(msg: str = "Press ENTER to continue...") -> None:
    try:
        input(f"\n  {M}→ {W}{msg}{RST} ")
    except (EOFError, KeyboardInterrupt):
        print()

# ══════════════════════════════════════════════════════════════════════════
#  SECTION 1 – Environment
# ══════════════════════════════════════════════════════════════════════════
def dump_environment() -> None:
    banner("1 · ENVIRONMENT & DISPLAY")

    section("Python / Platform")
    kv("Python",        sys.version.split()[0])
    kv("Executable",    sys.executable)
    kv("Platform",      sys.platform)
    kv("win32 present", str(_WIN32_OK))
    if not _WIN32_OK:
        err(f"pywin32 / psutil not importable: {_WIN32_MISSING}")
        err("Install with:  pip install pywin32 psutil")
        return

    section("Windows Version")
    try:
        vi = sys.getwindowsversion()
        kv("Major.Minor.Build", f"{vi.major}.{vi.minor}.{vi.build}")
        kv("Service Pack",      vi.service_pack or "(none)")
    except Exception as e:
        warn(f"Could not read Windows version: {e}")

    section("DPI / Virtual Screen")
    try:
        user32 = ctypes.windll.user32
        user32.SetProcessDPIAware()
        pw = user32.GetSystemMetrics(78)   # SM_CXVIRTUALSCREEN
        ph = user32.GetSystemMetrics(79)   # SM_CYVIRTUALSCREEN
        px = user32.GetSystemMetrics(76)   # SM_XVIRTUALSCREEN
        py = user32.GetSystemMetrics(77)   # SM_YVIRTUALSCREEN
        kv("Virtual screen origin",     f"({px}, {py})")
        kv("Virtual screen size",       f"{pw} × {ph}")
        kv("Primary W",                 user32.GetSystemMetrics(0))
        kv("Primary H",                 user32.GetSystemMetrics(1))
        kv("Monitor count",             user32.GetSystemMetrics(80))
    except Exception as e:
        warn(f"DPI query failed: {e}")

    section("Monitor Details (EnumDisplayMonitors)")
    try:
        monitors = win32api.EnumDisplayMonitors(None, None)
        for i, mon in enumerate(monitors):
            hmon, hdc, rect = mon
            try:
                info_struct = win32api.GetMonitorInfo(hmon)
                flags    = info_struct.get("Flags", 0)
                wrect    = info_struct.get("Work")
                mrect    = info_struct.get("Monitor")
                is_prim  = "PRIMARY" if flags & 1 else ""
                print(f"\n    {C}Monitor {i}{RST}  {is_prim}")
                kv("  Monitor rect",  mrect)
                kv("  Work area",     wrect)
            except Exception as me:
                warn(f"Monitor {i} info failed: {me}")
    except Exception as e:
        warn(f"EnumDisplayMonitors failed: {e}")


# ══════════════════════════════════════════════════════════════════════════
#  SECTION 2 – Raw window enumeration (everything, unfiltered)
# ══════════════════════════════════════════════════════════════════════════
def _raw_all_windows() -> List[Dict]:
    """Enumerate ALL top-level windows with zero filtering."""
    results = []

    def _cb(hwnd, _):
        try:
            title      = win32gui.GetWindowText(hwnd) or ""
            cls        = win32gui.GetClassName(hwnd) or ""
            visible    = bool(win32gui.IsWindowVisible(hwnd))
            iconic     = bool(win32gui.IsIconic(hwnd))
            parent     = win32gui.GetParent(hwnd)
            ex_style   = win32gui.GetWindowLong(hwnd, win32con.GWL_EXSTYLE)
            style      = win32gui.GetWindowLong(hwnd, win32con.GWL_STYLE)
            owner      = win32gui.GetWindow(hwnd, win32con.GW_OWNER)

            try:
                rect = win32gui.GetWindowRect(hwnd)
            except Exception:
                rect = (0, 0, 0, 0)

            try:
                placement = win32gui.GetWindowPlacement(hwnd)
                show_cmd   = placement[1]
                norm_rect  = placement[4]
            except Exception:
                show_cmd  = -1
                norm_rect = (0, 0, 0, 0)

            try:
                _, pid = win32process.GetWindowThreadProcessId(hwnd)
            except Exception:
                pid = 0

            try:
                p = psutil.Process(pid)
                proc_name = p.name()
                exe       = p.exe()
            except Exception:
                proc_name = ""
                exe       = ""

            tool_win  = bool(ex_style & win32con.WS_EX_TOOLWINDOW)
            app_win   = bool(ex_style & win32con.WS_EX_APPWINDOW)
            no_activ  = bool(ex_style & win32con.WS_EX_NOACTIVATE)

            w = rect[2] - rect[0]
            h = rect[3] - rect[1]

            results.append({
                "hwnd":       hwnd,
                "title":      title,
                "class":      cls,
                "pid":        pid,
                "proc":       proc_name,
                "exe":        exe,
                "visible":    visible,
                "iconic":     iconic,
                "parent":     parent,
                "owner":      owner,
                "show_cmd":   show_cmd,
                "rect":       rect,
                "norm_rect":  norm_rect,
                "w":          w,
                "h":          h,
                "ex_style":   ex_style,
                "style":      style,
                "tool_win":   tool_win,
                "app_win":    app_win,
                "no_activ":   no_activ,
            })
        except Exception:
            pass

    win32gui.EnumWindows(_cb, None)
    return results


SHOW_CMD_NAMES = {
    0: "SW_HIDE",
    1: "SW_SHOWNORMAL",
    2: "SW_SHOWMINIMIZED",
    3: "SW_SHOWMAXIMIZED",
    4: "SW_SHOWNOACTIVATE",
    5: "SW_SHOW",
    6: "SW_MINIMIZE",
    7: "SW_SHOWMINNOACTIVE",
    8: "SW_SHOWNA",
    9: "SW_RESTORE",
    10: "SW_SHOWDEFAULT",
}


def dump_all_windows(verbose: bool = True) -> List[Dict]:
    banner("2 · ALL TOP-LEVEL WINDOWS  (unfiltered)")

    all_wins = _raw_all_windows()
    section(f"Total windows enumerated: {len(all_wins)}")

    # Apply same filter logic as window_layout.py for comparison
    def _would_capture(w: Dict) -> bool:
        if not w["visible"]:              return False
        if w["parent"]:                   return False
        if not w["title"].strip():        return False
        if w["tool_win"] and not w["app_win"]: return False
        if w["owner"] and not w["app_win"]:    return False
        if w["w"] < 120 or w["h"] < 80:  return False
        return True

    captured   = [w for w in all_wins if _would_capture(w)]
    excluded   = [w for w in all_wins if not _would_capture(w)]

    ok(f"{len(captured)} windows WOULD be captured by window_layout.py")
    warn(f"{len(excluded)} windows excluded by filters")

    section("═══ CAPTURED WINDOWS ═══")
    for i, w in enumerate(captured):
        sc_name = SHOW_CMD_NAMES.get(w["show_cmd"], str(w["show_cmd"]))
        print(f"\n  {G}[{i:02d}]{RST}  {W}{w['title'][:70]}{RST}")
        kv("HWND",        hex(w["hwnd"]))
        kv("Class",       w["class"])
        kv("Process",     f"{w['proc']}  (PID {w['pid']})")
        kv("Exe",         w["exe"])
        kv("Rect",        f"L={w['rect'][0]} T={w['rect'][1]} R={w['rect'][2]} B={w['rect'][3]}  ({w['w']}×{w['h']})")
        kv("NormRect",    w["norm_rect"])
        kv("ShowCmd",     f"{w['show_cmd']} → {sc_name}")
        kv("Visible",     w["visible"])
        kv("Minimised",   w["iconic"])
        kv("ToolWindow",  w["tool_win"])
        kv("AppWindow",   w["app_win"])

    if verbose:
        section("═══ EXCLUDED WINDOWS (with reason) ═══")
        for w in excluded:
            if not w["visible"] and not w["title"]: continue   # skip truly empty
            reasons = []
            if not w["visible"]:  reasons.append("invisible")
            if w["parent"]:       reasons.append(f"has parent {hex(w['parent'])}")
            if not w["title"].strip(): reasons.append("no title")
            if w["tool_win"] and not w["app_win"]: reasons.append("TOOLWINDOW")
            if w["owner"] and not w["app_win"]:    reasons.append(f"owned+no APPWINDOW")
            if w["w"] < 120 or w["h"] < 80:        reasons.append(f"too small {w['w']}×{w['h']}")
            if not reasons: reasons = ["unknown"]
            title_short = (w["title"] or "(no title)")[:50]
            print(f"  {DIM}[excl]{RST}  {title_short}  {Y}← {', '.join(reasons)}{RST}  cls={w['class']}  proc={w['proc']}")

    return captured


# ══════════════════════════════════════════════════════════════════════════
#  SECTION 3 – Process snapshot
# ══════════════════════════════════════════════════════════════════════════
def dump_processes() -> None:
    banner("3 · RUNNING PROCESSES  (GUI-relevant)")

    gui_procs = [
        "msedge.exe", "chrome.exe", "firefox.exe", "opera.exe", "brave.exe",
        "explorer.exe", "code.exe", "notepad.exe", "notepad++.exe",
        "devenv.exe", "slack.exe", "teams.exe", "discord.exe",
        "outlook.exe", "excel.exe", "winword.exe", "powerpnt.exe",
        "cmd.exe", "powershell.exe", "windowsterminal.exe",
        "mspaint.exe", "calc.exe", "taskmgr.exe",
    ]

    found = {}
    for proc in psutil.process_iter(["pid", "name", "exe", "status", "memory_info", "create_time"]):
        try:
            name = (proc.info["name"] or "").lower()
            if name in gui_procs:
                found.setdefault(name, []).append(proc.info)
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

    if not found:
        info("No common GUI processes found running.")
        return

    for name, instances in sorted(found.items()):
        section(f"{name}  ({len(instances)} instance(s))")
        for p in instances:
            try:
                ct = datetime.fromtimestamp(p["create_time"]).strftime("%H:%M:%S")
            except Exception:
                ct = "?"
            mem_mb = round(p["memory_info"].rss / 1024 / 1024, 1) if p.get("memory_info") else "?"
            kv("  PID",    p["pid"])
            kv("  Status", p["status"])
            kv("  Mem MB", mem_mb)
            kv("  Start",  ct)
            kv("  Exe",    p.get("exe") or "(n/a)")


# ══════════════════════════════════════════════════════════════════════════
#  SECTION 4 – Edge tab detection via CDP
# ══════════════════════════════════════════════════════════════════════════
def _check_port_open(host: str, port: int, timeout: float = 1.0) -> bool:
    try:
        s = socket.create_connection((host, port), timeout=timeout)
        s.close()
        return True
    except Exception:
        return False


def dump_edge_tabs(port: int = 9222) -> None:
    banner("4 · MICROSOFT EDGE / BROWSER TABS  (CDP)")

    section("Checking if Edge remote debugging port is open")
    is_open = _check_port_open("127.0.0.1", port)
    if is_open:
        ok(f"Port {port} is OPEN  →  Edge debug session reachable")
    else:
        err(f"Port {port} is CLOSED")
        warn("Edge must be launched with --remote-debugging-port=9222 for tab capture.")
        warn("Use:  python window_layout.py edge-debug --port 9222")
        warn("Then re-open your normal Edge windows in THAT session.")
        section("Checking common debug ports (9222-9229)")
        for p in range(9222, 9230):
            if _check_port_open("127.0.0.1", p):
                ok(f"  Port {p} is open (try --edge-debug-port {p})")
        return

    section(f"Fetching tab list from  http://127.0.0.1:{port}/json/list")
    try:
        url  = f"http://127.0.0.1:{port}/json/list"
        with urllib.request.urlopen(url, timeout=3) as resp:
            raw = resp.read().decode("utf-8", errors="replace")
        tabs_raw = json.loads(raw)
    except Exception as e:
        err(f"CDP request failed: {e}")
        return

    pages = [t for t in tabs_raw if t.get("type") == "page"]
    other = [t for t in tabs_raw if t.get("type") != "page"]

    ok(f"Total CDP targets: {len(tabs_raw)}  ({len(pages)} pages,  {len(other)} other)")

    section("PAGE TABS")
    window_groups: Dict[int, List[Dict]] = {}
    for tab in pages:
        wid = tab.get("windowId") if isinstance(tab.get("windowId"), int) else -1
        window_groups.setdefault(wid, []).append(tab)

    for wid, tabs in sorted(window_groups.items()):
        label = f"WindowId={wid}" if wid >= 0 else "Unknown Window"
        print(f"\n  {C}{label}{RST}  ({len(tabs)} tabs)")
        for t in tabs:
            title = (t.get("title") or "")[:60]
            url   = (t.get("url")   or "")[:80]
            tid   = t.get("id", "?")
            print(f"    {DIM}[{tid}]{RST}  {W}{title}{RST}")
            print(f"           {DIM}{url}{RST}")

    if other:
        section("NON-PAGE CDP TARGETS (workers, extensions, etc.)")
        for t in other:
            print(f"  type={t.get('type')}  title={t.get('title','')[:50]}")

    section("⚠  IMPORTANT: Edge Tab Capture Caveats")
    warn("Tabs can only be captured when Edge is running in debug mode.")
    warn("Normal Edge and debug-Edge are SEPARATE profiles by default.")
    warn("You must set up your workflow INSIDE the debug-Edge session.")
    warn("window_layout.py matches tabs to Edge windows by title heuristic")
    warn("→ if you have 2 Edge windows, tab assignment can be ambiguous!")
    info(f"CDP window IDs seen: {sorted(window_groups.keys())}")


# ══════════════════════════════════════════════════════════════════════════
#  SECTION 5 – Explorer / File Windows
# ══════════════════════════════════════════════════════════════════════════
def dump_explorer_windows() -> None:
    banner("5 · EXPLORER / FILE MANAGER WINDOWS")
    try:
        shell   = win32com.client.Dispatch("Shell.Application")
        windows = shell.Windows()
        count   = 0
        for window in windows:
            try:
                hwnd     = int(getattr(window, "HWND", 0) or 0)
                location = str(getattr(window, "LocationURL", "") or "").strip()
                name     = str(getattr(window, "LocationName", "") or "").strip()
                path     = location.replace("file:///", "").replace("/", "\\") if location.startswith("file:///") else location
                print(f"  {G}HWND={hex(hwnd)}{RST}  {W}{name}{RST}")
                kv("  Path", path)
                count += 1
            except Exception as e:
                warn(f"Could not read shell window: {e}")
        if count == 0:
            info("No Explorer / Shell windows detected.")
    except Exception as e:
        err(f"Shell.Application dispatch failed: {e}")
        warn("This is needed to capture Explorer window paths for restore.")


# ══════════════════════════════════════════════════════════════════════════
#  SECTION 6 – Scoring / matching simulation
# ══════════════════════════════════════════════════════════════════════════
def _title_similarity(a: str, b: str) -> float:
    """Simple token overlap score (mirrors window_layout.py approach)."""
    a_tokens = set(a.lower().split())
    b_tokens = set(b.lower().split())
    if not a_tokens or not b_tokens:
        return 0.0
    intersection = a_tokens & b_tokens
    return len(intersection) / max(len(a_tokens), len(b_tokens))


def dump_match_simulation(captured: List[Dict]) -> None:
    banner("6 · MATCH / RESTORE SIMULATION")
    section("Simulating what would happen if you saved NOW then restored NOW")

    # Build targets from captured and run matching against itself
    problems = []
    for i, target in enumerate(captured):
        title    = target["title"]
        proc     = target["proc"]
        matches  = []
        for j, candidate in enumerate(captured):
            if i == j: continue
            if candidate["proc"] != proc: continue
            score = _title_similarity(title, candidate["title"])
            if score > 0.4:
                matches.append((score, j, candidate["title"]))
        if matches:
            matches.sort(reverse=True)
            problems.append((i, title, proc, matches))

    if not problems:
        ok("No ambiguous title matches detected — restore should be clean.")
    else:
        warn(f"{len(problems)} windows have similar titles within the same process:")
        warn("These could get mis-matched during restore!")
        for (i, title, proc, matches) in problems:
            print(f"\n  {R}[{i:02d}]{RST}  {W}{title[:60]}{RST}  ({proc})")
            for score, j, mtitle in matches[:3]:
                print(f"       ↔ score={score:.2f}  [{j:02d}] {mtitle[:60]}")

    section("Minimised windows (restore requires special handling)")
    minimised = [w for w in captured if w["iconic"]]
    if minimised:
        for w in minimised:
            warn(f"  {w['title'][:60]}  [{w['proc']}]")
        info("window_layout.py stores show_cmd and normal_rect to restore these correctly.")
        info("If restore fails, check that SW_SHOWMINIMIZED + SetWindowPlacement is used.")
    else:
        ok("No minimised windows in current capture.")

    section("Off-screen / partially-off-screen windows")
    for w in captured:
        l, t, r, b = w["rect"]
        off = (r < 0) or (b < 0) or (l > 7680) or (t > 4320)
        if off:
            warn(f"  FULLY off-screen: {w['title'][:50]}  rect={w['rect']}")


# ══════════════════════════════════════════════════════════════════════════
#  SECTION 7 – window_layout.py compatibility audit
# ══════════════════════════════════════════════════════════════════════════
def dump_compatibility_audit() -> None:
    banner("7 · COMPATIBILITY & KNOWN BOTTLENECKS")

    section("Elevated / protected processes")
    elevated_procs = []
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            p = psutil.Process(proc.info["pid"])
            p.memory_info()   # will raise AccessDenied for elevated
        except psutil.AccessDenied:
            elevated_procs.append(proc.info.get("name", "?"))
        except (psutil.NoSuchProcess, Exception):
            pass

    if elevated_procs:
        common_el = [e for e in elevated_procs if e.lower() not in (
            "system", "registry", "smss.exe", "csrss.exe", "wininit.exe",
            "services.exe", "lsass.exe", "svchost.exe"
        )]
        if common_el:
            warn("These user-visible processes run elevated (cannot be moved by a non-elevated script):")
            for e in sorted(set(common_el)):
                warn(f"  {e}")
            info("Run window_layout.py as Administrator to move these windows.")
        else:
            ok("No user-visible elevated processes detected.")
    else:
        ok("No AccessDenied processes detected.")

    section("Edge debug port connectivity")
    for port in [9222, 9223]:
        if _check_port_open("127.0.0.1", port):
            ok(f"  Port {port} open (Edge debug active)")
        else:
            info(f"  Port {port} closed")

    section("Known bottlenecks in window_layout.py")
    bottlenecks = [
        ("Edge tab assignment",
         "Tabs are assigned to windows by title heuristic. With 2+ Edge windows "
         "the mapping can be wrong. CDP windowId helps but requires debug mode."),
        ("Edge debug profile",
         "edge-debug creates a SEPARATE profile. You must run your browsing session "
         "inside it — your normal Edge profile is a different app instance."),
        ("Minimised windows",
         "Win32 GetWindowRect returns the last NORMAL rect for minimised windows. "
         "window_layout.py uses GetWindowPlacement to get the correct saved rect."),
        ("UWP / Store apps",
         "Many UWP windows are 'cloaked'. They enumerate but cannot be moved or "
         "their state may not be queryable via win32gui."),
        ("Multi-monitor DPI",
         "If monitors have different DPI scaling, rects can appear offset. "
         "Ensure the script is DPI-aware (SetProcessDPIAware)."),
        ("Window title changes",
         "If a window title changes between save and restore (e.g. browser tab "
         "changes, document name changes) the match score drops and it may be skipped."),
        ("Race condition on restore",
         "Windows launched by the restore path may not appear immediately. "
         "The launch_wait sleep may need tuning for slow apps."),
        ("Explorer windows",
         "Explorer paths rely on Shell.Application COM. If it's unavailable the "
         "folder path is lost and explorer restores to default location."),
    ]
    for (name, desc) in bottlenecks:
        print(f"\n  {Y}⚑  {W}{name}{RST}")
        for line in _wrap(desc, 72):
            print(f"     {DIM}{line}{RST}")


def _wrap(text: str, width: int) -> List[str]:
    words  = text.split()
    lines, cur = [], ""
    for w in words:
        if len(cur) + len(w) + 1 > width:
            if cur: lines.append(cur)
            cur = w
        else:
            cur = (cur + " " + w).strip()
    if cur: lines.append(cur)
    return lines


# ══════════════════════════════════════════════════════════════════════════
#  SECTION 8 – JSON snapshot output
# ══════════════════════════════════════════════════════════════════════════
def save_snapshot(captured: List[Dict], path: str) -> None:
    snapshot = {
        "timestamp":  datetime.now().isoformat(),
        "host":       socket.gethostname(),
        "windows":    captured,
    }
    with open(path, "w", encoding="utf-8") as f:
        json.dump(snapshot, f, indent=2)
    ok(f"Snapshot saved → {path}")


# ══════════════════════════════════════════════════════════════════════════
#  SECTION 9 – Before/After comparison
# ══════════════════════════════════════════════════════════════════════════
def compare_snapshots(path_a: str, path_b: str) -> None:
    banner(f"COMPARE  {os.path.basename(path_a)}  vs  {os.path.basename(path_b)}")

    def load(p):
        with open(p, encoding="utf-8") as f:
            return json.load(f)

    try:
        a = load(path_a)
        b = load(path_b)
    except Exception as e:
        err(f"Could not load files: {e}")
        return

    wins_a = {w["title"]: w for w in a.get("windows", [])}
    wins_b = {w["title"]: w for w in b.get("windows", [])}

    added   = set(wins_b) - set(wins_a)
    removed = set(wins_a) - set(wins_b)
    common  = set(wins_a) & set(wins_b)

    section(f"Added windows ({len(added)})")
    for t in sorted(added):
        ok(f"  + {t[:70]}")

    section(f"Removed windows ({len(removed)})")
    for t in sorted(removed):
        err(f"  - {t[:70]}")

    section(f"Changed windows ({len(common)} compared)")
    for title in sorted(common):
        wa = wins_a[title]
        wb = wins_b[title]
        diffs = []
        for field in ("rect", "norm_rect", "show_cmd", "iconic"):
            if wa.get(field) != wb.get(field):
                diffs.append(f"{field}: {wa.get(field)} → {wb.get(field)}")
        if diffs:
            print(f"\n  {Y}{title[:60]}{RST}")
            for d in diffs:
                print(f"    {d}")


# ══════════════════════════════════════════════════════════════════════════
#  GUIDED WALKTHROUGH
# ══════════════════════════════════════════════════════════════════════════
def guided_walkthrough(edge_port: int = 9222) -> None:
    banner("WINDOW LAYOUT TESTBENCH  –  Guided First-Run Walkthrough", "═", M)

    print(f"""
{W}This tool will walk you through everything needed to understand, debug,
and get the most out of window_layout.py.{RST}

{Y}What we'll do:{RST}
  1.  Check your environment and displays
  2.  Dump ALL running windows (with filter analysis)
  3.  List relevant running processes
  4.  Check Edge / browser tab capture
  5.  Check Explorer/File windows
  6.  Simulate save→restore matching
  7.  Audit for known bottlenecks
  8.  Save a debug snapshot JSON

{R}TIP:{RST} Run this BEFORE you save a layout, then AFTER restoring, to compare.
""")

    pause("Ready? Press ENTER to start Step 1")

    # ─ Step 1 ─────────────────────────────────────────────────────────────
    dump_environment()
    pause("Step 1 done. Press ENTER for Step 2 (window enumeration)")

    # ─ Step 2 ─────────────────────────────────────────────────────────────
    captured = dump_all_windows(verbose=True)
    pause(f"Step 2 done ({len(captured)} captured). Press ENTER for Step 3")

    # ─ Step 3 ─────────────────────────────────────────────────────────────
    dump_processes()
    pause("Step 3 done. Press ENTER for Step 4 (Edge tabs)")

    # ─ Step 4 ─────────────────────────────────────────────────────────────
    print(f"""
{Y}Step 4 – Edge tab capture{RST}

For tab capture to work Edge must be running with remote debugging enabled.
If you haven't done this yet, here's how:

  {G}python window_layout.py edge-debug --port {edge_port}{RST}

Then open your desired tabs INSIDE that Edge window before capturing.
""")
    dump_edge_tabs(port=edge_port)
    pause("Step 4 done. Press ENTER for Step 5 (Explorer windows)")

    # ─ Step 5 ─────────────────────────────────────────────────────────────
    dump_explorer_windows()
    pause("Step 5 done. Press ENTER for Step 6 (match simulation)")

    # ─ Step 6 ─────────────────────────────────────────────────────────────
    dump_match_simulation(captured)
    pause("Step 6 done. Press ENTER for Step 7 (compatibility audit)")

    # ─ Step 7 ─────────────────────────────────────────────────────────────
    dump_compatibility_audit()
    pause("Step 7 done. Press ENTER to save debug snapshot")

    # ─ Step 8 ─────────────────────────────────────────────────────────────
    banner("8 · SAVE DEBUG SNAPSHOT")
    snap_path = f"testbench_snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    save_snapshot(captured, snap_path)

    banner("ALL DONE", "─", G)
    print(f"""
{G}Next steps:{RST}

  {W}Save a layout:{RST}
    python window_layout.py save my_layout.json

  {W}Save WITH Edge tabs (requires edge-debug session):{RST}
    python window_layout.py save my_layout.json --edge-tabs

  {W}Restore:{RST}
    python window_layout.py restore my_layout.json --mode smart --diagnostics

  {W}Compare two snapshots from this tool:{RST}
    python testbench_window_layout.py --compare snap_before.json snap_after.json

  {W}Re-run this testbench (one-shot, no prompts):{RST}
    python testbench_window_layout.py --snapshot

{Y}Debug tip:{RST} run with --snapshot BEFORE and AFTER a restore to diff what changed.
""")


# ══════════════════════════════════════════════════════════════════════════
#  ENTRY POINT
# ══════════════════════════════════════════════════════════════════════════
def main() -> None:
    # Enable ANSI on Windows
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleMode(
                ctypes.windll.kernel32.GetStdHandle(-11), 7
            )
        except Exception:
            pass

    parser = argparse.ArgumentParser(
        description="Window Layout Testbench – verbose diagnostics and guided walkthrough"
    )
    parser.add_argument("--snapshot",   action="store_true",
                        help="One-shot dump of all sections, no prompts")
    parser.add_argument("--edge-only",  action="store_true",
                        help="Only dump Edge / browser tab info")
    parser.add_argument("--edge-port",  type=int, default=9222,
                        help="Edge remote debugging port (default 9222)")
    parser.add_argument("--compare",    nargs=2, metavar=("BEFORE", "AFTER"),
                        help="Compare two snapshot JSON files")
    parser.add_argument("--save-snap",  metavar="PATH",
                        help="Save a debug snapshot JSON to PATH")
    args = parser.parse_args()

    if not _WIN32_OK:
        print(f"ERROR: Required modules not available: {_WIN32_MISSING}")
        print("Install with:  pip install pywin32 psutil")
        sys.exit(1)

    if args.compare:
        compare_snapshots(args.compare[0], args.compare[1])
        return

    if args.edge_only:
        dump_edge_tabs(port=args.edge_port)
        return

    if args.snapshot:
        dump_environment()
        captured = dump_all_windows(verbose=True)
        dump_processes()
        dump_edge_tabs(port=args.edge_port)
        dump_explorer_windows()
        dump_match_simulation(captured)
        dump_compatibility_audit()
        snap = args.save_snap or f"testbench_snapshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
        save_snapshot(captured, snap)
        return

    # Default: full guided walkthrough
    guided_walkthrough(edge_port=args.edge_port)


if __name__ == "__main__":
    main()