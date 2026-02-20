import importlib
import json
import pathlib
import sys
import types


def _load_module(monkeypatch):
    repo_root = pathlib.Path(__file__).resolve().parents[1]
    if str(repo_root) not in sys.path:
        sys.path.insert(0, str(repo_root))

    fake_win32con = types.SimpleNamespace(
        SW_SHOWNORMAL=1,
        SW_SHOWMAXIMIZED=3,
        SW_SHOWMINIMIZED=2,
        SW_RESTORE=9,
        GWL_EXSTYLE=-20,
        GW_OWNER=4,
        WS_EX_TOOLWINDOW=0x80,
        WS_EX_APPWINDOW=0x40000,
    )
    fake_win32gui = types.SimpleNamespace(
        GetWindowLong=lambda *_: 0,
        GetWindow=lambda *_: 0,
        GetWindowRect=lambda *_: (0, 0, 100, 100),
    )

    monkeypatch.setitem(sys.modules, "win32con", fake_win32con)
    monkeypatch.setitem(sys.modules, "win32gui", fake_win32gui)
    monkeypatch.setitem(sys.modules, "win32api", types.SimpleNamespace())
    monkeypatch.setitem(sys.modules, "win32process", types.SimpleNamespace())
    monkeypatch.setitem(sys.modules, "psutil", types.SimpleNamespace(Process=lambda *_: None))

    win32com_module = types.ModuleType("win32com")
    client_module = types.ModuleType("win32com.client")
    client_module.Dispatch = lambda *_: None
    win32com_module.client = client_module
    monkeypatch.setitem(sys.modules, "win32com", win32com_module)
    monkeypatch.setitem(sys.modules, "win32com.client", client_module)

    sys.modules.pop("window_layout", None)
    return importlib.import_module("window_layout")


def test_assign_edge_tabs_keeps_tabs_grouped_by_window_id(monkeypatch):
    wl = _load_module(monkeypatch)
    windows = [
        {"process_name": "msedge.exe", "title": "Dashboard - Work - Microsoft Edge"},
        {"process_name": "msedge.exe", "title": "My Profile - Work - Microsoft Edge"},
    ]
    tabs = [
        {"title": "Dashboard", "url": "https://sn.example", "window_id": 55},
        {"title": "Incident", "url": "https://inc.example", "window_id": 55},
        {"title": "My Profile", "url": "https://profile.example", "window_id": 77},
    ]

    wl._assign_edge_tabs_to_windows(windows, tabs)

    assert [t["url"] for t in windows[0]["edge_tabs"]] == ["https://sn.example", "https://inc.example"]
    assert [t["url"] for t in windows[1]["edge_tabs"]] == ["https://profile.example"]


def test_save_layout_edge_tabs_only_persists_per_window(monkeypatch, tmp_path):
    wl = _load_module(monkeypatch)
    output_path = tmp_path / "layout.json"

    monkeypatch.setattr(wl, "capture_windows", lambda: [
        {"process_name": "msedge.exe", "title": "Edge A", "window_id": "w1"},
        {"process_name": "notepad.exe", "title": "Notes", "window_id": "n1"},
    ])
    monkeypatch.setattr(wl, "_fetch_edge_tabs", lambda _port: [
        {"title": "A", "url": "https://a.example", "window_id": 1},
    ])

    wl.save_layout(str(output_path), capture_edge_tabs=True, edge_debug_port=9222)

    saved = json.loads(output_path.read_text(encoding="utf-8"))
    assert "browser_tabs" not in saved
    assert "edge_sessions" not in saved
    assert "open_urls" not in saved
    edge_windows = [w for w in saved["windows"] if w["process_name"].lower() == "msedge.exe"]
    assert edge_windows[0]["edge_tabs"][0]["url"] == "https://a.example"


def test_restore_edge_tabs_preserves_per_window_mapping(monkeypatch, tmp_path):
    wl = _load_module(monkeypatch)
    layout = {
        "schema": "window-layout.v2",
        "windows": [
            {
                "process_name": "msedge.exe",
                "title": "Edge A",
                "normal_rect": [0, 0, 400, 400],
                "rect": [0, 0, 400, 400],
                "show_cmd": 1,
                "edge_tabs": [{"title": "A", "url": "https://a.example"}],
            },
            {
                "process_name": "msedge.exe",
                "title": "Edge B",
                "normal_rect": [0, 0, 100, 100],
                "rect": [0, 0, 100, 100],
                "show_cmd": 1,
                "edge_tabs": [{"title": "B", "url": "https://b.example"}],
            },
        ],
    }
    file_path = tmp_path / "layout.json"
    file_path.write_text(json.dumps(layout), encoding="utf-8")

    monkeypatch.setattr(wl, "_current_windows_with_hwnds", lambda: [])
    monkeypatch.setattr(wl, "_edge_exe_from_targets", lambda _targets: "C:/Edge/msedge.exe")

    launched = []

    def _launch(exe, tabs, dry_run=False, base_args=None):
        launched.append([t["url"] for t in tabs])
        return len(tabs)

    monkeypatch.setattr(wl, "_launch_edge_tabs", _launch)

    wl.restore_layout(str(file_path), mode="smart")

    assert launched == [["https://a.example"], ["https://b.example"]]


def test_main_edit_dispatch(monkeypatch):
    wl = _load_module(monkeypatch)
    called = {"path": None}
    monkeypatch.setattr(sys, "argv", ["window_layout.py", "edit", "layout.json"])
    monkeypatch.setattr(wl, "run_edit_wizard", lambda path: called.__setitem__("path", path))

    wl.main()

    assert called["path"] == "layout.json"


def test_stabilize_edge_window_sizes_reapplies_mismatched_size(monkeypatch):
    wl = _load_module(monkeypatch)
    target = {"rect": [0, 0, 500, 400], "normal_rect": [0, 0, 500, 400]}

    rect_calls = iter([(0, 0, 300, 200), (0, 0, 500, 400)])
    monkeypatch.setattr(wl.win32gui, "GetWindowRect", lambda _hwnd: next(rect_calls))

    applied = {"count": 0}

    def _apply(_hwnd, _target):
        applied["count"] += 1
        return True

    monkeypatch.setattr(wl, "_apply_window_position", _apply)
    fixes = wl._stabilize_edge_window_sizes([(1001, target)], retries=2, delay_s=0)

    assert fixes == 1
    assert applied["count"] == 1


def test_match_diagnostics_output_for_edge_deweighted_title(monkeypatch, capsys):
    wl = _load_module(monkeypatch)
    target = {
        "process_name": "msedge.exe",
        "title": "Some New Tab Title",
        "class_name": "Chrome_WidgetWin_1",
        "exe": "C:/Edge/msedge.exe",
        "rect": [0, 0, 1200, 900],
    }
    current = [
        {
            "hwnd": 100,
            "process_name": "msedge.exe",
            "title": "Completely Different",
            "class_name": "Chrome_WidgetWin_1",
            "exe": "C:/Edge/msedge.exe",
            "rect": [10, 10, 1210, 910],
        }
    ]

    ranked = wl._match_candidates_with_scores(target, current, used_hwnds=set())
    wl._print_match_diagnostics(target, ranked, top_n=1)
    out = capsys.readouterr().out

    assert "[DIAG] Target: proc=msedge.exe title=Some New Tab Title" in out
    assert "edge_title_deweighted=true" in out
    assert "title=0" in out


def test_restore_layout_diagnostics_prints_candidates(monkeypatch, tmp_path, capsys):
    wl = _load_module(monkeypatch)
    layout = {
        "schema": "window-layout.v2",
        "windows": [
            {
                "process_name": "notepad.exe",
                "title": "notes",
                "class_name": "Notepad",
                "exe": "C:/Windows/notepad.exe",
                "normal_rect": [0, 0, 200, 200],
                "rect": [0, 0, 200, 200],
                "show_cmd": 1,
            }
        ],
    }
    file_path = tmp_path / "layout.json"
    file_path.write_text(json.dumps(layout), encoding="utf-8")

    monkeypatch.setattr(wl, "_current_windows_with_hwnds", lambda: [
        {
            "hwnd": 77,
            "process_name": "notepad.exe",
            "title": "notes",
            "class_name": "Notepad",
            "exe": "C:/Windows/notepad.exe",
            "normal_rect": [0, 0, 200, 200],
            "rect": [0, 0, 200, 200],
        }
    ])
    monkeypatch.setattr(wl, "_apply_window_position", lambda _hwnd, _entry: True)
    monkeypatch.setattr(wl, "_stabilize_edge_window_sizes", lambda _matches: 0)

    wl.restore_layout(str(file_path), mode="basic", diagnostics=True, diagnostics_top_n=1)
    out = capsys.readouterr().out

    assert "[DIAG] Target: proc=notepad.exe title=notes" in out
    assert "[DIAG]   #1 hwnd=77" in out


def test_assign_edge_tabs_prefers_cdp_window_hint_when_titles_change(monkeypatch):
    wl = _load_module(monkeypatch)
    windows = [
        {
            "process_name": "msedge.exe",
            "title": "Current Different Title A - Microsoft Edge",
            "edge": {"cdp_window_hint": 101},
        },
        {
            "process_name": "msedge.exe",
            "title": "Current Different Title B - Microsoft Edge",
            "edge": {"cdp_window_hint": 202},
        },
    ]
    tabs = [
        {"title": "Old title no longer matching", "url": "https://a.example", "window_id": 202},
        {"title": "Another stale title", "url": "https://b.example", "window_id": 101},
    ]

    wl._assign_edge_tabs_to_windows(windows, tabs)

    assert [t["url"] for t in windows[0]["edge_tabs"]] == ["https://b.example"]
    assert [t["url"] for t in windows[1]["edge_tabs"]] == ["https://a.example"]
    assert windows[0]["edge"]["cdp_window_hint"] == 101
    assert windows[1]["edge"]["cdp_window_hint"] == 202


def test_smart_restore_uses_existing_edge_window_for_tabs(monkeypatch, tmp_path):
    wl = _load_module(monkeypatch)
    layout = {
        "schema": "window-layout.v2",
        "windows": [
            {
                "process_name": "msedge.exe",
                "title": "Target Edge",
                "class_name": "Chrome_WidgetWin_1",
                "exe": "C:/Edge/msedge.exe",
                "normal_rect": [0, 0, 450, 450],
                "rect": [0, 0, 450, 450],
                "show_cmd": 1,
                "edge_tabs": [{"title": "A", "url": "https://a.example"}],
            }
        ],
    }
    file_path = tmp_path / "layout.json"
    file_path.write_text(json.dumps(layout), encoding="utf-8")

    calls = {"existing": 0, "new": 0}

    current_windows = [
        {
            "hwnd": 1001,
            "process_name": "msedge.exe",
            "title": "Different Title",
            "class_name": "Chrome_WidgetWin_1",
            "exe": "C:/Edge/msedge.exe",
            "normal_rect": [0, 0, 100, 100],
            "rect": [0, 0, 100, 100],
        }
    ]
    monkeypatch.setattr(wl, "_current_windows_with_hwnds", lambda: list(current_windows))
    monkeypatch.setattr(wl, "_apply_window_position", lambda _hwnd, _entry: True)
    monkeypatch.setattr(wl, "_stabilize_edge_window_sizes", lambda _matches: 0)
    monkeypatch.setattr(wl, "_edge_exe_from_targets", lambda _targets: "C:/Edge/msedge.exe")

    def _launch_existing(_exe, _tabs, dry_run=False, base_args=None):
        calls["existing"] += 1
        return 1

    def _launch_new(_exe, _tabs, dry_run=False, base_args=None):
        calls["new"] += 1
        return 1

    monkeypatch.setattr(wl, "_launch_edge_tabs_existing", _launch_existing)
    monkeypatch.setattr(wl, "_launch_edge_tabs", _launch_new)

    wl.restore_layout(str(file_path), mode="smart")

    assert calls["existing"] == 1
    assert calls["new"] == 0


def test_smart_restore_skips_edge_tab_relaunch_when_windows_already_in_place(monkeypatch, tmp_path):
    wl = _load_module(monkeypatch)
    layout = {
        "schema": "window-layout.v2",
        "windows": [
            {
                "process_name": "msedge.exe",
                "title": "Edge One",
                "class_name": "Chrome_WidgetWin_1",
                "exe": "C:/Edge/msedge.exe",
                "normal_rect": [0, 0, 300, 300],
                "rect": [0, 0, 300, 300],
                "show_cmd": 1,
                "edge_tabs": [{"title": "A", "url": "https://a.example"}],
            },
            {
                "process_name": "msedge.exe",
                "title": "Edge Two",
                "class_name": "Chrome_WidgetWin_1",
                "exe": "C:/Edge/msedge.exe",
                "normal_rect": [400, 0, 700, 300],
                "rect": [400, 0, 700, 300],
                "show_cmd": 1,
                "edge_tabs": [{"title": "B", "url": "https://b.example"}],
            },
        ],
    }
    file_path = tmp_path / "layout.json"
    file_path.write_text(json.dumps(layout), encoding="utf-8")

    running = [
        {"hwnd": 2001, "process_name": "msedge.exe", "title": "Edge One", "class_name": "Chrome_WidgetWin_1", "exe": "C:/Edge/msedge.exe", "rect": [0, 0, 300, 300], "normal_rect": [0, 0, 300, 300]},
        {"hwnd": 2002, "process_name": "msedge.exe", "title": "Edge Two", "class_name": "Chrome_WidgetWin_1", "exe": "C:/Edge/msedge.exe", "rect": [400, 0, 700, 300], "normal_rect": [400, 0, 700, 300]},
    ]
    monkeypatch.setattr(wl, "_current_windows_with_hwnds", lambda: list(running))
    monkeypatch.setattr(wl, "_apply_window_position", lambda _hwnd, _entry: True)
    monkeypatch.setattr(wl, "_stabilize_edge_window_sizes", lambda _matches: 0)
    monkeypatch.setattr(wl, "_edge_exe_from_targets", lambda _targets: "C:/Edge/msedge.exe")

    calls = {"existing": 0, "new": 0}
    monkeypatch.setattr(wl, "_launch_edge_tabs_existing", lambda *_args, **_kwargs: calls.__setitem__("existing", calls["existing"] + 1) or 1)
    monkeypatch.setattr(wl, "_launch_edge_tabs", lambda *_args, **_kwargs: calls.__setitem__("new", calls["new"] + 1) or 1)

    wl.restore_layout(str(file_path), mode="smart")

    assert calls["existing"] == 0
    assert calls["new"] == 0
