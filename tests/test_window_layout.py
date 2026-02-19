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
        GWL_EXSTYLE=-20,
        GW_OWNER=4,
        WS_EX_TOOLWINDOW=0x80,
        WS_EX_APPWINDOW=0x40000,
    )
    fake_win32gui = types.SimpleNamespace(
        GetWindowLong=lambda *_: 0,
        GetWindow=lambda *_: 0,
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


def test_assign_edge_tabs_to_windows(monkeypatch):
    wl = _load_module(monkeypatch)
    windows = [
        {"process_name": "msedge.exe", "title": "My Profile - Work - Microsoft Edge"},
        {"process_name": "msedge.exe", "title": "NiCE CXone Mpower - Work - Microsoft Edge"},
    ]
    tabs = [
        {"title": "My Profile", "url": "https://a.example", "window_id": 1},
        {"title": "NiCE CXone Mpower", "url": "https://b.example", "window_id": 2},
    ]

    wl._assign_edge_tabs_to_windows(windows, tabs)

    assert windows[0]["edge_tabs"][0]["url"] == "https://a.example"
    assert windows[1]["edge_tabs"][0]["url"] == "https://b.example"




def test_assign_edge_tabs_keeps_tabs_grouped_by_window_id(monkeypatch):
    wl = _load_module(monkeypatch)
    windows = [
        {"process_name": "msedge.exe", "title": "TSD Dashboard V2 | ServiceNow - Work - Microsoft Edge"},
        {"process_name": "msedge.exe", "title": "My Profile - Work - Microsoft Edge"},
    ]
    tabs = [
        {"title": "TSD Dashboard V2 | ServiceNow", "url": "https://sn.example", "window_id": 55},
        {"title": "Incident INC123", "url": "https://inc.example", "window_id": 55},
        {"title": "My Profile", "url": "https://profile.example", "window_id": 77},
    ]

    wl._assign_edge_tabs_to_windows(windows, tabs)

    dashboard_urls = [t["url"] for t in windows[0]["edge_tabs"]]
    profile_urls = [t["url"] for t in windows[1]["edge_tabs"]]

    assert dashboard_urls == ["https://sn.example", "https://inc.example"]
    assert profile_urls == ["https://profile.example"]

def test_collect_edge_tabs_prefers_per_window(monkeypatch):
    wl = _load_module(monkeypatch)
    data = {
        "windows": [{"process_name": "msedge.exe", "edge_tabs": [{"title": "A", "url": "https://a"}]}],
        "browser_tabs": {"edge": {"tabs": [{"title": "B", "url": "https://b"}]}}
    }

    assert wl._collect_edge_tabs(data) == [{"title": "A", "url": "https://a"}]


def test_collect_edge_tabs_fallback_legacy(monkeypatch):
    wl = _load_module(monkeypatch)
    data = {
        "windows": [{"process_name": "msedge.exe"}],
        "browser_tabs": {"edge": {"tabs": [{"title": "B", "url": "https://b"}]}}
    }

    assert wl._collect_edge_tabs(data) == [{"title": "B", "url": "https://b"}]


def test_run_edit_wizard_updates_assignments(monkeypatch, tmp_path):
    wl = _load_module(monkeypatch)
    payload = {
        "schema": "window-layout.v1",
        "windows": [
            {"process_name": "msedge.exe", "title": "Window 1"},
            {"process_name": "msedge.exe", "title": "Window 2"},
        ],
        "browser_tabs": {"edge": {"tabs": [
            {"title": "Tab 1", "url": "https://1"},
            {"title": "Tab 2", "url": "https://2"},
        ]}},
    }
    file_path = tmp_path / "layout.json"
    file_path.write_text(json.dumps(payload), encoding="utf-8")

    answers = iter(["1", "2"])
    monkeypatch.setattr(wl, "_prompt", lambda *_args, **_kwargs: next(answers))

    wl.run_edit_wizard(str(file_path))

    updated = json.loads(file_path.read_text(encoding="utf-8"))
    assert updated["windows"][0]["edge_tabs"][0]["url"] == "https://1"
    assert updated["windows"][1]["edge_tabs"][0]["url"] == "https://2"


def test_is_taskbar_window_rejects_toolwindow(monkeypatch):
    wl = _load_module(monkeypatch)
    monkeypatch.setattr(wl.win32gui, "GetWindowLong", lambda *_: wl.win32con.WS_EX_TOOLWINDOW)
    monkeypatch.setattr(wl.win32gui, "GetWindow", lambda *_: 0)

    assert wl._is_taskbar_window(101) is False


def test_main_edit_dispatch(monkeypatch):
    wl = _load_module(monkeypatch)
    called = {"path": None}
    monkeypatch.setattr(sys, "argv", ["window_layout.py", "edit", "layout.json"])
    monkeypatch.setattr(wl, "run_edit_wizard", lambda path: called.__setitem__("path", path))

    wl.main()

    assert called["path"] == "layout.json"


def test_restore_uses_collect_edge_tabs(monkeypatch, tmp_path):
    wl = _load_module(monkeypatch)
    payload = {
        "schema": "window-layout.v1",
        "windows": [{"process_name": "msedge.exe", "title": "Window", "normal_rect": [0, 0, 200, 200], "rect": [0, 0, 200, 200], "show_cmd": 1}],
        "browser_tabs": {"edge": {"tabs": [{"title": "B", "url": "https://b"}]}}
    }
    file_path = tmp_path / "layout.json"
    file_path.write_text(json.dumps(payload), encoding="utf-8")

    monkeypatch.setattr(wl, "_current_windows_with_hwnds", lambda: [])
    monkeypatch.setattr(wl, "_edge_exe_from_targets", lambda _targets: "C:/Edge/msedge.exe")
    launched = {"count": 0, "tabs": []}
    monkeypatch.setattr(wl.os.path, "exists", lambda _p: True)

    def _launch(exe, tabs, dry_run=False):
        launched["count"] += 1
        launched["tabs"] = tabs
        return len(tabs)

    monkeypatch.setattr(wl, "_launch_edge_tabs", _launch)

    wl.restore_layout(str(file_path), dry_run=False, launch_missing=False, restore_edge_tabs=True)

    assert launched["count"] >= 1
    assert launched["tabs"][0]["url"] == "https://b"


def test_score_match_prefers_fingerprint_when_title_changes(monkeypatch):
    wl = _load_module(monkeypatch)
    target = {
        "title": "Old title - Microsoft Edge",
        "process_name": "msedge.exe",
        "class_name": "Chrome_WidgetWin_1",
        "exe": r"C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
        "rect": [0, 0, 1300, 900],
    }
    candidate = {
        "title": "Completely different title - Microsoft Edge",
        "process_name": "msedge.exe",
        "class_name": "Chrome_WidgetWin_1",
        "exe": r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "rect": [20, 20, 1320, 920],
    }
    target["match_fingerprint"] = wl._build_match_fingerprint(target)
    candidate["match_fingerprint"] = wl._build_match_fingerprint(candidate)

    score = wl._score_match(candidate, target)

    assert target["match_fingerprint"] == candidate["match_fingerprint"]
    assert score >= 160


def test_score_match_falls_back_without_fingerprint(monkeypatch):
    wl = _load_module(monkeypatch)
    target = {
        "title": "Visual Studio Code",
        "process_name": "Code.exe",
        "class_name": "Chrome_WidgetWin_1",
        "exe": r"C:\Users\me\AppData\Local\Programs\Microsoft VS Code\Code.exe",
    }
    candidate = {
        "title": "Visual Studio Code - workspace",
        "process_name": "code.exe",
        "class_name": "Chrome_WidgetWin_1",
        "exe": r"c:/users/me/AppData/Local/Programs/Microsoft VS Code/Code.exe",
    }

    score = wl._score_match(candidate, target)

    assert score == 100
