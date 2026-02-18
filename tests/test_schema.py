import pytest

pytest.importorskip("win32gui")

import window_layout as wl


def test_migrate_v1_to_v2_edge_session():
    v1 = {
        "schema": "window-layout.v1",
        "created_at": "2026-02-18 10:00:00",
        "windows": [
            {
                "title": "Edge",
                "class_name": "Chrome_WidgetWin_1",
                "pid": 123,
                "process_name": "msedge.exe",
                "exe": "C:/Edge/msedge.exe",
                "is_visible": True,
                "is_minimized": False,
                "is_maximized": False,
                "rect": [0, 0, 100, 100],
                "normal_rect": [0, 0, 100, 100],
                "show_cmd": 1,
                "edge_tabs": [{"title": "Tab", "url": "https://example.com"}],
            }
        ],
        "browser_tabs": {
            "edge": {
                "debug_port": 9222,
                "captured_at": "2026-02-18 10:00:00",
                "tabs": [{"title": "Tab", "url": "https://example.com"}],
            }
        },
    }

    v2 = wl._migrate_v1_to_v2(v1)
    assert v2["schema"] == "window-layout.v2"
    assert v2.get("edge_sessions")
    session = v2["edge_sessions"][0]
    assert session["port"] == 9222
    assert session["tabs"][0]["url"] == "https://example.com"
    assert v2["windows"][0].get("window_id")


def test_collect_edge_tabs_by_session_open_urls():
    data = {
        "schema": "window-layout.v2",
        "windows": [],
        "edge_sessions": [],
        "open_urls": {"edge": ["https://example.com", "https://openai.com"]},
    }
    sessions = wl._collect_edge_tabs_by_session(data)
    assert len(sessions) == 1
    assert len(sessions[0]["tabs"]) == 2


def test_collect_edge_tabs_by_session_from_sessions():
    data = {
        "schema": "window-layout.v2",
        "windows": [],
        "edge_sessions": [
            {"port": 9222, "profile_dir": "", "tabs": [{"url": "https://example.com"}]}
        ],
    }
    sessions = wl._collect_edge_tabs_by_session(data)
    assert len(sessions) == 1
    assert sessions[0]["port"] == 9222
    assert sessions[0]["tabs"][0]["url"] == "https://example.com"
