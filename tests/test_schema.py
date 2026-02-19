import pytest

pytest.importorskip("win32gui")

import window_layout as wl


def test_migrate_v1_legacy_tabs_to_per_window_edge_tabs():
    v1 = {
        "schema": "window-layout.v1",
        "windows": [
            {
                "window_id": "w-edge",
                "title": "Edge",
                "process_name": "msedge.exe",
                "edge_tabs": [],
            }
        ],
        "browser_tabs": {
            "edge": {
                "tabs": [{"title": "Legacy", "url": "https://legacy.example"}],
            }
        },
        "open_urls": {"edge": ["https://fallback.example"]},
    }

    v2 = wl._migrate_v1_to_v2(v1)

    assert v2["schema"] == "window-layout.v2"
    assert "browser_tabs" not in v2
    assert "edge_sessions" not in v2
    assert "open_urls" not in v2
    assert [t["url"] for t in v2["windows"][0]["edge_tabs"]] == [
        "https://legacy.example",
        "https://fallback.example",
    ]


def test_migrate_v2_mixed_legacy_payload_drops_legacy_keys_and_maps_tabs():
    data = {
        "schema": "window-layout.v2",
        "windows": [
            {"window_id": "w1", "title": "Edge A", "process_name": "msedge.exe", "edge_tabs": []},
            {"window_id": "w2", "title": "Edge B", "process_name": "msedge.exe", "edge_tabs": []},
        ],
        "edge_sessions": [
            {
                "window_ids": ["w2"],
                "tabs": [{"title": "B", "url": "https://b.example"}],
            }
        ],
        "browser_tabs": {
            "edge": {
                "tabs": [{"title": "A", "url": "https://a.example"}],
            }
        },
        "open_urls": {"edge": ["https://c.example"]},
    }

    migrated = wl._ensure_v2_layout(data)

    assert "browser_tabs" not in migrated
    assert "edge_sessions" not in migrated
    assert "open_urls" not in migrated
    assert [t["url"] for t in migrated["windows"][1]["edge_tabs"]] == ["https://b.example"]
    assert [t["url"] for t in migrated["windows"][0]["edge_tabs"]] == [
        "https://a.example",
        "https://c.example",
    ]
