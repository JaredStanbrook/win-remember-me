import pathlib
import sys

repo_root = pathlib.Path(__file__).resolve().parents[1]
if str(repo_root) not in sys.path:
    sys.path.insert(0, str(repo_root))

import gui_app


def test_build_cli_command_save_edge():
    cmd = gui_app.build_cli_command("save_edge", "layout.json")
    assert cmd.label == "Save Layout + Edge Tabs"
    assert cmd.args[-3:] == ["save", "layout.json", "--edge-tabs"]


def test_build_cli_command_restore_basic():
    cmd = gui_app.build_cli_command("restore_basic", "my-layout.json")
    assert cmd.args[1:] == ["window_layout.py", "restore", "my-layout.json", "--mode", "basic"]


def test_format_command_for_log_quotes_spaces():
    logged = gui_app.format_command_for_log(["python", "window_layout.py", "save", "my layout.json"])
    assert "my layout.json" in logged


def test_build_cli_command_unknown_action():
    try:
        gui_app.build_cli_command("nope", "layout.json")
    except ValueError as exc:
        assert "Unknown GUI action" in str(exc)
        return
    raise AssertionError("Expected ValueError for unknown action")
