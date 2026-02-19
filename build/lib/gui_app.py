"""Lightweight GUI for Window Layout CLI (PySide6)."""

from __future__ import annotations

import json
import os
import shlex
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import List, Optional

import window_layout as wl

DEFAULT_LAYOUT_PATH = "layout.json"
LAYOUTS_ROOT = "layouts"
CONFIG_PATH = "config.json"
MIN_SETTINGS_SIZE = (700, 700)
MIN_SPEED_SIZE = (420, 320)


@dataclass
class GuiCommand:
    label: str
    args: List[str]


@dataclass
class SpeedMenuItem:
    label: str
    emoji: str
    layout: str
    args: List[str]


def build_cli_command(action: str, layout_path: str, edge_port: Optional[int] = None, edge_profile_dir: str = "") -> GuiCommand:
    base = [sys.executable, "window_layout.py"]
    if action == "save":
        return GuiCommand("Save Layout", base + ["save", layout_path])
    if action == "save_edge":
        cmd = base + ["save", layout_path, "--edge-tabs"]
        if edge_port:
            cmd.extend(["--edge-debug-port", str(edge_port)])
        if edge_profile_dir:
            cmd.extend(["--edge-profile-dir", edge_profile_dir])
        return GuiCommand("Save Layout + Edge Tabs", cmd)
    if action == "restore":
        return GuiCommand("Restore Layout", base + ["restore", layout_path])
    if action == "restore_smart":
        return GuiCommand("Smart Restore", base + ["restore", layout_path, "--smart", "--restore-edge-tabs"])
    if action == "restore_dry":
        return GuiCommand("Restore (Dry Run)", base + ["restore", layout_path, "--dry-run"])
    if action == "restore_missing":
        return GuiCommand("Restore + Launch Missing", base + ["restore", layout_path, "--launch-missing"])
    if action == "restore_edge":
        return GuiCommand("Restore + Edge Tabs", base + ["restore", layout_path, "--restore-edge-tabs"])
    if action == "restore_simple":
        return GuiCommand("Restore + Open URLs", base + ["restore", layout_path, "--restore-edge-tabs"])
    if action == "edit":
        return GuiCommand("Edit Edge Tab Mapping", base + ["edit", layout_path])
    if action == "edge_debug":
        cmd = base + ["edge-debug"]
        if edge_port:
            cmd.extend(["--port", str(edge_port)])
        if edge_profile_dir:
            cmd.extend(["--profile-dir", edge_profile_dir])
        return GuiCommand("Edge Debug Session", cmd)
    if action == "edge_capture":
        cmd = base + ["edge-capture", layout_path]
        if edge_port:
            cmd.extend(["--port", str(edge_port)])
        if edge_profile_dir:
            cmd.extend(["--profile-dir", edge_profile_dir])
        return GuiCommand("Edge Capture Tabs", cmd)
    raise ValueError(f"Unknown GUI action: {action}")


def format_command_for_log(cmd: List[str]) -> str:
    return " ".join(shlex.quote(part) for part in cmd)


def _run_command_sync(cmd: List[str]) -> subprocess.CompletedProcess:
    return subprocess.run(cmd, capture_output=True, text=True)


def _load_json(path: str) -> Optional[dict]:
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return None


def _load_config() -> dict:
    data = _load_json(CONFIG_PATH)
    if isinstance(data, dict):
        return data
    return {}


def _get_edge_defaults() -> tuple[int, str]:
    data = _load_config()
    edge = data.get("edge") or {}
    try:
        port = int(edge.get("debug_port") or 9222)
    except (TypeError, ValueError):
        port = 9222
    profile_dir = str(edge.get("profile_dir") or "").strip()
    return port, profile_dir


def _save_edge_defaults(port: int, profile_dir: str) -> None:
    data = _load_config()
    if not isinstance(data, dict):
        data = {}
    data["edge"] = {
        "debug_port": int(port),
        "profile_dir": profile_dir,
    }
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _get_hotkeys_enabled() -> bool:
    data = _load_config()
    return bool(data.get("hotkeys_enabled", False))


def _set_hotkeys_enabled(enabled: bool) -> None:
    data = _load_config()
    if not isinstance(data, dict):
        data = {}
    data["hotkeys_enabled"] = bool(enabled)
    with open(CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)


def _get_layouts_root() -> str:
    data = _load_config()
    root = str(data.get("layouts_root") or "").strip() or LAYOUTS_ROOT
    if not os.path.isabs(root):
        root = os.path.abspath(root)
    return root


def _ensure_layouts_root() -> str:
    root = _get_layouts_root()
    os.makedirs(root, exist_ok=True)
    return root


def _parse_speed_menu(config_path: str) -> List[SpeedMenuItem]:
    data = _load_json(config_path)
    if not isinstance(data, dict):
        return []

    items: List[SpeedMenuItem] = []
    speed_menu = data.get("speed_menu") or {}
    if not isinstance(speed_menu, dict):
        return items
    buttons = speed_menu.get("buttons") or []
    if isinstance(buttons, list):
        for raw in buttons:
            if not isinstance(raw, dict):
                continue
            label = str(raw.get("label") or "").strip()
            emoji = str(raw.get("emoji") or "").strip()
            layout = str(raw.get("layout") or "").strip()
            args = raw.get("args") or []
            if isinstance(args, str):
                args = [args]
            elif not isinstance(args, list):
                args = []
            args = [str(a) for a in args if str(a).strip()]
            if not label and not layout:
                continue
            items.append(SpeedMenuItem(label=label, emoji=emoji, layout=layout, args=args))

    return items


def _resolve_speed_layout(target: str) -> str:
    if not target:
        return ""
    if os.path.isabs(target):
        return target

    candidates: List[str] = []
    layouts_root = _ensure_layouts_root()
    candidates.append(os.path.join(layouts_root, target))

    for candidate in candidates:
        if os.path.exists(candidate):
            return candidate
    return candidates[0] if candidates else target


def main() -> int:
    try:
        from PySide6.QtCore import QProcess, QEvent, Qt, QSignalBlocker, QObject, Signal
        from PySide6.QtGui import QColor, QFont, QPalette
        from PySide6.QtWidgets import (
            QApplication,
            QAbstractItemView,
            QFileDialog,
            QFrame,
            QGridLayout,
            QHBoxLayout,
            QLabel,
            QListWidget,
            QListWidgetItem,
            QComboBox,
            QInputDialog,
            QLineEdit,
            QMainWindow,
            QMessageBox,
            QPushButton,
            QPlainTextEdit,
            QCheckBox,
            QSpinBox,
            QScrollArea,
            QSizePolicy,
            QSystemTrayIcon,
            QMenu,
            QStyle,
            QTabWidget,
            QVBoxLayout,
            QWidget,
        )
    except ImportError:
        print("PySide6 is required for GUI mode. Install with: pip install PySide6")
        return 1

    def _apply_fluent_style(app: QApplication) -> None:
        app.setFont(QFont("Segoe UI", 11))
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#F7F8FA"))
        palette.setColor(QPalette.Base, QColor("#FFFFFF"))
        palette.setColor(QPalette.AlternateBase, QColor("#F3F5F7"))
        palette.setColor(QPalette.Text, QColor("#111111"))
        palette.setColor(QPalette.WindowText, QColor("#111111"))
        palette.setColor(QPalette.Button, QColor("#FFFFFF"))
        palette.setColor(QPalette.ButtonText, QColor("#111111"))
        palette.setColor(QPalette.Highlight, QColor("#2B7CD3"))
        palette.setColor(QPalette.HighlightedText, QColor("#FFFFFF"))
        app.setPalette(palette)

        app.setStyleSheet(
            """
            QWidget {
                background: #F7F8FA;
                color: #111111;
                font-size: 11pt;
            }
            QMainWindow {
                background: #F7F8FA;
            }
            QLineEdit, QPlainTextEdit, QTableWidget, QListWidget {
                background: #FFFFFF;
                border: 1px solid #D5D7DA;
                border-radius: 8px;
                padding: 6px 8px;
                selection-background-color: #2B7CD3;
            }
            QPlainTextEdit {
                padding: 8px;
            }
            QPushButton {
                background: #FFFFFF;
                border: 1px solid #D0D3D8;
                border-radius: 10px;
                padding: 7px 14px;
            }
            QPushButton:hover {
                background: #F1F4F8;
                border-color: #C5CBD3;
            }
            QPushButton:pressed {
                background: #E6EBF2;
                border-color: #BFC6D0;
            }
            QPushButton:disabled {
                background: #F4F5F7;
                color: #8A8F98;
                border-color: #E3E6EA;
            }
            QLabel {
                background: transparent;
            }
            QTabWidget::pane {
                border: 1px solid #E1E4E8;
                border-top-left-radius: 0px;
                border-top-right-radius: 10px;
                border-bottom-left-radius: 10px;
                border-bottom-right-radius: 10px;
                padding: 6px;
                background: #F7F8FA;
            }
            QTabBar::tab {
                background: #E9EDF3;
                border: 1px solid #D4D9E0;
                border-bottom: none;
                border-top-left-radius: 10px;
                border-top-right-radius: 10px;
                padding: 8px 16px;
                margin-right: 4px;
            }
            QTabBar::tab:selected {
                background: #FFFFFF;
                border-color: #D4D9E0;
            }
            QTableWidget::item:selected {
                background: #E7F0FB;
                color: #111111;
            }
            QListWidget::item:selected {
                background: #E7F0FB;
                color: #111111;
            }
            QHeaderView::section {
                background: #F1F3F6;
                border: 1px solid #D5D7DA;
                padding: 6px 8px;
                border-radius: 6px;
            }
            QScrollBar:vertical {
                background: transparent;
                width: 10px;
                margin: 4px;
            }
            QScrollBar::handle:vertical {
                background: #C7CDD6;
                border-radius: 5px;
                min-height: 30px;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: transparent;
            }
            """
        )

    class HotkeyEmitter(QObject):
        fired = Signal(str)

    class MainWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("Window Layout Manager")
            self.resize(860, 560)
            _ensure_layouts_root()

            self._proc = QProcess(self)
            self._proc.readyReadStandardOutput.connect(self._append_stdout)
            self._proc.readyReadStandardError.connect(self._append_stderr)
            self._proc.finished.connect(self._on_finished)
            self._speed_item_cache: dict[str, SpeedMenuItem] = {}

            root = QWidget(self)
            self.setCentralWidget(root)

            root_layout = QVBoxLayout(root)
            root_layout.setContentsMargins(16, 16, 16, 16)
            root_layout.setSpacing(14)
            tabs = QTabWidget()
            self._tabs = tabs
            tabs.currentChanged.connect(self._on_tab_changed)
            tabs.tabBarClicked.connect(self._on_tab_bar_clicked)
            root_layout.addWidget(tabs)

            settings_tab = QWidget()
            speed_tab = QWidget()
            editor_tab = QWidget()
            layout_editor_tab = QWidget()
            tabs.addTab(settings_tab, "Settings")
            tabs.addTab(speed_tab, "Speed Menu")
            tabs.addTab(editor_tab, "Speed Menu Editor")
            tabs.addTab(layout_editor_tab, "Layout Editor")

            settings_layout = QGridLayout(settings_tab)
            settings_layout.setHorizontalSpacing(12)
            settings_layout.setVerticalSpacing(12)
            settings_layout.setContentsMargins(14, 14, 14, 14)

            root_row = QHBoxLayout()
            self.layouts_root_input = QLineEdit()
            root_browse_btn = QPushButton("Browse...")
            root_browse_btn.clicked.connect(self._browse_layouts_root)
            self.layouts_root_input.editingFinished.connect(self._save_layouts_root)
            root_row.addWidget(QLabel("Layouts Root:"))
            root_row.addWidget(self.layouts_root_input, 1)
            root_row.addWidget(root_browse_btn)
            settings_layout.addLayout(root_row, 0, 0, 1, 2)

            hotkey_row = QHBoxLayout()
            self.hotkeys_enabled = QCheckBox("Enable Global Hotkeys")
            self.hotkeys_enabled.stateChanged.connect(self._toggle_hotkeys)
            hotkey_row.addWidget(self.hotkeys_enabled)
            hotkey_row.addStretch(1)
            settings_layout.addLayout(hotkey_row, 1, 0, 1, 2)

            path_row = QHBoxLayout()
            self.layout_select = QComboBox()
            self.layout_select.currentIndexChanged.connect(self._sync_layout_editor_choice)
            path_row.addWidget(QLabel("Layout JSON:"))
            path_row.addWidget(self.layout_select, 1)
            settings_layout.addLayout(path_row, 2, 0, 1, 2)

            new_row = QHBoxLayout()
            self.new_layout_input = QLineEdit()
            self.new_layout_input.setPlaceholderText("new-layout.json")
            create_btn = QPushButton("Create Layout")
            create_btn.clicked.connect(self._create_layout)
            new_row.addWidget(QLabel("New Layout:"))
            new_row.addWidget(self.new_layout_input, 1)
            new_row.addWidget(create_btn)
            settings_layout.addLayout(new_row, 3, 0, 1, 2)

            actions = [
                ("Save", "save"),
                ("Save + Edge Tabs", "save_edge"),
                ("Edge Debug Session", "edge_debug"),
                ("Edge Capture Tabs", "edge_capture"),
                ("Restore", "restore"),
                ("Smart Restore", "restore_smart"),
                ("Restore Dry Run", "restore_dry"),
                ("Restore + Launch Missing", "restore_missing"),
                ("Restore + Edge Tabs", "restore_edge"),
                ("Restore + Open URLs", "restore_simple"),
                ("Edit Edge Mapping", "edit"),
            ]

            self.layout_select.currentIndexChanged.connect(self._reload_speed_menu)

            actions_widget = QWidget()
            actions_grid = QGridLayout(actions_widget)
            actions_grid.setHorizontalSpacing(10)
            actions_grid.setVerticalSpacing(10)
            for idx, (title, action) in enumerate(actions):
                btn = QPushButton(title)
                btn.clicked.connect(lambda _=False, a=action: self._run(a))
                row = idx // 2
                col = idx % 2
                actions_grid.addWidget(btn, row, col)

            log_panel = QWidget()
            log_layout = QVBoxLayout(log_panel)
            log_layout.setSpacing(8)
            log_layout.setContentsMargins(0, 0, 0, 0)
            self.log = QPlainTextEdit()
            self.log.setReadOnly(True)
            self.status = QLabel("Ready")
            log_layout.addWidget(self.log, 1)
            log_layout.addWidget(self.status)

            action_log_row = QHBoxLayout()
            action_log_row.setSpacing(12)
            action_log_row.setAlignment(Qt.AlignTop)
            action_log_row.addWidget(actions_widget, 0, Qt.AlignTop)
            action_log_row.addWidget(log_panel, 1)
            action_log_widget = QWidget()
            action_log_widget.setLayout(action_log_row)

            actions_row = 4
            settings_layout.addWidget(action_log_widget, actions_row, 0, 1, 2)


            speed_layout = QVBoxLayout(speed_tab)
            speed_layout.setContentsMargins(14, 14, 14, 14)
            speed_layout.setSpacing(12)
            self.speed_menu_widget = QWidget()
            self.speed_menu_layout = QGridLayout(self.speed_menu_widget)
            self.speed_menu_layout.setHorizontalSpacing(8)
            self.speed_menu_layout.setVerticalSpacing(8)
            self.speed_menu_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
            speed_layout.addWidget(self.speed_menu_widget)
            self.speed_menu_widget.installEventFilter(self)

            editor_layout = QGridLayout(editor_tab)
            editor_layout.setHorizontalSpacing(12)
            editor_layout.setVerticalSpacing(12)
            editor_layout.setContentsMargins(14, 14, 14, 14)

            self.available_list = QListWidget()
            self.available_list.setSelectionMode(QAbstractItemView.SingleSelection)
            editor_layout.addWidget(QLabel("Layouts (Root)"), 0, 0)
            editor_layout.addWidget(self.available_list, 1, 0)

            self.speed_list = QListWidget()
            self.speed_list.setSelectionMode(QAbstractItemView.SingleSelection)
            self.speed_list.itemSelectionChanged.connect(self._speed_selection_changed)
            editor_layout.addWidget(QLabel("Speed Menu"), 0, 2)
            editor_layout.addWidget(self.speed_list, 1, 2)

            move_buttons = QVBoxLayout()
            add_to_speed_btn = QPushButton("Add ->")
            add_to_speed_btn.clicked.connect(self._move_available_selected)
            remove_from_speed_btn = QPushButton("<- Remove")
            remove_from_speed_btn.clicked.connect(self._move_speed_selected)
            move_buttons.addStretch(1)
            move_buttons.addWidget(add_to_speed_btn)
            move_buttons.addWidget(remove_from_speed_btn)
            move_buttons.addStretch(1)
            move_buttons_widget = QWidget()
            move_buttons_widget.setLayout(move_buttons)
            editor_layout.addWidget(move_buttons_widget, 1, 1)

            detail_row = QHBoxLayout()
            self.speed_label_input = QLineEdit()
            self.speed_label_input.setPlaceholderText("Label")
            self.speed_label_input.textChanged.connect(self._apply_speed_item_edits)
            self.speed_emoji_input = QLineEdit()
            self.speed_emoji_input.setPlaceholderText("Emoji")
            self.speed_emoji_input.textChanged.connect(self._apply_speed_item_edits)
            self.speed_args_input = QLineEdit()
            self.speed_args_input.setPlaceholderText("Args (e.g. --launch-missing)")
            self.speed_args_input.textChanged.connect(self._apply_speed_item_edits)
            self.speed_args_preset = QComboBox()
            self.speed_args_preset.addItems([
                "Custom",
                "Restore",
                "Restore + Launch Missing",
                "Restore + Edge Tabs",
                "Smart Restore + Edge Tabs",
                "Restore Dry Run",
                "Restore + Launch Missing + Edge Tabs",
            ])
            self.speed_args_preset.currentIndexChanged.connect(self._apply_speed_args_preset)
            detail_row.addWidget(self.speed_label_input, 2)
            detail_row.addWidget(self.speed_emoji_input, 1)
            detail_row.addWidget(self.speed_args_preset, 2)
            detail_row.addWidget(self.speed_args_input, 3)

            detail_widget = QWidget()
            detail_widget.setLayout(detail_row)
            editor_layout.addWidget(detail_widget, 2, 0, 1, 3)

            editor_buttons = QHBoxLayout()
            reload_btn = QPushButton("Reload")
            reload_btn.clicked.connect(self._reload_speed_menu)
            save_btn = QPushButton("Save Speed Menu")
            save_btn.clicked.connect(self._save_speed_menu)
            editor_buttons.addStretch(1)
            editor_buttons.addWidget(reload_btn)
            editor_buttons.addWidget(save_btn)

            editor_buttons_widget = QWidget()
            editor_buttons_widget.setLayout(editor_buttons)
            editor_layout.addWidget(editor_buttons_widget, 3, 0, 1, 3)

            layout_editor_layout = QGridLayout(layout_editor_tab)
            layout_editor_layout.setHorizontalSpacing(12)
            layout_editor_layout.setVerticalSpacing(12)
            layout_editor_layout.setContentsMargins(14, 14, 14, 14)

            layout_pick_row = QHBoxLayout()
            self.layout_editor_select = QComboBox()
            self.layout_editor_select.currentIndexChanged.connect(self._sync_layout_settings_choice)
            layout_pick_row.addWidget(QLabel("Layout JSON:"))
            layout_pick_row.addWidget(self.layout_editor_select, 1)
            layout_editor_layout.addLayout(layout_pick_row, 0, 0, 1, 3)

            self.layout_windows_list = QListWidget()
            self.layout_windows_list.setSelectionMode(QAbstractItemView.SingleSelection)
            self.layout_windows_list.itemSelectionChanged.connect(self._layout_window_selected)
            layout_editor_layout.addWidget(QLabel("Windows"), 1, 0)
            layout_editor_layout.addWidget(self.layout_windows_list, 2, 0, 1, 1)

            fields_panel = QWidget()
            fields_layout = QGridLayout(fields_panel)
            fields_layout.setHorizontalSpacing(8)
            fields_layout.setVerticalSpacing(8)

            self.le_title = QLineEdit()
            self.le_class = QLineEdit()
            self.le_process = QLineEdit()
            self.le_exe = QLineEdit()
            self.le_window_id = QLineEdit()
            self.le_window_id.setReadOnly(True)

            self.rect_left = QSpinBox()
            self.rect_top = QSpinBox()
            self.rect_right = QSpinBox()
            self.rect_bottom = QSpinBox()
            for w in (self.rect_left, self.rect_top, self.rect_right, self.rect_bottom):
                w.setRange(-20000, 20000)

            self.nrect_left = QSpinBox()
            self.nrect_top = QSpinBox()
            self.nrect_right = QSpinBox()
            self.nrect_bottom = QSpinBox()
            for w in (self.nrect_left, self.nrect_top, self.nrect_right, self.nrect_bottom):
                w.setRange(-20000, 20000)

            self.spin_show_cmd = QSpinBox()
            self.spin_show_cmd.setRange(0, 20)
            self.chk_visible = QCheckBox("Visible")
            self.chk_minimized = QCheckBox("Minimized")
            self.chk_maximized = QCheckBox("Maximized")

            self.le_launch_exe = QLineEdit()
            self.le_launch_args = QLineEdit()
            self.le_launch_cwd = QLineEdit()

            self.spin_edge_port = QSpinBox()
            self.spin_edge_port.setRange(0, 65535)

            row = 0
            fields_layout.addWidget(QLabel("Title"), row, 0)
            fields_layout.addWidget(self.le_title, row, 1, 1, 3)
            row += 1
            fields_layout.addWidget(QLabel("Class"), row, 0)
            fields_layout.addWidget(self.le_class, row, 1)
            fields_layout.addWidget(QLabel("Process"), row, 2)
            fields_layout.addWidget(self.le_process, row, 3)
            row += 1
            fields_layout.addWidget(QLabel("Exe"), row, 0)
            fields_layout.addWidget(self.le_exe, row, 1, 1, 3)
            row += 1
            fields_layout.addWidget(QLabel("Window ID"), row, 0)
            fields_layout.addWidget(self.le_window_id, row, 1, 1, 3)
            row += 1

            fields_layout.addWidget(QLabel("Rect L/T/R/B"), row, 0)
            rect_row = QHBoxLayout()
            rect_row.addWidget(self.rect_left)
            rect_row.addWidget(self.rect_top)
            rect_row.addWidget(self.rect_right)
            rect_row.addWidget(self.rect_bottom)
            rect_row_widget = QWidget()
            rect_row_widget.setLayout(rect_row)
            fields_layout.addWidget(rect_row_widget, row, 1, 1, 3)
            row += 1

            fields_layout.addWidget(QLabel("Normal L/T/R/B"), row, 0)
            nrect_row = QHBoxLayout()
            nrect_row.addWidget(self.nrect_left)
            nrect_row.addWidget(self.nrect_top)
            nrect_row.addWidget(self.nrect_right)
            nrect_row.addWidget(self.nrect_bottom)
            nrect_row_widget = QWidget()
            nrect_row_widget.setLayout(nrect_row)
            fields_layout.addWidget(nrect_row_widget, row, 1, 1, 3)
            row += 1

            fields_layout.addWidget(QLabel("Show Cmd"), row, 0)
            fields_layout.addWidget(self.spin_show_cmd, row, 1)
            flags_row = QHBoxLayout()
            flags_row.addWidget(self.chk_visible)
            flags_row.addWidget(self.chk_minimized)
            flags_row.addWidget(self.chk_maximized)
            flags_widget = QWidget()
            flags_widget.setLayout(flags_row)
            fields_layout.addWidget(flags_widget, row, 2, 1, 2)
            row += 1

            fields_layout.addWidget(QLabel("Launch Exe"), row, 0)
            fields_layout.addWidget(self.le_launch_exe, row, 1, 1, 3)
            row += 1
            fields_layout.addWidget(QLabel("Launch Args"), row, 0)
            fields_layout.addWidget(self.le_launch_args, row, 1, 1, 3)
            row += 1
            fields_layout.addWidget(QLabel("Launch CWD"), row, 0)
            fields_layout.addWidget(self.le_launch_cwd, row, 1, 1, 3)
            row += 1

            fields_layout.addWidget(QLabel("Edge Session Port"), row, 0)
            fields_layout.addWidget(self.spin_edge_port, row, 1)
            row += 1

            self.edge_tabs_list = QListWidget()
            self.edge_tabs_list.setSelectionMode(QAbstractItemView.SingleSelection)
            fields_layout.addWidget(QLabel("Edge Tabs"), row, 0)
            fields_layout.addWidget(self.edge_tabs_list, row, 1, 1, 3)
            row += 1

            edge_btns = QHBoxLayout()
            self.edge_tab_add = QPushButton("Add Tab")
            self.edge_tab_remove = QPushButton("Remove Tab")
            self.edge_tab_add.clicked.connect(self._add_edge_tab)
            self.edge_tab_remove.clicked.connect(self._remove_edge_tab)
            edge_btns.addWidget(self.edge_tab_add)
            edge_btns.addWidget(self.edge_tab_remove)
            edge_btns_widget = QWidget()
            edge_btns_widget.setLayout(edge_btns)
            fields_layout.addWidget(edge_btns_widget, row, 1, 1, 3)

            self.le_title.textChanged.connect(self._mark_layout_dirty)
            self.le_class.textChanged.connect(self._mark_layout_dirty)
            self.le_process.textChanged.connect(self._mark_layout_dirty)
            self.le_exe.textChanged.connect(self._mark_layout_dirty)
            self.rect_left.valueChanged.connect(self._mark_layout_dirty)
            self.rect_top.valueChanged.connect(self._mark_layout_dirty)
            self.rect_right.valueChanged.connect(self._mark_layout_dirty)
            self.rect_bottom.valueChanged.connect(self._mark_layout_dirty)
            self.nrect_left.valueChanged.connect(self._mark_layout_dirty)
            self.nrect_top.valueChanged.connect(self._mark_layout_dirty)
            self.nrect_right.valueChanged.connect(self._mark_layout_dirty)
            self.nrect_bottom.valueChanged.connect(self._mark_layout_dirty)
            self.spin_show_cmd.valueChanged.connect(self._mark_layout_dirty)
            self.chk_visible.stateChanged.connect(self._mark_layout_dirty)
            self.chk_minimized.stateChanged.connect(self._mark_layout_dirty)
            self.chk_maximized.stateChanged.connect(self._mark_layout_dirty)
            self.le_launch_exe.textChanged.connect(self._mark_layout_dirty)
            self.le_launch_args.textChanged.connect(self._mark_layout_dirty)
            self.le_launch_cwd.textChanged.connect(self._mark_layout_dirty)
            self.spin_edge_port.valueChanged.connect(self._mark_layout_dirty)

            fields_scroll = QScrollArea()
            fields_scroll.setWidgetResizable(True)
            fields_scroll.setFrameShape(QFrame.NoFrame)
            fields_scroll.setWidget(fields_panel)
            layout_editor_layout.addWidget(fields_scroll, 2, 1, 1, 2)

            layout_buttons = QHBoxLayout()
            self.layout_reload_btn = QPushButton("Reload")
            self.layout_save_btn = QPushButton("Save Layout")
            self.layout_remove_btn = QPushButton("Remove Window")
            self.layout_restore_btn = QPushButton("Restore Removed")
            self.layout_reload_btn.clicked.connect(self._load_layout_for_editing)
            self.layout_save_btn.clicked.connect(self._save_layout_edit)
            self.layout_remove_btn.clicked.connect(self._remove_selected_window)
            self.layout_restore_btn.clicked.connect(self._restore_removed_window)
            layout_buttons.addWidget(self.layout_reload_btn)
            layout_buttons.addWidget(self.layout_save_btn)
            layout_buttons.addStretch(1)
            layout_buttons.addWidget(self.layout_remove_btn)
            layout_buttons.addWidget(self.layout_restore_btn)
            layout_buttons_widget = QWidget()
            layout_buttons_widget.setLayout(layout_buttons)
            layout_editor_layout.addWidget(layout_buttons_widget, 3, 0, 1, 3)


            self._reload_speed_menu()
            self._on_tab_changed(tabs.currentIndex())
            self._load_layouts_root_field()
            self._reload_layout_choices()
            self._layout_edit_data = None
            self._layout_removed_cache = None
            self._layout_edit_name = ""
            self._speed_dirty = False
            self._layout_dirty = False
            self._last_tab_index = tabs.currentIndex()
            self._layout_edit_loading = False
            self._tab_change_guard = False
            self._hotkey_thread = None
            self._hotkey_thread_id = None
            self._tray_enabled = True
            self._tray_icon = None
            self._hotkey_emitter = HotkeyEmitter()
            self._hotkey_emitter.fired.connect(self._log_hotkey_fire)
            self.hotkeys_enabled.setChecked(_get_hotkeys_enabled())
            if self.hotkeys_enabled.isChecked():
                self._start_hotkeys()
            self._init_tray_icon()

        def _load_layouts_root_field(self) -> None:
            root = _get_layouts_root()
            self.layouts_root_input.setText(root)

        def _browse_layouts_root(self) -> None:
            selected = QFileDialog.getExistingDirectory(self, "Select layouts root", self.layouts_root_input.text())
            if selected:
                self.layouts_root_input.setText(selected)
                self._save_layouts_root()

        def _save_layouts_root(self) -> None:
            root = self.layouts_root_input.text().strip()
            if not root:
                return
            data = _load_config()
            data["layouts_root"] = root
            try:
                with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
            except Exception as exc:
                QMessageBox.warning(self, "Settings", f"Failed to save layouts root: {exc}")
                return
            _ensure_layouts_root()
            self._reload_layout_choices()
            self._reload_speed_menu()

        def _reload_layout_choices(self) -> None:
            root = _ensure_layouts_root()
            layouts = self._list_layout_files(root)
            current = self.layout_select.currentText()
            self.layout_select.blockSignals(True)
            self.layout_select.clear()
            for name in layouts:
                self.layout_select.addItem(name)
            if current and current in layouts:
                self.layout_select.setCurrentText(current)
            elif layouts:
                self.layout_select.setCurrentIndex(0)
            self.layout_select.blockSignals(False)

            current_editor = self.layout_editor_select.currentText()
            self.layout_editor_select.blockSignals(True)
            self.layout_editor_select.clear()
            for name in layouts:
                self.layout_editor_select.addItem(name)
            if current_editor and current_editor in layouts:
                self.layout_editor_select.setCurrentText(current_editor)
            elif layouts:
                self.layout_editor_select.setCurrentIndex(0)
            self.layout_editor_select.blockSignals(False)
            self._sync_layout_editor_choice()

        def _create_layout(self) -> None:
            name = self.new_layout_input.text().strip()
            if not name:
                QMessageBox.information(self, "New Layout", "Enter a layout name.")
                return
            if not name.lower().endswith(".json"):
                name += ".json"
            root = _ensure_layouts_root()
            path = os.path.join(root, name)
            if os.path.exists(path):
                QMessageBox.warning(self, "New Layout", "Layout already exists.")
                return
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(
                        {
                            "schema": "window-layout.v2",
                            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                            "windows": [],
                            "edge_sessions": [],
                            "open_urls": {"edge": []},
                        },
                        f,
                        indent=2,
                        ensure_ascii=False,
                    )
            except Exception as exc:
                QMessageBox.warning(self, "New Layout", f"Failed to create: {exc}")
                return
            self.new_layout_input.setText("")
            self._reload_layout_choices()
            self.layout_select.setCurrentText(name)

        def _on_tab_changed(self, index: int) -> None:
            # Index 0 = Settings, 1 = Speed Menu, 2 = Speed Menu Editor, 3 = Layout Editor
            if getattr(self, "_tab_change_guard", False):
                self._tab_change_guard = False
                self._last_tab_index = index
                return
            if hasattr(self, "_last_tab_index") and self._last_tab_index != index:
                leaving = self._last_tab_index
                if leaving == 2 and self._speed_dirty:
                    if not self._confirm_unsaved(
                        "Speed Menu Editor",
                        self._save_speed_menu,
                        self._discard_speed_menu_changes,
                    ):
                        blocker = QSignalBlocker(self._tabs)
                        self._tabs.setCurrentIndex(leaving)
                        del blocker
                        return
                if leaving == 3 and self._layout_dirty:
                    if not self._confirm_unsaved(
                        "Layout Editor",
                        self._save_layout_edit,
                        self._discard_layout_changes,
                    ):
                        blocker = QSignalBlocker(self._tabs)
                        self._tabs.setCurrentIndex(leaving)
                        del blocker
                        return
                self._last_tab_index = index
            if index in (1, 2):
                self.setMinimumSize(MIN_SPEED_SIZE[0], MIN_SPEED_SIZE[1])
            else:
                self.setMinimumSize(MIN_SETTINGS_SIZE[0], MIN_SETTINGS_SIZE[1])
            if index == 3:
                self._load_layout_for_editing()

        def changeEvent(self, event) -> None:
            if event.type() == QEvent.WindowStateChange:
                if self._tray_enabled and self._tray_icon is not None:
                    if self.isMinimized():
                        self._hide_to_tray()
                        event.ignore()
                        return
            super().changeEvent(event)

        def _on_tab_bar_clicked(self, index: int) -> None:
            current = self._tabs.currentIndex()
            if index == current:
                return
            if current == 2 and self._speed_dirty:
                if not self._confirm_unsaved(
                    "Speed Menu Editor",
                    self._save_speed_menu,
                    self._discard_speed_menu_changes,
                ):
                    blocker = QSignalBlocker(self._tabs)
                    self._tabs.setCurrentIndex(current)
                    del blocker
                    return
                self._tab_change_guard = True
            if current == 3 and self._layout_dirty:
                if not self._confirm_unsaved(
                    "Layout Editor",
                    self._save_layout_edit,
                    self._discard_layout_changes,
                ):
                    blocker = QSignalBlocker(self._tabs)
                    self._tabs.setCurrentIndex(current)
                    del blocker
                    return
                self._tab_change_guard = True

        def _reload_speed_menu(self, force: bool = False) -> None:
            if not force and getattr(self, "_speed_dirty", False):
                if not self._confirm_unsaved(
                    "Speed Menu Editor",
                    self._save_speed_menu,
                    self._discard_speed_menu_changes,
                ):
                    return
            items = _parse_speed_menu(CONFIG_PATH)
            self._speed_menu_items = items
            self._render_speed_menu()
            self._load_speed_menu_editor(items)
            self._speed_dirty = False

        def _render_speed_menu(self) -> None:
            if self.speed_list.count() > 0:
                self._speed_menu_items = self._collect_speed_items_from_list()
            while self.speed_menu_layout.count():
                item = self.speed_menu_layout.takeAt(0)
                widget = item.widget()
                if widget is not None:
                    widget.deleteLater()

            if not getattr(self, "_speed_menu_items", None):
                empty = QLabel("No speed menu configured in config.json.")
                self.speed_menu_layout.addWidget(empty, 0, 0, 1, 2)
                return

            columns = self._compute_speed_columns()
            last_columns = getattr(self, "_speed_menu_last_columns", 0)
            for col in range(max(columns, last_columns)):
                self.speed_menu_layout.setColumnStretch(col, 0)
            for col in range(columns):
                self.speed_menu_layout.setColumnStretch(col, 1)
            for idx, item in enumerate(self._speed_menu_items):
                label = item.label or Path(item.layout).stem or "Restore"
                if item.emoji:
                    label = f"{item.emoji} {label}"
                btn = QPushButton(label)
                btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                btn.setMinimumHeight(96)
                btn.setMinimumWidth(96)
                layout_target = _resolve_speed_layout(item.layout)
                tooltip_args = " ".join(item.args)
                btn.setToolTip(f"{layout_target} {tooltip_args}".strip())
                btn.clicked.connect(lambda _=False, entry=item: self._run_speed(entry))
                row = idx // columns
                col = idx % columns
                self.speed_menu_layout.setRowStretch(row, 1)
                self.speed_menu_layout.addWidget(btn, row, col)
            self._speed_menu_last_columns = columns

        def _compute_speed_columns(self) -> int:
            width = self.speed_menu_widget.width()
            height = self.speed_menu_widget.height()
            count = len(getattr(self, "_speed_menu_items", []) or [])
            if count <= 1:
                return 1
            tile_min = 96
            if width <= 0 or height <= 0:
                return max(1, min(count, 3))
            max_cols = max(1, width // tile_min)
            max_rows = max(1, height // tile_min)
            columns = (count + max_rows - 1) // max_rows
            return max(1, min(columns, max_cols))

        def eventFilter(self, obj, event) -> bool:
            if obj is self.speed_menu_widget and event.type() == QEvent.Resize:
                self._render_speed_menu()
            return super().eventFilter(obj, event)

        def _load_speed_menu_editor(self, items: List[SpeedMenuItem]) -> None:
            self._speed_edit_loading = True
            self.available_list.clear()
            self.speed_list.clear()
            self._speed_item_cache = {
                item.layout: SpeedMenuItem(item.label, item.emoji, item.layout, list(item.args))
                for item in items
                if item.layout
            }

            layouts_root = _ensure_layouts_root()
            available = self._list_layout_files(layouts_root)
            speed_layouts = {item.layout for item in items if item.layout}

            for layout in sorted(available):
                if layout in speed_layouts:
                    continue
                entry = QListWidgetItem(layout)
                entry.setData(Qt.UserRole, layout)
                self.available_list.addItem(entry)

            for item in items:
                self._add_speed_item_to_list(item)

            self.speed_label_input.setText("")
            self.speed_emoji_input.setText("")
            self.speed_args_input.setText("")
            self.speed_args_preset.setCurrentIndex(0)
            self._speed_edit_loading = False
            self._speed_dirty = False

        def _list_layout_files(self, root: str) -> List[str]:
            try:
                names = [
                    f
                    for f in os.listdir(root)
                    if f.lower().endswith(".json") and os.path.isfile(os.path.join(root, f))
                ]
                return sorted(names, key=str.lower)
            except FileNotFoundError:
                return []

        def _collect_speed_items_from_list(self) -> List[SpeedMenuItem]:
            items: List[SpeedMenuItem] = []
            for row in range(self.speed_list.count()):
                entry = self.speed_list.item(row)
                item = entry.data(Qt.UserRole)
                if isinstance(item, SpeedMenuItem):
                    items.append(item)
            return items

        def _add_speed_item_to_list(self, item: SpeedMenuItem) -> None:
            display = self._format_speed_item_label(item)
            entry = QListWidgetItem(display)
            entry.setData(Qt.UserRole, item)
            self.speed_list.addItem(entry)

        def _format_speed_item_label(self, item: SpeedMenuItem) -> str:
            base = item.label or Path(item.layout).stem or item.layout or "Layout"
            if item.emoji:
                return f"{item.emoji} {base}"
            return base

        def _move_available_to_speed(self, entry: QListWidgetItem) -> None:
            layout = entry.data(Qt.UserRole)
            if not layout:
                return
            cached = self._speed_item_cache.get(layout)
            if cached is not None:
                item = SpeedMenuItem(cached.label, cached.emoji, cached.layout, list(cached.args))
            else:
                item = SpeedMenuItem(label="", emoji="", layout=layout, args=[])
            self._add_speed_item_to_list(item)
            row = self.available_list.row(entry)
            self.available_list.takeItem(row)
            self.available_list.sortItems()
            self.speed_list.setCurrentRow(self.speed_list.count() - 1)
            self._speed_item_selected(self.speed_list.currentItem())
            self._render_speed_menu()
            self._speed_dirty = True

        def _move_speed_to_available(self, entry: QListWidgetItem) -> None:
            item = entry.data(Qt.UserRole)
            if not isinstance(item, SpeedMenuItem):
                return
            if item.layout:
                self._speed_item_cache[item.layout] = SpeedMenuItem(
                    item.label,
                    item.emoji,
                    item.layout,
                    list(item.args),
                )
            list_blocker = QSignalBlocker(self.speed_list)
            label_blocker = QSignalBlocker(self.speed_label_input)
            emoji_blocker = QSignalBlocker(self.speed_emoji_input)
            args_blocker = QSignalBlocker(self.speed_args_input)
            if item.layout:
                available_entry = QListWidgetItem(item.layout)
                available_entry.setData(Qt.UserRole, item.layout)
                self.available_list.addItem(available_entry)
                self.available_list.sortItems()
            row = self.speed_list.row(entry)
            self.speed_list.takeItem(row)
            self._speed_edit_loading = True
            self.speed_list.setCurrentRow(-1)
            self.speed_label_input.setText("")
            self.speed_emoji_input.setText("")
            self.speed_args_input.setText("")
            self._speed_edit_loading = False
            del list_blocker, label_blocker, emoji_blocker, args_blocker
            self._render_speed_menu()
            self._speed_dirty = True

        def _move_available_selected(self) -> None:
            entry = self.available_list.currentItem()
            if entry is not None:
                self._move_available_to_speed(entry)

        def _move_speed_selected(self) -> None:
            entry = self.speed_list.currentItem()
            if entry is not None:
                self._move_speed_to_available(entry)

        def _speed_item_selected(self, entry: QListWidgetItem) -> None:
            if entry is None:
                return
            item = entry.data(Qt.UserRole)
            if not isinstance(item, SpeedMenuItem):
                return
            self._speed_edit_loading = True
            self.speed_label_input.setText(item.label)
            self.speed_emoji_input.setText(item.emoji)
            self.speed_args_input.setText(" ".join(item.args))
            self._sync_args_preset(item.args)
            self._speed_edit_loading = False

        def _speed_selection_changed(self) -> None:
            if getattr(self, "_speed_edit_loading", False):
                return
            entry = self.speed_list.currentItem()
            if entry is None:
                self._speed_edit_loading = True
                self.speed_label_input.setText("")
                self.speed_emoji_input.setText("")
                self.speed_args_input.setText("")
                self.speed_args_preset.setCurrentIndex(0)
                self._speed_edit_loading = False
                return
            self._speed_item_selected(entry)

        def _apply_speed_item_edits(self) -> None:
            if getattr(self, "_speed_edit_loading", False):
                return
            entry = self.speed_list.currentItem()
            if entry is None:
                return
            item = entry.data(Qt.UserRole)
            if not isinstance(item, SpeedMenuItem):
                return
            item.label = self.speed_label_input.text().strip()
            item.emoji = self.speed_emoji_input.text().strip()
            raw_args = self.speed_args_input.text().strip()
            try:
                item.args = shlex.split(raw_args, posix=False) if raw_args else []
            except ValueError:
                item.args = raw_args.split()
            entry.setText(self._format_speed_item_label(item))
            entry.setData(Qt.UserRole, item)
            if item.layout:
                self._speed_item_cache[item.layout] = SpeedMenuItem(
                    item.label,
                    item.emoji,
                    item.layout,
                    list(item.args),
                )
            self._render_speed_menu()
            self._speed_dirty = True

        def _apply_speed_args_preset(self) -> None:
            if getattr(self, "_speed_edit_loading", False):
                return
            preset = self.speed_args_preset.currentText()
            mapping = {
                "Custom": "",
                "Restore": "",
                "Restore + Launch Missing": "--launch-missing",
                "Restore + Edge Tabs": "--restore-edge-tabs",
                "Smart Restore + Edge Tabs": "--smart --restore-edge-tabs",
                "Restore Dry Run": "--dry-run",
                "Restore + Launch Missing + Edge Tabs": "--launch-missing --restore-edge-tabs",
            }
            args = mapping.get(preset, "")
            if args:
                self._speed_edit_loading = True
                self.speed_args_input.setText(args)
                self._speed_edit_loading = False
            self._apply_speed_item_edits()

        def _sync_args_preset(self, args: List[str]) -> None:
            raw = " ".join(args).strip()
            mapping = {
                "": "Restore",
                "--launch-missing": "Restore + Launch Missing",
                "--restore-edge-tabs": "Restore + Edge Tabs",
                "--smart --restore-edge-tabs": "Smart Restore + Edge Tabs",
                "--dry-run": "Restore Dry Run",
                "--launch-missing --restore-edge-tabs": "Restore + Launch Missing + Edge Tabs",
            }
            label = mapping.get(raw, "Custom")
            idx = self.speed_args_preset.findText(label)
            if idx >= 0:
                self.speed_args_preset.setCurrentIndex(idx)

        def _mark_layout_dirty(self, *_args) -> None:
            if getattr(self, "_layout_edit_loading", False):
                return
            self._layout_dirty = True

        def _discard_speed_menu_changes(self) -> None:
            self._speed_dirty = False
            self._reload_speed_menu(force=True)

        def _discard_layout_changes(self) -> None:
            self._layout_dirty = False
            self._load_layout_for_editing(force=True)

        def _save_speed_menu(self) -> None:
            data = _load_json(CONFIG_PATH)
            if data is None:
                data = {}
            if not isinstance(data, dict):
                QMessageBox.warning(self, "Speed Menu", "Config JSON not found or invalid.")
                return

            items: List[SpeedMenuItem] = []
            for row in range(self.speed_list.count()):
                entry = self.speed_list.item(row)
                item = entry.data(Qt.UserRole)
                if isinstance(item, SpeedMenuItem):
                    items.append(item)
            data["speed_menu"] = {
                "buttons": [
                    {
                        "label": item.label,
                        "emoji": item.emoji,
                        "layout": item.layout,
                        "args": item.args,
                    }
                    for item in items
                ],
            }

            try:
                with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
            except Exception as exc:
                QMessageBox.warning(self, "Speed Menu", f"Failed to save: {exc}")
                return

            self._speed_dirty = False
            self._reload_speed_menu()

        def _load_layout_for_editing(self, force: bool = False) -> None:
            if not force and self._layout_dirty:
                if not self._confirm_unsaved(
                    "Layout Editor",
                    self._save_layout_edit,
                    self._discard_layout_changes,
                ):
                    if self._layout_edit_name:
                        blocker = QSignalBlocker(self.layout_editor_select)
                        self.layout_editor_select.setCurrentText(self._layout_edit_name)
                        del blocker
                        blocker = QSignalBlocker(self.layout_select)
                        self.layout_select.setCurrentText(self._layout_edit_name)
                        del blocker
                    return
            name = self.layout_editor_select.currentText().strip()
            if not name:
                name = self.layout_select.currentText().strip()
            if not name:
                return
            path = os.path.join(_ensure_layouts_root(), name)
            data = _load_json(path)
            if not isinstance(data, dict):
                QMessageBox.warning(self, "Layout Editor", "Failed to load layout JSON.")
                return
            self._layout_edit_data = data
            self._layout_removed_cache = None
            self._layout_edit_name = name
            self._reload_layout_windows_list()
            self._layout_dirty = False

        def _sync_layout_editor_choice(self) -> None:
            name = self.layout_select.currentText().strip()
            if not name:
                return
            if self.layout_editor_select.currentText().strip() == name:
                return
            blocker = QSignalBlocker(self.layout_editor_select)
            self.layout_editor_select.setCurrentText(name)
            del blocker
            self._load_layout_for_editing()

        def _sync_layout_settings_choice(self) -> None:
            name = self.layout_editor_select.currentText().strip()
            if not name:
                return
            if self.layout_select.currentText().strip() == name:
                return
            blocker = QSignalBlocker(self.layout_select)
            self.layout_select.setCurrentText(name)
            del blocker
            self._reload_speed_menu()
            self._load_layout_for_editing()

        def _reload_layout_windows_list(self) -> None:
            self.layout_windows_list.clear()
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            for idx, window in enumerate(windows):
                title = str(window.get("title") or "(untitled)")
                proc = str(window.get("process_name") or "")
                label = f"{title} [{proc}]"
                item = QListWidgetItem(label)
                item.setData(Qt.UserRole, idx)
                self.layout_windows_list.addItem(item)
            if self.layout_windows_list.count() > 0:
                self.layout_windows_list.setCurrentRow(0)

        def _layout_window_selected(self) -> None:
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            item = self.layout_windows_list.currentItem()
            if item is None:
                return
            idx = item.data(Qt.UserRole)
            if not isinstance(idx, int) or idx < 0 or idx >= len(windows):
                return
            window = windows[idx]
            self._layout_selected_index = idx
            self._load_window_fields(window)

        def _load_window_fields(self, window: dict) -> None:
            self._layout_edit_loading = True
            self.le_title.setText(str(window.get("title") or ""))
            self.le_class.setText(str(window.get("class_name") or ""))
            self.le_process.setText(str(window.get("process_name") or ""))
            self.le_exe.setText(str(window.get("exe") or ""))
            self.le_window_id.setText(str(window.get("window_id") or ""))

            rect = window.get("rect") or [0, 0, 0, 0]
            nrect = window.get("normal_rect") or [0, 0, 0, 0]
            if len(rect) == 4:
                self.rect_left.setValue(int(rect[0]))
                self.rect_top.setValue(int(rect[1]))
                self.rect_right.setValue(int(rect[2]))
                self.rect_bottom.setValue(int(rect[3]))
            if len(nrect) == 4:
                self.nrect_left.setValue(int(nrect[0]))
                self.nrect_top.setValue(int(nrect[1]))
                self.nrect_right.setValue(int(nrect[2]))
                self.nrect_bottom.setValue(int(nrect[3]))

            self.spin_show_cmd.setValue(int(window.get("show_cmd") or 0))
            self.chk_visible.setChecked(bool(window.get("is_visible", True)))
            self.chk_minimized.setChecked(bool(window.get("is_minimized", False)))
            self.chk_maximized.setChecked(bool(window.get("is_maximized", False)))

            launch = window.get("launch") or {}
            if isinstance(launch, dict):
                self.le_launch_exe.setText(str(launch.get("exe") or ""))
                args = launch.get("args") or []
                if isinstance(args, list):
                    self.le_launch_args.setText(" ".join(str(a) for a in args))
                else:
                    self.le_launch_args.setText(str(args or ""))
                self.le_launch_cwd.setText(str(launch.get("cwd") or ""))
            else:
                self.le_launch_exe.setText("")
                self.le_launch_args.setText("")
                self.le_launch_cwd.setText("")

            edge = window.get("edge") or {}
            if isinstance(edge, dict):
                try:
                    self.spin_edge_port.setValue(int(edge.get("session_port") or 0))
                except (TypeError, ValueError):
                    self.spin_edge_port.setValue(0)
            else:
                self.spin_edge_port.setValue(0)

            self.edge_tabs_list.clear()
            tabs = window.get("edge_tabs") or []
            for tab in tabs:
                url = str(tab.get("url") or "")
                title = str(tab.get("title") or "")
                label = f"{title} -> {url}" if title else url
                item = QListWidgetItem(label)
                item.setData(Qt.UserRole, tab)
                self.edge_tabs_list.addItem(item)
            self._layout_edit_loading = False

        def _apply_window_fields(self) -> None:
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            idx = getattr(self, "_layout_selected_index", None)
            if idx is None or idx < 0 or idx >= len(windows):
                return
            window = windows[idx]

            window["title"] = self.le_title.text().strip()
            window["class_name"] = self.le_class.text().strip()
            window["process_name"] = self.le_process.text().strip()
            window["exe"] = self.le_exe.text().strip()

            window["rect"] = [
                self.rect_left.value(),
                self.rect_top.value(),
                self.rect_right.value(),
                self.rect_bottom.value(),
            ]
            window["normal_rect"] = [
                self.nrect_left.value(),
                self.nrect_top.value(),
                self.nrect_right.value(),
                self.nrect_bottom.value(),
            ]
            window["show_cmd"] = int(self.spin_show_cmd.value())
            window["is_visible"] = bool(self.chk_visible.isChecked())
            window["is_minimized"] = bool(self.chk_minimized.isChecked())
            window["is_maximized"] = bool(self.chk_maximized.isChecked())

            launch_exe = self.le_launch_exe.text().strip()
            launch_args = self.le_launch_args.text().strip()
            launch_cwd = self.le_launch_cwd.text().strip()
            if launch_exe or launch_args or launch_cwd:
                window["launch"] = {
                    "exe": launch_exe,
                    "args": shlex.split(launch_args, posix=False) if launch_args else [],
                    "cwd": launch_cwd,
                }
            else:
                window.pop("launch", None)

            edge_port = int(self.spin_edge_port.value())
            if edge_port > 0:
                window["edge"] = {"session_port": edge_port}
            else:
                window.pop("edge", None)

        def _save_layout_edit(self) -> None:
            self._apply_window_fields()
            name = self.layout_editor_select.currentText().strip() or self.layout_select.currentText().strip()
            if not name:
                return
            path = os.path.join(_ensure_layouts_root(), name)
            try:
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(self._layout_edit_data or {}, f, indent=2, ensure_ascii=False)
            except Exception as exc:
                QMessageBox.warning(self, "Layout Editor", f"Failed to save: {exc}")
                return
            self._reload_layout_windows_list()
            self._layout_dirty = False

        def _remove_selected_window(self) -> None:
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            idx = getattr(self, "_layout_selected_index", None)
            if idx is None or idx < 0 or idx >= len(windows):
                return
            removed = windows.pop(idx)
            self._layout_removed_cache = (removed, idx)
            self._reload_layout_windows_list()
            self._layout_dirty = True

        def _restore_removed_window(self) -> None:
            if not self._layout_removed_cache:
                return
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            removed, idx = self._layout_removed_cache
            if idx < 0 or idx > len(windows):
                idx = len(windows)
            windows.insert(idx, removed)
            self._layout_removed_cache = None
            self._reload_layout_windows_list()
            self._layout_dirty = True

        def _add_edge_tab(self) -> None:
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            idx = getattr(self, "_layout_selected_index", None)
            if idx is None or idx < 0 or idx >= len(windows):
                return
            url, ok = QInputDialog.getText(self, "Add Edge Tab", "URL:")
            if not ok:
                return
            url = url.strip()
            if not url:
                return
            title, _ = QInputDialog.getText(self, "Add Edge Tab", "Title (optional):")
            window = windows[idx]
            tabs = window.get("edge_tabs") or []
            tabs.append({"title": title.strip(), "url": url})
            window["edge_tabs"] = tabs
            self._load_window_fields(window)
            self._layout_dirty = True

        def _remove_edge_tab(self) -> None:
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            idx = getattr(self, "_layout_selected_index", None)
            if idx is None or idx < 0 or idx >= len(windows):
                return
            window = windows[idx]
            tabs = window.get("edge_tabs") or []
            item = self.edge_tabs_list.currentItem()
            if item is None:
                return
            row = self.edge_tabs_list.row(item)
            if 0 <= row < len(tabs):
                tabs.pop(row)
                window["edge_tabs"] = tabs
                self._load_window_fields(window)
                self._layout_dirty = True

        def _confirm_unsaved(self, label: str, save_fn, discard_fn) -> bool:
            box = QMessageBox(self)
            box.setWindowTitle("Unsaved Changes")
            box.setText(f"Unsaved changes in {label}. Save before continuing?")
            box.setIcon(QMessageBox.Warning)
            save_btn = box.addButton("Save", QMessageBox.AcceptRole)
            discard_btn = box.addButton("Discard", QMessageBox.DestructiveRole)
            box.addButton("Cancel", QMessageBox.RejectRole)
            box.setDefaultButton(save_btn)
            box.exec()
            if box.clickedButton() == save_btn:
                save_fn()
                return True
            if box.clickedButton() == discard_btn:
                discard_fn()
                return True
            return False

        def closeEvent(self, event) -> None:
            if self._speed_dirty:
                if not self._confirm_unsaved(
                    "Speed Menu Editor",
                    self._save_speed_menu,
                    self._discard_speed_menu_changes,
                ):
                    event.ignore()
                    return
            if self._layout_dirty:
                if not self._confirm_unsaved(
                    "Layout Editor",
                    self._save_layout_edit,
                    self._discard_layout_changes,
                ):
                    event.ignore()
                    return
            if self._tray_enabled and self._tray_icon is not None:
                event.ignore()
                self._hide_to_tray()
                return
            self._stop_hotkeys()
            event.accept()

        def _run_speed(self, entry: SpeedMenuItem) -> None:
            if self._proc.state() != QProcess.NotRunning:
                QMessageBox.information(self, "Busy", "A command is already running.")
                return

            target_layout = _resolve_speed_layout(entry.layout)
            if not target_layout:
                QMessageBox.warning(self, "Speed Menu", "Speed menu entry is missing a layout path.")
                return

            cmd = [sys.executable, "window_layout.py", "restore", target_layout, *entry.args]
            title = entry.label or Path(target_layout).stem or "Speed Restore"
            if entry.emoji:
                title = f"{entry.emoji} {title}"
            self.status.setText(f"Running: {title}")
            self.log.appendPlainText(f"\n$ {format_command_for_log(cmd)}")

            program = cmd[0]
            args = cmd[1:]
            self._proc.start(program, args)

        def _current_layout_path(self) -> str:
            name = self.layout_select.currentText().strip()
            if not name:
                return os.path.join(_ensure_layouts_root(), DEFAULT_LAYOUT_PATH)
            return os.path.join(_ensure_layouts_root(), name)

        def _run(self, action: str) -> None:
            if self._proc.state() != QProcess.NotRunning:
                QMessageBox.information(self, "Busy", "A command is already running.")
                return

            layout_path = self._current_layout_path()
            edge_port, edge_profile_dir = _get_edge_defaults()
            if action in ("edge_debug", "edge_capture"):
                edge_settings = self._prompt_edge_settings(edge_port, edge_profile_dir)
                if not edge_settings:
                    return
                edge_port, edge_profile_dir = edge_settings
            cmd = build_cli_command(action, layout_path, edge_port=edge_port, edge_profile_dir=edge_profile_dir).args
            self.status.setText(f"Running: {action}")
            self.log.appendPlainText(f"\n$ {format_command_for_log(cmd)}")

            program = cmd[0]
            args = cmd[1:]
            self._proc.start(program, args)

        def _prompt_edge_settings(self, port: int, profile_dir: str) -> Optional[tuple[int, str]]:
            port_value, ok = QInputDialog.getInt(
                self,
                "Edge Debug Port",
                "Remote debugging port:",
                port,
                1,
                65535,
                1,
            )
            if not ok:
                return None
            profile_value, ok = QInputDialog.getText(
                self,
                "Edge Profile Dir",
                "Profile directory (optional):",
                text=profile_dir,
            )
            if not ok:
                return None
            profile_value = profile_value.strip()
            try:
                _save_edge_defaults(port_value, profile_value)
            except Exception:
                pass
            return port_value, profile_value

        def _init_tray_icon(self) -> None:
            if not QSystemTrayIcon.isSystemTrayAvailable():
                self._tray_enabled = False
                return
            self._tray_icon = QSystemTrayIcon(self)
            icon = self.windowIcon()
            if icon.isNull():
                icon = self.style().standardIcon(QStyle.SP_ComputerIcon)
                self.setWindowIcon(icon)
            self._tray_icon.setIcon(icon)
            self._tray_icon.setToolTip("Window Layout Manager")
            menu = QMenu()
            show_action = menu.addAction("Show")
            hide_action = menu.addAction("Hide")
            quit_action = menu.addAction("Quit")
            show_action.triggered.connect(self._show_from_tray)
            hide_action.triggered.connect(self._hide_to_tray)
            quit_action.triggered.connect(self._quit_from_tray)
            self._tray_icon.setContextMenu(menu)
            self._tray_icon.activated.connect(self._on_tray_activated)
            self._tray_icon.show()

        def _show_from_tray(self) -> None:
            self.show()
            self.raise_()
            self.activateWindow()

        def _hide_to_tray(self) -> None:
            self.hide()

        def _quit_from_tray(self) -> None:
            self._tray_enabled = False
            self._stop_hotkeys()
            QApplication.instance().quit()

        def _on_tray_activated(self, reason) -> None:
            if reason == QSystemTrayIcon.Trigger:
                if self.isVisible():
                    self.hide()
                else:
                    self._show_from_tray()

        def _start_hotkeys(self) -> None:
            if self._hotkey_thread is not None:
                return
            hotkeys = wl._load_hotkeys(CONFIG_PATH)
            if not hotkeys:
                return
            self._stop_hotkeys()

            def worker():
                import win32con
                import win32gui
                try:
                    win32gui.PeekMessage(None, 0, 0, win32con.PM_NOREMOVE)
                except Exception:
                    pass
                try:
                    import win32api
                    thread_id = win32api.GetCurrentThreadId()
                except Exception:
                    thread_id = 0
                self._hotkey_thread_id = thread_id
                self._hotkey_emitter.fired.emit(f"Hotkey listener running (thread_id={thread_id})")
                registered = {}
                next_id = 1
                for entry in hotkeys:
                    parsed = wl._parse_hotkey_keys(entry["keys"])
                    if not parsed:
                        continue
                    modifiers, vk = parsed
                    try:
                        win32gui.RegisterHotKey(None, next_id, modifiers, vk)
                        registered[next_id] = entry
                        label = f"Hotkey registered: {entry['keys']} -> {entry['action']} {' '.join(entry.get('args', []))}".strip()
                        self._hotkey_emitter.fired.emit(label)
                        next_id += 1
                    except Exception as exc:
                        label = f"Hotkey failed: {entry['keys']} ({exc})"
                        self._hotkey_emitter.fired.emit(label)
                        continue
                if not registered:
                    return
                debug_count = 0
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
                            label = f"Hotkey: {entry['keys']} -> {entry['action']} {' '.join(entry.get('args', []))}".strip()
                            self._hotkey_emitter.fired.emit(label)
                            wl._run_hotkey_action(entry["action"], entry.get("args", []))
                        else:
                            self._hotkey_emitter.fired.emit(f"Hotkey message received (id={wparam}) but no entry found")

            import threading

            self._hotkey_thread = threading.Thread(target=worker, daemon=True)
            self._hotkey_thread.start()

        def _stop_hotkeys(self) -> None:
            if self._hotkey_thread_id:
                try:
                    import win32con
                    try:
                        import win32gui
                        win32gui.PostThreadMessage(self._hotkey_thread_id, win32con.WM_QUIT, 0, 0)
                    except Exception:
                        import win32api
                        win32api.PostThreadMessage(self._hotkey_thread_id, win32con.WM_QUIT, 0, 0)
                except Exception:
                    pass
            self._hotkey_thread = None
            self._hotkey_thread_id = None

        def _log_hotkey_fire(self, message: str) -> None:
            self.log.appendPlainText(message)

        def _toggle_hotkeys(self) -> None:
            enabled = bool(self.hotkeys_enabled.isChecked())
            try:
                _set_hotkeys_enabled(enabled)
            except Exception:
                pass
            if enabled:
                self._start_hotkeys()
            else:
                self._stop_hotkeys()

        def _append_stdout(self) -> None:
            data = bytes(self._proc.readAllStandardOutput()).decode("utf-8", errors="replace")
            if data:
                self.log.appendPlainText(data.rstrip("\n"))

        def _append_stderr(self) -> None:
            data = bytes(self._proc.readAllStandardError()).decode("utf-8", errors="replace")
            if data:
                self.log.appendPlainText(data.rstrip("\n"))

        def _on_finished(self, code: int, _status) -> None:
            if code == 0:
                self.status.setText("Completed")
            else:
                self.status.setText(f"Failed (exit={code})")

    app = QApplication(sys.argv)
    _apply_fluent_style(app)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())

