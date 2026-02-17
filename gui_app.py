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


def build_cli_command(action: str, layout_path: str) -> GuiCommand:
    base = [sys.executable, "window_layout.py"]
    if action == "save":
        return GuiCommand("Save Layout", base + ["save", layout_path])
    if action == "save_edge":
        return GuiCommand("Save Layout + Edge Tabs", base + ["save", layout_path, "--edge-tabs"])
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
    if action == "edit":
        return GuiCommand("Edit Edge Tab Mapping", base + ["edit", layout_path])
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
        from PySide6.QtCore import QProcess, QEvent, Qt, QSignalBlocker
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
            QLineEdit,
            QMainWindow,
            QMessageBox,
            QPushButton,
            QPlainTextEdit,
            QSizePolicy,
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
            tabs.currentChanged.connect(self._on_tab_changed)
            root_layout.addWidget(tabs)

            settings_tab = QWidget()
            speed_tab = QWidget()
            editor_tab = QWidget()
            tabs.addTab(settings_tab, "Settings")
            tabs.addTab(speed_tab, "Speed Menu")
            tabs.addTab(editor_tab, "Speed Menu Editor")

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

            path_row = QHBoxLayout()
            self.layout_select = QComboBox()
            path_row.addWidget(QLabel("Layout JSON:"))
            path_row.addWidget(self.layout_select, 1)
            settings_layout.addLayout(path_row, 1, 0, 1, 2)

            new_row = QHBoxLayout()
            self.new_layout_input = QLineEdit()
            self.new_layout_input.setPlaceholderText("new-layout.json")
            create_btn = QPushButton("Create Layout")
            create_btn.clicked.connect(self._create_layout)
            new_row.addWidget(QLabel("New Layout:"))
            new_row.addWidget(self.new_layout_input, 1)
            new_row.addWidget(create_btn)
            settings_layout.addLayout(new_row, 2, 0, 1, 2)

            actions = [
                ("Save", "save"),
                ("Save + Edge Tabs", "save_edge"),
                ("Restore", "restore"),
                ("Smart Restore", "restore_smart"),
                ("Restore Dry Run", "restore_dry"),
                ("Restore + Launch Missing", "restore_missing"),
                ("Restore + Edge Tabs", "restore_edge"),
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

            actions_row = 3
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
            detail_row.addWidget(self.speed_label_input, 2)
            detail_row.addWidget(self.speed_emoji_input, 1)
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


            self._reload_speed_menu()
            self._on_tab_changed(tabs.currentIndex())
            self._load_layouts_root_field()
            self._reload_layout_choices()

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
                            "schema": "window-layout.v1",
                            "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                            "windows": [],
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
            # Index 0 = Settings, 1 = Speed Menu, 2 = Speed Menu Editor
            if index in (1, 2):
                self.setMinimumSize(MIN_SPEED_SIZE[0], MIN_SPEED_SIZE[1])
            else:
                self.setMinimumSize(MIN_SETTINGS_SIZE[0], MIN_SETTINGS_SIZE[1])

        def _reload_speed_menu(self) -> None:
            items = _parse_speed_menu(CONFIG_PATH)
            self._speed_menu_items = items
            self._render_speed_menu()
            self._load_speed_menu_editor(items)

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
            self._speed_edit_loading = False

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

            self._reload_speed_menu()

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
            cmd = build_cli_command(action, layout_path).args
            self.status.setText(f"Running: {action}")
            self.log.appendPlainText(f"\n$ {format_command_for_log(cmd)}")

            program = cmd[0]
            args = cmd[1:]
            self._proc.start(program, args)

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

