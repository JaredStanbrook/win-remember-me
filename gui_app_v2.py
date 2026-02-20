"""Lightweight GUI for Window Layout CLI (PySide6) — Dark redesign."""

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
MIN_SETTINGS_SIZE = (720, 580)


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
        return GuiCommand("Save Layout + Edge Tabs", cmd)
    if action == "restore_basic":
        return GuiCommand("Restore (Existing Only)", base + ["restore", layout_path])
    if action == "restore_launch":
        return GuiCommand("Restore + Launch Missing", base + ["restore", layout_path, "--launch-missing"])
    if action == "restore_edge":
        return GuiCommand("Restore + Edge Tabs", base + ["restore", layout_path, "--edge-tabs"])
    if action == "restore_edge_destructive":
        return GuiCommand("Restore + Edge Tabs (Destructive)", base + ["restore", layout_path, "--edge-tabs", "--destructive"])
    if action == "restore_launch_edge":
        return GuiCommand("Restore + Launch Missing + Edge Tabs", base + ["restore", layout_path, "--launch-missing", "--edge-tabs"])
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


def _load_hotkeys(path: str = CONFIG_PATH) -> List[dict]:
    data = _load_json(path)
    if not isinstance(data, dict):
        return []
    raw = data.get("hotkeys") or []
    if not isinstance(raw, list):
        return []
    entries: List[dict] = []
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
    cmd = [sys.executable, os.path.abspath("window_layout.py"), action, *args]
    try:
        subprocess.Popen(cmd)
    except Exception:
        pass


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


# ─────────────────────────────────────────────────────────────────────────────
# DARK THEME STYLESHEET
# ─────────────────────────────────────────────────────────────────────────────

DARK_STYLE = """
QWidget {
    background: #1a1d23;
    color: #d4d8e2;
    font-family: "Segoe UI", "Consolas", sans-serif;
    font-size: 10pt;
}

QMainWindow {
    background: #1a1d23;
}

/* ── Sidebar nav ── */
QWidget#sidebar {
    background: #13151a;
    border-right: 1px solid #2a2d35;
}

QPushButton#navBtn {
    background: transparent;
    border: none;
    border-radius: 8px;
    color: #7a8090;
    font-size: 10pt;
    padding: 10px 14px;
    text-align: left;
}

QPushButton#navBtn:hover {
    background: #21252e;
    color: #c8cdd8;
}

QPushButton#navBtn[active="true"] {
    background: #252932;
    color: #6eb3f7;
    border-left: 3px solid #4d9cf5;
}

/* ── Speed Menu popup ── */
QWidget#speedPopup {
    background: #16181f;
    border: 1px solid #2e3340;
    border-radius: 12px;
}

QWidget#speedPopupHeader {
    background: transparent;
}

QPushButton#speedBtn {
    background: #1e2230;
    border: 1px solid #2a2f3d;
    border-radius: 10px;
    color: #c8d0e8;
    font-size: 11pt;
    padding: 10px 8px;
    min-height: 64px;
    min-width: 90px;
}

QPushButton#speedBtn:hover {
    background: #252b3d;
    border-color: #4d7fc4;
    color: #e8edf8;
}

QPushButton#speedBtn:pressed {
    background: #1c2235;
    border-color: #3a6ab0;
}

QPushButton#speedLaunchBtn {
    background: #1a2640;
    border: 1px solid #2c4270;
    border-radius: 8px;
    color: #6eb3f7;
    font-size: 9pt;
    padding: 6px 12px;
    min-height: 28px;
}

QPushButton#speedLaunchBtn:hover {
    background: #1f2f50;
    border-color: #4d7fc4;
}

QLabel#speedTitle {
    color: #6a7085;
    font-size: 9pt;
    background: transparent;
}

QLabel#speedHandle {
    color: #3a3f50;
    background: transparent;
    font-size: 12pt;
}

/* ── Inputs ── */
QLineEdit, QPlainTextEdit, QListWidget, QComboBox {
    background: #1e2128;
    border: 1px solid #2a2e3a;
    border-radius: 7px;
    padding: 5px 8px;
    color: #c8cdd8;
    selection-background-color: #2d5c9e;
}

QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus {
    border-color: #4d7fc4;
    outline: none;
}

QLineEdit:hover, QComboBox:hover {
    border-color: #3a3f50;
}

QComboBox {
    padding-right: 26px;
    min-height: 28px;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 24px;
    border-left: 1px solid #2a2e3a;
    background: #252830;
    border-top-right-radius: 7px;
    border-bottom-right-radius: 7px;
}

QComboBox QAbstractItemView {
    background: #1e2128;
    border: 1px solid #2a2e3a;
    selection-background-color: #2d5c9e;
    outline: 0;
}

QPlainTextEdit {
    font-family: "Consolas", "Courier New", monospace;
    font-size: 9pt;
    padding: 8px;
    line-height: 1.4;
}

/* ── Standard buttons ── */
QPushButton {
    background: #252830;
    border: 1px solid #2e323f;
    border-radius: 7px;
    color: #c0c6d4;
    padding: 6px 14px;
    min-height: 28px;
}

QPushButton:hover {
    background: #2c3040;
    border-color: #404760;
    color: #dde3f0;
}

QPushButton:pressed {
    background: #202430;
    border-color: #3a4060;
}

QPushButton:disabled {
    background: #1e2028;
    color: #484e5e;
    border-color: #252830;
}

QPushButton#primaryBtn {
    background: #1f3b6e;
    border: 1px solid #2d5499;
    color: #a8ccf5;
    font-weight: 600;
}

QPushButton#primaryBtn:hover {
    background: #254680;
    border-color: #4d7fc4;
    color: #cde4ff;
}

QPushButton#accentBtn {
    background: #1e3d2a;
    border: 1px solid #2a5c3a;
    color: #7bc99a;
}

QPushButton#accentBtn:hover {
    background: #244a32;
    border-color: #3d8050;
    color: #9de0b8;
}

QPushButton#dangerBtn {
    background: #3d1e1e;
    border: 1px solid #5a2929;
    color: #d47a7a;
}

QPushButton#dangerBtn:hover {
    background: #4a2222;
    border-color: #7a3333;
}

/* ── Labels ── */
QLabel {
    background: transparent;
    color: #8a92a8;
}

QLabel#sectionHeader {
    color: #4a5068;
    font-size: 8pt;
    font-weight: 600;
    letter-spacing: 1px;
    background: transparent;
}

QLabel#pageTitle {
    color: #9aa4c0;
    font-size: 13pt;
    font-weight: 600;
    background: transparent;
}

QLabel#statusLabel {
    color: #5a7a5a;
    font-size: 9pt;
    background: transparent;
}

QLabel#statusLabel[error="true"] {
    color: #9a5a5a;
}

/* ── Tab widget ── */
QTabWidget::pane {
    border: 1px solid #2a2e3a;
    border-radius: 8px;
    background: #1a1d23;
}

QTabBar::tab {
    background: #1a1d23;
    border: 1px solid #2a2e3a;
    border-bottom: none;
    border-top-left-radius: 7px;
    border-top-right-radius: 7px;
    padding: 7px 16px;
    margin-right: 3px;
    color: #5a6070;
}

QTabBar::tab:selected {
    background: #1e2230;
    color: #a0aac0;
    border-color: #2e3450;
}

QTabBar::tab:hover:!selected {
    color: #8090b0;
    background: #1c1f28;
}

/* ── Lists ── */
QListWidget {
    border-radius: 7px;
    padding: 2px;
}

QListWidget::item {
    padding: 6px 8px;
    border-radius: 5px;
}

QListWidget::item:selected {
    background: #1f3860;
    color: #a8c8f0;
}

QListWidget::item:hover:!selected {
    background: #1e2230;
}

/* ── Checkboxes ── */
QCheckBox {
    color: #9098b0;
    spacing: 8px;
}

QCheckBox::indicator {
    width: 16px;
    height: 16px;
    border-radius: 4px;
    border: 1px solid #3a3f50;
    background: #1e2128;
}

QCheckBox::indicator:checked {
    background: #2d5c9e;
    border-color: #4d7fc4;
}

QCheckBox::indicator:hover {
    border-color: #505878;
}

/* ── SpinBox ── */
QSpinBox {
    background: #1e2128;
    border: 1px solid #2a2e3a;
    border-radius: 7px;
    padding: 5px 8px;
    color: #c8cdd8;
}

QSpinBox:focus {
    border-color: #4d7fc4;
}

QSpinBox::up-button, QSpinBox::down-button {
    background: #252830;
    border: none;
    width: 18px;
}

QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background: #2e3340;
}

/* ── Scrollbars ── */
QScrollBar:vertical {
    background: transparent;
    width: 8px;
    margin: 2px;
}

QScrollBar::handle:vertical {
    background: #2e3348;
    border-radius: 4px;
    min-height: 28px;
}

QScrollBar::handle:vertical:hover {
    background: #3a4060;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
    background: transparent;
}

QScrollBar:horizontal {
    background: transparent;
    height: 8px;
    margin: 2px;
}

QScrollBar::handle:horizontal {
    background: #2e3348;
    border-radius: 4px;
    min-width: 28px;
}

/* ── Frame ── */
QFrame[frameShape="4"],
QFrame[frameShape="5"] {
    color: #2a2e3a;
}

QFrame#card {
    background: #1e2230;
    border: 1px solid #272b38;
    border-radius: 10px;
}

/* ── Divider line ── */
QFrame#divider {
    background: #242830;
    max-height: 1px;
    border: none;
}

/* ── Stacked widget transparency ── */
QStackedWidget {
    background: transparent;
}

/* ── Header bar ── */
QWidget#headerBar {
    background: #13151a;
    border-bottom: 1px solid #21252e;
}

/* ── Table widget ── */
QHeaderView::section {
    background: #1e2230;
    border: 1px solid #2a2e3a;
    padding: 5px 8px;
    color: #7080a0;
    font-size: 9pt;
}

QTableWidget::item:selected {
    background: #1f3860;
    color: #a8c8f0;
}

/* ── Tooltip ── */
QToolTip {
    background: #1e2230;
    color: #a8b0c8;
    border: 1px solid #2e3450;
    padding: 4px 8px;
    border-radius: 5px;
}
"""


def main() -> int:
    try:
        from PySide6.QtCore import (
            QProcess, QEvent, Qt, QSignalBlocker, QObject, Signal, QPoint, QRect, QSize
        )
        from PySide6.QtGui import QColor, QFont, QPalette, QCursor
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
            QStackedWidget,
            QTabWidget,
            QVBoxLayout,
            QWidget,
        )
    except ImportError:
        print("PySide6 is required for GUI mode. Install with: pip install PySide6")
        return 1

    def _apply_dark_style(app: QApplication) -> None:
        app.setFont(QFont("Segoe UI", 10))
        palette = QPalette()
        palette.setColor(QPalette.Window, QColor("#1a1d23"))
        palette.setColor(QPalette.Base, QColor("#1e2128"))
        palette.setColor(QPalette.AlternateBase, QColor("#1a1d23"))
        palette.setColor(QPalette.Text, QColor("#c8cdd8"))
        palette.setColor(QPalette.WindowText, QColor("#c8cdd8"))
        palette.setColor(QPalette.Button, QColor("#252830"))
        palette.setColor(QPalette.ButtonText, QColor("#c0c6d4"))
        palette.setColor(QPalette.Highlight, QColor("#2d5c9e"))
        palette.setColor(QPalette.HighlightedText, QColor("#d0e4ff"))
        palette.setColor(QPalette.BrightText, QColor("#ffffff"))
        palette.setColor(QPalette.Link, QColor("#4d9cf5"))
        app.setPalette(palette)
        app.setStyleSheet(DARK_STYLE)

    class HotkeyEmitter(QObject):
        fired = Signal(str)

    def _qt_key_name(key: int, text: str) -> str:
        if text and len(text) == 1 and text.isalnum():
            return text.upper()
        if Qt.Key_F1 <= key <= Qt.Key_F24:
            return f"F{key - Qt.Key_F1 + 1}"
        mapping = {
            Qt.Key_Tab: "TAB",
            Qt.Key_Return: "ENTER",
            Qt.Key_Enter: "ENTER",
            Qt.Key_Escape: "ESC",
            Qt.Key_Space: "SPACE",
            Qt.Key_Backspace: "BACKSPACE",
            Qt.Key_Delete: "DELETE",
            Qt.Key_Home: "HOME",
            Qt.Key_End: "END",
            Qt.Key_PageUp: "PGUP",
            Qt.Key_PageDown: "PGDN",
            Qt.Key_Left: "LEFT",
            Qt.Key_Right: "RIGHT",
            Qt.Key_Up: "UP",
            Qt.Key_Down: "DOWN",
        }
        return mapping.get(key, "")

    class HotkeyCaptureLineEdit(QLineEdit):
        def keyPressEvent(self, event) -> None:
            key = event.key()
            if key in (Qt.Key_Control, Qt.Key_Shift, Qt.Key_Alt, Qt.Key_Meta):
                return
            mods = event.modifiers()
            parts = []
            if mods & Qt.ControlModifier:
                parts.append("Ctrl")
            if mods & Qt.AltModifier:
                parts.append("Alt")
            if mods & Qt.ShiftModifier:
                parts.append("Shift")
            if mods & Qt.MetaModifier:
                parts.append("Win")
            key_name = _qt_key_name(key, event.text())
            if not key_name:
                return
            parts.append(key_name)
            self.setText("+".join(parts))
            event.accept()

    # ─────────────────────────────────────────────────────────────────────────
    # FLOATING SPEED MENU POPUP WINDOW
    # ─────────────────────────────────────────────────────────────────────────

    class SpeedMenuPopup(QWidget):
        """Frameless, draggable floating speed menu window."""

        def __init__(self, parent=None):
            super().__init__(parent, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
            self.setObjectName("speedPopup")
            self.setAttribute(Qt.WA_TranslucentBackground, False)
            self._drag_pos = None
            self._run_callback = None
            self._items: List[SpeedMenuItem] = []

            outer = QVBoxLayout(self)
            outer.setContentsMargins(1, 1, 1, 1)
            outer.setSpacing(0)

            # Header / drag handle
            header = QWidget()
            header.setObjectName("speedPopupHeader")
            header.setFixedHeight(34)
            h_lay = QHBoxLayout(header)
            h_lay.setContentsMargins(10, 0, 10, 0)

            handle_lbl = QLabel("⠿")
            handle_lbl.setObjectName("speedHandle")
            handle_lbl.setFixedWidth(20)
            h_lay.addWidget(handle_lbl)

            title_lbl = QLabel("SPEED MENU")
            title_lbl.setObjectName("speedTitle")
            h_lay.addWidget(title_lbl)
            h_lay.addStretch(1)

            self._pin_btn = QPushButton("📌")
            self._pin_btn.setObjectName("speedLaunchBtn")
            self._pin_btn.setFixedSize(28, 24)
            self._pin_btn.setToolTip("Pin / unpin on top")
            self._pin_btn.clicked.connect(self._toggle_pin)
            h_lay.addWidget(self._pin_btn)

            close_btn = QPushButton("✕")
            close_btn.setObjectName("speedLaunchBtn")
            close_btn.setFixedSize(28, 24)
            close_btn.setToolTip("Close speed menu")
            close_btn.clicked.connect(self.hide)
            h_lay.addWidget(close_btn)

            outer.addWidget(header)

            # Thin separator
            sep = QFrame()
            sep.setFrameShape(QFrame.HLine)
            sep.setObjectName("divider")
            outer.addWidget(sep)

            # Buttons container (scrollable)
            self._scroll = QScrollArea()
            self._scroll.setWidgetResizable(True)
            self._scroll.setFrameShape(QFrame.NoFrame)
            self._scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
            self._btn_widget = QWidget()
            self._btn_layout = QGridLayout(self._btn_widget)
            self._btn_layout.setContentsMargins(8, 8, 8, 8)
            self._btn_layout.setHorizontalSpacing(7)
            self._btn_layout.setVerticalSpacing(7)
            self._scroll.setWidget(self._btn_widget)
            outer.addWidget(self._scroll)

            # Bottom status bar
            self._status_lbl = QLabel("Ready")
            self._status_lbl.setObjectName("speedTitle")
            self._status_lbl.setContentsMargins(10, 4, 10, 4)
            outer.addWidget(self._status_lbl)

            self.resize(320, 260)
            self._pinned = True

        def _toggle_pin(self):
            self._pinned = not self._pinned
            flags = Qt.Tool | Qt.FramelessWindowHint
            if self._pinned:
                flags |= Qt.WindowStaysOnTopHint
            self.setWindowFlags(flags)
            self.show()

        def set_run_callback(self, fn):
            self._run_callback = fn

        def set_status(self, text: str):
            self._status_lbl.setText(text)

        def populate(self, items: List[SpeedMenuItem]):
            self._items = items
            # Clear
            while self._btn_layout.count():
                w = self._btn_layout.takeAt(0).widget()
                if w:
                    w.deleteLater()

            if not items:
                lbl = QLabel("No speed menu items.\nConfigure in Speed Menu Editor.")
                lbl.setAlignment(Qt.AlignCenter)
                lbl.setObjectName("speedTitle")
                self._btn_layout.addWidget(lbl, 0, 0)
                return

            cols = 3 if len(items) > 4 else (2 if len(items) > 1 else 1)
            for idx, item in enumerate(items):
                label = item.label or Path(item.layout).stem or "Restore"
                btn_text = f"{item.emoji}\n{label}" if item.emoji else label
                btn = QPushButton(btn_text)
                btn.setObjectName("speedBtn")
                btn.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
                btn.setMinimumHeight(60)
                target = _resolve_speed_layout(item.layout)
                tip = f"{target} {' '.join(item.args)}".strip()
                btn.setToolTip(tip)
                btn.clicked.connect(lambda _=False, e=item: self._on_click(e))
                r, c = divmod(idx, cols)
                self._btn_layout.addWidget(btn, r, c)

            for c in range(cols):
                self._btn_layout.setColumnStretch(c, 1)

        def _on_click(self, entry: SpeedMenuItem):
            if self._run_callback:
                self._run_callback(entry)

        # ── drag support ──
        def mousePressEvent(self, event):
            if event.button() == Qt.LeftButton:
                self._drag_pos = event.globalPosition().toPoint() - self.frameGeometry().topLeft()
                event.accept()

        def mouseMoveEvent(self, event):
            if self._drag_pos is not None and event.buttons() & Qt.LeftButton:
                self.move(event.globalPosition().toPoint() - self._drag_pos)
                event.accept()

        def mouseReleaseEvent(self, event):
            self._drag_pos = None

    # ─────────────────────────────────────────────────────────────────────────
    # MAIN WINDOW
    # ─────────────────────────────────────────────────────────────────────────

    class MainWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("Window Layout Manager")
            self.resize(820, 600)
            self.setMinimumSize(*MIN_SETTINGS_SIZE)
            _ensure_layouts_root()

            self._proc = QProcess(self)
            self._proc.readyReadStandardOutput.connect(self._append_stdout)
            self._proc.readyReadStandardError.connect(self._append_stderr)
            self._proc.finished.connect(self._on_finished)
            self._speed_item_cache: dict[str, SpeedMenuItem] = {}

            # ── Floating speed popup (lives independently) ──
            self._speed_popup = SpeedMenuPopup()
            self._speed_popup.set_run_callback(self._run_speed)

            # ── Root layout: sidebar + content ──
            root = QWidget(self)
            self.setCentralWidget(root)
            root_h = QHBoxLayout(root)
            root_h.setContentsMargins(0, 0, 0, 0)
            root_h.setSpacing(0)

            # Sidebar
            sidebar = QWidget()
            sidebar.setObjectName("sidebar")
            sidebar.setFixedWidth(170)
            sb_lay = QVBoxLayout(sidebar)
            sb_lay.setContentsMargins(8, 12, 8, 12)
            sb_lay.setSpacing(3)

            app_title = QLabel("⚡ Layout Mgr")
            app_title.setStyleSheet("color: #5a7ab8; font-size: 11pt; font-weight: 700; padding: 4px 6px 12px 6px; background: transparent;")
            sb_lay.addWidget(app_title)

            self._nav_btns: List[QPushButton] = []
            nav_items = [
                ("🗂  Workspace", "workspace"),
                ("▶  Actions", "actions"),
                ("⌨  Hotkeys", "hotkeys"),
                ("✏  Speed Editor", "speed_editor"),
                ("🔧  Layout Editor", "layout_editor"),
            ]
            for label, name in nav_items:
                btn = QPushButton(label)
                btn.setObjectName("navBtn")
                btn.setProperty("page", name)
                btn.clicked.connect(lambda _=False, n=name: self._nav_to(n))
                sb_lay.addWidget(btn)
                self._nav_btns.append(btn)

            sb_lay.addStretch(1)

            # Speed Menu toggle button at bottom of sidebar
            speed_toggle_btn = QPushButton("⚡  Speed Menu")
            speed_toggle_btn.setObjectName("navBtn")
            speed_toggle_btn.setToolTip("Open / close the floating speed menu")
            speed_toggle_btn.clicked.connect(self._toggle_speed_popup)
            sb_lay.addWidget(speed_toggle_btn)

            self._tray_label = QLabel()
            self._tray_label.setStyleSheet("color: #3a4050; font-size: 8pt; padding: 4px 6px 0 6px; background: transparent;")
            sb_lay.addWidget(self._tray_label)

            root_h.addWidget(sidebar)

            # Content area (stacked pages)
            self._stack = QStackedWidget()
            root_h.addWidget(self._stack, 1)

            # ── Build pages ──
            self._page_workspace = self._build_workspace_page()
            self._page_actions = self._build_actions_page()
            self._page_hotkeys = self._build_hotkeys_page()
            self._page_speed_editor = self._build_speed_editor_page()
            self._page_layout_editor = self._build_layout_editor_page()

            self._stack.addWidget(self._page_workspace)
            self._stack.addWidget(self._page_actions)
            self._stack.addWidget(self._page_hotkeys)
            self._stack.addWidget(self._page_speed_editor)
            self._stack.addWidget(self._page_layout_editor)

            self._page_names = ["workspace", "actions", "hotkeys", "speed_editor", "layout_editor"]

            # ── Init state ──
            self._speed_dirty = False
            self._layout_dirty = False
            self._last_page = "workspace"
            self._layout_edit_data = None
            self._layout_removed_cache = None
            self._layout_edit_name = ""
            self._layout_edit_loading = False
            self._speed_edit_loading = False
            self._hotkey_thread = None
            self._hotkey_thread_id = None
            self._tray_enabled = True
            self._tray_icon = None
            self._hotkey_emitter = HotkeyEmitter()
            self._hotkey_emitter.fired.connect(self._log_hotkey_fire)

            self._load_layouts_root_field()
            self._reload_layout_choices()
            self._sync_hotkey_fields()
            self._reload_speed_menu()
            self.hotkeys_enabled.setChecked(_get_hotkeys_enabled())
            if self.hotkeys_enabled.isChecked():
                self._start_hotkeys()
            self._init_tray_icon()
            self._nav_to("workspace")

        # ─────────────────────────────────────────────────────────────────────
        # PAGE BUILDERS
        # ─────────────────────────────────────────────────────────────────────

        def _page_wrap(self, title: str) -> tuple[QWidget, QVBoxLayout]:
            page = QWidget()
            lay = QVBoxLayout(page)
            lay.setContentsMargins(20, 18, 20, 16)
            lay.setSpacing(12)
            hdr = QLabel(title)
            hdr.setObjectName("pageTitle")
            lay.addWidget(hdr)
            sep = QFrame()
            sep.setObjectName("divider")
            sep.setFixedHeight(1)
            lay.addWidget(sep)
            return page, lay

        def _section(self, text: str) -> QLabel:
            lbl = QLabel(text.upper())
            lbl.setObjectName("sectionHeader")
            return lbl

        def _card(self, layout: QVBoxLayout | QGridLayout | QHBoxLayout) -> QFrame:
            card = QFrame()
            card.setObjectName("card")
            card.setLayout(layout)
            return card

        def _build_workspace_page(self) -> QWidget:
            page, lay = self._page_wrap("Workspace")

            # Layouts root row
            root_card_lay = QVBoxLayout()
            root_card_lay.setSpacing(8)
            root_card_lay.addWidget(self._section("Layouts Folder"))
            root_row = QHBoxLayout()
            self.layouts_root_input = QLineEdit()
            root_browse_btn = QPushButton("Browse…")
            root_browse_btn.setFixedWidth(80)
            root_browse_btn.clicked.connect(self._browse_layouts_root)
            self.layouts_root_input.editingFinished.connect(self._save_layouts_root)
            root_row.addWidget(self.layouts_root_input)
            root_row.addWidget(root_browse_btn)
            root_card_lay.addLayout(root_row)
            lay.addWidget(self._card(root_card_lay))

            # Active layout
            layout_card_lay = QVBoxLayout()
            layout_card_lay.setSpacing(8)
            layout_card_lay.addWidget(self._section("Active Layout"))
            self.layout_select = QComboBox()
            self.layout_select.currentIndexChanged.connect(self._sync_layout_editor_choice)
            layout_card_lay.addWidget(self.layout_select)
            lay.addWidget(self._card(layout_card_lay))

            # New layout
            new_card_lay = QVBoxLayout()
            new_card_lay.setSpacing(8)
            new_card_lay.addWidget(self._section("Create New Layout"))
            new_row = QHBoxLayout()
            self.new_layout_input = QLineEdit()
            self.new_layout_input.setPlaceholderText("my-layout.json")
            create_btn = QPushButton("Create")
            create_btn.setObjectName("accentBtn")
            create_btn.setFixedWidth(70)
            create_btn.clicked.connect(self._create_layout)
            new_row.addWidget(self.new_layout_input)
            new_row.addWidget(create_btn)
            new_card_lay.addLayout(new_row)
            lay.addWidget(self._card(new_card_lay))

            lay.addStretch(1)
            return page

        def _build_actions_page(self) -> QWidget:
            page, lay = self._page_wrap("Actions")

            actions = [
                ("💾  Save", "save", "primaryBtn"),
                ("💾  Save + Edge Tabs", "save_edge", "primaryBtn"),
                ("🔄  Restore (Existing)", "restore_basic", "accentBtn"),
                ("🔄  Restore + Launch", "restore_launch", "accentBtn"),
                ("🔄  Restore + Edge", "restore_edge", "accentBtn"),
                ("🔄  Restore + Edge (Dest.)", "restore_edge_destructive", "dangerBtn"),
                ("🔄  Restore + Launch + Edge", "restore_launch_edge", "accentBtn"),
                ("🌐  Edge Debug Session", "edge_debug", None),
                ("📷  Edge Capture Tabs", "edge_capture", None),
            ]

            grid_card_lay = QGridLayout()
            grid_card_lay.setSpacing(8)
            for idx, (title, action, obj_name) in enumerate(actions):
                btn = QPushButton(title)
                if obj_name:
                    btn.setObjectName(obj_name)
                btn.setMinimumHeight(36)
                btn.clicked.connect(lambda _=False, a=action: self._run(a))
                r, c = divmod(idx, 2)
                grid_card_lay.addWidget(btn, r, c)

            lay.addWidget(self._card(grid_card_lay))

            # Log area
            log_card_lay = QVBoxLayout()
            log_card_lay.setSpacing(6)
            log_hdr = QHBoxLayout()
            log_hdr.addWidget(self._section("Output Log"))
            log_hdr.addStretch(1)
            clear_btn = QPushButton("Clear")
            clear_btn.setFixedWidth(60)
            clear_btn.clicked.connect(lambda: self.log.clear())
            log_hdr.addWidget(clear_btn)
            log_card_lay.addLayout(log_hdr)
            self.log = QPlainTextEdit()
            self.log.setReadOnly(True)
            self.log.setMinimumHeight(140)
            log_card_lay.addWidget(self.log)
            self.status = QLabel("Ready")
            self.status.setObjectName("statusLabel")
            log_card_lay.addWidget(self.status)
            lay.addWidget(self._card(log_card_lay), 1)

            return page

        def _build_hotkeys_page(self) -> QWidget:
            page, lay = self._page_wrap("Hotkeys")

            # Enable toggle
            toggle_card_lay = QVBoxLayout()
            self.hotkeys_enabled = QCheckBox("Enable Global Hotkeys")
            self.hotkeys_enabled.stateChanged.connect(self._toggle_hotkeys)
            toggle_card_lay.addWidget(self.hotkeys_enabled)
            lay.addWidget(self._card(toggle_card_lay))

            # New hotkey entry
            entry_card_lay = QVBoxLayout()
            entry_card_lay.setSpacing(8)
            entry_card_lay.addWidget(self._section("Register Hotkey"))

            row1 = QHBoxLayout()
            row1.addWidget(QLabel("Action:"))
            self.hotkey_action_select = QComboBox()
            self.hotkey_action_select.addItems(["save", "restore"])
            self.hotkey_action_select.currentIndexChanged.connect(self._sync_hotkey_fields)
            row1.addWidget(self.hotkey_action_select, 1)
            row1.addWidget(QLabel("Layout:"))
            self.hotkey_layout_select = QComboBox()
            row1.addWidget(self.hotkey_layout_select, 2)
            entry_card_lay.addLayout(row1)

            row2 = QHBoxLayout()
            row2.addWidget(QLabel("Args:"))
            self.hotkey_args_select = QComboBox()
            self.hotkey_args_select.addItems([
                "Existing Only", "Launch Missing", "Edge Tabs",
                "Edge Tabs (Destructive)", "Launch Missing + Edge Tabs",
            ])
            row2.addWidget(self.hotkey_args_select, 2)
            row2.addWidget(QLabel("Keys:"))
            self.hotkey_input = HotkeyCaptureLineEdit()
            self.hotkey_input.setPlaceholderText("Click and press keys…")
            row2.addWidget(self.hotkey_input, 2)
            entry_card_lay.addLayout(row2)

            save_hk_btn = QPushButton("Save Hotkey")
            save_hk_btn.setObjectName("primaryBtn")
            save_hk_btn.setFixedWidth(120)
            save_hk_btn.clicked.connect(self._save_hotkey_entry)
            entry_card_lay.addWidget(save_hk_btn, 0, Qt.AlignRight)
            lay.addWidget(self._card(entry_card_lay))

            lay.addStretch(1)
            return page

        def _build_speed_editor_page(self) -> QWidget:
            page, lay = self._page_wrap("Speed Menu Editor")

            # Lists
            lists_card_lay = QHBoxLayout()
            lists_card_lay.setSpacing(10)

            avail_col = QVBoxLayout()
            avail_col.addWidget(self._section("Available Layouts"))
            self.available_list = QListWidget()
            self.available_list.setSelectionMode(QAbstractItemView.SingleSelection)
            avail_col.addWidget(self.available_list, 1)
            lists_card_lay.addLayout(avail_col, 1)

            mid_col = QVBoxLayout()
            mid_col.addStretch(1)
            add_btn = QPushButton("→")
            add_btn.setFixedWidth(36)
            add_btn.setObjectName("accentBtn")
            add_btn.clicked.connect(self._move_available_selected)
            rem_btn = QPushButton("←")
            rem_btn.setFixedWidth(36)
            rem_btn.clicked.connect(self._move_speed_selected)
            mid_col.addWidget(add_btn)
            mid_col.addSpacing(4)
            mid_col.addWidget(rem_btn)
            mid_col.addStretch(1)
            lists_card_lay.addLayout(mid_col)

            speed_col = QVBoxLayout()
            speed_col.addWidget(self._section("Speed Menu"))
            self.speed_list = QListWidget()
            self.speed_list.setSelectionMode(QAbstractItemView.SingleSelection)
            self.speed_list.itemSelectionChanged.connect(self._speed_selection_changed)
            speed_col.addWidget(self.speed_list, 1)
            lists_card_lay.addLayout(speed_col, 1)

            lay.addWidget(self._card(lists_card_lay), 2)

            # Item detail editor
            detail_card_lay = QVBoxLayout()
            detail_card_lay.setSpacing(8)
            detail_card_lay.addWidget(self._section("Item Properties"))
            detail_row = QHBoxLayout()
            self.speed_label_input = QLineEdit()
            self.speed_label_input.setPlaceholderText("Label")
            self.speed_label_input.textChanged.connect(self._apply_speed_item_edits)
            self.speed_emoji_input = QLineEdit()
            self.speed_emoji_input.setPlaceholderText("Emoji")
            self.speed_emoji_input.setFixedWidth(70)
            self.speed_emoji_input.textChanged.connect(self._apply_speed_item_edits)
            self.speed_args_preset = QComboBox()
            self.speed_args_preset.addItems([
                "Custom", "Existing Only", "Launch Missing",
                "Edge Tabs", "Edge Tabs (Destructive)",
            ])
            self.speed_args_preset.currentIndexChanged.connect(self._apply_speed_args_preset)
            self.speed_args_input = QLineEdit()
            self.speed_args_input.setPlaceholderText("Args (e.g. --edge-tabs)")
            self.speed_args_input.textChanged.connect(self._apply_speed_item_edits)
            detail_row.addWidget(self.speed_label_input, 2)
            detail_row.addWidget(self.speed_emoji_input)
            detail_row.addWidget(self.speed_args_preset, 2)
            detail_row.addWidget(self.speed_args_input, 2)
            detail_card_lay.addLayout(detail_row)
            lay.addWidget(self._card(detail_card_lay))

            # Buttons
            btn_row = QHBoxLayout()
            btn_row.addStretch(1)
            reload_btn = QPushButton("Reload")
            reload_btn.clicked.connect(self._reload_speed_menu)
            save_btn = QPushButton("Save Speed Menu")
            save_btn.setObjectName("primaryBtn")
            save_btn.clicked.connect(self._save_speed_menu)
            btn_row.addWidget(reload_btn)
            btn_row.addWidget(save_btn)
            lay.addLayout(btn_row)

            return page

        def _build_layout_editor_page(self) -> QWidget:
            page, lay = self._page_wrap("Layout Editor")

            # Layout selection + view mode
            sel_card_lay = QHBoxLayout()
            sel_card_lay.addWidget(QLabel("Layout:"))
            self.layout_editor_select = QComboBox()
            self.layout_editor_select.currentIndexChanged.connect(self._sync_layout_settings_choice)
            sel_card_lay.addWidget(self.layout_editor_select, 2)
            sel_card_lay.addSpacing(16)
            sel_card_lay.addWidget(QLabel("View:"))
            self.layout_view_select = QComboBox()
            self.layout_view_select.addItems(["Simple", "Advanced"])
            self.layout_view_select.currentIndexChanged.connect(self._set_layout_view)
            sel_card_lay.addWidget(self.layout_view_select)
            lay.addWidget(self._card(sel_card_lay))

            # Split: window list + fields
            split = QHBoxLayout()
            split.setSpacing(10)

            left_col = QVBoxLayout()
            left_col.addWidget(self._section("Windows"))
            self.layout_windows_list = QListWidget()
            self.layout_windows_list.setSelectionMode(QAbstractItemView.SingleSelection)
            self.layout_windows_list.itemSelectionChanged.connect(self._layout_window_selected)
            self.layout_windows_list.setMinimumWidth(180)
            left_col.addWidget(self.layout_windows_list, 1)
            split.addLayout(left_col)

            # Fields panel (scrollable)
            fields_panel = QWidget()
            fields_layout = QGridLayout(fields_panel)
            fields_layout.setHorizontalSpacing(10)
            fields_layout.setVerticalSpacing(7)
            fields_layout.setContentsMargins(0, 0, 0, 0)

            self.le_title = QLineEdit()
            self.le_class = QLineEdit()
            self.le_process = QLineEdit()
            self.le_exe = QLineEdit()
            self.le_window_id = QLineEdit()
            self.le_window_id.setReadOnly(True)

            self.rect_left = QSpinBox(); self.rect_top = QSpinBox()
            self.rect_right = QSpinBox(); self.rect_bottom = QSpinBox()
            self.nrect_left = QSpinBox(); self.nrect_top = QSpinBox()
            self.nrect_right = QSpinBox(); self.nrect_bottom = QSpinBox()
            for w in (self.rect_left, self.rect_top, self.rect_right, self.rect_bottom,
                      self.nrect_left, self.nrect_top, self.nrect_right, self.nrect_bottom):
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
            self.chk_destructive = QCheckBox("Destructive")

            self.edge_tabs_list = QListWidget()
            self.edge_tabs_list.setSelectionMode(QAbstractItemView.SingleSelection)
            self.edge_tabs_list.setMaximumHeight(100)

            def _add_row(lay, row, lbl_text, widget, col_span=3):
                lbl = QLabel(lbl_text)
                lbl.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
                lay.addWidget(lbl, row, 0)
                if isinstance(widget, QWidget):
                    lay.addWidget(widget, row, 1, 1, col_span)
                else:
                    lay.addLayout(widget, row, 1, 1, col_span)
                return lbl

            r = 0
            _add_row(fields_layout, r, "Title:", self.le_title); r += 1

            label_class = _add_row(fields_layout, r, "Class:", self.le_class)
            label_process = QLabel("Process:"); label_process.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            fields_layout.addWidget(label_process, r, 2)
            fields_layout.addWidget(self.le_process, r, 3); r += 1

            label_exe = _add_row(fields_layout, r, "Exe:", self.le_exe); r += 1
            label_window_id = _add_row(fields_layout, r, "Window ID:", self.le_window_id); r += 1

            rect_row = QHBoxLayout()
            for w in (self.rect_left, self.rect_top, self.rect_right, self.rect_bottom):
                rect_row.addWidget(w)
            label_rect = QLabel("Rect L/T/R/B:")
            label_rect.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            fields_layout.addWidget(label_rect, r, 0)
            rect_wid = QWidget(); rect_wid.setLayout(rect_row)
            fields_layout.addWidget(rect_wid, r, 1, 1, 3); r += 1

            nrect_row = QHBoxLayout()
            for w in (self.nrect_left, self.nrect_top, self.nrect_right, self.nrect_bottom):
                nrect_row.addWidget(w)
            label_nrect = QLabel("Normal L/T/R/B:")
            label_nrect.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            fields_layout.addWidget(label_nrect, r, 0)
            nrect_wid = QWidget(); nrect_wid.setLayout(nrect_row)
            fields_layout.addWidget(nrect_wid, r, 1, 1, 3); r += 1

            label_show_cmd = QLabel("Show Cmd:"); label_show_cmd.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            fields_layout.addWidget(label_show_cmd, r, 0)
            fields_layout.addWidget(self.spin_show_cmd, r, 1)
            flags_row = QHBoxLayout()
            flags_row.addWidget(self.chk_visible); flags_row.addWidget(self.chk_minimized); flags_row.addWidget(self.chk_maximized)
            flags_wid = QWidget(); flags_wid.setLayout(flags_row)
            fields_layout.addWidget(flags_wid, r, 2, 1, 2); r += 1

            label_launch_exe = _add_row(fields_layout, r, "Launch Exe:", self.le_launch_exe); r += 1
            label_launch_args = _add_row(fields_layout, r, "Launch Args:", self.le_launch_args); r += 1
            label_launch_cwd = _add_row(fields_layout, r, "Launch CWD:", self.le_launch_cwd); r += 1
            label_edge_port = _add_row(fields_layout, r, "Edge Port:", self.spin_edge_port); r += 1

            label_destructive = QLabel("Destructive:"); label_destructive.setAlignment(Qt.AlignRight | Qt.AlignVCenter)
            fields_layout.addWidget(label_destructive, r, 0)
            fields_layout.addWidget(self.chk_destructive, r, 1); r += 1

            label_edge_tabs = _add_row(fields_layout, r, "Edge Tabs:", self.edge_tabs_list); r += 1

            edge_btns_row = QHBoxLayout()
            self.edge_tab_add = QPushButton("Add Tab")
            self.edge_tab_remove = QPushButton("Remove Tab")
            self.edge_tab_add.clicked.connect(self._add_edge_tab)
            self.edge_tab_remove.clicked.connect(self._remove_edge_tab)
            edge_btns_row.addWidget(self.edge_tab_add); edge_btns_row.addWidget(self.edge_tab_remove); edge_btns_row.addStretch(1)
            edge_btns_wid = QWidget(); edge_btns_wid.setLayout(edge_btns_row)
            fields_layout.addWidget(edge_btns_wid, r, 1, 1, 3); r += 1

            self._advanced_widgets = [
                label_class, self.le_class, label_process, self.le_process,
                label_exe, self.le_exe, label_window_id, self.le_window_id,
                label_rect, rect_wid, label_nrect, nrect_wid,
                label_show_cmd, self.spin_show_cmd, flags_wid,
                label_launch_args, self.le_launch_args, label_launch_cwd, self.le_launch_cwd,
                label_edge_port, self.spin_edge_port,
            ]
            self._set_layout_view()

            # Connect dirty signals
            for sig_widget in [self.le_title, self.le_class, self.le_process, self.le_exe,
                                self.le_launch_exe, self.le_launch_args, self.le_launch_cwd]:
                sig_widget.textChanged.connect(self._mark_layout_dirty)
            for sb in [self.rect_left, self.rect_top, self.rect_right, self.rect_bottom,
                       self.nrect_left, self.nrect_top, self.nrect_right, self.nrect_bottom,
                       self.spin_show_cmd, self.spin_edge_port]:
                sb.valueChanged.connect(self._mark_layout_dirty)
            for chk in [self.chk_visible, self.chk_minimized, self.chk_maximized, self.chk_destructive]:
                chk.stateChanged.connect(self._mark_layout_dirty)

            fields_scroll = QScrollArea()
            fields_scroll.setWidgetResizable(True)
            fields_scroll.setFrameShape(QFrame.NoFrame)
            fields_scroll.setWidget(fields_panel)
            split.addWidget(fields_scroll, 2)

            lay.addLayout(split, 1)

            # Bottom buttons
            edit_btns = QHBoxLayout()
            self.layout_reload_btn = QPushButton("Reload")
            self.layout_save_btn = QPushButton("Save Layout")
            self.layout_save_btn.setObjectName("primaryBtn")
            self.layout_remove_btn = QPushButton("Remove Window")
            self.layout_remove_btn.setObjectName("dangerBtn")
            self.layout_restore_btn = QPushButton("Restore Removed")
            for btn in (self.layout_reload_btn, self.layout_save_btn):
                edit_btns.addWidget(btn)
            edit_btns.addStretch(1)
            for btn in (self.layout_remove_btn, self.layout_restore_btn):
                edit_btns.addWidget(btn)

            self.layout_reload_btn.clicked.connect(self._load_layout_for_editing)
            self.layout_save_btn.clicked.connect(self._save_layout_edit)
            self.layout_remove_btn.clicked.connect(self._remove_selected_window)
            self.layout_restore_btn.clicked.connect(self._restore_removed_window)

            lay.addLayout(edit_btns)
            return page

        # ─────────────────────────────────────────────────────────────────────
        # NAVIGATION
        # ─────────────────────────────────────────────────────────────────────

        def _nav_to(self, name: str) -> None:
            if name == self._last_page:
                pass
            else:
                # Dirty checks when leaving editors
                if self._last_page == "speed_editor" and self._speed_dirty:
                    if not self._confirm_unsaved("Speed Menu Editor", self._save_speed_menu, self._discard_speed_menu_changes):
                        return
                if self._last_page == "layout_editor" and self._layout_dirty:
                    if not self._confirm_unsaved("Layout Editor", self._save_layout_edit, self._discard_layout_changes):
                        return

            if name == "layout_editor":
                self._load_layout_for_editing()

            idx = self._page_names.index(name) if name in self._page_names else 0
            self._stack.setCurrentIndex(idx)
            self._last_page = name

            for btn in self._nav_btns:
                btn.setProperty("active", btn.property("page") == name)
                btn.style().unpolish(btn)
                btn.style().polish(btn)

        def _toggle_speed_popup(self):
            if self._speed_popup.isVisible():
                self._speed_popup.hide()
            else:
                # Position near the main window bottom-left if no saved pos
                geo = self.geometry()
                self._speed_popup.move(geo.left() + 10, geo.bottom() - self._speed_popup.height() - 40)
                self._speed_popup.show()
                self._speed_popup.raise_()

        # ─────────────────────────────────────────────────────────────────────
        # TRAY
        # ─────────────────────────────────────────────────────────────────────

        def _init_tray_icon(self) -> None:
            if not QSystemTrayIcon.isSystemTrayAvailable():
                self._tray_enabled = False
                return
            self._tray_icon = QSystemTrayIcon(self)
            icon = self.style().standardIcon(QStyle.SP_ComputerIcon)
            self.setWindowIcon(icon)
            self._tray_icon.setIcon(icon)
            self._tray_icon.setToolTip("Window Layout Manager")
            menu = QMenu()
            show_action = menu.addAction("Show")
            hide_action = menu.addAction("Hide")
            speed_action = menu.addAction("Speed Menu")
            menu.addSeparator()
            quit_action = menu.addAction("Quit")
            show_action.triggered.connect(self._show_from_tray)
            hide_action.triggered.connect(self._hide_to_tray)
            speed_action.triggered.connect(self._toggle_speed_popup)
            quit_action.triggered.connect(self._quit_from_tray)
            self._tray_icon.setContextMenu(menu)
            self._tray_icon.activated.connect(self._on_tray_activated)
            self._tray_icon.show()
            self._tray_label.setText("↓ Tray active")

        def _show_from_tray(self) -> None:
            self.show(); self.raise_(); self.activateWindow()

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

        def changeEvent(self, event) -> None:
            if event.type() == QEvent.WindowStateChange:
                if self._tray_enabled and self._tray_icon is not None:
                    if self.isMinimized():
                        self._hide_to_tray()
                        event.ignore()
                        return
            super().changeEvent(event)

        # ─────────────────────────────────────────────────────────────────────
        # WORKSPACE
        # ─────────────────────────────────────────────────────────────────────

        def _load_layouts_root_field(self) -> None:
            self.layouts_root_input.setText(_get_layouts_root())

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

            for combo, attr in [
                (self.layout_select, None),
                (self.layout_editor_select, None),
                (self.hotkey_layout_select, None),
            ]:
                current = combo.currentText()
                combo.blockSignals(True)
                combo.clear()
                for name in layouts:
                    combo.addItem(name)
                if current in layouts:
                    combo.setCurrentText(current)
                elif layouts:
                    combo.setCurrentIndex(0)
                combo.blockSignals(False)

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
                    json.dump({
                        "schema": "window-layout.v2",
                        "created_at": time.strftime("%Y-%m-%d %H:%M:%S"),
                        "windows": [], "edge_sessions": [],
                        "open_urls": {"edge": []},
                    }, f, indent=2, ensure_ascii=False)
            except Exception as exc:
                QMessageBox.warning(self, "New Layout", f"Failed to create: {exc}")
                return
            self.new_layout_input.setText("")
            self._reload_layout_choices()
            self.layout_select.setCurrentText(name)

        # ─────────────────────────────────────────────────────────────────────
        # HOTKEYS
        # ─────────────────────────────────────────────────────────────────────

        def _sync_hotkey_fields(self) -> None:
            is_restore = self.hotkey_action_select.currentText().strip().lower() == "restore"
            self.hotkey_args_select.setEnabled(is_restore)

        def _save_hotkey_entry(self) -> None:
            keys = self.hotkey_input.text().strip()
            if not keys:
                QMessageBox.information(self, "Hotkeys", "Press a hotkey first.")
                return
            action = self.hotkey_action_select.currentText().strip().lower()
            layout = self.hotkey_layout_select.currentText().strip()
            if not layout:
                QMessageBox.information(self, "Hotkeys", "Select a layout.")
                return
            layout_path = _resolve_speed_layout(layout)
            args_map = {
                "Existing Only": [],
                "Launch Missing": ["--launch-missing"],
                "Edge Tabs": ["--edge-tabs"],
                "Edge Tabs (Destructive)": ["--edge-tabs", "--destructive"],
                "Launch Missing + Edge Tabs": ["--launch-missing", "--edge-tabs"],
            }
            restore_args = args_map.get(self.hotkey_args_select.currentText(), [])
            entry = {"keys": keys, "action": action, "args": [layout_path]}
            if action == "restore":
                entry["args"] = [layout_path, *restore_args]
            data = _load_config()
            hotkeys = data.get("hotkeys") or []
            if not isinstance(hotkeys, list):
                hotkeys = []
            replaced = False
            for i, existing in enumerate(hotkeys):
                if not isinstance(existing, dict):
                    continue
                if str(existing.get("keys") or "") == keys:
                    hotkeys[i] = entry
                    replaced = True
                    break
            if not replaced:
                hotkeys.append(entry)
            data["hotkeys"] = hotkeys
            try:
                with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
            except Exception as exc:
                QMessageBox.warning(self, "Hotkeys", f"Failed to save hotkey: {exc}")
                return
            self.status.setText(f"Saved hotkey: {keys} → {action}")

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

        def _start_hotkeys(self) -> None:
            if self._hotkey_thread is not None:
                return
            hotkeys = _load_hotkeys(CONFIG_PATH)
            if not hotkeys:
                return
            self._stop_hotkeys()

            def worker():
                import win32con, win32gui
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
                    parsed = wl._parse_hotkey(entry["keys"])
                    if not parsed:
                        continue
                    modifiers, vk = parsed
                    try:
                        win32gui.RegisterHotKey(None, next_id, modifiers, vk)
                        registered[next_id] = entry
                        self._hotkey_emitter.fired.emit(f"Hotkey registered: {entry['keys']} → {entry['action']}")
                        next_id += 1
                    except Exception as exc:
                        self._hotkey_emitter.fired.emit(f"Hotkey failed: {entry['keys']} ({exc})")
                        continue
                if not registered:
                    return
                while True:
                    msg = win32gui.GetMessage(None, 0, 0)
                    if not msg:
                        continue
                    payload = msg
                    if isinstance(msg, (list, tuple)) and len(msg) == 2:
                        payload = msg[1]
                    message = wparam = None
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
                            self._hotkey_emitter.fired.emit(f"Hotkey: {entry['keys']} → {entry['action']}")
                            _run_hotkey_action(entry["action"], entry.get("args", []))

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

        # ─────────────────────────────────────────────────────────────────────
        # SPEED MENU
        # ─────────────────────────────────────────────────────────────────────

        def _reload_speed_menu(self, force: bool = False) -> None:
            if not force and getattr(self, "_speed_dirty", False):
                if not self._confirm_unsaved("Speed Menu Editor", self._save_speed_menu, self._discard_speed_menu_changes):
                    return
            items = _parse_speed_menu(CONFIG_PATH)
            self._speed_menu_items = items
            self._speed_popup.populate(items)
            self._load_speed_menu_editor(items)
            self._speed_dirty = False

        def _load_speed_menu_editor(self, items: List[SpeedMenuItem]) -> None:
            self._speed_edit_loading = True
            self.available_list.clear()
            self.speed_list.clear()
            self._speed_item_cache = {
                item.layout: SpeedMenuItem(item.label, item.emoji, item.layout, list(item.args))
                for item in items if item.layout
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
                return sorted(
                    [f for f in os.listdir(root) if f.lower().endswith(".json") and os.path.isfile(os.path.join(root, f))],
                    key=str.lower
                )
            except FileNotFoundError:
                return []

        def _collect_speed_items_from_list(self) -> List[SpeedMenuItem]:
            items = []
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
            return f"{item.emoji} {base}" if item.emoji else base

        def _move_available_to_speed(self, entry: QListWidgetItem) -> None:
            layout = entry.data(Qt.UserRole)
            if not layout:
                return
            cached = self._speed_item_cache.get(layout)
            item = SpeedMenuItem(cached.label, cached.emoji, cached.layout, list(cached.args)) if cached else SpeedMenuItem("", "", layout, [])
            self._add_speed_item_to_list(item)
            self.available_list.takeItem(self.available_list.row(entry))
            self.available_list.sortItems()
            self.speed_list.setCurrentRow(self.speed_list.count() - 1)
            self._speed_item_selected(self.speed_list.currentItem())
            self._update_speed_popup_preview()
            self._speed_dirty = True

        def _move_speed_to_available(self, entry: QListWidgetItem) -> None:
            item = entry.data(Qt.UserRole)
            if not isinstance(item, SpeedMenuItem):
                return
            if item.layout:
                self._speed_item_cache[item.layout] = SpeedMenuItem(item.label, item.emoji, item.layout, list(item.args))
            list_blocker = QSignalBlocker(self.speed_list)
            label_blocker = QSignalBlocker(self.speed_label_input)
            emoji_blocker = QSignalBlocker(self.speed_emoji_input)
            args_blocker = QSignalBlocker(self.speed_args_input)
            if item.layout:
                available_entry = QListWidgetItem(item.layout)
                available_entry.setData(Qt.UserRole, item.layout)
                self.available_list.addItem(available_entry)
                self.available_list.sortItems()
            self.speed_list.takeItem(self.speed_list.row(entry))
            self._speed_edit_loading = True
            self.speed_list.setCurrentRow(-1)
            self.speed_label_input.setText("")
            self.speed_emoji_input.setText("")
            self.speed_args_input.setText("")
            self._speed_edit_loading = False
            del list_blocker, label_blocker, emoji_blocker, args_blocker
            self._update_speed_popup_preview()
            self._speed_dirty = True

        def _move_available_selected(self) -> None:
            entry = self.available_list.currentItem()
            if entry:
                self._move_available_to_speed(entry)

        def _move_speed_selected(self) -> None:
            entry = self.speed_list.currentItem()
            if entry:
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
                self._speed_item_cache[item.layout] = SpeedMenuItem(item.label, item.emoji, item.layout, list(item.args))
            self._update_speed_popup_preview()
            self._speed_dirty = True

        def _update_speed_popup_preview(self) -> None:
            items = self._collect_speed_items_from_list()
            self._speed_popup.populate(items)

        def _apply_speed_args_preset(self) -> None:
            if getattr(self, "_speed_edit_loading", False):
                return
            preset = self.speed_args_preset.currentText()
            mapping = {
                "Custom": "",
                "Existing Only": "",
                "Launch Missing": "--launch-missing",
                "Edge Tabs": "--edge-tabs",
                "Edge Tabs (Destructive)": "--edge-tabs --destructive",
            }
            args = mapping.get(preset, "")
            self._speed_edit_loading = True
            self.speed_args_input.setText(args)
            self._speed_edit_loading = False
            self._apply_speed_item_edits()

        def _sync_args_preset(self, args: List[str]) -> None:
            raw = " ".join(args).strip()
            mapping = {
                "": "Existing Only",
                "--launch-missing": "Launch Missing",
                "--edge-tabs": "Edge Tabs",
                "--edge-tabs --destructive": "Edge Tabs (Destructive)",
            }
            label = mapping.get(raw, "Custom")
            idx = self.speed_args_preset.findText(label)
            if idx >= 0:
                self.speed_args_preset.setCurrentIndex(idx)

        def _save_speed_menu(self) -> None:
            data = _load_json(CONFIG_PATH)
            if data is None:
                data = {}
            if not isinstance(data, dict):
                QMessageBox.warning(self, "Speed Menu", "Config JSON not found or invalid.")
                return
            items = self._collect_speed_items_from_list()
            data["speed_menu"] = {
                "buttons": [{"label": i.label, "emoji": i.emoji, "layout": i.layout, "args": i.args} for i in items]
            }
            try:
                with open(CONFIG_PATH, "w", encoding="utf-8") as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
            except Exception as exc:
                QMessageBox.warning(self, "Speed Menu", f"Failed to save: {exc}")
                return
            self._speed_dirty = False
            self._reload_speed_menu()

        def _discard_speed_menu_changes(self) -> None:
            self._speed_dirty = False
            self._reload_speed_menu(force=True)

        # ─────────────────────────────────────────────────────────────────────
        # LAYOUT EDITOR
        # ─────────────────────────────────────────────────────────────────────

        def _sync_layout_editor_choice(self) -> None:
            name = self.layout_select.currentText().strip()
            if not name or self.layout_editor_select.currentText().strip() == name:
                return
            blocker = QSignalBlocker(self.layout_editor_select)
            self.layout_editor_select.setCurrentText(name)
            del blocker
            self._load_layout_for_editing()

        def _sync_layout_settings_choice(self) -> None:
            name = self.layout_editor_select.currentText().strip()
            if not name or self.layout_select.currentText().strip() == name:
                return
            blocker = QSignalBlocker(self.layout_select)
            self.layout_select.setCurrentText(name)
            del blocker
            self._reload_speed_menu()
            self._load_layout_for_editing()

        def _load_layout_for_editing(self, force: bool = False) -> None:
            if not force and self._layout_dirty:
                if not self._confirm_unsaved("Layout Editor", self._save_layout_edit, self._discard_layout_changes):
                    if self._layout_edit_name:
                        blocker = QSignalBlocker(self.layout_editor_select)
                        self.layout_editor_select.setCurrentText(self._layout_edit_name)
                        del blocker
                        blocker = QSignalBlocker(self.layout_select)
                        self.layout_select.setCurrentText(self._layout_edit_name)
                        del blocker
                    return
            name = self.layout_editor_select.currentText().strip() or self.layout_select.currentText().strip()
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

        def _reload_layout_windows_list(self) -> None:
            self.layout_windows_list.clear()
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            for idx, window in enumerate(windows):
                title = str(window.get("title") or "(untitled)")
                proc = str(window.get("process_name") or "")
                item = QListWidgetItem(f"{title} [{proc}]")
                item.setData(Qt.UserRole, idx)
                self.layout_windows_list.addItem(item)
            if self.layout_windows_list.count() > 0:
                self.layout_windows_list.setCurrentRow(0)

        def _set_layout_view(self) -> None:
            view = self.layout_view_select.currentText().strip().lower()
            simple = view == "simple"
            for widget in getattr(self, "_advanced_widgets", []):
                widget.setVisible(not simple)

        def _layout_window_selected(self) -> None:
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            item = self.layout_windows_list.currentItem()
            if item is None:
                return
            idx = item.data(Qt.UserRole)
            if not isinstance(idx, int) or idx < 0 or idx >= len(windows):
                return
            self._layout_selected_index = idx
            self._load_window_fields(windows[idx])

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
                self.rect_left.setValue(int(rect[0])); self.rect_top.setValue(int(rect[1]))
                self.rect_right.setValue(int(rect[2])); self.rect_bottom.setValue(int(rect[3]))
            if len(nrect) == 4:
                self.nrect_left.setValue(int(nrect[0])); self.nrect_top.setValue(int(nrect[1]))
                self.nrect_right.setValue(int(nrect[2])); self.nrect_bottom.setValue(int(nrect[3]))
            self.spin_show_cmd.setValue(int(window.get("show_cmd") or 0))
            self.chk_visible.setChecked(bool(window.get("is_visible", True)))
            self.chk_minimized.setChecked(bool(window.get("is_minimized", False)))
            self.chk_maximized.setChecked(bool(window.get("is_maximized", False)))
            self.chk_destructive.setChecked(bool(window.get("destructive", False)))
            launch = window.get("launch") or {}
            if isinstance(launch, dict):
                self.le_launch_exe.setText(str(launch.get("exe") or ""))
                args = launch.get("args") or []
                self.le_launch_args.setText(" ".join(str(a) for a in args) if isinstance(args, list) else str(args or ""))
                self.le_launch_cwd.setText(str(launch.get("cwd") or ""))
            else:
                self.le_launch_exe.setText(""); self.le_launch_args.setText(""); self.le_launch_cwd.setText("")
            edge = window.get("edge") or {}
            if isinstance(edge, dict):
                try:
                    self.spin_edge_port.setValue(int(edge.get("session_port") or 0))
                except (TypeError, ValueError):
                    self.spin_edge_port.setValue(0)
            else:
                self.spin_edge_port.setValue(0)
            self.edge_tabs_list.clear()
            for tab in (window.get("edge_tabs") or []):
                url = str(tab.get("url") or "")
                title = str(tab.get("title") or "")
                lbl = f"{title} → {url}" if title else url
                item = QListWidgetItem(lbl)
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
            window["rect"] = [self.rect_left.value(), self.rect_top.value(), self.rect_right.value(), self.rect_bottom.value()]
            window["normal_rect"] = [self.nrect_left.value(), self.nrect_top.value(), self.nrect_right.value(), self.nrect_bottom.value()]
            window["show_cmd"] = int(self.spin_show_cmd.value())
            window["is_visible"] = bool(self.chk_visible.isChecked())
            window["is_minimized"] = bool(self.chk_minimized.isChecked())
            window["is_maximized"] = bool(self.chk_maximized.isChecked())
            if self.chk_destructive.isChecked():
                window["destructive"] = True
            else:
                window.pop("destructive", None)
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

        def _mark_layout_dirty(self, *_args) -> None:
            if getattr(self, "_layout_edit_loading", False):
                return
            self._layout_dirty = True

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

        def _discard_layout_changes(self) -> None:
            self._layout_dirty = False
            self._load_layout_for_editing(force=True)

        def _remove_selected_window(self) -> None:
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            idx = getattr(self, "_layout_selected_index", None)
            if idx is None or idx < 0 or idx >= len(windows):
                return
            self._layout_removed_cache = (windows.pop(idx), idx)
            self._reload_layout_windows_list()
            self._layout_dirty = True

        def _restore_removed_window(self) -> None:
            if not self._layout_removed_cache:
                return
            data = self._layout_edit_data or {}
            windows = data.get("windows") or []
            removed, idx = self._layout_removed_cache
            idx = max(0, min(idx, len(windows)))
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

        # ─────────────────────────────────────────────────────────────────────
        # ACTIONS / RUN
        # ─────────────────────────────────────────────────────────────────────

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
            if action == "edge_debug":
                edge_settings = self._prompt_edge_settings(edge_port, edge_profile_dir)
                if not edge_settings:
                    return
                edge_port, edge_profile_dir = edge_settings
            elif action == "edge_capture":
                edge_port = self._prompt_edge_port(edge_port)
                if edge_port is None:
                    return
            cmd = build_cli_command(action, layout_path, edge_port=edge_port, edge_profile_dir=edge_profile_dir).args
            self.status.setText(f"Running: {action}")
            self.log.appendPlainText(f"\n$ {format_command_for_log(cmd)}")
            self._proc.start(cmd[0], cmd[1:])

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
            self._speed_popup.set_status(f"Running: {title}")
            self.log.appendPlainText(f"\n$ {format_command_for_log(cmd)}")
            self._proc.start(cmd[0], cmd[1:])

        def _prompt_edge_settings(self, port: int, profile_dir: str) -> Optional[tuple[int, str]]:
            port_value, ok = QInputDialog.getInt(self, "Edge Debug Port", "Remote debugging port:", port, 1, 65535, 1)
            if not ok:
                return None
            profile_value, ok = QInputDialog.getText(self, "Edge Profile Dir", "Profile directory (optional):", text=profile_dir)
            if not ok:
                return None
            try:
                _save_edge_defaults(port_value, profile_value.strip())
            except Exception:
                pass
            return port_value, profile_value.strip()

        def _prompt_edge_port(self, port: int) -> Optional[int]:
            port_value, ok = QInputDialog.getInt(self, "Edge Debug Port", "Remote debugging port:", port, 1, 65535, 1)
            if not ok:
                return None
            try:
                _save_edge_defaults(port_value, _get_edge_defaults()[1])
            except Exception:
                pass
            return port_value

        def _append_stdout(self) -> None:
            data = bytes(self._proc.readAllStandardOutput()).decode("utf-8", errors="replace")
            if data:
                self.log.appendPlainText(data.rstrip("\n"))

        def _append_stderr(self) -> None:
            data = bytes(self._proc.readAllStandardError()).decode("utf-8", errors="replace")
            if data:
                self.log.appendPlainText(data.rstrip("\n"))

        def _on_finished(self, code: int, _status) -> None:
            msg = "Completed" if code == 0 else f"Failed (exit={code})"
            self.status.setText(msg)
            self._speed_popup.set_status(msg)

        # ─────────────────────────────────────────────────────────────────────
        # UNSAVED CHANGES / CLOSE
        # ─────────────────────────────────────────────────────────────────────

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
                if not self._confirm_unsaved("Speed Menu Editor", self._save_speed_menu, self._discard_speed_menu_changes):
                    event.ignore()
                    return
            if self._layout_dirty:
                if not self._confirm_unsaved("Layout Editor", self._save_layout_edit, self._discard_layout_changes):
                    event.ignore()
                    return
            if self._tray_enabled and self._tray_icon is not None:
                event.ignore()
                self._hide_to_tray()
                return
            self._stop_hotkeys()
            self._speed_popup.close()
            event.accept()

    app = QApplication(sys.argv)
    _apply_dark_style(app)
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
