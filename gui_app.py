"""Lightweight GUI for Window Layout CLI (PySide6)."""

from __future__ import annotations

import shlex
import subprocess
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import List


DEFAULT_LAYOUT_PATH = "layout.json"


@dataclass
class GuiCommand:
    label: str
    args: List[str]


def build_cli_command(action: str, layout_path: str) -> GuiCommand:
    base = [sys.executable, "window_layout.py"]
    if action == "save":
        return GuiCommand("Save Layout", base + ["save", layout_path])
    if action == "save_edge":
        return GuiCommand("Save Layout + Edge Tabs", base + ["save", layout_path, "--edge-tabs"])
    if action == "restore":
        return GuiCommand("Restore Layout", base + ["restore", layout_path])
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


def main() -> int:
    try:
        from PySide6.QtCore import QProcess
        from PySide6.QtWidgets import (
            QApplication,
            QFileDialog,
            QGridLayout,
            QHBoxLayout,
            QLabel,
            QLineEdit,
            QMainWindow,
            QMessageBox,
            QPushButton,
            QPlainTextEdit,
            QWidget,
        )
    except ImportError:
        print("PySide6 is required for GUI mode. Install with: pip install PySide6")
        return 1

    class MainWindow(QMainWindow):
        def __init__(self) -> None:
            super().__init__()
            self.setWindowTitle("Window Layout Manager")
            self.resize(860, 560)

            self._proc = QProcess(self)
            self._proc.readyReadStandardOutput.connect(self._append_stdout)
            self._proc.readyReadStandardError.connect(self._append_stderr)
            self._proc.finished.connect(self._on_finished)

            root = QWidget(self)
            self.setCentralWidget(root)

            main_layout = QGridLayout(root)
            main_layout.setHorizontalSpacing(10)
            main_layout.setVerticalSpacing(8)

            path_row = QHBoxLayout()
            self.path_input = QLineEdit(DEFAULT_LAYOUT_PATH)
            browse_btn = QPushButton("Browse…")
            browse_btn.clicked.connect(self._browse)
            path_row.addWidget(QLabel("Layout JSON:"))
            path_row.addWidget(self.path_input, 1)
            path_row.addWidget(browse_btn)
            main_layout.addLayout(path_row, 0, 0, 1, 2)

            actions = [
                ("Save", "save"),
                ("Save + Edge Tabs", "save_edge"),
                ("Restore", "restore"),
                ("Restore Dry Run", "restore_dry"),
                ("Restore + Launch Missing", "restore_missing"),
                ("Restore + Edge Tabs", "restore_edge"),
                ("Edit Edge Mapping", "edit"),
            ]

            for idx, (title, action) in enumerate(actions):
                btn = QPushButton(title)
                btn.clicked.connect(lambda _=False, a=action: self._run(a))
                row = 1 + idx // 2
                col = idx % 2
                main_layout.addWidget(btn, row, col)

            self.log = QPlainTextEdit()
            self.log.setReadOnly(True)
            main_layout.addWidget(self.log, 5, 0, 1, 2)

            self.status = QLabel("Ready")
            main_layout.addWidget(self.status, 6, 0, 1, 2)

        def _browse(self) -> None:
            selected, _ = QFileDialog.getSaveFileName(self, "Select layout JSON", self.path_input.text(), "JSON (*.json)")
            if selected:
                self.path_input.setText(selected)

        def _run(self, action: str) -> None:
            if self._proc.state() != QProcess.NotRunning:
                QMessageBox.information(self, "Busy", "A command is already running.")
                return

            layout_path = self.path_input.text().strip() or DEFAULT_LAYOUT_PATH
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
    win = MainWindow()
    win.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
