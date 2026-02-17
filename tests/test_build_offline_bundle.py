import importlib.util
from pathlib import Path


MODULE_PATH = Path(__file__).resolve().parents[1] / "scripts" / "build_offline_bundle.py"
spec = importlib.util.spec_from_file_location("build_offline_bundle", MODULE_PATH)
mod = importlib.util.module_from_spec(spec)
assert spec and spec.loader
spec.loader.exec_module(mod)


def test_python_tag_sanitizes_command_tokens():
    assert mod._python_tag(["py", "-3.13"]) == "py_3.13"


def test_write_offline_wrapper_scripts(tmp_path):
    mod._write_offline_wrapper_scripts(tmp_path)

    assert (tmp_path / "offline-install.cmd").exists()
    assert (tmp_path / "offline-install.ps1").exists()
    assert (tmp_path / "offline-install.zsh").exists()


def test_zip_dir_creates_zip(tmp_path):
    src = tmp_path / "src"
    src.mkdir()
    (src / "a.txt").write_text("hello", encoding="utf-8")
    out = tmp_path / "out.zip"

    mod._zip_dir(src, out)

    assert out.exists()
