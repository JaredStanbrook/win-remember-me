# Repository Guidelines

## Project Structure & Module Organization
This repository contains a purpose-built Windows CLI tool for fast session state management. It captures window positions, PIDs, titles, and related metadata so you can restore your daily workspace as quickly as possible. The structure is flat:
- `window_layout.py` is the main CLI script for capture and restore.
- `layout.json` is an example output file produced by the script.
- `.venv` is a local virtual environment (optional).

If the tool grows, add new modules under `src/` and keep the CLI entry point at the root.

## Build, Test, and Development Commands
This project has no build step. Typical usage in PowerShell:
- Create a virtual environment (optional): `python -m venv .venv`
- Install dependencies: `pip install psutil pywin32`
- Save a layout at end-of-day: `python window_layout.py save layout.json`
- Save layout + Edge tabs (requires Edge debug port): `python window_layout.py save layout.json --edge-tabs`
- Launch Edge debug instance: `python window_layout.py edge-debug`
- Launch Edge debug on a custom port: `python window_layout.py edge-debug --port 9223`
- First-time setup wizard: `python window_layout.py wizard`
- Restore a layout at start-of-day: `python window_layout.py restore layout.json`
- Dry-run restore (safe preview): `python window_layout.py restore layout.json --dry-run`
- Restore and launch missing apps: `python window_layout.py restore layout.json --launch-missing`
- Tune launch wait time: `python window_layout.py restore layout.json --launch-missing --launch-wait 8`
- Restore captured Edge tabs: `python window_layout.py restore layout.json --restore-edge-tabs`

## Coding Style & Naming Conventions
- Indentation: 4 spaces.
- Python: follow PEP 8 for naming (`snake_case` for functions, `CapWords` for classes).
- Keep docstrings brief and focused on behavior and assumptions.
- Prefer small, single-purpose helpers (see `_safe_get_text`, `_window_placement`).

No formatter or linter is configured yet. If you add one, document it here and include the config file in the repo.

## Testing Guidelines
No test framework is currently set up. If you add tests:
- Prefer `pytest`.
- Put tests in `tests/` and name them `test_*.py`.
- Add a quick-start command, e.g., `pytest -q`.

## Commit & Pull Request Guidelines
This repository does not include Git history, so there is no established commit message convention. If you initialize Git, use concise, imperative messages (e.g., `Add restore dry-run output`).

For pull requests:
- Describe behavior changes and edge cases.
- Include reproduction steps for window layout issues.
- Attach screenshots only if the UI behavior is relevant.

## Security & Configuration Tips
The script interacts with Windows window handles and process metadata. Avoid running it with elevated privileges unless necessary, and be cautious when sharing `layout.json` because it includes window titles, PIDs, and executable paths.

## Usage Notes
This tool is tuned for quick, reliable restore of your daily TSD workspace. Matching is heuristic-based (process, class, title, exe), so stability improves if you keep app names and window titles consistent. If restores skip windows, raise the `--min-score` threshold only after you understand which metadata is changing.

When `--launch-missing` is used, the tool attempts to start apps using absolute executable paths saved in `layout.json`. If an app cannot be launched (missing path or unsupported window type), it will be reported as skipped.

Edge tab capture requires running Edge with remote debugging enabled (example: `msedge.exe --remote-debugging-port=9222`). The tool will only save and restore URLs it can read from the debug endpoint, and it will skip internal pages like `edge://settings`.
