#!/usr/bin/env bash
set -euo pipefail

task="${1:-}"
layout="${2:-layout.json}"

if [[ -z "$task" ]]; then
  echo "Usage: ./scripts/dev.sh <task> [layout.json]"
  exit 1
fi

case "$task" in
  save) python window_layout.py save "$layout" ;;
  restore) python window_layout.py restore "$layout" ;;
  restore-missing) python window_layout.py restore "$layout" --launch-missing ;;
  edge-debug) python window_layout.py edge-debug ;;
  edge-save) python window_layout.py save "$layout" --edge-tabs ;;
  edge-restore) python window_layout.py restore "$layout" --restore-edge-tabs ;;
  wizard) python window_layout.py wizard ;;
  help) python window_layout.py help ;;
  download-wheels) python -m pip download -r requirements.txt -d wheels ;;
  build-wheels) python -m pip wheel . -w dist --no-deps ;;
  *)
    echo "Unknown task: $task"
    exit 1
    ;;
esac
