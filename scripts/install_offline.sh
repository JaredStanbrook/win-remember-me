#!/usr/bin/env bash
set -euo pipefail

package="${1:-window-layout}"
dist="${2:-dist}"
wheels="${3:-wheels}"

python -m pip install --no-index --find-links "$wheels" --find-links "$dist" "$package"
