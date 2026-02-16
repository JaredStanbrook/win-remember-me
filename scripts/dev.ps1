param(
    [Parameter(Mandatory = $true)]
    [ValidateSet(
        "save",
        "restore",
        "restore-missing",
        "edge-debug",
        "edge-save",
        "edge-restore",
        "wizard",
        "help",
        "build-wheels",
        "download-wheels"
    )]
    [string]$Task,

    [string]$Layout = "layout.json"
)

switch ($Task) {
    "save" { python window_layout.py save $Layout }
    "restore" { python window_layout.py restore $Layout }
    "restore-missing" { python window_layout.py restore $Layout --launch-missing }
    "edge-debug" { python window_layout.py edge-debug }
    "edge-save" { python window_layout.py save $Layout --edge-tabs }
    "edge-restore" { python window_layout.py restore $Layout --restore-edge-tabs }
    "wizard" { python window_layout.py wizard }
    "help" { python window_layout.py help }
    "download-wheels" { python -m pip download -r requirements.txt -d wheels }
    "build-wheels" { python -m pip wheel . -w dist --no-deps }
}
