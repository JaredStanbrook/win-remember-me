param(
    [string[]]$PythonVersions = @("3.13", "3.12", "3.11"),
    [switch]$RequireAll
)

$cmd = @("scripts/build_offline_bundle.py", "--python-versions") + $PythonVersions
if ($RequireAll) {
    $cmd += "--require-all"
}

python @cmd
