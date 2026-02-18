param(
    [string[]]$PythonVersions = @("3.13", "3.12", "3.11"),
    [switch]$RequireAll,
    [string[]]$Extras = @(),
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ExtraArgs = @()
)

$cmd = @("scripts/build_offline_bundle.py", "--python-versions") + $PythonVersions
if ($RequireAll) {
    $cmd += "--require-all"
}

if ($Extras.Count -gt 0) {
    $cmd += @("--extras") + $Extras
}

if ($ExtraArgs.Count -gt 0) {
    $cmd += $ExtraArgs
}

python @cmd
