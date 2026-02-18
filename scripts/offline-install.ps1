param(
    [string]$Package = "window-layout-cli",
    [string]$DistDir = "dist",
    [string]$WheelsDir = "wheels"
)

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
& "$ScriptDir\install_offline.ps1" -Package $Package -DistDir $DistDir -WheelsDir $WheelsDir
