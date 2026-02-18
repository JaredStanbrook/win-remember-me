param(
    [string]$Package = "window-layout-cli",
    [string]$DistDir = "dist",
    [string]$WheelsDir = "wheels"
)

python -m pip install --no-index --find-links $WheelsDir --find-links $DistDir $Package
