@echo off
setlocal

REM Pass through all args to the Python build script.
python scripts\build_offline_bundle.py %*

endlocal
