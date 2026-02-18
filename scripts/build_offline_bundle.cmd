@echo off
setlocal

set REQUIRE_ALL=
if /I "%1"=="--require-all" set REQUIRE_ALL=--require-all

python scripts\build_offline_bundle.py --python-versions 3.13 3.12 3.11 %REQUIRE_ALL%

endlocal
