@echo off
setlocal
set SCRIPT_DIR=%~dp0
call "%SCRIPT_DIR%install_offline.cmd" %*
endlocal
