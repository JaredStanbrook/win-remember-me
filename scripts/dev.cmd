@echo off
setlocal

if "%~1"=="" (
  echo Usage: scripts\dev.cmd ^<task^> [layout.json]
  exit /b 1
)

set TASK=%~1
set LAYOUT=%~2
if "%LAYOUT%"=="" set LAYOUT=layout.json

if /i "%TASK%"=="save" python window_layout.py save "%LAYOUT%" & goto :done
if /i "%TASK%"=="restore" python window_layout.py restore "%LAYOUT%" & goto :done
if /i "%TASK%"=="restore-missing" python window_layout.py restore "%LAYOUT%" --launch-missing & goto :done
if /i "%TASK%"=="edge-debug" python window_layout.py edge-debug & goto :done
if /i "%TASK%"=="edge-save" python window_layout.py save "%LAYOUT%" --edge-tabs & goto :done
if /i "%TASK%"=="edge-restore" python window_layout.py restore "%LAYOUT%" --restore-edge-tabs & goto :done
if /i "%TASK%"=="wizard" python window_layout.py wizard & goto :done
if /i "%TASK%"=="help" python window_layout.py help & goto :done
if /i "%TASK%"=="download-wheels" python -m pip download -r requirements.txt -d wheels & goto :done
if /i "%TASK%"=="build-wheels" python -m pip wheel . -w dist --no-deps & goto :done

echo Unknown task: %TASK%
echo Usage: scripts\dev.cmd ^<task^> [layout.json]
exit /b 1

:done

endlocal
