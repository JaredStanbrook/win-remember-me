@echo off
setlocal

set PACKAGE=%~1
if "%PACKAGE%"=="" set PACKAGE=window-layout

set DISTDIR=%~2
if "%DISTDIR%"=="" set DISTDIR=dist

set WHEELSDIR=%~3
if "%WHEELSDIR%"=="" set WHEELSDIR=wheels

python -m pip install --no-index --find-links %WHEELSDIR% --find-links %DISTDIR% %PACKAGE%

endlocal
