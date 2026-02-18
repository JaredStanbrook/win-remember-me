@echo off
setlocal

set PACKAGE=window-layout-cli
set DISTDIR=dist
set WHEELSDIR=wheels

python -m pip install --no-index --find-links %WHEELSDIR% --find-links %DISTDIR% %PACKAGE%

endlocal
