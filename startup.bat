
@echo off
setlocal EnableDelayedExpansion
set My_PATH=%cd%
set PATH=%PATH%;%My_PATH%
python %cd%\uiall.cpython-36.pyc
echo %cd%
pause
@echo off
