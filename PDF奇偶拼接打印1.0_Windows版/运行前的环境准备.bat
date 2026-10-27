@echo off
rem ============================================================
rem  PDF odd-even merge print tool v1.0 - pre-run check
rem
rem  This file is deliberately pure ASCII.
rem  See comments in the main .bat for the reason.
rem ============================================================
setlocal
pushd "%~dp0"

set "PYEXE="

rem ---- 1) prefer "py" launcher ----
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PYEXE=py -3"
if defined PYEXE goto RUN

rem ---- 2) fall back to "python" ----
python -c "import sys" >nul 2>&1
if not errorlevel 1 set "PYEXE=python"
if defined PYEXE goto RUN

echo.
echo   [!] Python 3 was not found on this computer.
echo       Install from https://www.python.org/downloads/windows/
echo       TICK "Add python.exe to PATH" during setup.
echo.
pause
popd
exit /b 1

:RUN
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
%PYEXE% "%~dp0setup_env.py" --setup
set "RC=%ERRORLEVEL%"
echo.
pause
popd
exit /b %RC%
