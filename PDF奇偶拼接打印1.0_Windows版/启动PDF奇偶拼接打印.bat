@echo off
rem ============================================================
rem  PDF odd-even merge print tool v1.0
rem
rem  This file is deliberately pure ASCII.  Batch files are decoded
rem  by cmd.exe with the console codepage, which differs between
rem  machines; all Chinese text is printed by setup_env.py, where
rem  the encoding is guaranteed.  So: never add Chinese here.
rem ============================================================
setlocal
pushd "%~dp0"

set "PYEXE="

rem ---- 1) prefer the official "py" launcher from python.org ----
py -3 -c "import sys" >nul 2>&1
if not errorlevel 1 set "PYEXE=py -3"
if defined PYEXE goto RUN

rem ---- 2) fall back to "python" on PATH ----
python -c "import sys" >nul 2>&1
if not errorlevel 1 set "PYEXE=python"
if defined PYEXE goto RUN

rem ---- 3) no Python at all ----
echo.
echo   [!] Python 3 was not found on this computer.
echo.
echo       Install it from:
echo         https://www.python.org/downloads/windows/
echo.
echo       During setup, TICK this box:
echo         [x] Add python.exe to PATH
echo.
echo       Then double-click this file again.
echo.
pause
popd
exit /b 1

:RUN
set "PYTHONIOENCODING=utf-8"
set "PYTHONUTF8=1"
%PYEXE% "%~dp0setup_env.py" %*
set "RC=%ERRORLEVEL%"

if "%RC%"=="0" goto DONE
echo.
echo   [!] Stopped with exit code %RC%. See the messages above.
echo.
pause

:DONE
popd
exit /b %RC%
