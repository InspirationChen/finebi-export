@echo off
setlocal
rem ============================================================
rem  FineBI preflight check (no data export, finishes fast).
rem
rem    preflight.bat                      check first enabled job
rem    preflight.bat --save before.json   save baseline before upgrade
rem    preflight.bat --compare before.json  diff after upgrade
rem    preflight.bat --all-jobs           check all enabled jobs
rem
rem  Exit code: 0 = compatible (warnings ok), 1 = incompatible.
rem  NOTE: pure ASCII on purpose -- cmd on Chinese Windows
rem  parses .bat files as GBK, so keep it that way.
rem ============================================================

set "WORKDIR=%~dp0"

if not defined PYTHON set "PYTHON=python"

"%PYTHON%" -V >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found: %PYTHON%
    echo         Install Python 3.8/3.11, or edit this file and
    echo         set PYTHON to the full path of python.exe.
    exit /b 2
)

"%PYTHON%" "%WORKDIR%preflight.py" --config "%WORKDIR%config.xlsx" %*

if errorlevel 1 (
    echo.
    echo [ERROR] Blocking problems found. See [FAIL] lines above.
)

echo.
if "%~1"=="" pause
