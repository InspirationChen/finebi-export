@echo off
setlocal enabledelayedexpansion
rem ============================================================
rem  Run one collection round, then exit.
rem  Good for double-click / scheduled task / scheduler calls.
rem
rem  Usage:
rem    run_once.bat                     run all enabled jobs
rem    run_once.bat --job <name>        run one job by name
rem    run_once.bat --max-pages 2       smoke test (2 pages each)
rem
rem  Exit code: 0 ok, 1 some job failed, 2 config/arg error.
rem  NOTE: pure ASCII on purpose -- cmd on Chinese Windows
rem  parses .bat files as GBK, so keep it that way.
rem ============================================================

set "WORKDIR=%~dp0"

rem If python is not in PATH, edit the line below to a full path:
rem   set "PYTHON=C:\Python38\python.exe"
if not defined PYTHON set "PYTHON=python"

"%PYTHON%" -V >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found: %PYTHON%
    echo         Install Python 3.8/3.11, or edit this file and
    echo         set PYTHON to the full path of python.exe.
    exit /b 2
)

"%PYTHON%" "%WORKDIR%collector.py" --config "%WORKDIR%config.xlsx" --once --json %*
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo [ERROR] Failed with exit code %RC%. Check log above or logs\collector.log
)

exit /b %RC%
