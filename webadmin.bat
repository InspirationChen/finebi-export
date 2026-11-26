@echo off
rem ============================================================
rem  Start the FineBI web admin console (webadmin.py).
rem  Browser opens http://127.0.0.1:8787 automatically.
rem
rem  Usage:
rem    webadmin.bat                     default 127.0.0.1:8787
rem    webadmin.bat --port 9000         use another port
rem    webadmin.bat --host 0.0.0.0      allow LAN access
rem
rem  Stop: close this window, or press Ctrl+C.
rem  NOTE: this file is pure ASCII on purpose -- cmd on Chinese
rem  Windows parses .bat files as GBK, so keep it that way.
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
    pause
    exit /b 2
)

title FineBI Admin Console
echo Starting admin console... browser will open http://127.0.0.1:8787
echo If it does not open, visit that address manually.
echo Close this window (or Ctrl+C) to stop.
echo.
"%PYTHON%" "%WORKDIR%webadmin.py" %*
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo [ERROR] Admin console exited with code %RC%.
    echo         Common cause: port 8787 already in use --
    echo         try: webadmin.bat --port 9000
)

pause
exit /b %RC%
