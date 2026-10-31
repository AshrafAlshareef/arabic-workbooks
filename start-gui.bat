@echo off
rem ---------------------------------------------------------------------------
rem  start-gui.bat - opens the Arabic workbook generator window.
rem  Double-click this file. No install and no server: it only starts the
rem  PowerShell + WPF interface, which in turn runs the normal Node build.
rem ---------------------------------------------------------------------------
setlocal
cd /d "%~dp0"

where node >nul 2>nul
if errorlevel 1 (
    echo Node.js was not found in PATH.
    echo Install Node.js from https://nodejs.org and run this file again.
    pause
    exit /b 1
)

start "" powershell.exe -NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File "%~dp0ui\workbooks-gui.ps1"
exit /b 0
