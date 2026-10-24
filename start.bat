@echo off
setlocal
cd /d "%~dp0"

powershell.exe -NoLogo -NoProfile -ExecutionPolicy Bypass -STA -File "%~dp0SymlinkManager.ps1"

if errorlevel 1 (
    echo.
    echo The program exited with an error.
    echo See symlink-manager.log for details.
    pause
)
