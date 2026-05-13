@echo off
REM Double-click this file to apply your edits from the updates\ folder.
REM This is the Windows equivalent of update-site.command (which is for Mac).

REM Always operate from the project root, no matter where this file lives
cd /d "%~dp0\.."

echo ================================================
echo   MUSE Lab: Update Site
echo ================================================
echo.

REM First-run setup: create venv and install dependencies if missing.
if not exist ".venv" (
    echo First-time setup: creating a Python virtual environment...
    echo (this only happens once and takes about 30 seconds)
    echo.

    where python >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Python is not installed.
        echo Install Python 3 from https://www.python.org/downloads/
        echo Make sure to check "Add Python to PATH" during installation.
        echo.
        pause
        exit /b 1
    )

    python -m venv .venv
    call .venv\Scripts\python.exe -m pip install --quiet --upgrade pip
    call .venv\Scripts\pip.exe install --quiet -r scripts\requirements.txt
    echo Setup complete.
    echo.
)

echo Reading your Excel sheets and applying changes...
echo.
call .venv\Scripts\python.exe scripts\update_site.py

echo.
echo ================================================
echo   Done.
echo.
echo   Next steps:
echo     1. Open this folder in GitHub Desktop (or your
echo        favorite git tool) to see what changed.
echo     2. Commit and push to publish your changes.
echo ================================================
echo.
pause
