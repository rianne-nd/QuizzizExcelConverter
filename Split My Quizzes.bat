@echo off
rem Double-click this to split a quiz .txt into parts.
rem The actual script lives in tools\.
setlocal
cd /d "%~dp0"
python "%~dp0tools\Splitter.py"
if errorlevel 1 (
    echo.
    echo If Python is not installed, get it from https://www.python.org/downloads/
    echo During setup, tick "Add python.exe to PATH".
)
pause
