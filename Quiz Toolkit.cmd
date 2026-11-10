@echo off
setlocal
title Quiz Toolkit
cd /d "%~dp0"

set "TOOLKIT=%~dp0tools\QuizToolkit.ps1"

if not exist "%TOOLKIT%" (
    echo.
    echo   ERROR: could not find the toolkit script at:
    echo     %TOOLKIT%
    echo.
    echo   Keep this file inside the DataComs_Reviewer folder, in the
    echo   same place as the "tools" folder.
    echo.
    pause
    exit /b 1
)

where powershell.exe >nul 2>nul
if errorlevel 1 (
    echo.
    echo   ERROR: Windows PowerShell was not found on this PC.
    echo.
    pause
    exit /b 1
)

powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -File "%TOOLKIT%" %*
set "RC=%ERRORLEVEL%"

rem Keep the window open after a normal double-click, and after any error.
if "%~1"=="" (
    echo.
    pause
) else (
    if not "%RC%"=="0" (
        echo.
        pause
    )
)

exit /b %RC%
