@echo off
setlocal
title Quizizz Converter

set "SCRIPT=%~dp0tools\Convert-Quizzes.ps1"

if not exist "%SCRIPT%" (
    echo.
    echo   ERROR: could not find the converter script at:
    echo     %SCRIPT%
    echo.
    echo   Keep this file inside the DataComs_Reviewer folder, in the
    echo   same place as the "tools" folder.
    echo.
    pause
    exit /b 1
)

set "TARGET="
set "RECURSE="

rem --- a folder was dragged onto this file: just use it ---
if not "%~1"=="" (
    if exist "%~1\*" (
        set "TARGET=%~1"
        goto run
    )
)

:menu
cls
echo ============================================================
echo    QUIZIZZ CONVERTER
echo ============================================================
echo.
echo    Turns your raw .txt quiz files into Excel files that
echo    Quizizz can import. One Excel file per .txt file.
echo.
echo    Which folder holds your raw .txt files?
echo.
echo      [1]  DataCom quizzes
echo           %~dp0Quizzes\DataComs
echo.
echo      [2]  A different folder
echo.
echo      [0]  Quit
echo.
echo    Tip: you can also drag a folder onto this file.
echo.
set "CHOICE="
set /p "CHOICE=Type 1, 2 or 0 and press Enter: "

if "%CHOICE%"=="1" (
    set "TARGET=%~dp0Quizzes\DataComs"
    goto run
)
if "%CHOICE%"=="2" goto ask
if "%CHOICE%"=="0" exit /b 0
goto menu

:ask
echo.
set "TARGET="
set /p "TARGET=Paste the folder path, then press Enter: "
set "TARGET=%TARGET:"=%"
if not exist "%TARGET%\" (
    echo.
    echo    That folder was not found.
    echo      %TARGET%
    echo.
    pause
    goto menu
)
echo.
set "SUB="
set /p "SUB=Also include sub-folders inside it? Type Y or N then Enter: "
if /i "%SUB%"=="Y" set "RECURSE=-Recurse"
goto run

:run
cls
echo ============================================================
echo    QUIZIZZ CONVERTER
echo ============================================================
echo.
echo    Reading   : %TARGET%
echo    Please wait...
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" -InputFolder "%TARGET%" %RECURSE%

echo.
echo ============================================================
echo    Finished. Your Excel files are in the folder called
echo    _QuizizzOutput (opening it now).
echo.
echo    Press any key to close this window.
echo ============================================================
if exist "%TARGET%\_QuizizzOutput\" start "" "%TARGET%\_QuizizzOutput"
pause >nul
exit /b 0
