@echo off
:: Drag and drop a folder onto this batch file
if "%~1"=="" (
    echo Please drag and drop a folder onto this script.
    pause
    exit /b
)

:: Get full folder path
set "FOLDER=%~1"

:: Run the PowerShell script with the folder path
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0convert-ppt-to-pdf.ps1" -FolderPath "%FOLDER%"

pause