@echo off
setlocal

set "INSTALLER=%~dp0scripts\install.ps1"
if not exist "%INSTALLER%" (
    echo.
    echo The deployment package is incomplete. Extract the entire ZIP before running this file.
    echo.
    pause
    exit /b 1
)

cd /d "%TEMP%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%INSTALLER%" -PromptForHiddenExcel %*
if errorlevel 1 (
    echo.
    echo Installation failed. Review the error above.
    echo.
    pause
    exit /b 1
)

echo.
pause
endlocal
