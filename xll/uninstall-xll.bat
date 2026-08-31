@echo off
setlocal

set "UNINSTALLER=%~dp0scripts\uninstall.ps1"
if not exist "%UNINSTALLER%" (
    echo.
    echo The uninstaller is missing.
    echo.
    pause
    exit /b 1
)

cd /d "%TEMP%"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%UNINSTALLER%" -PromptForHiddenExcel %*
if errorlevel 1 (
    echo.
    echo Uninstall failed. Review the error above.
    echo.
    pause
    exit /b 1
)

echo.
pause
endlocal
