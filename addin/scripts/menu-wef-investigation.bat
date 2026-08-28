@echo off
setlocal
for /f "tokens=2 delims=:" %%A in ('chcp') do set "DDOT_OLD_CP=%%A"
chcp 65001 >nul
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0run-wef-investigation.ps1" menu
set "DDOT_EXIT=%ERRORLEVEL%"
if defined DDOT_OLD_CP chcp %DDOT_OLD_CP% >nul
endlocal & exit /b %DDOT_EXIT%
