@echo off
rem The only file a user is expected to double-click. Installation stays in the invoking
rem user's HKCU and LOCALAPPDATA. Never elevate this launcher into a different administrator
rem profile: that account's Excel cannot see the invoking user's registration or files.
setlocal
for /f "tokens=2 delims=:." %%a in ('chcp') do set "ORIGINAL_CP=%%a"
chcp 65001 >nul

set "MENU=%~dp0scripts\menu.ps1"
if not exist "%MENU%" (
    echo.
    echo   ZIP 파일 안에서 바로 실행할 수 없습니다.
    echo   ZIP을 폴더에 전체 압축 해제한 뒤 다시 실행하세요.
    echo.
    pause
    chcp %ORIGINAL_CP% >nul
    exit /b 1
)

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%MENU%" -InvokedBy "%USERDOMAIN%\%USERNAME%"
if errorlevel 1 (
    echo.
    echo   메뉴가 정상적으로 끝나지 않았습니다.
    echo   PowerShell 실행이 회사 보안 정책으로 차단되었는지 확인하세요.
    echo.
    pause
)

chcp %ORIGINAL_CP% >nul
endlocal
