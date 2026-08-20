@echo off
rem The only file a user is expected to double-click. It clears three things the console
rem puts in front of menu.ps1 -- the cp949 code page, the execution policy, and the lack
rem of an elevated token -- and then gets out of the way. Every Korean message below must
rem stay after the chcp call, or it renders as mojibake.
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

rem fltmc needs an elevated token and fails cleanly without one. `net session` would do
rem the same job but reports failure when the Server service is stopped, which would send
rem an already-elevated run back through UAC forever.
fltmc >nul 2>&1
if not errorlevel 1 goto :elevated

echo.
echo   관리자 권한을 요청합니다. UAC 창에서 [예]를 선택하세요.
echo.
rem The elevated process starts in system32 and inherits no environment, so the script
rem path is passed absolute and the invoking account is passed explicitly: the installer
rem writes to HKCU and %%LOCALAPPDATA%%, and elevating into a *different* admin account
rem would install into that account's profile where Excel never looks.
powershell.exe -NoProfile -ExecutionPolicy Bypass -Command ^
 "Start-Process powershell.exe -Verb RunAs -ArgumentList @('-NoProfile','-ExecutionPolicy','Bypass','-File',[char]34+'%MENU%'+[char]34,'-InvokedBy',[char]34+'%USERDOMAIN%\%USERNAME%'+[char]34)"
if errorlevel 1 (
    echo.
    echo   관리자 권한 상승이 취소되었거나 차단되었습니다.
    echo   회사 보안 정책을 확인하세요.
    echo.
    pause
)
chcp %ORIGINAL_CP% >nul
endlocal
exit /b 0

:elevated
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
