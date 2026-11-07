@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
echo.
echo ===============================================
echo    Marble Delivery Pro 10.0.5 Setup Builder
echo ===============================================
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install_Build_Tools.ps1"
if errorlevel 1 goto :failed
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Build_All.ps1"
if errorlevel 1 goto :failed
echo.
echo Setup created in the dist folder.
start "" "%~dp0dist"
pause
exit /b 0
:failed
echo.
echo Build failed. See output and fix errors before retrying.
pause
exit /b 1
