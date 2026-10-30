@echo off
setlocal EnableExtensions
pushd "%~dp0"
title Office COM Bridge

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] .venv was not found.
    echo         Run install.bat first.
    goto :fail
)

".venv\Scripts\python.exe" -c "import win32com.client" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] pywin32 is not available in .venv.
    echo         Run install.bat again.
    goto :fail
)

echo Starting Office COM Bridge ...
echo   Default URL:    http://127.0.0.1:8733
echo   Default hotkey: Ctrl+Alt+O
echo   Stop:           Ctrl+C
echo.
echo Logs: %%LOCALAPPDATA%%\office-bridge\logs
echo.

".venv\Scripts\python.exe" -m office_bridge.server %*
set "RC=%ERRORLEVEL%"

echo.
if "%RC%"=="0" (
    echo Bridge stopped.
) else (
    echo Bridge stopped with exit code %RC%.
    echo Check %%LOCALAPPDATA%%\office-bridge\logs\bridge.log for details.
    pause
)

popd
endlocal
exit /b %RC%

:fail
popd
endlocal
pause
exit /b 1
