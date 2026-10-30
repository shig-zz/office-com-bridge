@echo off
setlocal EnableExtensions
pushd "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] .venv was not found.
    echo         Run install.bat first.
    goto :fail
)

echo Starting the configured agent with the Office skill ...
echo The launcher will start the bridge if it is not already running.
echo Use configure.bat to switch agents, or pass -Agent hermes.
echo.

powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\start-agent.ps1" %*
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo start-agent.ps1 exited with code %RC%.
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
