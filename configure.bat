@echo off
setlocal EnableExtensions
pushd "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] .venv was not found.
    echo         Run install.bat first.
    goto :fail
)

if "%~1"=="" (
    ".venv\Scripts\python.exe" -m office_bridge.agent_config setup
) else (
    ".venv\Scripts\python.exe" -m office_bridge.agent_config %*
)
set "RC=%ERRORLEVEL%"

if not "%RC%"=="0" (
    echo.
    echo configure.bat exited with code %RC%.
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
