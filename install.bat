@echo off
setlocal EnableExtensions
pushd "%~dp0"

echo === Office COM Bridge installer ===
echo.

rem Find a supported Python interpreter: 3.11 or 3.12.
set "PYTHON_CMD="
py -3.12 -c "import sys" >nul 2>nul
if not errorlevel 1 set "PYTHON_CMD=py -3.12"

if not defined PYTHON_CMD (
    py -3.11 -c "import sys" >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=py -3.11"
)

if not defined PYTHON_CMD (
    python -c "import sys; raise SystemExit(0 if (3, 11) <= sys.version_info[:2] <= (3, 12) else 1)" >nul 2>nul
    if not errorlevel 1 set "PYTHON_CMD=python"
)

if not defined PYTHON_CMD (
    echo [ERROR] Python 3.11 or 3.12 was not found.
    echo         Install Python 3.11 or 3.12, then run this file again.
    echo         Make sure "py" or "python" is available on PATH.
    goto :fail
)

echo Using Python: %PYTHON_CMD%
echo.

if exist ".venv\Scripts\python.exe" (
    echo [1/3] Reusing existing .venv
) else (
    echo [1/3] Creating .venv ...
    %PYTHON_CMD% -m venv --system-site-packages .venv
    if errorlevel 1 (
        echo [ERROR] Could not create the virtual environment.
        goto :fail
    )
)

".venv\Scripts\python.exe" -c "import sys; raise SystemExit(0 if (3, 11) <= sys.version_info[:2] <= (3, 12) else 1)" >nul 2>nul
if errorlevel 1 (
    echo [ERROR] The existing .venv uses an unsupported Python version.
    echo         Delete the .venv directory and run this installer again.
    goto :fail
)

echo [2/3] Installing dependencies ...
".venv\Scripts\python.exe" -m pip install -e ".[dev]"
if errorlevel 1 (
    echo [ERROR] pip install failed.
    goto :fail
)

echo [3/3] Installing the Office skill ...
powershell -NoProfile -ExecutionPolicy Bypass -File "scripts\install-skill.ps1"
if errorlevel 1 (
    echo [WARN] Skill installation failed. The bridge still works, but the
    echo        configured agent will not see the Office skill.
)

echo.
echo Verifying the installation ...
".venv\Scripts\python.exe" -c "import office_bridge, office_cli; print('office-com-bridge', office_bridge.__version__)"
if errorlevel 1 (
    echo [ERROR] The installed packages could not be imported.
    goto :fail
)

echo.
echo Installation complete.
echo.
echo   1. Start the bridge with start.bat
echo   2. Start the configured agent with start-agent.bat
echo   3. Focus a Word, Excel, or PowerPoint document and press Ctrl+Alt+O
echo.
echo To use Hermes, Claude Code, Codex, or a custom command:
echo   configure.bat
echo.
popd
endlocal
exit /b 0

:fail
echo.
echo Installation failed.
popd
endlocal
pause
exit /b 1
