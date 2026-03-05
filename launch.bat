@echo off
title QI Screening Email Pipeline

:: Change to the folder where this script lives
cd /d "%~dp0"

:: ── Locate Python ────────────────────────────────────────────────────────────
:: Searches (in order):
::   1. winpython\python.exe                          (flat)
::   2. winpython\python-*\python.exe                 (versioned flat)
::   3. winpython\WPy64-*\python-*\python.exe         (standard WinPython layout)
::   4. System Python on PATH
set PYTHON=

if exist "winpython\python.exe" (
    set PYTHON=winpython\python.exe
    goto :found
)

for /d %%d in ("winpython\python-*") do (
    if exist "%%d\python.exe" (
        set PYTHON=%%d\python.exe
        goto :found
    )
)

for /d %%w in ("winpython\WPy64-*") do (
    for /d %%d in ("%%w\python-*") do (
        if exist "%%d\python.exe" (
            set PYTHON=%%d\python.exe
            goto :found
        )
    )
)

where python >nul 2>&1
if %errorlevel%==0 (
    set PYTHON=python
    goto :found
)

echo ERROR: Python not found.
echo   Run INSTALL.bat first, or place WinPython in the 'winpython' sub-folder.
pause
exit /b 1

:found
echo Using Python: %PYTHON%

:: ── Install / update dependencies ────────────────────────────────────────────
"%PYTHON%" -m pip install -r requirements.txt --quiet

:: ── Launch app ───────────────────────────────────────────────────────────────
echo.
echo Starting app — your browser will open automatically.
echo Close this window to stop the app.
echo.
"%PYTHON%" -m streamlit run app.py --server.headless false

pause
