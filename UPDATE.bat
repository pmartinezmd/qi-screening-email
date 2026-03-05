@echo off
title QI Screening Email Pipeline — Updater
cd /d "%~dp0"

echo.
echo ============================================================
echo   QI Screening Email Pipeline — Update
echo ============================================================
echo.
echo This will download the latest version from GitHub and
echo replace the app files. Your data folder will NOT be touched.
echo.
set /p CONFIRM=Continue? (Y/N):
if /i not "%CONFIRM%"=="Y" (
    echo Cancelled — no changes made.
    pause
    exit /b 0
)

:: ── Download latest zip from GitHub ──────────────────────────────────────────
echo.
echo Downloading latest update...

set ZIPURL=https://github.com/pmartinezmd/qi-screening-email/archive/refs/heads/main.zip
set ZIPFILE=%TEMP%\qi_update.zip
set EXTRACTDIR=%TEMP%\qi_update_extract

powershell -NoProfile -Command ^
  "[Net.ServicePointManager]::SecurityProtocol = 'Tls12';" ^
  "Invoke-WebRequest -Uri '%ZIPURL%' -OutFile '%ZIPFILE%' -UseBasicParsing"

if not exist "%ZIPFILE%" (
    echo.
    echo ERROR: Download failed. Check your internet connection and try again.
    echo.
    pause
    exit /b 1
)

:: ── Extract zip ──────────────────────────────────────────────────────────────
echo Extracting...

if exist "%EXTRACTDIR%" rd /s /q "%EXTRACTDIR%"
powershell -NoProfile -Command ^
  "Expand-Archive -Path '%ZIPFILE%' -DestinationPath '%EXTRACTDIR%' -Force"

del "%ZIPFILE%" 2>nul

:: ── Copy app files (preserve data\ and winpython\) ───────────────────────────
echo Applying update...

set SRC=%EXTRACTDIR%\qi-screening-email-main

copy /y "%SRC%\app.py"           "%~dp0app.py"           >nul
copy /y "%SRC%\send_emails.py"   "%~dp0send_emails.py"   >nul
copy /y "%SRC%\process_data.py"  "%~dp0process_data.py"  >nul
copy /y "%SRC%\preview.py"       "%~dp0preview.py"       >nul
copy /y "%SRC%\requirements.txt" "%~dp0requirements.txt" >nul
copy /y "%SRC%\launch.bat"       "%~dp0launch.bat"       >nul
copy /y "%SRC%\INSTALL.bat"      "%~dp0INSTALL.bat"      >nul
copy /y "%SRC%\UPDATE.bat"       "%~dp0UPDATE.bat"       >nul
copy /y "%SRC%\templates\email_template.html" "%~dp0templates\email_template.html" >nul

:: Clean up temp files
rd /s /q "%EXTRACTDIR%" 2>nul

:: ── Re-run pip to pick up any new packages ────────────────────────────────────
echo Updating packages...

set PYTHON=
for /d %%d in ("winpython\python-*") do (
    if exist "%%d\python.exe" ( set PYTHON=%%d\python.exe & goto :pipupdate )
)
for /d %%w in ("winpython\WPy64-*") do (
    for /d %%d in ("%%w\python-*") do (
        if exist "%%d\python.exe" ( set PYTHON=%%d\python.exe & goto :pipupdate )
    )
)
where python >nul 2>&1 && set PYTHON=python

:pipupdate
if not "%PYTHON%"=="" (
    "%PYTHON%" -m pip install -r requirements.txt --quiet
)

echo.
echo ============================================================
echo   Update complete! Launch the app as usual.
echo ============================================================
echo.
pause
