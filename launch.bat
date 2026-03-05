@echo off
title QI Screening Email Pipeline

:: Change to the folder where this script lives
cd /d "%~dp0"

:: Install / update dependencies silently
echo Installing dependencies...
pip install -r requirements.txt --quiet

:: Open the browser and start the app
echo.
echo Starting app — your browser will open automatically.
echo Close this window to stop the app.
echo.
streamlit run app.py --server.headless false

pause
