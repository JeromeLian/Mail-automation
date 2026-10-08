@echo off
rem Starts the app. The first time (or if .setup_done is deleted) it runs setup.bat to install the packages.
cd /d "%~dp0"

if not exist ".setup_done" (
    call setup.bat
    if errorlevel 1 exit /b 1
)

start "" pythonw app.py
