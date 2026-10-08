@echo off
rem Installs the Python packages the app needs. run.bat calls this the first time only.
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Install Python 3.12 from python.org, tick "Add python.exe to PATH", then run run.bat again.
    pause
    exit /b 1
)

echo Installing the packages this app needs. This can take a minute...
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo The install failed. Check your internet connection and read the messages above.
    pause
    exit /b 1
)

python -c "import win32com.client, PIL" >nul 2>nul
if errorlevel 1 (
    echo.
    echo The packages installed, but the app could not load them. Send the messages above to whoever set this up.
    pause
    exit /b 1
)

echo Done. > ".setup_done"
echo Setup finished.
exit /b 0
