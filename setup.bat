@echo off
cd /d "%~dp0"
echo ================================================
echo  CubeLogReader - Setup
echo ================================================
echo.
echo Installing Python packages, please wait...
echo.
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo ERROR: setup failed.
    echo - Is Python installed? Install Python 3.10+ from https://python.org/downloads
    echo - When installing, tick the "Add Python to PATH" box.
    echo.
    pause
    exit /b 1
)
echo.
echo ================================================
echo  Setup complete!
echo ================================================
echo.
echo Now double-click "CubeLogReader.bat" (or the exe) to start the program.
echo On first launch it will ask for your API key — get a free one from
echo aistudio.google.com/apikey.
echo.
pause
