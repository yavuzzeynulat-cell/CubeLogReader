@echo off
cd /d "%~dp0"
echo ================================================
echo  CubeLogReader - building EXE (launcher mode)
echo ================================================
echo.

if not exist "launcher.py" (
    echo ERROR: launcher.py not found
    pause
    exit /b 1
)
if not exist "main.py" (
    echo ERROR: main.py not found
    pause
    exit /b 1
)

pyinstaller --noconfirm CubeLogReader.spec

if errorlevel 1 (
    echo.
    echo ERROR: build failed
    pause
    exit /b 1
)

echo.
echo Copying src/ files into dist...
set SRCDIR=dist\CubeLogReader\src
if not exist "%SRCDIR%" mkdir "%SRCDIR%"
copy /Y main.py     "%SRCDIR%\" >nul
copy /Y reader.py   "%SRCDIR%\" >nul
copy /Y writer.py   "%SRCDIR%\" >nul
copy /Y updater.py  "%SRCDIR%\" >nul
copy /Y version.txt "%SRCDIR%\" >nul

echo.
echo ================================================
echo  Build complete!
echo ================================================
echo.
echo Exe folder: dist\CubeLogReader\
echo Main file:  dist\CubeLogReader\CubeLogReader.exe
echo Source dir: dist\CubeLogReader\src\
echo.
echo Copy this ENTIRE folder to the other computer and
echo double-click the exe. Python is not required.
echo.
pause
