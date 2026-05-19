@echo off
cd /d "%~dp0"
echo ================================================
echo  CubeLogReader - building EXE
echo ================================================
echo.

if not exist "main.py" (
    echo ERROR: main.py not found
    pause
    exit /b 1
)

pyinstaller --noconfirm ^
    --onedir ^
    --windowed ^
    --name CubeLogReader ^
    --collect-all google ^
    --collect-all grpc ^
    --hidden-import pythoncom ^
    --hidden-import win32com.client ^
    --hidden-import win32timezone ^
    main.py

if errorlevel 1 (
    echo.
    echo ERROR: build failed
    pause
    exit /b 1
)

echo.
echo ================================================
echo  Build complete!
echo ================================================
echo.
echo Exe folder: dist\CubeLogReader\
echo Main file:  dist\CubeLogReader\CubeLogReader.exe
echo.
echo Copy this ENTIRE folder to the other computer and
echo double-click the exe. Python is not required.
echo.
pause
