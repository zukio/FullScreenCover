@echo off
setlocal

:: --- Activate virtual environment ---
call venv\Scripts\activate.bat

:: --- FullScreenCover Project Settings ---
set PROJECT_NAME=FullScreenCover
set ENTRY_SCRIPT=main.py

:: --- PyInstaller Options (without --windowed for console output) ---
set OPTIONS=--name %PROJECT_NAME% ^
 --onefile ^
 --windowed ^
 --add-data "assets;assets" ^
 --add-data "modules;modules" ^
 --icon "assets\icon.ico" ^
 --hidden-import "asyncio" ^
 --hidden-import "tkinter" ^
 --hidden-import "tkinter.ttk" ^
 --hidden-import "tkinter.filedialog" ^
 --hidden-import "tkinter.messagebox" ^
 --hidden-import "pystray" ^
 --hidden-import "PIL" ^
 --hidden-import "cv2" ^
 --hidden-import "pynput" ^
 --hidden-import "psutil" ^
 --hidden-import "win32gui" ^
 --hidden-import "win32api" ^
 --hidden-import "win32con" ^
 --hidden-import "pycaw.pycaw" ^
 --hidden-import "comtypes"

:: --- Clean old builds ---
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist %PROJECT_NAME%.spec del %PROJECT_NAME%.spec

:: --- Build with PyInstaller ---
echo Building %PROJECT_NAME%...
echo Options: %OPTIONS%
echo.
pyinstaller %OPTIONS% %ENTRY_SCRIPT%

:: --- Build completion message ---
echo.
echo =====================================
echo Build complete! Check /dist folder.
echo =====================================
echo.

:: --- Copy config.json to dist folder ---
echo Copying config.json to dist folder...
if exist config.json (
    copy config.json dist\
    echo config.json copied successfully.
) else (
    echo Warning: config.json not found!
)
echo.

echo Built files:
if exist dist\%PROJECT_NAME%.exe (
    dir dist\%PROJECT_NAME%.exe
    if exist dist\config.json (
        dir dist\config.json
    )
) else (
    echo Build failed!
)
echo.

pause
endlocal
