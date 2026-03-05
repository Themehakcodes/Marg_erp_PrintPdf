Copy

@echo off
REM ================================================================
REM  Marg ERP Auto Printer — Build Script
REM  Developed by Mehak Singh | TheMehakCodes
REM ================================================================

echo.
echo ============================================================
echo  Marg ERP Auto Printer — Build Script
echo  Developed by Mehak Singh ^| TheMehakCodes
echo ============================================================
echo.

echo [1/5] Cleaning old build files...
echo ------------------------------------------------
if exist build rmdir /s /q build
if exist marg_auto_printer.spec del marg_auto_printer.spec
if exist marg_updater.spec       del marg_updater.spec

echo.
echo [2/5] Installing / upgrading dependencies...
echo ------------------------------------------------
pip install --upgrade pyinstaller pywin32 pystray pillow requests

echo.
echo [3/5] Building main application (onefile)...
echo ------------------------------------------------
pyinstaller ^
--noconfirm ^
--clean ^
--onefile ^
--windowed ^
--name marg_auto_printer ^
--icon logo.ico ^
--add-data "logo.ico;." ^
--add-data "logo.png;." ^
--hidden-import win32print ^
--hidden-import win32api ^
--hidden-import win32con ^
--hidden-import pystray._win32 ^
--hidden-import PIL._imaging ^
--hidden-import PIL.Image ^
--hidden-import requests ^
--hidden-import urllib3 ^
--hidden-import certifi ^
--hidden-import charset_normalizer ^
--hidden-import idna ^
--collect-all win32print ^
--collect-all win32api ^
--collect-all pystray ^
--collect-all PIL ^
marg_auto_printer.py

echo.
echo [4/5] Building updater EXE (tiny onefile)...
echo ------------------------------------------------
pyinstaller ^
--noconfirm ^
--clean ^
--onefile ^
--windowed ^
--name marg_updater ^
--icon logo.ico ^
marg_updater.py

echo.
echo [5/5] Checking build results...
echo ------------------------------------------------
if exist "dist\marg_auto_printer.exe" (
    if exist "dist\marg_updater.exe" (
        echo.
        echo BUILD SUCCESSFUL
        echo.
        echo Output files:
        echo   dist\marg_auto_printer.exe   -- main app
        echo   dist\marg_updater.exe        -- updater (bundle with installer)
        echo.
        echo Next steps:
        echo   1. Copy BOTH EXEs to your InstallerBuild folder
        echo   2. Compile installer using Inno Setup
        echo.
    ) else (
        echo Main app built but updater FAILED - check errors above.
    )
) else (
    echo.
    echo BUILD FAILED - Check errors above.
)

echo.
pause