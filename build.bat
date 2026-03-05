@echo off
REM ================================================================
REM  Marg ERP Auto Printer — Build Script
REM  Developed by Mehak Singh | TheMehakCodes
REM  Using --onedir mode to avoid PyInstaller DLL extraction issues
REM  in protected directories (Program Files)
REM ================================================================

echo.
echo ============================================================
echo  Marg ERP Auto Printer — Build Script
echo  Developed by Mehak Singh ^| TheMehakCodes
echo ============================================================
echo.

echo [1/4] Cleaning old build files...
echo ------------------------------------------------
if exist build      rmdir /s /q build
if exist marg_auto_printer.spec del marg_auto_printer.spec

echo.
echo [2/4] Installing / upgrading dependencies...
echo ------------------------------------------------
pip install --upgrade pyinstaller pywin32 pystray pillow requests

echo.
echo [3/4] Building application (onedir mode)...
echo ------------------------------------------------

REM  --onedir  = all DLLs sit beside the EXE permanently
REM             no _MEI extraction folder, no DLL-load race on startup
REM             this is the correct solution for apps installed in
REM             Program Files where the EXE cannot write to its own dir

pyinstaller ^
--noconfirm ^
--clean ^
--onedir ^
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
echo [4/4] Checking build result...
echo ------------------------------------------------

if exist "dist\marg_auto_printer\marg_auto_printer.exe" (
    echo.
    echo BUILD SUCCESSFUL
    echo.
    echo Output folder:
    echo dist\marg_auto_printer\
    echo.
    echo IMPORTANT - Inno Setup change needed ^(see below^)
    echo.
) else (
    echo.
    echo BUILD FAILED - Check errors above.
)

echo.
pause