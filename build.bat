@echo off
REM ================================================================
REM  Marg ERP Auto Printer — Build Script (CLEANED)
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
if exist marg_auto_printer.spec del marg_auto_printer.spec
if exist marg_updater.spec del marg_updater.spec
if exist "dist\marg_auto_printer.exe" del "dist\marg_auto_printer.exe"
if exist "dist\marg_updater.exe" del "dist\marg_updater.exe"

echo.
echo [2/5] Installing / upgrading dependencies...
echo ------------------------------------------------
pip install --upgrade pyinstaller pywin32 pystray pillow requests PyMuPDF

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
--add-data "SumatraPDF.exe;." ^
--add-data "PDFtoPrinter_m.exe;." ^
--hidden-import win32api ^
--hidden-import win32con ^
--hidden-import win32event ^
--hidden-import win32file ^
--hidden-import win32gui ^
--hidden-import win32process ^
--hidden-import pystray ^
--hidden-import pystray._win32 ^
--hidden-import PIL ^
--hidden-import PIL._imaging ^
--hidden-import PIL.Image ^
--hidden-import PIL.ImageDraw ^
--hidden-import fitz ^
--hidden-import queue ^
--hidden-import threading ^
--hidden-import logging ^
--hidden-import logging.handlers ^
--hidden-import urllib.request ^
--hidden-import hashlib ^
--hidden-import json ^
--hidden-import datetime ^
--hidden-import tkinter ^
--hidden-import tkinter.ttk ^
--hidden-import tkinter.filedialog ^
--hidden-import tkinter.messagebox ^
--collect-all win32print ^
--collect-all win32api ^
--collect-all win32event ^
--collect-all pystray ^
--collect-all PIL ^
--collect-all tkinter ^
--collect-all fitz ^
--exclude-module matplotlib ^
--exclude-module numpy ^
--exclude-module pandas ^
--exclude-module scipy ^
--exclude-module PyQt5 ^
--exclude-module PySide2 ^
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
        echo ========== BUILD SUCCESSFUL ==========
        echo.
        echo Output files:
        echo   dist\marg_auto_printer.exe   -- main app
        echo   dist\marg_updater.exe        -- updater
        echo.
        echo File sizes:
        for %%I in ("dist\marg_auto_printer.exe") do echo   main app: %%~zI bytes
        for %%I in ("dist\marg_updater.exe") do echo   updater:  %%~zI bytes
        echo.
        echo Next steps:
        echo   1. Copy BOTH EXEs to your InstallerBuild folder
        echo   2. Test the EXE: dist\marg_auto_printer.exe
        echo   3. If working, compile installer with Inno Setup
        echo.
    ) else (
        echo.
        echo ⚠️  Main app built but updater FAILED
        echo    Check errors above.
        echo.
    )
) else (
    echo.
    echo ========== BUILD FAILED ==========
    echo.
    echo Check errors above. Common issues:
    echo   - Missing PyMuPDF: pip install PyMuPDF
    echo   - Missing logo.ico file
    echo   - Missing SumatraPDF.exe
    echo   - Missing PDFtoPrinter_m.exe
    echo.
)

echo.
pause