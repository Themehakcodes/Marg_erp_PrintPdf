@echo off
REM ================================================================
REM  Marg ERP Auto Printer — Build Script
REM  Developed by Mehak Singh | TheMehakCodes
REM
REM  Run this bat file from inside InstallerBuild\
REM  Output EXE will be in:  InstallerBuild\dist\marg_auto_printer.exe
REM  Copy that EXE back into InstallerBuild\ before running Inno Setup.
REM ================================================================

echo.
echo  Installing / upgrading Python dependencies...
echo  -----------------------------------------------
pip install --upgrade pyinstaller pywin32 pystray Pillow

echo.
echo  Building windowless EXE with PyInstaller...
echo  -----------------------------------------------

pyinstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "marg_auto_printer" ^
    --icon "logo.ico" ^
    --add-data "logo.ico;." ^
    --add-data "logo.png;." ^
    --hidden-import "win32print" ^
    --hidden-import "win32api" ^
    --hidden-import "win32con" ^
    --hidden-import "pystray._win32" ^
    --hidden-import "PIL._imaging" ^
    --hidden-import "PIL.Image" ^
    marg_auto_printer.py

echo.
echo  ================================================================
echo  BUILD COMPLETE
echo  EXE location:  dist\marg_auto_printer.exe
echo.
echo  Next steps:
echo    1. Copy dist\marg_auto_printer.exe  →  InstallerBuild\
echo    2. Open Inno Setup and compile installer.iss
echo  ================================================================
echo.
pause