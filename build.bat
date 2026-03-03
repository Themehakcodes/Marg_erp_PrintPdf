@echo off
REM ================================================================
REM  Marg ERP Auto Printer — Build Script
REM  Developed by Mehak Singh | TheMehakCodes
REM
REM  Run this bat file from inside your project root folder
REM  (same folder that contains marg_auto_printer.py)
REM
REM  Output EXE:  dist\marg_auto_printer.exe
REM  Copy that EXE into InstallerBuild\ before running Inno Setup.
REM ================================================================

echo.
echo  ============================================================
echo   Marg ERP Auto Printer — Build Script
echo   Developed by Mehak Singh ^| TheMehakCodes
echo  ============================================================
echo.

echo  [1/3]  Installing / upgrading Python dependencies...
echo  -----------------------------------------------
pip install --upgrade pyinstaller pywin32 pystray Pillow requests

echo.
echo  [2/3]  Building windowless single-file EXE with PyInstaller...
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
    --hidden-import "requests" ^
    --hidden-import "urllib3" ^
    --hidden-import "certifi" ^
    --hidden-import "charset_normalizer" ^
    --hidden-import "idna" ^
    marg_auto_printer.py

echo.
echo  [3/3]  Checking output...
if exist "dist\marg_auto_printer.exe" (
    echo  ✔  BUILD SUCCESSFUL
    echo.
    echo  EXE location:  dist\marg_auto_printer.exe
    echo.
    echo  Next steps:
    echo    1. Copy dist\marg_auto_printer.exe  →  InstallerBuild\
    echo    2. Open Inno Setup and compile installer.iss
) else (
    echo  ✖  BUILD FAILED — check errors above
)

echo.
echo  ================================================================
echo.
pause