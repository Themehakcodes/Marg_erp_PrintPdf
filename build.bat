@echo off
REM ================================================================
REM  Marg ERP Auto Printer — Build Script
REM  Developed by Mehak Singh | TheMehakCodes
REM ================================================================

echo.
echo  Installing / upgrading Python dependencies...
echo  -----------------------------------------------
pip install --upgrade pyinstaller pywin32 pystray Pillow

echo.
echo  Building main application EXE...
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
echo  Building updater EXE...
echo  -----------------------------------------------
pyinstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name "updater" ^
    --icon "logo.ico" ^
    updater.py

echo.
echo  Moving files to InstallerBuild...
echo  -----------------------------------------------
copy /Y dist\marg_auto_printer.exe InstallerBuild\
copy /Y dist\updater.exe InstallerBuild\
copy /Y version.txt InstallerBuild\

echo.
echo  ================================================================
echo  BUILD COMPLETE
echo  Main EXE:  InstallerBuild\marg_auto_printer.exe
echo  Updater:   InstallerBuild\updater.exe
echo  Version:   InstallerBuild\version.txt
echo.
echo  Next steps:
echo    1. Test the application
echo    2. Create new GitHub release with these files
echo    3. Update version.txt for next release
echo  ================================================================
echo.
pause