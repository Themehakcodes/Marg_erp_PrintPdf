@echo off
setlocal
:wait
tasklist /FI "PID eq 14812" 2>nul | find /I "14812" >nul
if not errorlevel 1 (
    ping 127.0.0.1 -n 1 -w 500 >nul
    goto wait
)
move /Y "D:\projects\Marg_erp_PrintPdf\dist\update_new.exe" "D:\projects\Marg_erp_PrintPdf\dist\marg_auto_printer.exe" >nul 2>&1
start "" "D:\projects\Marg_erp_PrintPdf\dist\marg_auto_printer.exe"
(goto) 2>nul & del "%~f0"
