@echo off
setlocal
:wait
tasklist /FI "PID eq 25780" 2>nul | find /I "25780" >nul
if not errorlevel 1 (
    ping 127.0.0.1 -n 1 -w 500 >nul
    goto wait
)
ping 127.0.0.1 -n 4 >nul
powershell.exe -NoProfile -WindowStyle Hidden -Command "Expand-Archive -Path \"D:\projects\Marg_erp_PrintPdf\dist\marg_auto_printer\update_new.zip\" -DestinationPath \"D:\projects\Marg_erp_PrintPdf\dist\marg_auto_printer\" -Force"
ping 127.0.0.1 -n 2 >nul
start "" "D:\projects\Marg_erp_PrintPdf\dist\marg_auto_printer\marg_auto_printer.exe"
del /Q "D:\projects\Marg_erp_PrintPdf\dist\marg_auto_printer\update_new.zip" 2>nul
(goto) 2>nul & del "%~f0"
