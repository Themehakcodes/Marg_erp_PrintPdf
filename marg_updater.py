# ============================================================
#  Marg ERP Auto Printer — Updater
#  Developed by Mehak Singh | TheMehakCodes
#
#  This is a SEPARATE tiny script built into its own EXE.
#  It is launched by the main app when an update is ready.
#  It waits for the main app to exit, replaces the EXE, restarts.
#
#  Build command:
#    pyinstaller --onefile --windowed --name marg_updater
#                --icon logo.ico marg_updater.py
# ============================================================

import sys
import os
import time
import ctypes
import subprocess

def is_process_running(pid: int) -> bool:
    """Return True if a process with the given PID is still alive."""
    try:
        import ctypes
        handle = ctypes.windll.kernel32.OpenProcess(0x0400, False, pid)
        if not handle:
            return False
        exit_code = ctypes.c_ulong(0)
        ctypes.windll.kernel32.GetExitCodeProcess(handle, ctypes.byref(exit_code))
        ctypes.windll.kernel32.CloseHandle(handle)
        return exit_code.value == 259  # STILL_ACTIVE = 259
    except Exception:
        return False


def main():
    # Args: <old_pid> <update_new.exe path> <target_exe path>
    if len(sys.argv) < 4:
        return

    old_pid     = int(sys.argv[1])
    update_path = sys.argv[2]   # e.g. C:\Program Files\...\update_new.exe
    target_path = sys.argv[3]   # e.g. C:\Program Files\...\marg_auto_printer.exe

    # ── Wait for main app to fully exit (max 30 seconds) ──────────
    for _ in range(60):
        if not is_process_running(old_pid):
            break
        time.sleep(0.5)

    # ── Extra buffer for Windows to release file handles ──────────
    time.sleep(3)

    # ── Replace the EXE ───────────────────────────────────────────
    try:
        if os.path.exists(target_path):
            os.remove(target_path)
        os.rename(update_path, target_path)
    except Exception as e:
        # If rename fails (cross-device), try copy+delete
        try:
            import shutil
            shutil.copy2(update_path, target_path)
            os.remove(update_path)
        except Exception:
            pass

    # ── 1 second buffer after replace ────────────────────────────
    time.sleep(1)

    # ── Restart the new EXE ───────────────────────────────────────
    try:
        subprocess.Popen(
            [target_path],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
        )
    except Exception:
        pass


if __name__ == "__main__":
    main()