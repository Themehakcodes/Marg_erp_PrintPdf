# ============================================================
#  Marg ERP Auto Printer
#  Developed by Mehak Singh | TheMehakCodes
#  Version: 1.2.2 (Enhanced Stability Release)
# ============================================================
#
#  IMPORTANT — HOW THE WINDOW IS HIDDEN:
#  1. Built with PyInstaller --windowed  → no console allocated at all
#  2. As a safety net, the very first thing we do (before ANY import
#     that might flash a window) is call the Win32 API to hide the
#     console window if one somehow exists.
#
#  UPDATE TYPES (controlled by version.json → "release_type"):
#  • "direct"    → new .exe replaces current .exe silently, auto-restarts
#  • "installer" → downloaded Setup.exe is launched for the user to run
# ============================================================

import ctypes
import sys

# ── Hide console window immediately (belt-and-suspenders) ──────────
def _hide_console():
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)   # SW_HIDE = 0
            ctypes.windll.kernel32.FreeConsole()
    except Exception:
        pass

_hide_console()

# ── Single instance check (prevents multiple copies) ───────────────
def ensure_single_instance():
    """Ensure only one instance of the application runs"""
    try:
        import win32event
        import win32api
        import winerror
        
        mutex_name = "MargERPAutoPrinter_SingleInstance_Mutex"
        mutex = win32event.CreateMutex(None, False, mutex_name)
        last_error = win32api.GetLastError()
        
        if last_error == winerror.ERROR_ALREADY_EXISTS:
            # Another instance is running - silently exit
            sys.exit(0)
    except:
        pass

ensure_single_instance()

# ── Now safe to import everything else ─────────────────────────────
import os
import time
import json
import subprocess
import threading
import queue
import win32print
import win32api
import win32event
#import win32spool
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
import hashlib
import urllib.request
from PIL import Image, ImageDraw
import pystray
import logging
from logging.handlers import RotatingFileHandler

# ==============================
# VERSION
# ==============================

APP_VERSION        = "1.2.3"
UPDATE_VERSION_URL = "https://raw.githubusercontent.com/Themehakcodes/Marg_erp_PrintPdf/main/version.json"

# ==============================
# BASE PATH  (works frozen & raw)
# ==============================

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE  = os.path.join(BASE_DIR, "config.json")
SUMATRA_PATH = os.path.join(BASE_DIR, "SumatraPDF.exe")
ICON_PATH    = os.path.join(BASE_DIR, "logo.ico")
LOG_FILE     = os.path.join(BASE_DIR, "marg_printer.log")

# ==============================
# GLOBAL VARIABLES
# ==============================

log_queue          = queue.Queue()
log_lines          = []              # (level, text)  — full in-memory history
MAX_LOGS           = 500

# Printer queue management (AUTO ONLY - no user options)
LAST_PRINTER_CLEAR = 0
PRINTER_CLEAR_COOLDOWN = 300  # 5 minutes between automatic clears
MAX_CONSECUTIVE_FAILURES = 3

# Processed files tracking (prevents duplicates)
PROCESSED_FILES = {}  # filename -> timestamp
PROCESSED_EXPIRY = 30  # seconds
MAX_PROCESSED_FILES = 1000  # Maximum entries

# Thread management
_print_queue = queue.Queue()
_queued_files = set()
_queued_lock = threading.Lock()
stop_event = threading.Event()

# ==============================
# FILE LOGGING SETUP
# ==============================

def setup_file_logging():
    """Setup file-based logging with rotation"""
    logger = logging.getLogger('MargPrinter')
    logger.setLevel(logging.INFO)
    
    # Remove existing handlers
    logger.handlers.clear()
    
    # Create rotating file handler
    handler = RotatingFileHandler(
        LOG_FILE, maxBytes=5*1024*1024, backupCount=3  # 5MB per file, 3 backups
    )
    handler.setFormatter(logging.Formatter(
        '%(asctime)s - %(levelname)s - %(message)s'
    ))
    logger.addHandler(handler)
    return logger

file_logger = setup_file_logging()

# ==============================
# THEME
# ==============================

THEME = {
    "bg":       "#0F1117",
    "surface":  "#1A1D27",
    "surface2": "#22263A",
    "accent":   "#4F8EF7",
    "accent2":  "#7C5CF6",
    "success":  "#2ECC71",
    "warning":  "#F39C12",
    "danger":   "#E74C3C",
    "text":     "#E8ECF4",
    "text_dim": "#7A8099",
    "border":   "#2E3450",
    "input_bg": "#13161F",
}

FONT_LABEL   = ("Segoe UI", 10)
FONT_LABEL_SM= ("Segoe UI",  9)
FONT_MONO    = ("Consolas",  9)
FONT_BTN     = ("Segoe UI", 10, "bold")
FONT_CREDIT  = ("Segoe UI",  8, "italic")

LEVEL_COLORS = {
    "INFO":    "#7ABAFF",
    "SUCCESS": "#2ECC71",
    "WARN":    "#F39C12",
    "ERROR":   "#E74C3C",
    "PRINT":   "#B388FF",
    "WATCH":   "#4F8EF7",
    "UPDATE":  "#FF79C6",
}

# ==============================
# LOGGER  (thread-safe, no print)
# ==============================

def log(msg, level="INFO"):
    now   = datetime.now().strftime("%H:%M:%S")
    icons = {
        "INFO":    "ℹ",  "SUCCESS": "✔",  "WARN":  "⚠",
        "ERROR":   "✖",  "PRINT":   "🖨",  "WATCH": "👁",
        "UPDATE":  "🔄",
    }
    entry = f"[{now}]  {level:<7}  {icons.get(level, '•')}  {msg}"
    log_lines.append((level, entry))
    if len(log_lines) > MAX_LOGS:
        log_lines.pop(0)
    log_queue.put((level, entry))
    
    # Also log to file
    try:
        if level == "ERROR":
            file_logger.error(msg)
        elif level == "WARN":
            file_logger.warning(msg)
        else:
            file_logger.info(msg)
    except:
        pass

# ==============================
# DEFAULT CONFIG
# ==============================

DEFAULT_CONFIG = {
    "printer":        "",
    "watch_folder":   os.path.join(os.path.expanduser("~"), "Downloads"),
    "file_prefix":    "Marg_erp",
    "check_interval": 3,
    "silent_mode":    True,
}

# ==============================
# HIDDEN TK ROOT  (keeps Tk alive)
# ==============================

_tk_root: tk.Tk = None   # type: ignore

def get_root() -> tk.Tk:
    global _tk_root
    if _tk_root is None:
        _tk_root = tk.Tk()
        _tk_root.withdraw()
        _tk_root.title("MargERPAutoPrinter")
        if os.path.exists(ICON_PATH):
            try: _tk_root.iconbitmap(ICON_PATH)
            except Exception: pass
    return _tk_root

# ==============================
# ENTRY-STYLE HELPER
# ==============================

def _entry_kw():
    return dict(
        bg=THEME["input_bg"], fg=THEME["text"],
        insertbackground=THEME["accent"], relief="flat", bd=0,
        font=FONT_LABEL, highlightthickness=1,
        highlightbackground=THEME["border"],
        highlightcolor=THEME["accent"],
    )

def _combo_style():
    s = ttk.Style()
    s.theme_use("clam")
    s.configure("Dark.TCombobox",
        fieldbackground=THEME["input_bg"], background=THEME["input_bg"],
        foreground=THEME["text"], arrowcolor=THEME["accent"],
        bordercolor=THEME["border"], lightcolor=THEME["border"],
        darkcolor=THEME["border"], selectbackground=THEME["accent2"],
        selectforeground=THEME["text"], padding=6,
    )

# ==============================
# PRINTER UTILITY FUNCTIONS (AUTO ONLY)
# ==============================

def get_printer_status_detailed(printer_name):
    """Check if printer is ready and get detailed status (internal use only)"""
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        printer_info = win32print.GetPrinter(hprinter, 2)
        
        # Get job count
        jobs = win32print.EnumJobs(hprinter, 0, -1, 1)
        job_count = len(jobs) if jobs else 0
        
        win32print.ClosePrinter(hprinter)
        
        status = printer_info['Status']
        if status == 0:
            return True, "Ready", job_count
        elif status & win32print.PRINTER_STATUS_PAUSED:
            return False, "Paused", job_count
        elif status & win32print.PRINTER_STATUS_ERROR:
            return False, "Error", job_count
        elif status & win32print.PRINTER_STATUS_PAPER_JAM:
            return False, "Paper Jam", job_count
        elif status & win32print.PRINTER_STATUS_PAPER_OUT:
            return False, "Out of Paper", job_count
        elif status & win32print.PRINTER_STATUS_OFFLINE:
            return False, "Offline", job_count
        elif status & win32print.PRINTER_STATUS_BUSY:
            return False, "Busy", job_count
        else:
            return True, "Ready", job_count
    except:
        return False, "Cannot Access Printer", 0

def clear_windows_spooler():
    """Clear Windows print spooler service cache (auto internal only)"""
    try:
        # Stop spooler service
        subprocess.run(
            ["net", "stop", "spooler", "/y"],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        time.sleep(2)
        
        # Clear spooler folder
        spool_path = os.path.join(os.environ.get('SYSTEMROOT', 'C:\\Windows'), 
                                  'System32', 'spool', 'PRINTERS')
        if os.path.exists(spool_path):
            for file in os.listdir(spool_path):
                try:
                    os.remove(os.path.join(spool_path, file))
                except:
                    pass
        
        # Start spooler service
        subprocess.run(
            ["net", "start", "spooler"],
            capture_output=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        
        time.sleep(2)
        return True
    except:
        return False

def auto_clear_printer_queue(printer_name):
    """
    Automatically clear printer queue when needed (internal only - no user interaction)
    """
    global LAST_PRINTER_CLEAR
    
    current_time = time.time()
    if (current_time - LAST_PRINTER_CLEAR) < PRINTER_CLEAR_COOLDOWN:
        return False
    
    try:
        hprinter = win32print.OpenPrinter(printer_name)
        jobs = win32print.EnumJobs(hprinter, 0, -1, 1)
        
        if jobs:
            cleared_count = 0
            for job in jobs:
                try:
                    win32print.SetJob(hprinter, job['JobId'], 0, None, win32print.JOB_CONTROL_DELETE)
                    cleared_count += 1
                except:
                    pass
            
            win32print.ClosePrinter(hprinter)
            
            if cleared_count > 0:
                clear_windows_spooler()
                log(f"Auto-cleared {cleared_count} stuck jobs from printer", "INFO")
                LAST_PRINTER_CLEAR = current_time
                return True
        else:
            win32print.ClosePrinter(hprinter)
        
    except Exception as e:
        log(f"Auto-clear attempt failed: {e}", "WARN")
    
    return False

# ==============================
# FILE UTILITY FUNCTIONS
# ==============================

def is_file_locked(filepath):
    """Check if file is locked by another process"""
    if not os.path.exists(filepath):
        return False
    
    try:
        with open(filepath, 'rb') as f:
            f.read(1)
        return False
    except (IOError, OSError):
        return True

def is_file_stable(filepath, wait_seconds=1):
    """Check if file is completely written and stable"""
    if not os.path.exists(filepath):
        return False
    
    try:
        if is_file_locked(filepath):
            return False
        
        size1 = os.path.getsize(filepath)
        if size1 == 0:
            return False
            
        time.sleep(wait_seconds)
        size2 = os.path.getsize(filepath)
        return size1 == size2
    except:
        return False

def is_valid_pdf(filepath):
    """Quick check if file is a valid PDF"""
    try:
        with open(filepath, 'rb') as f:
            header = f.read(5)
            return header == b'%PDF-'
    except:
        return False

def is_network_path(path):
    """Check if path is on network drive"""
    try:
        drive = os.path.splitdrive(path)[0]
        if drive:
            DRIVE_REMOTE = 4
            return ctypes.windll.kernel32.GetDriveTypeW(drive) == DRIVE_REMOTE
    except:
        pass
    return False

# ==============================
# AUTO-UPDATER (YOUR ORIGINAL CODE - UNCHANGED)
# ==============================

def _parse_version(v: str):
    """Convert '1.2.3' → (1, 2, 3) for safe semantic comparison."""
    try:
        return tuple(int(x) for x in str(v).strip().split("."))
    except Exception:
        return (0, 0, 0)

def _apply_direct_update(exe_url: str, remote_ver: str,
                          remote_sha: str, silent: bool, parent_win):
    """
    DIRECT EXE update flow  (release_type == "direct")
    """
    current_exe = sys.executable if getattr(sys, "frozen", False) \
                   else os.path.abspath(__file__)
    app_dir      = os.path.dirname(current_exe)
    update_path  = os.path.join(app_dir, "update_new.exe")
    updater_exe  = os.path.join(app_dir, "marg_updater.exe")

    log("Downloading update…", "UPDATE")

    # Download
    dl_req = urllib.request.Request(
        exe_url,
        headers={"User-Agent": "MargERPAutoPrinter-Updater"}
    )
    hasher = hashlib.sha256()
    with urllib.request.urlopen(dl_req, timeout=180) as dl_resp, \
         open(update_path, "wb") as out_f:
        while True:
            chunk = dl_resp.read(65536)
            if not chunk:
                break
            out_f.write(chunk)
            hasher.update(chunk)

    # Size guard
    if os.path.getsize(update_path) < 100_000:
        log("Downloaded file too small — aborting.", "ERROR")
        try: os.remove(update_path)
        except Exception: pass
        return

    # Optional SHA-256 check
    if remote_sha:
        actual_sha = hasher.hexdigest().lower()
        if actual_sha != remote_sha.lower():
            log(f"Checksum mismatch — aborting. Got: {actual_sha}", "ERROR")
            try: os.remove(update_path)
            except Exception: pass
            return
        log("SHA-256 checksum verified ✔", "SUCCESS")

    log(f"Download complete. Launching updater for v{remote_ver}…", "UPDATE")

    if not os.path.exists(updater_exe):
        log("marg_updater.exe not found — cannot apply update.", "ERROR")
        return

    pid = os.getpid()

    # Detect protected install directory
    protected = (
        os.environ.get("ProgramFiles",      "C:\\Program Files"),
        os.environ.get("ProgramFiles(x86)", "C:\\Program Files (x86)"),
        os.environ.get("SystemRoot",        "C:\\Windows"),
    )
    needs_elevation = any(
        app_dir.lower().startswith(p.lower()) for p in protected if p
    )

    if needs_elevation:
        # Use PowerShell Start-Process -Verb RunAs to elevate updater
        log("Protected dir — requesting UAC elevation for updater…", "UPDATE")
        args = f'{pid} "{update_path}" "{current_exe}"'
        ps_cmd = (
            f'Start-Process "{updater_exe}" '
            f'-ArgumentList \'{args}\' '
            f'-Verb RunAs '
            f'-WindowStyle Hidden'
        )
        subprocess.Popen(
            ["powershell.exe", "-NoProfile", "-WindowStyle", "Hidden",
             "-Command", ps_cmd],
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
    else:
        subprocess.Popen(
            [updater_exe, str(pid), update_path, current_exe],
            creationflags=subprocess.DETACHED_PROCESS | subprocess.CREATE_NO_WINDOW,
        )

    log("Updater launched — shutting down now…", "UPDATE")
    stop_event.set()
    get_root().after(0, get_root().destroy)

def _apply_installer_update(exe_url: str, remote_ver: str,
                             remote_sha: str, silent: bool, parent_win):
    """
    INSTALLER EXE update flow  (release_type == "installer")
    """
    current_exe = sys.executable if getattr(sys, "frozen", False) \
                  else os.path.abspath(__file__)
    app_dir      = os.path.dirname(current_exe)
    setup_path   = os.path.join(app_dir, "update_setup.exe")

    log("Downloading installer update…", "UPDATE")

    # Download
    dl_req = urllib.request.Request(
        exe_url,
        headers={"User-Agent": "MargERPAutoPrinter-Updater"}
    )
    hasher = hashlib.sha256()
    with urllib.request.urlopen(dl_req, timeout=180) as dl_resp, \
         open(setup_path, "wb") as out_f:
        while True:
            chunk = dl_resp.read(65536)
            if not chunk:
                break
            out_f.write(chunk)
            hasher.update(chunk)

    # Size guard
    if os.path.getsize(setup_path) < 100_000:
        log("Downloaded installer too small — aborting.", "ERROR")
        try: os.remove(setup_path)
        except Exception: pass
        return

    # Optional SHA-256 validation
    if remote_sha:
        actual_sha = hasher.hexdigest().lower()
        if actual_sha != remote_sha.lower():
            log(f"Checksum mismatch — aborting. Got: {actual_sha}", "ERROR")
            try: os.remove(setup_path)
            except Exception: pass
            return
        log("SHA-256 checksum verified ✔", "SUCCESS")

    log(f"Installer downloaded for v{remote_ver}. Prompting user…", "UPDATE")

    def _prompt():
        answer = messagebox.askyesno(
            "Update Available — Installer Ready",
            f"✔  A new version is ready to install.\n\n"
            f"Current version : v{APP_VERSION}\n"
            f"New version     : v{remote_ver}\n\n"
            f"The installer has been downloaded to:\n{setup_path}\n\n"
            f"Click YES to launch the installer now.\n"
            f"(You can also run it manually from the above path later.)",
            parent=parent_win or get_root()
        )
        if answer:
            log("User confirmed — launching installer…", "UPDATE")
            try:
                # ShellExecute with "runas" triggers proper UAC elevation
                ctypes.windll.shell32.ShellExecuteW(
                    None, "runas", setup_path, None, None, 1
                )
                log("Installer launched. Close this app if prompted.", "INFO")
            except Exception as e:
                log(f"Failed to launch installer: {e}", "ERROR")
                messagebox.showerror(
                    "Launch Failed",
                    f"Could not launch the installer automatically.\n\n"
                    f"Please run it manually:\n{setup_path}",
                    parent=parent_win or get_root()
                )
        else:
            log(f"User postponed installer update. File kept at: {setup_path}", "WARN")

    get_root().after(0, _prompt)

def _do_update_check(silent: bool = True, parent_win=None):
    """
    Core update logic — runs in a background thread.
    """
    try:
        log("Checking for updates…", "UPDATE")

        # Fetch version manifest
        req = urllib.request.Request(
            UPDATE_VERSION_URL,
            headers={"User-Agent": "MargERPAutoPrinter-Updater"}
        )
        with urllib.request.urlopen(req, timeout=10) as resp:
            data = json.loads(resp.read().decode())

        remote_ver    = data.get("version",      "0.0.0")
        exe_url       = data.get("exe_url",       "")
        remote_sha    = data.get("sha256",        "").lower()
        release_type  = data.get("release_type", "direct").lower()
        
        if release_type not in ("direct", "installer"):
            release_type = "direct"

        # Already up to date?
        if _parse_version(remote_ver) <= _parse_version(APP_VERSION):
            log(f"Already up to date  (v{APP_VERSION})", "SUCCESS")
            if not silent:
                def _show():
                    messagebox.showinfo(
                        "No Update Available",
                        f"✔  You are running the latest version.\n\nCurrent version: v{APP_VERSION}",
                        parent=parent_win or get_root()
                    )
                get_root().after(0, _show)
            return

        log(f"Update available!  v{APP_VERSION}  →  v{remote_ver}  "
            f"[type: {release_type}]", "UPDATE")

        # Route to the correct update handler
        if release_type == "installer":
            _apply_installer_update(exe_url, remote_ver, remote_sha,
                                    silent, parent_win)
        else:
            _apply_direct_update(exe_url, remote_ver, remote_sha,
                                 silent, parent_win)

    except Exception as e:
        err_type = type(e).__name__
        if "URLError" in err_type or "ConnectionRefused" in err_type:
            log("Update check skipped — no internet connection.", "INFO")
        elif "timeout" in str(e).lower() or "Timeout" in err_type:
            log("Update check timed out.", "WARN")
        else:
            log(f"Update check failed: {e}", "WARN")

        if not silent:
            def _err():
                messagebox.showwarning(
                    "Update Check Failed",
                    f"Could not check for updates.\n\nReason: {e}\n\nPlease check your internet connection.",
                    parent=parent_win or get_root()
                )
            get_root().after(0, _err)

def start_update_check(silent: bool = True, parent_win=None):
    """Kick off update check in a non-blocking daemon thread."""
    t = threading.Thread(
        target=_do_update_check,
        args=(silent, parent_win),
        daemon=True,
        name="Updater"
    )
    t.start()

# ==============================
# FIRST-TIME SETUP WINDOW (YOUR ORIGINAL - UNCHANGED)
# ==============================

def first_time_setup() -> dict:
    result = {}
    root   = get_root()

    win = tk.Toplevel(root)
    win.title("Marg ERP Auto Printer — Setup")
    win.geometry("520x570")
    win.resizable(False, False)
    win.configure(bg=THEME["bg"])
    win.grab_set()
    if os.path.exists(ICON_PATH):
        try: win.iconbitmap(ICON_PATH)
        except Exception: pass

    tk.Frame(win, bg=THEME["accent2"], height=6).pack(fill="x")
    hdr = tk.Frame(win, bg=THEME["bg"], pady=18); hdr.pack(fill="x")
    tk.Label(hdr, text="🖨  Marg ERP Auto Printer",
             font=("Segoe UI", 16, "bold"), fg=THEME["text"], bg=THEME["bg"]).pack()
    tk.Label(hdr, text="Initial Configuration",
             font=FONT_LABEL_SM, fg=THEME["text_dim"], bg=THEME["bg"]).pack()
    tk.Frame(win, bg=THEME["border"], height=1).pack(fill="x", padx=30)

    form = tk.Frame(win, bg=THEME["bg"], padx=30, pady=8)
    form.pack(fill="both", expand=True)

    def lbl(parent, text):
        tk.Label(parent, text=text, font=FONT_LABEL, fg=THEME["text_dim"],
                 bg=THEME["bg"], anchor="w").pack(fill="x", pady=(10, 2))

    lbl(form, "🖨  Printer")
    printers    = [p[2] for p in win32print.EnumPrinters(2)]
    printer_var = tk.StringVar(value=printers[0] if printers else "")
    _combo_style()
    ttk.Combobox(form, textvariable=printer_var, values=printers,
                 style="Dark.TCombobox", state="readonly").pack(fill="x")

    lbl(form, "📁  Watch Folder")
    folder_var = tk.StringVar(value=DEFAULT_CONFIG["watch_folder"])
    fr = tk.Frame(form, bg=THEME["bg"]); fr.pack(fill="x")
    tk.Entry(fr, textvariable=folder_var, **_entry_kw()).pack(
        side="left", fill="x", expand=True, ipady=6)
    tk.Button(fr, text=" Browse ",
              command=lambda: folder_var.set(filedialog.askdirectory() or folder_var.get()),
              bg=THEME["surface2"], fg=THEME["accent"], relief="flat",
              font=FONT_LABEL, cursor="hand2", bd=0,
              activebackground=THEME["accent"], activeforeground=THEME["bg"],
              padx=10, pady=6).pack(side="left", padx=(6,0))

    row = tk.Frame(form, bg=THEME["bg"]); row.pack(fill="x")
    lc  = tk.Frame(row,  bg=THEME["bg"]); lc.pack(side="left", fill="x", expand=True, padx=(0,10))
    rc  = tk.Frame(row,  bg=THEME["bg"]); rc.pack(side="left", fill="x", expand=True)
    lbl(lc, "🏷  File Prefix")
    prefix_var = tk.StringVar(value=DEFAULT_CONFIG["file_prefix"])
    tk.Entry(lc, textvariable=prefix_var, **_entry_kw()).pack(fill="x", ipady=6)
    lbl(rc, "⏱  Interval (sec)")
    interval_var = tk.StringVar(value=str(DEFAULT_CONFIG["check_interval"]))
    tk.Entry(rc, textvariable=interval_var, **_entry_kw()).pack(fill="x", ipady=6)

    silent_var = tk.BooleanVar(value=True)
    sf = tk.Frame(form, bg=THEME["bg"]); sf.pack(fill="x", pady=(14,0))
    tk.Label(sf, text="🔇  Silent Print Mode",
             font=FONT_LABEL, fg=THEME["text"], bg=THEME["bg"]).pack(side="left")
    ind = tk.Label(sf, text="ON ", bg=THEME["success"], fg=THEME["bg"],
                   font=("Segoe UI", 8, "bold"), padx=6, pady=2)
    ind.pack(side="right")
    def _tog():
        ind.config(bg=THEME["success"] if silent_var.get() else THEME["text_dim"],
                   text="ON " if silent_var.get() else "OFF")
    tk.Checkbutton(sf, variable=silent_var, command=_tog,
                   bg=THEME["bg"], selectcolor=THEME["surface2"],
                   relief="flat", bd=0, activebackground=THEME["bg"]).pack(side="right", padx=4)

    tk.Frame(win, bg=THEME["border"], height=1).pack(fill="x", padx=30)

    def _save():
        try:
            if not printer_var.get():
                messagebox.showerror("Error", "Please select a printer.", parent=win); return
            result.update({
                "printer":        printer_var.get(),
                "watch_folder":   folder_var.get(),
                "file_prefix":    prefix_var.get(),
                "check_interval": int(interval_var.get()),
                "silent_mode":    silent_var.get(),
            })
            with open(CONFIG_FILE, "w") as f:
                json.dump(result, f, indent=4)
            messagebox.showinfo("Saved", "✔  Configuration saved!", parent=win)
            win.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=win)

    bf = tk.Frame(win, bg=THEME["bg"], pady=14); bf.pack()
    tk.Button(bf, text="  Save Configuration  ", command=_save,
              bg=THEME["accent"], fg=THEME["bg"], font=FONT_BTN,
              relief="flat", bd=0, cursor="hand2", padx=20, pady=10,
              activebackground=THEME["accent2"], activeforeground=THEME["text"]).pack()

    tk.Label(win, text="Developed by Mehak Singh | TheMehakCodes",
             font=FONT_CREDIT, fg=THEME["text_dim"], bg=THEME["bg"]).pack(pady=(0,10))

    win.wait_window()
    return result

# ==============================
# LOAD CONFIG (YOUR ORIGINAL - UNCHANGED)
# ==============================

def load_config() -> dict:
    if "--config" in sys.argv or not os.path.exists(CONFIG_FILE):
        return first_time_setup()
    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except Exception:
        return first_time_setup()

_root_ready      = get_root()
CONFIG           = load_config()
SELECTED_PRINTER = CONFIG.get("printer",        DEFAULT_CONFIG["printer"])
WATCH_FOLDER     = CONFIG.get("watch_folder",   DEFAULT_CONFIG["watch_folder"])
FILE_PREFIX      = CONFIG.get("file_prefix",    DEFAULT_CONFIG["file_prefix"])
CHECK_INTERVAL   = CONFIG.get("check_interval", DEFAULT_CONFIG["check_interval"])
USE_SILENT_MODE  = CONFIG.get("silent_mode",    DEFAULT_CONFIG["silent_mode"])

# ==============================
# PRINT FUNCTIONS (YOUR ORIGINAL - ENHANCED BUT PRESERVED)
# ==============================

def print_pdf_legacy(fp: str):
    try:
        win32print.SetDefaultPrinter(SELECTED_PRINTER)
        win32api.ShellExecute(0, "print", fp, None, ".", 0)
        time.sleep(5)
        if os.path.exists(fp): 
            try:
                os.remove(fp)
            except:
                pass
    except Exception as e:
        log(f"Legacy print error: {e}", "ERROR")

def print_pdf_silent(fp: str):
    try:
        if not os.path.exists(SUMATRA_PATH):
            log("SumatraPDF not found — legacy fallback", "WARN")
            print_pdf_legacy(fp)
            return
        
        # Kill any hanging Sumatra processes first (added for stability)
        subprocess.run(["taskkill", "/F", "/IM", "SumatraPDF.exe"], 
                      capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
        time.sleep(1)
        
        # Validate PDF before printing (added for stability)
        if not is_valid_pdf(fp):
            log(f"Invalid or corrupted PDF: {os.path.basename(fp)}", "ERROR")
            try:
                os.remove(fp)
            except:
                pass
            return
        
        # Original print command
        process = subprocess.Popen(
            [
                SUMATRA_PATH,
                "-print-to", SELECTED_PRINTER,
                "-print-settings", "noscale",
                "-silent",
                "-exit-on-print",
                fp
            ],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        
        # Added timeout protection
        try:
            process.wait(timeout=30)
        except subprocess.TimeoutExpired:
            process.kill()
            log(f"Print timeout for: {os.path.basename(fp)}", "ERROR")
            return
        
        time.sleep(2)
        if os.path.exists(fp):
            try:
                os.remove(fp)
            except Exception as e:
                log(f"Could not delete {os.path.basename(fp)}: {e}", "WARN")
                
    except Exception as e:
        log(f"Silent print error: {e}", "ERROR")
        print_pdf_legacy(fp)

def print_pdf(fp: str):
    name = os.path.basename(fp)
    log(f"Sending → {name}", "PRINT")
    print_pdf_silent(fp) if USE_SILENT_MODE else print_pdf_legacy(fp)
    log(f"Done — {name}", "SUCCESS")

# ==============================
# PRINT QUEUE MANAGEMENT (ENHANCED - AUTO ONLY)
# ==============================

def _enqueue_file(fp: str):
    """Thread-safe: add a file to the print queue only once with deduplication"""
    with _queued_lock:
        filename = os.path.basename(fp)
        current_time = time.time()
        
        # Clean old entries - prevent memory leak
        if len(PROCESSED_FILES) > MAX_PROCESSED_FILES:
            sorted_files = sorted(PROCESSED_FILES.items(), key=lambda x: x[1])
            for f, _ in sorted_files[:200]:
                del PROCESSED_FILES[f]
        
        # Remove expired entries
        expired = [f for f, ts in PROCESSED_FILES.items() 
                   if current_time - ts > PROCESSED_EXPIRY]
        for f in expired:
            del PROCESSED_FILES[f]
        
        # Check if recently processed
        if filename in PROCESSED_FILES:
            log(f"Skipping duplicate: {filename} (processed {current_time - PROCESSED_FILES[filename]:.1f}s ago)", "WATCH")
            return
        
        if fp not in _queued_files:
            _queued_files.add(fp)
            PROCESSED_FILES[filename] = current_time
            _print_queue.put(fp)
            log(f"Queued  → {filename} (queue depth: {_print_queue.qsize()})", "WATCH")

def print_worker():
    """Worker thread that processes the print queue (enhanced with auto-clear)"""
    log("Print worker started — ready for jobs.", "INFO")
    consecutive_errors = 0
    consecutive_printer_failures = 0
    
    while not stop_event.is_set():
        try:
            fp = _print_queue.get(timeout=1)
        except queue.Empty:
            continue
            
        try:
            if os.path.exists(fp):
                # Check printer status (internal only)
                printer_ready, printer_status, printer_jobs = get_printer_status_detailed(SELECTED_PRINTER)
                
                # AUTO-CLEAR: If printer has too many stuck jobs
                if printer_jobs > 5:
                    log(f"Printer has {printer_jobs} stuck jobs, auto-clearing...", "WARN")
                    auto_clear_printer_queue(SELECTED_PRINTER)
                    time.sleep(3)
                
                if not printer_ready:
                    log(f"Printer not ready: {printer_status}", "ERROR")
                    consecutive_printer_failures += 1
                    
                    # AUTO-CLEAR: After multiple failures
                    if consecutive_printer_failures >= MAX_CONSECUTIVE_FAILURES:
                        log(f"Multiple printer failures ({consecutive_printer_failures}), auto-clearing...", "WARN")
                        auto_clear_printer_queue(SELECTED_PRINTER)
                        time.sleep(5)
                        consecutive_printer_failures = 0
                    
                    # Requeue the file
                    time.sleep(5)
                    with _queued_lock:
                        _queued_files.discard(fp)
                        _queued_files.add(fp)
                        _print_queue.put(fp)
                    continue
                
                consecutive_printer_failures = 0
                
                # Original print
                print_pdf(fp)
                consecutive_errors = 0
                        
            else:
                log(f"File gone before printing: {os.path.basename(fp)}", "WARN")
                
        except Exception as e:
            log(f"Print worker error ({os.path.basename(fp)}): {e}", "ERROR")
            consecutive_errors += 1
            
            # AUTO-CLEAR: After too many errors
            if consecutive_errors > 3:
                log("Too many consecutive errors, auto-clearing printer...", "ERROR")
                auto_clear_printer_queue(SELECTED_PRINTER)
                time.sleep(10)
                consecutive_errors = 0
            
        finally:
            with _queued_lock:
                _queued_files.discard(fp)
            _print_queue.task_done()

# ==============================
# FILE WATCHER (ENHANCED - AUTO ONLY)
# ==============================

def get_marg_files():
    if not os.path.exists(WATCH_FOLDER):
        log(f"Watch folder missing: {WATCH_FOLDER}", "ERROR"); return []
    
    files = [
        os.path.join(WATCH_FOLDER, f)
        for f in os.listdir(WATCH_FOLDER)
        if f.startswith(FILE_PREFIX) and f.lower().endswith(".pdf")
    ]
    
    is_network = is_network_path(WATCH_FOLDER)
    
    processable_files = []
    for f in files:
        if is_file_locked(f):
            continue
        
        stability_wait = 2 if is_network else 1
        if not is_file_stable(f, stability_wait):
            continue
        
        processable_files.append(f)
    
    processable_files.sort(key=os.path.getctime)
    return processable_files

def watcher_loop():
    log("Auto Printer started — watching for PDFs…", "WATCH")
    log(f"Printer : {SELECTED_PRINTER}", "INFO")
    log(f"Folder  : {WATCH_FOLDER}",     "INFO")
    log(f"Prefix  : {FILE_PREFIX}",      "INFO")
    log(f"Version : v{APP_VERSION}",     "INFO")
    
    last_scan_time = 0
    min_scan_interval = 1
    
    while not stop_event.is_set():
        current_time = time.time()
        
        if current_time - last_scan_time >= min_scan_interval:
            files = get_marg_files()
            for fp in files[:10]:
                if stop_event.is_set():
                    break
                _enqueue_file(fp)
            last_scan_time = current_time
        
        queue_size = _print_queue.qsize()
        if queue_size > 20:
            sleep_time = max(0.5, CHECK_INTERVAL / 2)
        elif queue_size > 10:
            sleep_time = CHECK_INTERVAL
        else:
            sleep_time = CHECK_INTERVAL * 1.5
        
        stop_event.wait(sleep_time)

# ==============================
# LOG WINDOW (YOUR ORIGINAL - EXACTLY AS BEFORE)
# ==============================

_log_win_open = False
_log_win_ref  = None

def open_log_window():
    global _log_win_open, _log_win_ref
    if _log_win_open:
        try:
            if _log_win_ref and _log_win_ref.winfo_exists():
                _log_win_ref.lift()
                _log_win_ref.focus_force()
                return
        except Exception:
            pass
    _log_win_open = True

    root = get_root()
    win  = tk.Toplevel(root)
    _log_win_ref = win
    win.title("Marg ERP Auto Printer (Beta Version) — Live Logs")
    win.geometry("860x530")
    win.configure(bg=THEME["bg"])
    if os.path.exists(ICON_PATH):
        try: win.iconbitmap(ICON_PATH)
        except Exception: pass

    def _on_close():
        global _log_win_open, _log_win_ref
        _log_win_open = False
        _log_win_ref  = None
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", _on_close)

    tk.Frame(win, bg=THEME["accent2"], height=5).pack(fill="x")
    hdr = tk.Frame(win, bg=THEME["surface"], pady=8); hdr.pack(fill="x")
    tk.Label(hdr, text="🖨  Marg ERP Auto Printer  —  Live Logs",
             font=("Segoe UI", 11, "bold"), fg=THEME["text"],
             bg=THEME["surface"]).pack(side="left", padx=16)
    tk.Label(hdr, text=f"v{APP_VERSION}", font=("Segoe UI", 9),
             fg=THEME["text_dim"], bg=THEME["surface"]).pack(side="right", padx=6)
    tk.Label(hdr, text="● RUNNING", font=("Segoe UI", 9, "bold"),
             fg=THEME["success"], bg=THEME["surface"]).pack(side="right", padx=10)

    info = tk.Frame(win, bg=THEME["surface2"], pady=4); info.pack(fill="x")
    tk.Label(info,
             text=(f"  Printer: {SELECTED_PRINTER}   │   Folder: {WATCH_FOLDER}"
                   f"   │   Prefix: {FILE_PREFIX}   │   "
                   f"Interval: {CHECK_INTERVAL}s   │   "
                   f"Silent: {'ON' if USE_SILENT_MODE else 'OFF'}"),
             font=("Consolas", 8), fg=THEME["text_dim"],
             bg=THEME["surface2"], anchor="w").pack(fill="x", padx=12)

    tf = tk.Frame(win, bg=THEME["bg"]); tf.pack(fill="both", expand=True, padx=8, pady=8)
    sb = tk.Scrollbar(tf, bg=THEME["surface2"], troughcolor=THEME["bg"],
                      activebackground=THEME["accent"])
    sb.pack(side="right", fill="y")
    txt = tk.Text(tf, bg=THEME["bg"], fg=THEME["text"], font=FONT_MONO,
                  relief="flat", bd=0, state="disabled", wrap="word",
                  yscrollcommand=sb.set, selectbackground=THEME["accent2"],
                  pady=4, padx=8)
    txt.pack(fill="both", expand=True)
    sb.config(command=txt.yview)
    for lvl, col in LEVEL_COLORS.items():
        txt.tag_config(lvl, foreground=col)

    def _append(level, entry):
        txt.config(state="normal")
        txt.insert("end", entry + "\n", level if level in LEVEL_COLORS else "INFO")
        txt.config(state="disabled")
        txt.see("end")

    for lvl, entry in log_lines:
        _append(lvl, entry)

    bb = tk.Frame(win, bg=THEME["surface"], pady=8); bb.pack(fill="x")
    bkw = dict(bg=THEME["surface2"], fg=THEME["accent"], font=("Segoe UI", 9),
               relief="flat", bd=0, cursor="hand2", padx=12, pady=5,
               activebackground=THEME["accent"], activeforeground=THEME["bg"])

    def _clear():
        log_lines.clear()
        txt.config(state="normal"); txt.delete("1.0","end"); txt.config(state="disabled")

    def _copy():
        win.clipboard_clear(); win.clipboard_append(txt.get("1.0","end"))
        messagebox.showinfo("Copied", "Log copied to clipboard.", parent=win)

    def _check_update_manual():
        log("Manual update check requested…", "UPDATE")
        start_update_check(silent=False, parent_win=win)

    # NO PRINTER CLEAR BUTTON - REMOVED

    tk.Button(bb, text="🗑  Clear",         command=_clear,               **bkw).pack(side="left",  padx=(12,4))
    tk.Button(bb, text="📋  Copy",          command=_copy,                **bkw).pack(side="left",  padx=4)
    tk.Button(bb, text="🔄  Check Updates", command=_check_update_manual,
              bg=THEME["surface2"], fg=THEME["warning"],
              font=("Segoe UI", 9), relief="flat", bd=0, cursor="hand2",
              padx=12, pady=5,
              activebackground=THEME["warning"],
              activeforeground=THEME["bg"]).pack(side="left", padx=4)
    tk.Button(bb, text="✖  Close",          command=_on_close,            **bkw).pack(side="right", padx=12)
    tk.Label(bb, text="Developed by Mehak Singh | TheMehakCodes",
             font=FONT_CREDIT, fg=THEME["text_dim"], bg=THEME["surface"]).pack(side="right", padx=16)

    def _poll():
        if not _log_win_open: return
        try:
            while True:
                lvl, entry = log_queue.get_nowait()
                _append(lvl, entry)
        except queue.Empty:
            pass
        win.after(300, _poll)
    _poll()

# ==============================
# CONFIG EDIT WINDOW (YOUR ORIGINAL - UNCHANGED)
# ==============================

def open_config_window():
    root = get_root()
    win  = tk.Toplevel(root)
    win.title("Edit Configuration")
    win.geometry("520x560")
    win.configure(bg=THEME["bg"])
    if os.path.exists(ICON_PATH):
        try: win.iconbitmap(ICON_PATH)
        except Exception: pass

    tk.Frame(win, bg=THEME["accent2"], height=6).pack(fill="x")
    hdr = tk.Frame(win, bg=THEME["bg"], pady=14); hdr.pack(fill="x")
    tk.Label(hdr, text="⚙️  Edit Configuration",
             font=("Segoe UI", 14, "bold"), fg=THEME["text"], bg=THEME["bg"]).pack()
    tk.Frame(win, bg=THEME["border"], height=1).pack(fill="x", padx=30)

    form = tk.Frame(win, bg=THEME["bg"], padx=30, pady=10); form.pack(fill="both", expand=True)

    def lbl(text):
        tk.Label(form, text=text, font=FONT_LABEL, fg=THEME["text_dim"],
                 bg=THEME["bg"], anchor="w").pack(fill="x", pady=(10,2))

    lbl("🖨  Printer")
    printers    = [p[2] for p in win32print.EnumPrinters(2)]
    printer_var = tk.StringVar(value=SELECTED_PRINTER)
    _combo_style()
    ttk.Combobox(form, textvariable=printer_var, values=printers,
                 style="Dark.TCombobox", state="readonly").pack(fill="x")

    lbl("📁  Watch Folder")
    folder_var = tk.StringVar(value=WATCH_FOLDER)
    fr = tk.Frame(form, bg=THEME["bg"]); fr.pack(fill="x")
    tk.Entry(fr, textvariable=folder_var, **_entry_kw()).pack(
        side="left", fill="x", expand=True, ipady=6)
    tk.Button(fr, text=" Browse ",
              command=lambda: folder_var.set(filedialog.askdirectory() or folder_var.get()),
              bg=THEME["surface2"], fg=THEME["accent"], relief="flat",
              font=FONT_LABEL, cursor="hand2", bd=0,
              activebackground=THEME["accent"], activeforeground=THEME["bg"],
              padx=10, pady=6).pack(side="left", padx=(6,0))

    lbl("🏷  File Prefix")
    prefix_var = tk.StringVar(value=FILE_PREFIX)
    tk.Entry(form, textvariable=prefix_var, **_entry_kw()).pack(fill="x", ipady=6)

    lbl("⏱  Check Interval (sec)")
    interval_var = tk.StringVar(value=str(CHECK_INTERVAL))
    tk.Entry(form, textvariable=interval_var, **_entry_kw()).pack(fill="x", ipady=6)

    silent_var = tk.BooleanVar(value=USE_SILENT_MODE)
    sf = tk.Frame(form, bg=THEME["bg"]); sf.pack(fill="x", pady=(14,0))
    tk.Label(sf, text="🔇  Silent Mode", font=FONT_LABEL,
             fg=THEME["text"], bg=THEME["bg"]).pack(side="left")
    tk.Checkbutton(sf, variable=silent_var, bg=THEME["bg"],
                   selectcolor=THEME["surface2"], relief="flat", bd=0,
                   activebackground=THEME["bg"]).pack(side="right")

    def _save():
        try:
            cfg = {
                "printer":        printer_var.get(),
                "watch_folder":   folder_var.get(),
                "file_prefix":    prefix_var.get(),
                "check_interval": int(interval_var.get()),
                "silent_mode":    silent_var.get(),
            }
            with open(CONFIG_FILE, "w") as f:
                json.dump(cfg, f, indent=4)
            log("Config saved — restart app to apply changes.", "WARN")
            messagebox.showinfo("Saved",
                "✔ Config saved.\nRestart the app for changes to take effect.",
                parent=win)
            win.destroy()
        except Exception as e:
            messagebox.showerror("Error", str(e), parent=win)

    tk.Frame(win, bg=THEME["border"], height=1).pack(fill="x", padx=30)
    bf = tk.Frame(win, bg=THEME["bg"], pady=14); bf.pack()
    tk.Button(bf, text="  Save  ", command=_save,
              bg=THEME["accent"], fg=THEME["bg"], font=FONT_BTN,
              relief="flat", bd=0, cursor="hand2",
              padx=18, pady=8,
              activebackground=THEME["accent2"],
              activeforeground=THEME["text"]).pack()
    tk.Label(win, text="Developed by Mehak Singh | TheMehakCodes",
             font=FONT_CREDIT, fg=THEME["text_dim"], bg=THEME["bg"]).pack(pady=(0,10))

# ==============================
# ABOUT / VERSION WINDOW (YOUR ORIGINAL - UNCHANGED)
# ==============================

def open_about_window():
    root = get_root()
    win  = tk.Toplevel(root)
    win.title("About — Marg ERP Auto Printer")
    win.geometry("420x320")
    win.resizable(False, False)
    win.configure(bg=THEME["bg"])
    if os.path.exists(ICON_PATH):
        try: win.iconbitmap(ICON_PATH)
        except Exception: pass

    tk.Frame(win, bg=THEME["accent2"], height=6).pack(fill="x")

    body = tk.Frame(win, bg=THEME["bg"], pady=30, padx=40)
    body.pack(fill="both", expand=True)

    tk.Label(body, text="🖨  Marg ERP Auto Printer",
             font=("Segoe UI", 14, "bold"), fg=THEME["text"], bg=THEME["bg"]).pack()
    tk.Label(body, text=f"Version  v{APP_VERSION}",
             font=("Segoe UI", 10), fg=THEME["accent"], bg=THEME["bg"]).pack(pady=(6,0))
    tk.Frame(body, bg=THEME["border"], height=1).pack(fill="x", pady=14)
    tk.Label(body, text="Developed by  Mehak Singh",
             font=FONT_LABEL, fg=THEME["text_dim"], bg=THEME["bg"]).pack()
    tk.Label(body, text="TheMehakCodes",
             font=("Segoe UI", 10, "bold"), fg=THEME["accent2"], bg=THEME["bg"]).pack(pady=(2,0))
    tk.Label(body, text="https://themehakcodes.com",
             font=FONT_LABEL_SM, fg=THEME["text_dim"], bg=THEME["bg"]).pack(pady=(2,14))

    bkw = dict(bg=THEME["surface2"], fg=THEME["warning"], font=("Segoe UI", 9),
               relief="flat", bd=0, cursor="hand2", padx=16, pady=7,
               activebackground=THEME["warning"], activeforeground=THEME["bg"])

    def _check():
        log("Manual update check from About window…", "UPDATE")
        start_update_check(silent=False, parent_win=win)

    tk.Button(body, text="🔄  Check for Updates", command=_check, **bkw).pack()

    tk.Frame(win, bg=THEME["border"], height=1).pack(fill="x", padx=30)
    tk.Label(win, text="Developed by Mehak Singh | TheMehakCodes",
             font=FONT_CREDIT, fg=THEME["text_dim"], bg=THEME["bg"]).pack(pady=8)

# ==============================
# SYSTEM TRAY ICON (YOUR ORIGINAL - MINIMAL ENHANCEMENT)
# ==============================

def _make_tray_image() -> Image.Image:
    if os.path.exists(ICON_PATH):
        try:
            return Image.open(ICON_PATH).convert("RGBA").resize((64, 64))
        except Exception:
            pass
    img  = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([2, 2, 62, 62], fill="#4F8EF7")
    draw.text((20, 18), "M", fill="white")
    return img

def _tray_show_logs(icon, item):
    get_root().after(0, open_log_window)

def _tray_edit_config(icon, item):
    get_root().after(0, open_config_window)

def _tray_about(icon, item):
    get_root().after(0, open_about_window)

def _tray_check_update(icon, item):
    def _run():
        log("Manual update check from system tray…", "UPDATE")
        start_update_check(silent=False, parent_win=None)
    get_root().after(0, _run)

# NO PRINTER CLEAR MENU ITEM - REMOVED

def _tray_exit(icon, item):
    log("Shutting down…", "WARN")
    stop_event.set()
    icon.stop()
    get_root().after(0, get_root().destroy)

def build_tray() -> pystray.Icon:
    menu = pystray.Menu(
        pystray.MenuItem(
            f"🖨  Marg ERP Auto Printer  v{APP_VERSION}",
            None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("📋  Show Logs",       _tray_show_logs,    default=True),
        pystray.MenuItem("⚙️   Edit Config",     _tray_edit_config),
        pystray.MenuItem("ℹ️   About",           _tray_about),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("🔄  Check for Updates", _tray_check_update),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("✖  Exit",             _tray_exit),
    )
    return pystray.Icon(
        name  = "MargERPAutoPrinter",
        icon  = _make_tray_image(),
        title = f"Marg ERP Auto Printer  v{APP_VERSION}",
        menu  = menu,
    )

# ==============================
# MAIN (YOUR ORIGINAL - UNCHANGED)
# ==============================

def main():
    log(f"Marg ERP Auto Printer v{APP_VERSION} starting...", "INFO")
    
    start_update_check(silent=True)
    threading.Thread(target=print_worker, daemon=True, name="PrintWorker").start()
    threading.Thread(target=watcher_loop, daemon=True, name="Watcher").start()
    tray = build_tray()
    threading.Thread(target=tray.run, daemon=True, name="Tray").start()
    
    log("Application started successfully", "SUCCESS")
    get_root().mainloop()

if __name__ == "__main__":
    main()