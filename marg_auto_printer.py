# ============================================================
#  Marg ERP Auto Printer
#  Developed by Mehak Singh | TheMehakCodes
#  Version: 1.0.0
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

import ctypes, sys

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

# ── Now safe to import everything else ─────────────────────────────
import os
import time
import json
import subprocess
import threading
import queue
import win32print
import win32api
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

import pystray
from PIL import Image, ImageDraw

# ==============================
# VERSION
# ==============================

APP_VERSION        = "1.0.3"
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

# ==============================
# GLOBAL LOG QUEUE  (thread-safe)
# ==============================

log_queue  = queue.Queue()
log_lines  = []              # (level, text)  — full in-memory history
MAX_LOGS   = 500

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
# AUTO-UPDATER
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
    ─────────────────────────────────────────────────
    1. Download new .exe → update_new.exe (same folder)
    2. Validate size + optional SHA-256
    3. Write _marg_updater.bat that:
         • waits for THIS process to exit (PID-poll, no fixed delay)
         • moves update_new.exe over current exe
         • restarts the new exe silently
         • self-deletes
    4. Launch bat detached → clean sys.exit() so bat's wait fires
    """
    import urllib.request, hashlib

    current_exe = sys.executable if getattr(sys, "frozen", False) \
                  else os.path.abspath(__file__)
    app_dir     = os.path.dirname(current_exe)
    update_path = os.path.join(app_dir, "update_new.exe")

    log("Downloading direct EXE update silently…", "UPDATE")

    # ── Download ──────────────────────────────────────────────────
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

    # ── Size guard ────────────────────────────────────────────────
    if os.path.getsize(update_path) < 100_000:
        log("Downloaded file too small — aborting update.", "ERROR")
        try: os.remove(update_path)
        except Exception: pass
        return

    # ── Optional SHA-256 validation ───────────────────────────────
    if remote_sha:
        actual_sha = hasher.hexdigest().lower()
        if actual_sha != remote_sha.lower():
            log(f"Checksum mismatch — aborting. Got: {actual_sha}", "ERROR")
            try: os.remove(update_path)
            except Exception: pass
            return
        log("SHA-256 checksum verified ✔", "SUCCESS")

    log(f"Download complete. Installing v{remote_ver} silently…", "UPDATE")

    # ── Write updater batch ───────────────────────────────────────
    bat_path = os.path.join(app_dir, "_marg_updater.bat")
    pid      = os.getpid()
    bat_contents = (
        "@echo off\n"
        "setlocal\n"
        # Wait until our PID disappears (poll every 500 ms, max ~30 s)
        f":wait\n"
        f'tasklist /FI "PID eq {pid}" 2>nul | find /I "{pid}" >nul\n'
        "if not errorlevel 1 (\n"
        "    ping 127.0.0.1 -n 1 -w 500 >nul\n"
        "    goto wait\n"
        ")\n"
        # Replace exe
        f'move /Y "{update_path}" "{current_exe}" >nul 2>&1\n'
        # Restart silently
        f'start "" "{current_exe}"\n'
        # Self-delete
        "(goto) 2>nul & del \"%~f0\"\n"
    )
    with open(bat_path, "w") as f:
        f.write(bat_contents)

    # ── Launch bat fully detached & hidden ────────────────────────
    subprocess.Popen(
        ["cmd.exe", "/C", bat_path],
        creationflags=(subprocess.DETACHED_PROCESS |
                       subprocess.CREATE_NO_WINDOW),
    )

    log("Updater launched — restarting now…", "UPDATE")

    # ── Clean shutdown so the bat's PID-wait loop actually fires ──
    stop_event.set()
    get_root().after(0, get_root().destroy)


def _apply_installer_update(exe_url: str, remote_ver: str,
                             remote_sha: str, silent: bool, parent_win):
    """
    INSTALLER EXE update flow  (release_type == "installer")
    ─────────────────────────────────────────────────────────
    1. Download Setup exe → update_setup.exe (same folder)
    2. Validate size + optional SHA-256
    3. Prompt the user with a messagebox (even in silent startup mode,
       because an installer always needs the user to run it)
    4. If user confirms → launch the installer normally (visible UAC prompt)
       Current app keeps running until user closes it / installer replaces it.
    """
    import urllib.request, hashlib

    current_exe = sys.executable if getattr(sys, "frozen", False) \
                  else os.path.abspath(__file__)
    app_dir      = os.path.dirname(current_exe)
    setup_path   = os.path.join(app_dir, "update_setup.exe")

    log("Downloading installer update…", "UPDATE")

    # ── Download ──────────────────────────────────────────────────
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

    # ── Size guard ────────────────────────────────────────────────
    if os.path.getsize(setup_path) < 100_000:
        log("Downloaded installer too small — aborting.", "ERROR")
        try: os.remove(setup_path)
        except Exception: pass
        return

    # ── Optional SHA-256 validation ───────────────────────────────
    if remote_sha:
        actual_sha = hasher.hexdigest().lower()
        if actual_sha != remote_sha.lower():
            log(f"Checksum mismatch — aborting. Got: {actual_sha}", "ERROR")
            try: os.remove(setup_path)
            except Exception: pass
            return
        log("SHA-256 checksum verified ✔", "SUCCESS")

    log(f"Installer downloaded for v{remote_ver}. Prompting user…", "UPDATE")

    # ── Prompt user (always shown — installer needs human interaction) ─
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

    version.json schema expected:
    {
        "version":      "1.2.0",
        "release_type": "direct",       ← "direct" | "installer"
        "exe_url":      "https://...",
        "sha256":       "abc123..."     ← optional
    }

    release_type == "direct"
        → silent download + bat-swap + auto-restart (no user prompt needed)

    release_type == "installer"
        → download Setup.exe → always prompt user → user runs installer
    """
    try:
        import urllib.request

        log("Checking for updates…", "UPDATE")

        # ── Fetch version manifest ────────────────────────────────────
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
        # Normalise: any value other than "installer" is treated as "direct"
        if release_type not in ("direct", "installer"):
            release_type = "direct"

        # ── Already up to date? ───────────────────────────────────────
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

        # ── Route to the correct update handler ───────────────────────
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
# FIRST-TIME SETUP WINDOW
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
# LOAD CONFIG
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
# PRINT FUNCTIONS
# ==============================

def print_pdf_legacy(fp: str):
    try:
        win32print.SetDefaultPrinter(SELECTED_PRINTER)
        win32api.ShellExecute(0, "print", fp, None, ".", 0)
        time.sleep(5)
        if os.path.exists(fp): os.remove(fp)
    except Exception as e:
        log(f"Legacy print error: {e}", "ERROR")

def print_pdf_silent(fp: str):
    try:
        if not os.path.exists(SUMATRA_PATH):
            log("SumatraPDF not found — legacy fallback", "WARN")
            print_pdf_legacy(fp)
            return
        subprocess.Popen(
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
        time.sleep(3)
        if os.path.exists(fp):
            os.remove(fp)
    except Exception as e:
        log(f"Silent print error: {e}", "ERROR")
        print_pdf_legacy(fp)

def print_pdf(fp: str):
    name = os.path.basename(fp)
    log(f"Sending → {name}", "PRINT")
    print_pdf_silent(fp) if USE_SILENT_MODE else print_pdf_legacy(fp)
    log(f"Done — {name}", "SUCCESS")

# ==============================
# PRINT QUEUE  (guaranteed ordered, no missed files)
# ==============================

_print_queue   = queue.Queue()
_queued_files  = set()
_queued_lock   = threading.Lock()


def _enqueue_file(fp: str):
    """Thread-safe: add a file to the print queue only once."""
    with _queued_lock:
        if fp not in _queued_files:
            _queued_files.add(fp)
            _print_queue.put(fp)
            log(f"Queued  → {os.path.basename(fp)}  "
                f"(queue depth: {_print_queue.qsize()})", "WATCH")


def print_worker():
    log("Print worker started — ready for jobs.", "INFO")
    while not stop_event.is_set():
        try:
            fp = _print_queue.get(timeout=1)
        except queue.Empty:
            continue
        try:
            if os.path.exists(fp):
                print_pdf(fp)
            else:
                log(f"File gone before printing: {os.path.basename(fp)}", "WARN")
        except Exception as e:
            log(f"Print worker error ({os.path.basename(fp)}): {e}", "ERROR")
        finally:
            with _queued_lock:
                _queued_files.discard(fp)
            _print_queue.task_done()

# ==============================
# FILE WATCHER
# ==============================

def get_marg_files():
    if not os.path.exists(WATCH_FOLDER):
        log(f"Watch folder missing: {WATCH_FOLDER}", "ERROR"); return []
    files = [
        os.path.join(WATCH_FOLDER, f)
        for f in os.listdir(WATCH_FOLDER)
        if f.startswith(FILE_PREFIX) and f.lower().endswith(".pdf")
    ]
    files.sort(key=os.path.getctime)
    return files

stop_event = threading.Event()

def watcher_loop():
    log("Auto Printer started — watching for PDFs…", "WATCH")
    log(f"Printer : {SELECTED_PRINTER}", "INFO")
    log(f"Folder  : {WATCH_FOLDER}",     "INFO")
    log(f"Prefix  : {FILE_PREFIX}",      "INFO")
    log(f"Version : v{APP_VERSION}",     "INFO")
    while not stop_event.is_set():
        for fp in get_marg_files():
            if stop_event.is_set():
                break
            _enqueue_file(fp)
        stop_event.wait(CHECK_INTERVAL)

# ==============================
# LOG WINDOW
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
    win.title("Marg ERP Auto Printer — Live Logs")
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
# CONFIG EDIT WINDOW
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
# ABOUT / VERSION WINDOW
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
# SYSTEM TRAY ICON
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
# MAIN
# ==============================

def main():
    start_update_check(silent=True)
    threading.Thread(target=print_worker, daemon=True, name="PrintWorker").start()
    threading.Thread(target=watcher_loop, daemon=True, name="Watcher").start()
    tray = build_tray()
    threading.Thread(target=tray.run, daemon=True, name="Tray").start()
    get_root().mainloop()

if __name__ == "__main__":
    main()