# ============================================================
#  Marg ERP Auto Printer
#  Developed by Mehak Singh | TheMehakCodes
# ============================================================
#
#  IMPORTANT — HOW THE WINDOW IS HIDDEN:
#  1. Built with PyInstaller --windowed  → no console allocated at all
#  2. As a safety net, the very first thing we do (before ANY import
#     that might flash a window) is call the Win32 API to hide the
#     console window if one somehow exists.
# ============================================================

import ctypes, sys

# ── Hide console window immediately (belt-and-suspenders) ──────────
def _hide_console():
    """
    Works whether the process was launched from Explorer, Task Scheduler,
    startup folder, or accidentally double-clicked in a terminal.
    """
    try:
        hwnd = ctypes.windll.kernel32.GetConsoleWindow()
        if hwnd:
            ctypes.windll.user32.ShowWindow(hwnd, 0)   # SW_HIDE = 0
            ctypes.windll.kernel32.FreeConsole()        # detach completely
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
}

# ==============================
# LOGGER  (thread-safe, no print)
# ==============================

def log(msg, level="INFO"):
    now   = datetime.now().strftime("%H:%M:%S")
    icons = {
        "INFO":    "ℹ",  "SUCCESS": "✔",  "WARN":  "⚠",
        "ERROR":   "✖",  "PRINT":   "🖨",  "WATCH": "👁",
    }
    entry = f"[{now}]  {level:<7}  {icons.get(level, '•')}  {msg}"
    log_lines.append((level, entry))
    if len(log_lines) > MAX_LOGS:
        log_lines.pop(0)
    log_queue.put((level, entry))
    # intentionally NO print() — we are windowless

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
        _tk_root.withdraw()                    # invisible, never shown
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

    # header accent bar
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

    # Printer
    lbl(form, "🖨  Printer")
    printers    = [p[2] for p in win32print.EnumPrinters(2)]
    printer_var = tk.StringVar(value=printers[0] if printers else "")
    _combo_style()
    ttk.Combobox(form, textvariable=printer_var, values=printers,
                 style="Dark.TCombobox", state="readonly").pack(fill="x")

    # Watch folder
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

    # Prefix + interval
    row = tk.Frame(form, bg=THEME["bg"]); row.pack(fill="x")
    lc  = tk.Frame(row,  bg=THEME["bg"]); lc.pack(side="left", fill="x", expand=True, padx=(0,10))
    rc  = tk.Frame(row,  bg=THEME["bg"]); rc.pack(side="left", fill="x", expand=True)
    lbl(lc, "🏷  File Prefix")
    prefix_var = tk.StringVar(value=DEFAULT_CONFIG["file_prefix"])
    tk.Entry(lc, textvariable=prefix_var, **_entry_kw()).pack(fill="x", ipady=6)
    lbl(rc, "⏱  Interval (sec)")
    interval_var = tk.StringVar(value=str(DEFAULT_CONFIG["check_interval"]))
    tk.Entry(rc, textvariable=interval_var, **_entry_kw()).pack(fill="x", ipady=6)

    # Silent mode
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

    win.wait_window()   # block until closed
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

# ── load (may open setup window) ───────────────────────────────────
_root_ready = get_root()          # create hidden root FIRST
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
        
    try:
        if not os.path.exists(SUMATRA_PATH):
            log("SumatraPDF not found — legacy fallback", "WARN")
            print_pdf_legacy(fp); return
        subprocess.Popen(
            [SUMATRA_PATH, "-print-to", SELECTED_PRINTER,
             "-silent", "-exit-on-print", fp],
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        time.sleep(3)
        if os.path.exists(fp): os.remove(fp)
    except Exception as e:
        log(f"Silent print error: {e}", "ERROR")
        print_pdf_legacy(fp)

def print_pdf(fp: str):
    name = os.path.basename(fp)
    log(f"Sending → {name}", "PRINT")
    print_pdf_silent(fp) if USE_SILENT_MODE else print_pdf_legacy(fp)
    log(f"Done — {name}", "SUCCESS")

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
    while not stop_event.is_set():
        for fp in get_marg_files():
            if stop_event.is_set(): break
            time.sleep(2)
            print_pdf(fp)
        stop_event.wait(CHECK_INTERVAL)

# ==============================
# LOG WINDOW
# ==============================

_log_win_open = False

def open_log_window():
    """Called on main Tk thread via root.after()"""
    global _log_win_open
    if _log_win_open:
        return
    _log_win_open = True

    root = get_root()
    win  = tk.Toplevel(root)
    win.title("Marg ERP Auto Printer — Live Logs")
    win.geometry("820x500")
    win.configure(bg=THEME["bg"])
    if os.path.exists(ICON_PATH):
        try: win.iconbitmap(ICON_PATH)
        except Exception: pass

    def _on_close():
        global _log_win_open
        _log_win_open = False
        win.destroy()
    win.protocol("WM_DELETE_WINDOW", _on_close)

    # header
    tk.Frame(win, bg=THEME["accent2"], height=5).pack(fill="x")
    hdr = tk.Frame(win, bg=THEME["surface"], pady=8); hdr.pack(fill="x")
    tk.Label(hdr, text="🖨  Marg ERP Auto Printer  —  Live Logs",
             font=("Segoe UI", 11, "bold"), fg=THEME["text"],
             bg=THEME["surface"]).pack(side="left", padx=16)
    tk.Label(hdr, text="● RUNNING", font=("Segoe UI", 9, "bold"),
             fg=THEME["success"], bg=THEME["surface"]).pack(side="right", padx=16)

    # info strip
    info = tk.Frame(win, bg=THEME["surface2"], pady=4); info.pack(fill="x")
    tk.Label(info,
             text=(f"  Printer: {SELECTED_PRINTER}   │   Folder: {WATCH_FOLDER}"
                   f"   │   Prefix: {FILE_PREFIX}   │   "
                   f"Interval: {CHECK_INTERVAL}s   │   "
                   f"Silent: {'ON' if USE_SILENT_MODE else 'OFF'}"),
             font=("Consolas", 8), fg=THEME["text_dim"],
             bg=THEME["surface2"], anchor="w").pack(fill="x", padx=12)

    # text widget
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

    # replay history
    for lvl, entry in log_lines:
        _append(lvl, entry)

    # toolbar
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

    tk.Button(bb, text="🗑  Clear",  command=_clear,    **bkw).pack(side="left",  padx=(12,4))
    tk.Button(bb, text="📋  Copy",   command=_copy,     **bkw).pack(side="left",  padx=4)
    tk.Button(bb, text="✖  Close",  command=_on_close, **bkw).pack(side="right", padx=12)
    tk.Label(bb, text="Developed by Mehak Singh | TheMehakCodes",
             font=FONT_CREDIT, fg=THEME["text_dim"], bg=THEME["surface"]).pack(side="right", padx=16)

    # poll queue for new entries
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
# SYSTEM TRAY ICON
# ==============================

def _make_tray_image() -> Image.Image:
    if os.path.exists(ICON_PATH):
        try:
            return Image.open(ICON_PATH).convert("RGBA").resize((64, 64))
        except Exception:
            pass
    # fallback — draw a coloured circle with "M"
    img  = Image.new("RGBA", (64, 64), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)
    draw.ellipse([2, 2, 62, 62], fill="#4F8EF7")
    draw.text((20, 18), "M", fill="white")
    return img

# Tray callbacks must be plain functions (pystray calls them from its thread).
# We schedule Tk work onto the main thread via root.after().

def _tray_show_logs(icon, item):
    get_root().after(0, open_log_window)

def _tray_edit_config(icon, item):
    get_root().after(0, open_config_window)

def _tray_exit(icon, item):
    log("Shutting down…", "WARN")
    stop_event.set()
    icon.stop()
    get_root().after(0, get_root().destroy)

def build_tray() -> pystray.Icon:
    menu = pystray.Menu(
        pystray.MenuItem(
            "🖨  Marg ERP Auto Printer",
            None,
            enabled=False,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            "📋  Show Logs",
            _tray_show_logs,
            default=True,       # ← double-click opens logs
        ),
        pystray.MenuItem(
            "⚙️   Edit Config",
            _tray_edit_config,
        ),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem(
            "✖  Exit",
            _tray_exit,
        ),
    )
    return pystray.Icon(
        name  = "MargERPAutoPrinter",
        icon  = _make_tray_image(),
        title = "Marg ERP Auto Printer",
        menu  = menu,
    )

# ==============================
# MAIN
# ==============================

def main():
    # 1 — start file watcher (daemon thread)
    threading.Thread(target=watcher_loop, daemon=True, name="Watcher").start()

    # 2 — start tray icon (daemon thread — pystray has its own loop)
    tray = build_tray()
    threading.Thread(target=tray.run, daemon=True, name="Tray").start()

    # 3 — Tk event loop on main thread (required by Windows)
    #     This keeps the process alive and handles all GUI callbacks.
    get_root().mainloop()

if __name__ == "__main__":
    main()