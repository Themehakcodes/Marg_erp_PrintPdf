import os
import time
import json
import sys
import subprocess
import win32print
import win32api
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime

# ==============================
# BASE PATH (WORKS WITH EXE)
# ==============================

if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

CONFIG_FILE = os.path.join(BASE_DIR, "config.json")
SUMATRA_PATH = os.path.join(BASE_DIR, "SumatraPDF.exe")

# ==============================
# TERMINAL COLORS (ANSI)
# ==============================

class Color:
    RESET   = "\033[0m"
    BOLD    = "\033[1m"
    DIM     = "\033[2m"

    BLACK   = "\033[30m"
    RED     = "\033[91m"
    GREEN   = "\033[92m"
    YELLOW  = "\033[93m"
    BLUE    = "\033[94m"
    MAGENTA = "\033[95m"
    CYAN    = "\033[96m"
    WHITE   = "\033[97m"

    BG_BLACK  = "\033[40m"
    BG_BLUE   = "\033[44m"
    BG_GREEN  = "\033[42m"
    BG_RED    = "\033[41m"

def cprint(text, color=Color.WHITE, bold=False, end="\n"):
    prefix = Color.BOLD if bold else ""
    print(f"{prefix}{color}{text}{Color.RESET}", end=end)

def log(msg, level="INFO"):
    now = datetime.now().strftime("%H:%M:%S")
    icons = {
        "INFO":    (Color.CYAN,    "ℹ"),
        "SUCCESS": (Color.GREEN,   "✔"),
        "WARN":    (Color.YELLOW,  "⚠"),
        "ERROR":   (Color.RED,     "✖"),
        "PRINT":   (Color.MAGENTA, "🖨"),
        "WATCH":   (Color.BLUE,    "👁"),
    }
    color, icon = icons.get(level, (Color.WHITE, "•"))
    time_str = f"{Color.DIM}{Color.WHITE}[{now}]{Color.RESET}"
    level_str = f"{Color.BOLD}{color}[{level}]{Color.RESET}"
    msg_str = f"{color}{msg}{Color.RESET}"
    print(f"  {time_str} {level_str} {icon}  {msg_str}")

# ==============================
# DEFAULT CONFIG
# ==============================

DEFAULT_CONFIG = {
    "printer": "",
    "watch_folder": os.path.join(os.path.expanduser("~"), "Downloads"),
    "file_prefix": "Marg_erp",
    "check_interval": 3,
    "silent_mode": True
}

# ==============================
# THEME & STYLES
# ==============================

THEME = {
    "bg":           "#0F1117",
    "surface":      "#1A1D27",
    "surface2":     "#22263A",
    "accent":       "#4F8EF7",
    "accent2":      "#7C5CF6",
    "success":      "#2ECC71",
    "warning":      "#F39C12",
    "danger":       "#E74C3C",
    "text":         "#E8ECF4",
    "text_dim":     "#7A8099",
    "border":       "#2E3450",
    "input_bg":     "#13161F",
}

FONT_HEADING  = ("Segoe UI", 15, "bold")
FONT_LABEL    = ("Segoe UI", 10)
FONT_LABEL_SM = ("Segoe UI", 9)
FONT_MONO     = ("Consolas", 9)
FONT_BTN      = ("Segoe UI", 10, "bold")
FONT_CREDIT   = ("Segoe UI", 8, "italic")

# ==============================
# FIRST TIME GUI SETUP
# ==============================

def first_time_setup():
    config_data = {}

    root = tk.Tk()
    root.title("Marg ERP Auto Printer — Setup")
    root.geometry("520x560")
    root.resizable(False, False)
    root.configure(bg=THEME["bg"])

    # ── header ──────────────────────────────────────────────
    header = tk.Frame(root, bg=THEME["accent2"], height=6)
    header.pack(fill="x")

    title_frame = tk.Frame(root, bg=THEME["bg"], pady=20)
    title_frame.pack(fill="x")

    tk.Label(title_frame, text="🖨  Marg ERP Auto Printer",
             font=("Segoe UI", 16, "bold"),
             fg=THEME["text"], bg=THEME["bg"]).pack()

    tk.Label(title_frame, text="Initial Configuration",
             font=FONT_LABEL_SM, fg=THEME["text_dim"], bg=THEME["bg"]).pack()

    # ── divider ──────────────────────────────────────────────
    tk.Frame(root, bg=THEME["border"], height=1).pack(fill="x", padx=30)

    # ── form ─────────────────────────────────────────────────
    form = tk.Frame(root, bg=THEME["bg"], padx=30, pady=10)
    form.pack(fill="both", expand=True)

    def field_label(parent, text):
        tk.Label(parent, text=text, font=FONT_LABEL,
                 fg=THEME["text_dim"], bg=THEME["bg"],
                 anchor="w").pack(fill="x", pady=(12, 2))

    entry_style = dict(
        bg=THEME["input_bg"], fg=THEME["text"],
        insertbackground=THEME["accent"],
        relief="flat", bd=0,
        font=FONT_LABEL,
        highlightthickness=1,
        highlightbackground=THEME["border"],
        highlightcolor=THEME["accent"]
    )

    # Printer
    field_label(form, "🖨  Printer")
    printers = [p[2] for p in win32print.EnumPrinters(2)]
    printer_var = tk.StringVar(value=printers[0] if printers else "")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Dark.TCombobox",
        fieldbackground=THEME["input_bg"],
        background=THEME["input_bg"],
        foreground=THEME["text"],
        arrowcolor=THEME["accent"],
        bordercolor=THEME["border"],
        lightcolor=THEME["border"],
        darkcolor=THEME["border"],
        selectbackground=THEME["accent2"],
        selectforeground=THEME["text"],
        padding=6
    )
    ttk.Combobox(form, textvariable=printer_var, values=printers,
                 style="Dark.TCombobox", state="readonly").pack(fill="x")

    # Watch Folder
    field_label(form, "📁  Watch Folder")
    folder_var = tk.StringVar(value=DEFAULT_CONFIG["watch_folder"])
    folder_row = tk.Frame(form, bg=THEME["bg"])
    folder_row.pack(fill="x")

    folder_entry = tk.Entry(folder_row, textvariable=folder_var, **entry_style)
    folder_entry.pack(side="left", fill="x", expand=True, ipady=6)

    def browse_folder():
        path = filedialog.askdirectory()
        if path:
            folder_var.set(path)

    browse_btn = tk.Button(
        folder_row, text=" Browse ",
        command=browse_folder,
        bg=THEME["surface2"], fg=THEME["accent"],
        relief="flat", font=FONT_LABEL,
        cursor="hand2", bd=0,
        activebackground=THEME["accent"], activeforeground=THEME["bg"],
        padx=10, pady=6
    )
    browse_btn.pack(side="left", padx=(6, 0))

    # Row: prefix + interval
    row2 = tk.Frame(form, bg=THEME["bg"])
    row2.pack(fill="x")

    left_col = tk.Frame(row2, bg=THEME["bg"])
    left_col.pack(side="left", fill="x", expand=True, padx=(0, 10))

    right_col = tk.Frame(row2, bg=THEME["bg"])
    right_col.pack(side="left", fill="x", expand=True)

    field_label(left_col, "🏷  File Prefix")
    prefix_var = tk.StringVar(value=DEFAULT_CONFIG["file_prefix"])
    tk.Entry(left_col, textvariable=prefix_var, **entry_style).pack(fill="x", ipady=6)

    field_label(right_col, "⏱  Check Interval (sec)")
    interval_var = tk.StringVar(value=str(DEFAULT_CONFIG["check_interval"]))
    tk.Entry(right_col, textvariable=interval_var, **entry_style).pack(fill="x", ipady=6)

    # Silent mode toggle
    silent_var = tk.BooleanVar(value=True)
    toggle_frame = tk.Frame(form, bg=THEME["bg"])
    toggle_frame.pack(fill="x", pady=(16, 0))

    def toggle_silent():
        if silent_var.get():
            toggle_indicator.config(bg=THEME["success"], text="ON ")
        else:
            toggle_indicator.config(bg=THEME["text_dim"], text="OFF")

    tk.Label(toggle_frame, text="🔇  Silent Print Mode",
             font=FONT_LABEL, fg=THEME["text"], bg=THEME["bg"]).pack(side="left")

    toggle_indicator = tk.Label(
        toggle_frame, text="ON ",
        bg=THEME["success"], fg=THEME["bg"],
        font=("Segoe UI", 8, "bold"), padx=6, pady=2
    )
    toggle_indicator.pack(side="right")

    tk.Checkbutton(
        toggle_frame, variable=silent_var,
        command=toggle_silent,
        bg=THEME["bg"], fg=THEME["text"],
        activebackground=THEME["bg"],
        selectcolor=THEME["surface2"],
        relief="flat", bd=0
    ).pack(side="right", padx=4)

    # ── save button ───────────────────────────────────────────
    tk.Frame(root, bg=THEME["border"], height=1).pack(fill="x", padx=30)

    btn_frame = tk.Frame(root, bg=THEME["bg"], pady=16)
    btn_frame.pack()

    def save_config():
        try:
            selected_printer = printer_var.get()
            watch_folder     = folder_var.get()
            file_prefix      = prefix_var.get()
            silent_mode      = silent_var.get()
            check_interval   = int(interval_var.get())

            if not selected_printer:
                messagebox.showerror("Error", "Please select a printer.", parent=root)
                return

            config_data.update({
                "printer":        selected_printer,
                "watch_folder":   watch_folder,
                "file_prefix":    file_prefix,
                "check_interval": check_interval,
                "silent_mode":    silent_mode
            })

            with open(CONFIG_FILE, "w") as f:
                json.dump(config_data, f, indent=4)

            messagebox.showinfo("Saved", "✔  Configuration saved successfully!", parent=root)
            root.destroy()

        except Exception as e:
            messagebox.showerror("Error", str(e), parent=root)

    save_btn = tk.Button(
        btn_frame, text="  Save Configuration  ",
        command=save_config,
        bg=THEME["accent"], fg=THEME["bg"],
        font=FONT_BTN, relief="flat", bd=0,
        cursor="hand2", padx=20, pady=10,
        activebackground=THEME["accent2"],
        activeforeground=THEME["text"]
    )
    save_btn.pack()

    # ── footer credit ─────────────────────────────────────────
    tk.Label(root,
             text="Developed by Mehak Singh",
             font=FONT_CREDIT,
             fg=THEME["text_dim"], bg=THEME["bg"]).pack(pady=(0, 10))

    root.mainloop()
    return config_data


# ==============================
# LOAD CONFIG
# ==============================

def load_config():
    if "--config" in sys.argv:
        return first_time_setup()

    if not os.path.exists(CONFIG_FILE):
        return first_time_setup()

    try:
        with open(CONFIG_FILE, "r") as f:
            return json.load(f)
    except:
        return first_time_setup()


CONFIG = load_config()

SELECTED_PRINTER = CONFIG["printer"]
WATCH_FOLDER     = CONFIG["watch_folder"]
FILE_PREFIX      = CONFIG["file_prefix"]
CHECK_INTERVAL   = CONFIG["check_interval"]
USE_SILENT_MODE  = CONFIG["silent_mode"]

# ==============================
# PRINT FUNCTIONS
# ==============================

def print_pdf_legacy(file_path):
    try:
        win32print.SetDefaultPrinter(SELECTED_PRINTER)
        win32api.ShellExecute(0, "print", file_path, None, ".", 0)
        time.sleep(5)
        if os.path.exists(file_path):
            os.remove(file_path)
    except Exception as e:
        log(f"Legacy print error: {e}", "ERROR")


def print_pdf_silent(file_path):
    try:
        if not os.path.exists(SUMATRA_PATH):
            log("SumatraPDF not found — falling back to legacy mode.", "WARN")
            print_pdf_legacy(file_path)
            return

        subprocess.Popen(
            [SUMATRA_PATH, "-print-to", SELECTED_PRINTER, "-silent", "-exit-on-print", file_path],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            creationflags=subprocess.CREATE_NO_WINDOW
        )

        time.sleep(3)

        if os.path.exists(file_path):
            os.remove(file_path)

    except Exception as e:
        log(f"Silent print error: {e}", "ERROR")
        print_pdf_legacy(file_path)


def print_pdf(file_path):
    filename = os.path.basename(file_path)
    log(f"Sending to printer → {filename}", "PRINT")

    if USE_SILENT_MODE:
        print_pdf_silent(file_path)
    else:
        print_pdf_legacy(file_path)

    log(f"Done — {filename}", "SUCCESS")


# ==============================
# FILE DETECTION
# ==============================

def get_marg_files():
    if not os.path.exists(WATCH_FOLDER):
        log(f"Watch folder not found: {WATCH_FOLDER}", "ERROR")
        return []

    files = [
        os.path.join(WATCH_FOLDER, f)
        for f in os.listdir(WATCH_FOLDER)
        if f.startswith(FILE_PREFIX) and f.lower().endswith(".pdf")
    ]
    files.sort(key=lambda x: os.path.getctime(x))
    return files


# ==============================
# STARTUP BANNER
# ==============================

os.system("color")  # enable ANSI on Windows CMD

banner_lines = [
    "",
    f"  {Color.BOLD}{Color.CYAN}╔══════════════════════════════════════════════╗{Color.RESET}",
    f"  {Color.BOLD}{Color.CYAN}║   🖨  Marg ERP Auto Printer                  ║{Color.RESET}",
    f"  {Color.BOLD}{Color.CYAN}╚══════════════════════════════════════════════╝{Color.RESET}",
    "",
]
for line in banner_lines:
    print(line)

cprint(f"  {'Developed by':<16}", Color.DIM, end="")
cprint("Mehak Singh", Color.MAGENTA, bold=True)
print()

info_rows = [
    ("Printer",        SELECTED_PRINTER, Color.CYAN),
    ("Watch Folder",   WATCH_FOLDER,     Color.BLUE),
    ("File Prefix",    FILE_PREFIX,      Color.YELLOW),
    ("Check Interval", f"{CHECK_INTERVAL}s", Color.GREEN),
    ("Silent Mode",    "ON" if USE_SILENT_MODE else "OFF",
                       Color.GREEN if USE_SILENT_MODE else Color.RED),
]

cprint(f"  {'─'*44}", Color.DIM)
for label, value, vc in info_rows:
    cprint(f"  {Color.DIM}{label:<18}{Color.RESET}  {vc}{Color.BOLD}{value}{Color.RESET}")
cprint(f"  {'─'*44}", Color.DIM)
print()

log("Auto Printer is running. Watching for new files...", "WATCH")
print()

# ==============================
# MAIN LOOP
# ==============================

while True:
    marg_files = get_marg_files()

    for file_path in marg_files:
        time.sleep(2)
        print_pdf(file_path)

    time.sleep(CHECK_INTERVAL)