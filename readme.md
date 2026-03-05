# 🖨 Marg ERP Auto Printer

**Automatic PDF printing for Marg ERP** — monitors a folder and silently sends matching PDFs to your printer the moment they appear.

Developed by **Mehak Singh** | [TheMehakCodes](https://themehakcodes.com)

---

## ✨ Features

- 👁 **Folder Watcher** — continuously monitors a folder for new PDF files
- 🏷 **Prefix Filter** — only prints files matching your configured prefix (e.g. `MC_PRINT`)
- 🔇 **Silent Mode** — uses SumatraPDF for zero-dialog background printing
- 🖥 **System Tray** — runs silently in the background with a tray icon
- 📋 **Live Logs** — real-time log window with colour-coded levels
- ⚙️ **Config Editor** — edit printer, folder, prefix, interval and silent mode at any time
- 🔄 **Auto-Updater** — checks GitHub for new versions on startup and updates silently

---

## 📦 Project Files

| File | Purpose |
|---|---|
| `marg_auto_printer.py` | Main application |
| `marg_updater.py` | Standalone updater EXE (replaces main EXE after update) |
| `build.bat` | Builds both EXEs with PyInstaller |
| `installer.iss` | Inno Setup installer script |
| `version.json` | Hosted on GitHub — controls update delivery |
| `SumatraPDF.exe` | Bundled PDF printer (silent mode) |
| `logo.ico` / `logo.png` | App icon |

---

## 🚀 Getting Started

### Requirements

- Windows 10 / 11
- Python 3.11
- PyInstaller, pywin32, pystray, Pillow, requests

### Install dependencies

```bat
pip install pyinstaller pywin32 pystray pillow requests
```

### Build

```bat
build.bat
```

This produces two files in `dist\`:

```
dist\marg_auto_printer.exe   ← main application
dist\marg_updater.exe        ← updater helper (must ship alongside main)
```

### Create Installer

1. Copy both EXEs + `SumatraPDF.exe` + `logo.ico` + `logo.png` to your `InstallerBuild\` folder
2. Open `installer.iss` in [Inno Setup](https://jrsoftware.org/isinfo.php)
3. Compile → produces `Output\Marg_ERP_Auto_Printer_Setup.exe`

---

## ⚙️ Configuration

Configuration is written to `config.json` in the install directory. It is created automatically during first run or via the Inno Setup wizard.

```json
{
    "printer":        "Your Printer Name",
    "watch_folder":   "C:\\Users\\You\\Downloads",
    "file_prefix":    "MC_PRINT",
    "check_interval": 5,
    "silent_mode":    true
}
```

| Key | Description |
|---|---|
| `printer` | Exact Windows printer name |
| `watch_folder` | Folder to monitor for new PDFs |
| `file_prefix` | Only print files whose name starts with this (leave blank for all PDFs) |
| `check_interval` | Seconds between folder scans |
| `silent_mode` | `true` = use SumatraPDF silently, `false` = use Windows shell print |

You can edit these at any time from the tray icon → **Edit Config**.  
Changes take effect after restarting the app.

---

## 🔄 Auto-Update System

The app checks for updates on startup by fetching `version.json` from GitHub.

### `version.json` format

```json
{
    "version":      "1.0.9",
    "release_type": "direct",
    "exe_url":      "https://github.com/YourRepo/releases/download/v1.0.9/marg_auto_printer.exe",
    "sha256":       ""
}
```

| Field | Description |
|---|---|
| `version` | Latest version string — compared against `APP_VERSION` in the running app |
| `release_type` | `"direct"` = silent EXE swap via `marg_updater.exe` · `"installer"` = prompts user to run a Setup EXE |
| `exe_url` | Direct download URL for the new EXE (or Setup EXE for installer type) |
| `sha256` | Optional SHA-256 checksum of the download — leave `""` to skip verification |

### How `direct` updates work

```
App detects newer version
    ↓
Downloads new marg_auto_printer.exe → update_new.exe
    ↓
Launches marg_updater.exe (elevated if in Program Files)
    ↓
App shuts down cleanly
    ↓
marg_updater.exe waits for PID to exit
    ↓
Replaces marg_auto_printer.exe with update_new.exe
    ↓
Restarts the new EXE
```

> ⚠️ `marg_updater.exe` **must** be present in the same folder as `marg_auto_printer.exe` for direct updates to work. The Inno Setup installer handles this automatically.

### Releasing a new version

1. Bump `APP_VERSION` in `marg_auto_printer.py`
2. Run `build.bat`
3. Upload `dist\marg_auto_printer.exe` to GitHub Releases
4. Update `version.json` with the new version and download URL
5. Commit and push `version.json` — users will receive the update on next startup

---

## 🖥 System Tray Menu

| Item | Action |
|---|---|
| 📋 Show Logs | Opens the live log window |
| ⚙️ Edit Config | Opens the configuration editor |
| ℹ️ About | Shows version and developer info |
| 🔄 Check for Updates | Manually triggers an update check |
| ✖ Exit | Stops the watcher and exits |

---

## 📋 Log Levels

| Level | Colour | Meaning |
|---|---|---|
| `INFO` | Blue | General status messages |
| `SUCCESS` | Green | Operation completed successfully |
| `WARN` | Orange | Non-critical warnings |
| `ERROR` | Red | Something went wrong |
| `PRINT` | Purple | A PDF is being sent to the printer |
| `WATCH` | Blue | File detected in watch folder |
| `UPDATE` | Pink | Update checker activity |

---

## 🏗 Architecture

```
marg_auto_printer.exe
├── Tray icon thread          (pystray)
├── File watcher thread       (polls watch_folder every check_interval seconds)
├── Print worker thread       (ordered queue, one file at a time)
├── Updater thread            (daemon, runs once at startup)
└── Tkinter root              (hidden, keeps GUI alive for Toplevel windows)

marg_updater.exe              (separate process, launched only during update)
├── Waits for main PID to exit
├── Replaces EXE file
└── Restarts new EXE
```

---

## 🔒 Silent Printing (SumatraPDF)

When `silent_mode` is `true`, printing is handled by the bundled `SumatraPDF.exe`:

```
SumatraPDF.exe -print-to "Printer Name" -print-settings noscale -silent -exit-on-print file.pdf
```

If `SumatraPDF.exe` is missing, the app falls back to the Windows shell `print` verb automatically.

---

## 📁 Install Directory Layout (after install)

```
C:\Program Files (x86)\Marg ERP Auto Printer\
├── marg_auto_printer.exe
├── marg_updater.exe
├── SumatraPDF.exe
├── logo.ico
├── config.json               ← created on first run
└── update_new.exe            ← temporary, only present during an update
```

---

## 🛠 Troubleshooting

**App doesn't print anything**
- Check the watch folder path and file prefix in Edit Config
- Ensure the printer name matches exactly what Windows shows
- Open Show Logs to see if files are being detected

**Silent mode not working**
- Verify `SumatraPDF.exe` is in the same folder as the main EXE
- A `WARN` log entry will appear if it falls back to legacy mode

**Update loops / keeps re-updating**
- Make sure `APP_VERSION` in the built EXE matches or exceeds the `version` in `version.json`
- Ensure `exe_url` points to a plain `.exe`, not a ZIP file

**UAC prompt on every update**
- Expected when installed in `Program Files` — the updater needs elevation to replace the EXE
- Install to a user-writable directory to avoid UAC prompts

---

## 👤 Developer

**Mehak Singh**  
TheMehakCodes  
[https://themehakcodes.com](https://themehakcodes.com)

---

*Built with Python 3.11 · PyInstaller · pystray · Pillow · SumatraPDF · Inno Setup*