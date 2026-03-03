# ============================================================
#  Marg ERP Auto Printer - Updater
#  Developed by Mehak Singh | TheMehakCodes
# ============================================================

import os
import sys
import time
import shutil
import subprocess

def main():
    """Replace current executable with downloaded update"""
    # Wait for main app to fully close
    time.sleep(3)
    
    if len(sys.argv) < 2:
        # Log error silently
        with open("updater_error.log", "w") as f:
            f.write("No update file specified")
        sys.exit(1)
    
    temp_exe = sys.argv[1]
    
    # Get the path of the main executable
    if getattr(sys, 'frozen', False):
        current_dir = os.path.dirname(sys.executable)
    else:
        current_dir = os.path.dirname(os.path.abspath(__file__))
    
    current_exe = os.path.join(current_dir, "marg_auto_printer.exe")
    
    try:
        # Verify temp file exists
        if not os.path.exists(temp_exe):
            raise Exception(f"Temp file not found: {temp_exe}")
        
        # Replace executable
        if os.path.exists(current_exe):
            os.remove(current_exe)
        
        shutil.copy2(temp_exe, current_exe)
        
        # Clean up temp file
        if os.path.exists(temp_exe):
            os.remove(temp_exe)
        
        # Update version.txt if it exists
        version_file = os.path.join(current_dir, "version.txt")
        if os.path.exists(version_file):
            # You could update version here if needed
            pass
        
        # Restart the application
        subprocess.Popen([current_exe])
        
    except Exception as e:
        # Log error
        with open(os.path.join(current_dir, "update_error.log"), "w") as f:
            f.write(f"Update failed: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()