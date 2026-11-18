"""
Stops the MCTM server that MCTM.pyw started in the background.

Only stops the process whose ID the launcher saved in data/server.pid, and only
if that process is still a Python process -- never anything else.

    pythonw installer/stop_server.pyw            # shows a small confirmation box
    python  installer/stop_server.pyw --quiet    # no dialogs (used by the uninstaller)
"""

import os
import sys
import subprocess

APP_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PID_FILE = os.path.join(APP_DIR, "data", "server.pid")
CREATE_NO_WINDOW = 0x08000000


def say(message):
    if "--quiet" in sys.argv:
        print(message)
        return
    try:
        import tkinter
        from tkinter import messagebox
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo("MCTM Tracker", message)
        root.destroy()
    except Exception:
        print(message)


def process_image(pid):
    """Executable name of a running PID, or None if it isn't running."""
    out = subprocess.run(
        ["tasklist", "/FI", f"PID eq {pid}", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, creationflags=CREATE_NO_WINDOW,
    ).stdout.strip()
    if not out or out.startswith("INFO"):
        return None
    return out.split(",")[0].strip('"').lower()


def main():
    try:
        with open(PID_FILE) as f:
            pid = int(f.read().strip())
    except (OSError, ValueError):
        say("The MCTM server isn't running (or wasn't started by the MCTM app).")
        return 0

    image = process_image(pid)
    if image is None:
        os.remove(PID_FILE)
        say("The MCTM server was already stopped.")
        return 0
    if not image.startswith("python"):
        # The PID was reused by some other program after a restart. Leave it alone.
        os.remove(PID_FILE)
        say("The MCTM server was already stopped.")
        return 0

    subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, creationflags=CREATE_NO_WINDOW)
    try:
        os.remove(PID_FILE)
    except OSError:
        pass
    say("The MCTM server has stopped. Other PCs can't use the tracker until it's started again.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
