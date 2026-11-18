"""
Stops the MCTM server that MCTM.pyw started in the background.

Stops the process whose ID the launcher saved in data/server.pid, plus any
other Python process still listening on the tracker's port (a server left over
from an earlier start). Only ever stops Python processes -- never anything else.

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


PORT = 5000


def pids_listening_on_port(port=PORT):
    """Process IDs listening on the tracker's port (from netstat)."""
    out = subprocess.run(["netstat", "-ano", "-p", "TCP"], capture_output=True, text=True,
                         creationflags=CREATE_NO_WINDOW).stdout
    pids = set()
    for line in out.splitlines():
        parts = line.split()
        if len(parts) >= 5 and parts[3].upper() == "LISTENING" and parts[1].endswith(f":{port}"):
            try:
                pids.add(int(parts[4]))
            except ValueError:
                pass
    return pids


def main():
    pids = set()
    try:
        with open(PID_FILE) as f:
            pids.add(int(f.read().strip()))
    except (OSError, ValueError):
        pass
    pids |= pids_listening_on_port()

    stopped = 0
    for pid in sorted(pids):
        image = process_image(pid)
        if image is None or not image.startswith("python"):
            continue  # gone already, or the ID now belongs to another program: leave it alone
        subprocess.run(["taskkill", "/PID", str(pid), "/T", "/F"], capture_output=True, creationflags=CREATE_NO_WINDOW)
        stopped += 1
    try:
        os.remove(PID_FILE)
    except OSError:
        pass
    if stopped:
        say("The MCTM server has stopped. Other PCs can't use the tracker until it's started again.")
    else:
        say("The MCTM server wasn't running.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
