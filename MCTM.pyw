"""
MCTM Tracker - one-click launcher.

Double-click (or use the desktop / Start menu shortcut the installer makes):
  1. creates the Excel database on the very first run,
  2. starts the web server quietly in the background if it isn't running yet,
  3. opens the tracker in its own app window (Edge/Chrome "app mode": no tabs,
     no address bar), or the default browser if neither is installed.

Closing the window does NOT stop the server, on purpose: the other PCs on the
network keep working. Use "Stop MCTM server" from the Start menu to stop it.

The very first launch on a PC runs "Setup MCTM.cmd" (Python packages,
Tesseract, optional Gemini key, shortcuts) and waits for it to finish.

Command-line flags (used by the installer's self-test):
  --no-window    start/check the server but don't open a window
  --skip-setup   don't run first-time setup even if it hasn't run yet
"""

import os
import sys
import json
import time
import shutil
import subprocess
import urllib.error
import urllib.request
import webbrowser

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
DB_PATH = os.path.join(DATA_DIR, "SchoolDisciplineSystem.xlsx")
PID_FILE = os.path.join(DATA_DIR, "server.pid")
LOG_FILE = os.path.join(DATA_DIR, "server.log")
PORT = 5000
URL = f"http://127.0.0.1:{PORT}/"
START_TIMEOUT = 40  # seconds to wait for the server to answer

CREATE_NO_WINDOW = 0x08000000
DETACHED_PROCESS = 0x00000008
CREATE_NEW_PROCESS_GROUP = 0x00000200


def show_error(message):
    """A real dialog box: this script runs without a console window."""
    try:
        import tkinter
        from tkinter import messagebox
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showerror("MCTM Tracker", message)
        root.destroy()
    except Exception:
        print(message, file=sys.stderr)


def port_in_use():
    """True when something is listening on the port -- even a server too busy
    to answer an HTTP request right now."""
    import socket
    try:
        with socket.create_connection(("127.0.0.1", PORT), timeout=1):
            return True
    except OSError:
        return False


def wait_for_busy_server(seconds=45):
    """A server is listening but slow: wait for it instead of starting a
    second one (two servers on one port fight over requests)."""
    deadline = time.time() + seconds
    while time.time() < deadline:
        status = server_status()
        if status:
            return status
        time.sleep(1)
    return None


def server_status():
    """'ours' if MCTM answers on the port, 'other' if something else does,
    None if nothing is listening."""
    try:
        with urllib.request.urlopen(URL, timeout=2) as resp:
            return "ours" if b"MCTM" in resp.read(4000) else "other"
    except urllib.error.HTTPError:
        return "other"
    except Exception:
        return None


SETUP_DONE = os.path.join(DATA_DIR, "setup-done.json")
SETUP_CMD = os.path.join(BASE_DIR, "Setup MCTM.cmd")
CREATE_NEW_CONSOLE = 0x00000010


def setup_python():
    """The Python that setup installed the packages into, if setup has run."""
    try:
        with open(SETUP_DONE, encoding="utf-8-sig") as f:
            path = json.load(f).get("python") or ""
        return path if os.path.exists(path) else None
    except (OSError, ValueError):
        return None


def run_first_time_setup():
    """First launch on this PC: run Setup MCTM.cmd in its own window and wait.
    Returns True when setup finished (data/setup-done.json written)."""
    if os.path.exists(SETUP_DONE):
        return True
    if not os.path.exists(SETUP_CMD):
        return True  # running from a copy without the setup file: carry on
    subprocess.run(["cmd", "/c", SETUP_CMD], cwd=BASE_DIR,
                   creationflags=CREATE_NEW_CONSOLE if os.name == "nt" else 0)
    return os.path.exists(SETUP_DONE)


def background_python():
    """pythonw.exe when we can find it, so the server never shows a console.
    Prefers the Python setup installed the packages into: a PC can have
    several, and double-clicking this file may open a different one."""
    exe = setup_python() or sys.executable
    if os.path.basename(exe).lower() == "python.exe":
        candidate = os.path.join(os.path.dirname(exe), "pythonw.exe")
        if os.path.exists(candidate):
            return candidate
    return exe


def ensure_database():
    if os.path.exists(DB_PATH):
        return
    sys.path.insert(0, BASE_DIR)
    import build_excel_template
    build_excel_template.main()


def start_server():
    os.makedirs(DATA_DIR, exist_ok=True)
    log = open(LOG_FILE, "a", encoding="utf-8")
    log.write(f"\n--- started by launcher {time.strftime('%Y-%m-%d %H:%M:%S')} ---\n")
    log.flush()
    flags = CREATE_NO_WINDOW | DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
    proc = subprocess.Popen(
        [background_python(), "-u", os.path.join(BASE_DIR, "app.py")],
        cwd=BASE_DIR, stdout=log, stderr=subprocess.STDOUT, stdin=subprocess.DEVNULL,
        creationflags=flags, close_fds=True,
    )
    with open(PID_FILE, "w") as f:
        f.write(str(proc.pid))

    deadline = time.time() + START_TIMEOUT
    while time.time() < deadline:
        if proc.poll() is not None:
            return False  # crashed on start-up; the log says why
        if server_status() == "ours":
            return True
        time.sleep(0.4)
    return False


def find_app_browser():
    roots = [os.environ.get(v) for v in ("ProgramFiles(x86)", "ProgramFiles", "LOCALAPPDATA")]
    candidates = []
    for root in filter(None, roots):
        candidates.append(os.path.join(root, "Microsoft", "Edge", "Application", "msedge.exe"))
        candidates.append(os.path.join(root, "Google", "Chrome", "Application", "chrome.exe"))
    for path in candidates:
        if os.path.exists(path):
            return path
    return shutil.which("msedge") or shutil.which("chrome")


def open_window():
    browser = find_app_browser()
    if browser:
        subprocess.Popen(
            [browser, f"--app={URL}", "--window-size=1440,900", "--new-window"],
            creationflags=DETACHED_PROCESS if os.name == "nt" else 0, close_fds=True,
        )
    else:
        webbrowser.open(URL)


def main():
    no_window = "--no-window" in sys.argv

    status = server_status()
    if status is None and port_in_use():
        status = wait_for_busy_server()
        if status is None:
            show_error(
                "The tracker is already running but isn't answering right now. "
                "Wait a minute and open it again. If it stays stuck, use "
                "Start menu > MCTM Tracker > Stop MCTM server, then open it again."
            )
            return 1
    if status == "other":
        show_error(
            f"Another program is already using port {PORT} on this PC, so the tracker can't start.\n\n"
            "Close that program (or restart the PC) and try again."
        )
        return 1

    if status is None:
        if "--skip-setup" not in sys.argv and not run_first_time_setup():
            show_error(
                "Setup didn't finish, so the tracker can't start yet.\n\n"
                f"Run \"{os.path.basename(SETUP_CMD)}\" in {BASE_DIR} and read what it says."
            )
            return 1
        try:
            ensure_database()
        except Exception as exc:
            show_error(f"Couldn't create the Excel database:\n\n{exc}")
            return 1
        if not start_server():
            tail = ""
            try:
                with open(LOG_FILE, encoding="utf-8", errors="replace") as f:
                    tail = "".join(f.readlines()[-12:])
            except OSError:
                pass
            show_error(
                "The tracker's server didn't start.\n\n"
                f"Last lines of {LOG_FILE}:\n\n{tail or '(empty)'}"
            )
            return 1

    if not no_window:
        open_window()
    return 0


if __name__ == "__main__":
    sys.exit(main())
