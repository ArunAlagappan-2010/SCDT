"""
MCTM Tracker - one-click launcher.

Double-click (or use the desktop / Start menu shortcut the installer makes):
  1. creates the Excel database on the very first run,
  2. starts the web server quietly in the background if it isn't running yet,
  3. opens the tracker in its own app window (Edge/Chrome "app mode": no tabs,
     no address bar), or the default browser if neither is installed.

Closing the window does NOT stop the server, on purpose: the other PCs on the
network keep working. Use "Stop MCTM server" from the Start menu to stop it.

Command-line flags (used by the installer's self-test):
  --no-window   start/check the server but don't open a window
"""

import os
import sys
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


def background_python():
    """pythonw.exe when we can find it, so the server never shows a console."""
    exe = sys.executable
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
    if status == "other":
        show_error(
            f"Another program is already using port {PORT} on this PC, so the tracker can't start.\n\n"
            "Close that program (or restart the PC) and try again."
        )
        return 1

    if status is None:
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
