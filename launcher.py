"""
Launcher for the Mandats TXT→XLSX Streamlit app.
Works both as a plain Python script and as a PyInstaller-frozen exe.
"""
import os
import socket
import subprocess
import sys
import threading
import time
import webbrowser
from pathlib import Path


def free_port(preferred: int = 8501) -> int:
    """Return *preferred* if available, otherwise any free port."""
    for port in (preferred, 0):
        try:
            with socket.socket() as s:
                s.bind(("127.0.0.1", port))
                return s.getsockname()[1]
        except OSError:
            continue
    raise RuntimeError("No free port found")


def main() -> None:
    # When frozen by PyInstaller, data files live in sys._MEIPASS
    if getattr(sys, "frozen", False):
        base = Path(sys._MEIPASS)
        python = Path(sys.executable).parent / "python.exe"
    else:
        base = Path(__file__).parent
        python = Path(sys.executable)

    app_py = base / "app.py"
    port = free_port()
    url = f"http://127.0.0.1:{port}"

    cmd = [
        str(python), "-m", "streamlit", "run", str(app_py),
        "--server.port", str(port),
        "--server.headless", "true",
        "--server.address", "127.0.0.1",
        "--browser.gatherUsageStats", "false",
    ]

    # Open browser after streamlit has had time to start
    def _open():
        time.sleep(3)
        webbrowser.open(url)

    threading.Thread(target=_open, daemon=True).start()

    proc = subprocess.Popen(cmd)
    try:
        proc.wait()
    except KeyboardInterrupt:
        proc.terminate()


if __name__ == "__main__":
    main()
