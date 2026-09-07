"""watcher.py
==============
Watches the 'Weekly Reports' folder for a new 'Weekly Sales Report (main).xlsx'.
When the file is created (or replaced), waits a few seconds for the copy to
finish, then runs touchbistro.py automatically.

Usage
-----
  python watcher.py          — run in foreground (Ctrl+C to stop)
  run_watcher.bat            — launched by Task Scheduler at logon

The watcher stays running in the background until you close it.
"""

import os
import sys
import time
import subprocess
from datetime import datetime
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# ── Config ──────────────────────────────────────────────────────
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
WATCH_DIR   = os.path.normpath(os.path.join(SCRIPT_DIR, "..", "Weekly Reports"))
TARGET_FILE = "Weekly Sales Report (main).xlsx"
DEBOUNCE_SECONDS = 10   # wait for the file copy to fully complete
TOUCHBISTRO_SCRIPT = os.path.join(SCRIPT_DIR, "touchbistro.py")
# ────────────────────────────────────────────────────────────────


def log(msg: str):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {msg}", flush=True)


class ReportHandler(FileSystemEventHandler):
    """Triggers touchbistro.py when the target file is created or replaced."""

    def __init__(self):
        super().__init__()
        self._last_trigger = 0  # timestamp of last trigger (for debounce)

    def _is_target(self, event) -> bool:
        """Check if the event is for our target file (not a directory)."""
        if event.is_directory:
            return False
        basename = os.path.basename(event.src_path)
        return basename.lower() == TARGET_FILE.lower()

    def _maybe_trigger(self, event):
        if not self._is_target(event):
            return

        now = time.time()
        if now - self._last_trigger < DEBOUNCE_SECONDS:
            return  # still within debounce window

        self._last_trigger = now
        log(f"Detected: {os.path.basename(event.src_path)} ({event.event_type})")
        log(f"Waiting {DEBOUNCE_SECONDS}s for file copy to complete...")
        time.sleep(DEBOUNCE_SECONDS)

        log("Launching touchbistro.py...")
        try:
            result = subprocess.run(
                [sys.executable, TOUCHBISTRO_SCRIPT],
                cwd=SCRIPT_DIR,
                timeout=600,  # 10 min max
            )
            if result.returncode == 0:
                log("✅ touchbistro.py completed successfully.")
            else:
                log(f"⚠️  touchbistro.py exited with code {result.returncode}")
        except subprocess.TimeoutExpired:
            log("⚠️  touchbistro.py timed out after 10 minutes.")
        except Exception as exc:
            log(f"⚠️  Error running touchbistro.py: {exc}")

    def on_created(self, event):
        self._maybe_trigger(event)

    def on_modified(self, event):
        self._maybe_trigger(event)


def main():
    log(f"Watching: {WATCH_DIR}")
    log(f"Trigger file: {TARGET_FILE}")
    log("Waiting for file to appear...\n")

    handler = ReportHandler()
    observer = Observer()
    observer.schedule(handler, WATCH_DIR, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        log("Stopping watcher...")
        observer.stop()
    observer.join()
    log("Watcher stopped.")


if __name__ == "__main__":
    main()
