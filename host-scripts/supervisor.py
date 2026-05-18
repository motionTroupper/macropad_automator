"""Supervisor for macro-daemon.py.

Spawns the daemon and restarts it on (a) process death and (b) hang — detected
by the staleness of the daemon's own `macro-daemon.watchdog` file, which it
refreshes every 500 ms.

Run this instead of running macro-daemon.py directly:

    pythonw supervisor.py

Logs supervisor activity (spawns, exits, hangs, backoff) to macro-supervisor.log.
"""

import os
import sys
import time
import subprocess
import datetime
from pathlib import Path


base_path = Path(sys.argv[0]).resolve().parent
os.chdir(base_path)

LOG_PATH = base_path / "macro-supervisor.log"
DAEMON_SCRIPT = base_path / "macro-daemon.py"
WATCHDOG_PATH = base_path / "macro-daemon.watchdog"

## Tolerance for the daemon's heartbeat watchdog file. If the daemon stops
## updating it for longer than WATCHDOG_MAX_AGE seconds (after an initial grace
## period), we consider it hung and kill+respawn it.
WATCHDOG_MAX_AGE = 15
WATCHDOG_GRACE_AFTER_START = 30

## Backoff between respawns; doubles on quick deaths to avoid thrash loops.
MIN_BACKOFF = 2
MAX_BACKOFF = 60
RAPID_WINDOW = 30


def log(msg):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(f"[{ts}] {msg}\n")
            f.flush()
            try:
                os.fsync(f.fileno())
            except Exception:
                pass
    except Exception:
        pass


def watchdog_age():
    try:
        if not WATCHDOG_PATH.exists():
            return None
        return time.time() - WATCHDOG_PATH.stat().st_mtime
    except Exception:
        return None


def supervise_one_run():
    """Spawn the daemon and supervise until it exits or hangs.
    Returns the elapsed wall-clock seconds of this run."""
    start = time.monotonic()
    log(f"Launching daemon: {DAEMON_SCRIPT}")
    try:
        proc = subprocess.Popen(
            [sys.executable, str(DAEMON_SCRIPT)],
            cwd=str(base_path),
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )
    except Exception as e:
        log(f"Failed to spawn daemon: {e}")
        return time.monotonic() - start

    log(f"Daemon started (pid={proc.pid})")

    while True:
        rc = proc.poll()
        if rc is not None:
            elapsed = time.monotonic() - start
            log(f"Daemon exited rc={rc} after {elapsed:.1f}s")
            return elapsed

        elapsed = time.monotonic() - start
        if elapsed > WATCHDOG_GRACE_AFTER_START:
            age = watchdog_age()
            if age is not None and age > WATCHDOG_MAX_AGE:
                log(f"Daemon hung (watchdog stale {age:.1f}s); terminating pid={proc.pid}")
                try:
                    proc.terminate()
                    try:
                        proc.wait(timeout=5)
                    except subprocess.TimeoutExpired:
                        log(f"Terminate timed out; killing pid={proc.pid}")
                        proc.kill()
                        try:
                            proc.wait(timeout=5)
                        except Exception:
                            pass
                except Exception as e:
                    log(f"Error terminating daemon: {e}")
                return time.monotonic() - start

        time.sleep(1)


def main():
    log(f"=== supervisor starting (pid={os.getpid()}) executable={sys.executable} ===")

    backoff = MIN_BACKOFF
    rapid_restarts = 0

    while True:
        elapsed = supervise_one_run()

        if elapsed < RAPID_WINDOW:
            rapid_restarts += 1
            backoff = min(MAX_BACKOFF, MIN_BACKOFF * (2 ** rapid_restarts))
            log(f"Quick exit ({elapsed:.1f}s); rapid_restarts={rapid_restarts}, backoff={backoff}s")
        else:
            rapid_restarts = 0
            backoff = MIN_BACKOFF

        log(f"Sleeping {backoff}s before respawn")
        time.sleep(backoff)


if __name__ == "__main__":
    main()
