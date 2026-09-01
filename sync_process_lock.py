"""One process-wide lock shared by every order synchronization batch."""

from __future__ import annotations

from contextlib import contextmanager
import os


SYNC_LOCK_FILE = os.environ.get(
    "WOO_ANALYSIS_SYNC_LOCK_FILE",
    "/tmp/woo-analysis-order-sync.lock",
)


class SyncAlreadyRunning(RuntimeError):
    """Raised when another automatic or full synchronization owns the lock."""


@contextmanager
def exclusive_sync_lock():
    """Acquire the shared non-blocking Linux process lock."""
    import fcntl

    lock_handle = open(SYNC_LOCK_FILE, "a+", encoding="ascii")
    acquired = False
    try:
        try:
            fcntl.flock(lock_handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            acquired = True
        except BlockingIOError as exc:
            raise SyncAlreadyRunning("another synchronization is running") from exc
        lock_handle.seek(0)
        lock_handle.truncate()
        lock_handle.write(str(os.getpid()))
        lock_handle.flush()
        yield
    finally:
        if acquired:
            try:
                fcntl.flock(lock_handle.fileno(), fcntl.LOCK_UN)
            except OSError:
                pass
        lock_handle.close()
