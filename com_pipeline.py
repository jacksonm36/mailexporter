"""
Dedicated STA COM worker thread for Outlook import.

Overlaps disk prep with COM import on a single session (safe for one PST).
True multi-thread COM to the same PST is NOT supported (Outlook MAPI is STA).
"""

from __future__ import annotations

import logging
import os
import queue
import threading
import time
from dataclasses import dataclass
from typing import Any, Callable

logger = logging.getLogger(__name__)

_SENTINEL = object()


def _env_int(name: str, default: int) -> int:
    try:
        return max(0, int(os.environ.get(name, str(default)) or default))
    except (TypeError, ValueError):
        return default


COM_PIPELINE_ENABLED = os.environ.get("EML2PST_COM_PIPELINE", "").strip().lower() in (
    "1",
    "true",
    "yes",
    "on",
)
COM_WORKERS = _env_int("EML2PST_COM_WORKERS", 0)


def com_pipeline_should_run() -> bool:
    """Use dedicated COM thread when pipeline flag set or COM_WORKERS == 1."""
    if COM_PIPELINE_ENABLED:
        return True
    if COM_WORKERS == 1:
        return True
    if COM_WORKERS > 1:
        logger.warning(
            "EML2PST_COM_WORKERS=%s: parallel COM to one PST is unsupported; "
            "using single COM pipeline thread (set COM_WORKERS=0 or 1).",
            COM_WORKERS,
        )
        return True
    return False


@dataclass
class ImportJob:
    path: str
    prep: Any | None = None


@dataclass
class ImportResult:
    path: str
    success: bool
    status: str = "error"
    detail: str = ""
    target_label: str = ""
    exception: BaseException | None = None


class ComImportPipeline:
    """
    Producer (export loop) enqueues jobs; one COM STA worker imports into Outlook.
    """

    def __init__(
        self,
        *,
        import_one: Callable[[str, Any | None], ImportResult],
        on_worker_start: Callable[[], Any] | None = None,
        on_worker_stop: Callable[[], None] | None = None,
        queue_size: int = 64,
    ):
        self._import_one = import_one
        self._on_worker_start = on_worker_start
        self._on_worker_stop = on_worker_stop
        self._in_q: queue.Queue[Any] = queue.Queue(maxsize=max(8, queue_size))
        self._out_q: queue.Queue[ImportResult] = queue.Queue()
        self._thread: threading.Thread | None = None
        self._stop = threading.Event()
        self._started = threading.Event()
        self._error: BaseException | None = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._stop.clear()
        self._thread = threading.Thread(
            target=self._worker_main,
            name="outlook-com-pipeline",
            daemon=True,
        )
        self._thread.start()
        if not self._started.wait(timeout=120.0):
            raise RuntimeError("COM pipeline worker failed to start within 120s")

    def _worker_main(self) -> None:
        ctx = None
        try:
            if self._on_worker_start:
                ctx = self._on_worker_start()
            self._started.set()
            while not self._stop.is_set():
                try:
                    job = self._in_q.get(timeout=0.25)
                except queue.Empty:
                    continue
                if job is _SENTINEL:
                    break
                if not isinstance(job, ImportJob):
                    continue
                try:
                    result = self._import_one(job.path, job.prep)
                except Exception as exc:
                    result = ImportResult(
                        path=job.path,
                        success=False,
                        detail=str(exc),
                        exception=exc,
                    )
                self._out_q.put(result)
        except Exception as exc:
            self._error = exc
            logger.exception("COM pipeline worker failed: %s", exc)
            self._started.set()
        finally:
            if self._on_worker_stop:
                try:
                    self._on_worker_stop(ctx)
                except Exception:
                    logger.debug("COM pipeline stop hook failed", exc_info=True)

    def submit(self, path: str, prep: Any | None = None, *, block: bool = True, timeout: float = 300.0) -> None:
        job = ImportJob(path=path, prep=prep)
        if block:
            self._in_q.put(job, timeout=timeout)
        else:
            self._in_q.put_nowait(job)

    def get_result(self, timeout: float = 600.0) -> ImportResult | None:
        try:
            return self._out_q.get(timeout=timeout)
        except queue.Empty:
            return None

    def drain_pending_results(self) -> list[ImportResult]:
        out: list[ImportResult] = []
        while True:
            try:
                out.append(self._out_q.get_nowait())
            except queue.Empty:
                break
        return out

    def shutdown(self, *, wait: bool = True, timeout: float = 30.0) -> None:
        self._stop.set()
        try:
            self._in_q.put_nowait(_SENTINEL)
        except queue.Full:
            self._in_q.put(_SENTINEL, timeout=5.0)
        if self._thread and wait:
            self._thread.join(timeout=timeout)
        self._thread = None

    @property
    def worker_error(self) -> BaseException | None:
        return self._error
