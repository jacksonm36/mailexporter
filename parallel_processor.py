"""
Parallel helpers for Mail Exporter.

Outlook COM import must stay on a single thread (STA). Parallelism here is limited to
disk/CPU prep work (fingerprints, optional MIME parse) ahead of the COM import loop.
"""

from __future__ import annotations

import concurrent.futures
import logging
import multiprocessing as mp
import os
import threading
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from typing import Any

logger = logging.getLogger(__name__)

try:
    import pythoncom
except ImportError:
    pythoncom = None  # type: ignore


def _env_int(name: str, default: int, *, lo: int, hi: int) -> int:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default
    try:
        return max(lo, min(int(raw), hi))
    except ValueError:
        return default


PARALLEL_PREP_WORKERS = _env_int("EML2PST_PARALLEL_WORKERS", 0, lo=0, hi=16)
PARALLEL_PREP_BATCH = _env_int("EML2PST_PARALLEL_PREP_BATCH", 50, lo=1, hi=500)
PARALLEL_COM_SLOTS = _env_int("EML2PST_PARALLEL_COM_SLOTS", 1, lo=1, hi=4)
PARALLEL_BATCH_TIMEOUT = _env_int("EML2PST_PARALLEL_BATCH_TIMEOUT", 300, lo=30, hi=3600)


@dataclass
class EmlPrepResult:
    """Per-file prep computed without Outlook COM."""

    file_path: str
    norm_path: str
    size: int
    dedup_key: str | None = None
    header_date: str = ""
    email_data: dict | None = None
    prep_error: str | None = None


class ParallelEmailProcessor:
    """
    Thread-pool batch runner with optional COM slot limiting.

    For Mail Exporter, ``process_func`` must NOT call Outlook unless
    ``outlook_com_slots=1`` and you accept stability risk. Default export uses
    ``EmlPrepPrefetcher`` instead (prep only, single-threaded COM import).
    """

    def __init__(self, max_workers: int | None = None, outlook_com_slots: int = 1):
        cpu = mp.cpu_count() or 4
        self.max_workers = max_workers if max_workers is not None else min(cpu, 8)
        self.outlook_semaphore = threading.Semaphore(max(1, outlook_com_slots))
        self._executor: concurrent.futures.ThreadPoolExecutor | None = None

    def process_batch_parallel(
        self,
        eml_files: list[str],
        process_func: Callable[[str], Any],
        *,
        batch_size: int = 100,
    ) -> Iterator[list[Any]]:
        """Process paths in parallel batches; yields each batch result list."""
        total_batches = (len(eml_files) + batch_size - 1) // batch_size
        with concurrent.futures.ThreadPoolExecutor(
            max_workers=self.max_workers
        ) as executor:
            futures: list[concurrent.futures.Future] = []
            for i in range(0, len(eml_files), batch_size):
                batch = eml_files[i : i + batch_size]
                futures.append(
                    executor.submit(
                        self._process_batch_with_limit, batch, process_func
                    )
                )
            done_batches = 0
            for future in concurrent.futures.as_completed(futures):
                try:
                    batch_results = future.result(timeout=PARALLEL_BATCH_TIMEOUT)
                    done_batches += 1
                    yield batch_results
                except Exception as e:
                    logger.error(
                        "Parallel batch failed (%d/%d): %s",
                        done_batches,
                        total_batches,
                        e,
                    )
                    yield []

    def _process_batch_with_limit(
        self, batch: list[str], process_func: Callable[[str], Any]
    ) -> list[Any]:
        with self.outlook_semaphore:
            com_initialized = False
            if pythoncom is not None:
                try:
                    pythoncom.CoInitialize()
                    com_initialized = True
                except Exception:
                    pass
            try:
                return [process_func(path) for path in batch]
            finally:
                if com_initialized and pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass


class EmlPrepPrefetcher:
    """
    Prefetch dedup keys (and optional parse summaries) on a thread pool while
    the main thread imports into Outlook sequentially.
    """

    def __init__(
        self,
        prep_func: Callable[[str], EmlPrepResult],
        *,
        max_workers: int,
        max_in_flight: int | None = None,
    ):
        self._prep_func = prep_func
        self._max_workers = max(1, max_workers)
        self._max_in_flight = max_in_flight or self._max_workers * 4
        self._executor: concurrent.futures.ThreadPoolExecutor | None = None
        self._futures: dict[str, concurrent.futures.Future] = {}
        self._lock = threading.Lock()
        self._paths_queue: list[str] = []
        self._queue_idx = 0

    def start(self, file_paths: list[str]) -> None:
        self._paths_queue = list(file_paths)
        self._queue_idx = 0
        self._executor = concurrent.futures.ThreadPoolExecutor(
            max_workers=self._max_workers,
            thread_name_prefix="eml-prep",
        )
        self._schedule_ahead()

    def _schedule_ahead(self) -> None:
        if not self._executor:
            return
        with self._lock:
            while (
                self._queue_idx < len(self._paths_queue)
                and len(self._futures) < self._max_in_flight
            ):
                path = self._paths_queue[self._queue_idx]
                self._queue_idx += 1
                norm = os.path.normpath(os.path.abspath(path))
                if norm in self._futures:
                    continue
                self._futures[norm] = self._executor.submit(self._prep_func, path)

    def take(self, norm_path: str, *, timeout: float | None = 120.0) -> EmlPrepResult | None:
        """Block until prep for norm_path finishes (or return None if not scheduled)."""
        self._schedule_ahead()
        with self._lock:
            future = self._futures.pop(norm_path, None)
        if future is None:
            return None
        try:
            return future.result(timeout=timeout)
        except concurrent.futures.TimeoutError:
            logger.warning("Prep timeout for %s", norm_path)
            return EmlPrepResult(
                file_path=norm_path,
                norm_path=norm_path,
                size=0,
                prep_error="Prep timeout",
            )
        except Exception as e:
            return EmlPrepResult(
                file_path=norm_path,
                norm_path=norm_path,
                size=0,
                prep_error=str(e),
            )
        finally:
            self._schedule_ahead()

    def shutdown(self, *, cancel_pending: bool = False) -> None:
        if self._executor is None:
            return
        with self._lock:
            pending = list(self._futures.values())
            self._futures.clear()
        if cancel_pending:
            for fut in pending:
                fut.cancel()
        try:
            self._executor.shutdown(wait=not cancel_pending, cancel_futures=cancel_pending)
        except TypeError:
            self._executor.shutdown(wait=not cancel_pending)
        self._executor = None


class AsyncEmlPrepPrefetcher(EmlPrepPrefetcher):
    """
    Same prefetch API as EmlPrepPrefetcher but schedules prep on a background asyncio loop.
    Enable with EML2PST_ASYNC_IO=1 (optional aiofiles for header reads in converter).
    """

    def __init__(
        self,
        prep_func: Callable[[str], EmlPrepResult],
        *,
        max_workers: int,
        max_in_flight: int | None = None,
    ):
        self._prep_func = prep_func
        self._max_workers = max(1, max_workers)
        self._max_in_flight = max_in_flight or self._max_workers * 4
        self._runner: Any = None
        self._paths_queue: list[str] = []
        self._queue_idx = 0
        self._scheduled: set[str] = set()
        self._lock = threading.Lock()

    def start(self, file_paths: list[str]) -> None:
        from async_processor import AsyncPrepLoopRunner

        self._paths_queue = list(file_paths)
        self._queue_idx = 0
        self._scheduled = set()
        self._runner = AsyncPrepLoopRunner(
            self._prep_func, concurrency=self._max_in_flight
        )
        self._runner.start()
        self._schedule_ahead()

    def _schedule_ahead(self) -> None:
        if not self._runner:
            return
        with self._lock:
            while (
                self._queue_idx < len(self._paths_queue)
                and len(self._scheduled) < self._max_in_flight
            ):
                path = self._paths_queue[self._queue_idx]
                self._queue_idx += 1
                norm = os.path.normpath(os.path.abspath(path))
                if norm in self._scheduled:
                    continue
                self._scheduled.add(norm)
                self._runner.submit(path)

    def take(self, norm_path: str, *, timeout: float | None = 120.0) -> EmlPrepResult | None:
        self._schedule_ahead()
        if not self._runner:
            return None
        result = self._runner.take(norm_path, timeout=timeout or 120.0)
        with self._lock:
            self._scheduled.discard(norm_path)
        self._schedule_ahead()
        if result is None:
            return EmlPrepResult(
                file_path=norm_path,
                norm_path=norm_path,
                size=0,
                prep_error="Prep timeout",
            )
        if isinstance(result, EmlPrepResult):
            return result
        if isinstance(result, BaseException):
            return EmlPrepResult(
                file_path=norm_path,
                norm_path=norm_path,
                size=0,
                prep_error=str(result),
            )
        return result

    def shutdown(self, *, cancel_pending: bool = False) -> None:
        if self._runner:
            self._runner.shutdown()
        self._runner = None


def create_eml_prep_prefetcher(
    prep_func: Callable[[str], EmlPrepResult],
    *,
    max_workers: int,
) -> EmlPrepPrefetcher:
    """Thread-pool prefetcher, or asyncio-backed when EML2PST_ASYNC_IO=1."""
    try:
        from async_processor import ASYNC_IO_ENABLED
    except ImportError:
        ASYNC_IO_ENABLED = False
    if ASYNC_IO_ENABLED and max_workers > 0:
        logger.info(
            "Async I/O prep enabled (concurrency ~%d; install aiofiles for async reads)",
            max_workers * 4,
        )
        return AsyncEmlPrepPrefetcher(prep_func, max_workers=max_workers)
    return EmlPrepPrefetcher(prep_func, max_workers=max_workers)
