"""
Async file I/O for parallel .eml prep (optional aiofiles; falls back to thread pool).
"""

from __future__ import annotations

import asyncio
import logging
import os
import threading
from collections.abc import Callable
from typing import Any

logger = logging.getLogger(__name__)

try:
    import aiofiles

    HAS_AIOFILES = True
except ImportError:
    aiofiles = None  # type: ignore
    HAS_AIOFILES = False


def _env_flag(name: str, default: bool = False) -> bool:
    raw = os.environ.get(name, "").strip().lower()
    if not raw:
        return default
    return raw in ("1", "true", "yes", "on")


ASYNC_IO_ENABLED = _env_flag("EML2PST_ASYNC_IO", False)
ASYNC_IO_CONCURRENCY = max(4, min(128, int(os.environ.get("EML2PST_ASYNC_CONCURRENCY", "32") or 32)))


async def _read_bytes_aio(path: str, max_bytes: int | None = None) -> bytes:
    if not HAS_AIOFILES:
        raise RuntimeError("aiofiles not installed")
    async with aiofiles.open(path, "rb") as handle:
        if max_bytes is None:
            return await handle.read()
        return await handle.read(max_bytes)


async def read_header_block_async(path: str, limit: int = 65536) -> bytes:
    """Read until blank line or limit (aiofiles when available)."""
    if HAS_AIOFILES:
        chunks: list[bytes] = []
        total = 0
        async with aiofiles.open(path, "rb") as handle:
            while total < limit:
                block = await handle.read(min(8192, limit - total))
                if not block:
                    break
                chunks.append(block)
                total += len(block)
                if b"\n\n" in block or b"\r\n\r\n" in block:
                    break
        return b"".join(chunks)
    loop = asyncio.get_running_loop()
    return await loop.run_in_executor(None, read_header_block_sync, path, limit)


def read_header_block_sync(path: str, limit: int = 65536) -> bytes:
    try:
        with open(path, "rb") as handle:
            return handle.read(limit)
    except OSError:
        return b""


async def prep_paths_async(
    paths: list[str],
    prep_func: Callable[[str], Any],
    *,
    concurrency: int | None = None,
) -> list[Any]:
    limit = concurrency or ASYNC_IO_CONCURRENCY
    sem = asyncio.Semaphore(limit)
    loop = asyncio.get_running_loop()

    async def _one(path: str) -> Any:
        async with sem:
            return await loop.run_in_executor(None, prep_func, path)

    results = await asyncio.gather(*[_one(p) for p in paths], return_exceptions=True)
    out: list[Any] = []
    for item in results:
        if isinstance(item, BaseException):
            logger.debug("Async prep error: %s", item)
            out.append(item)
        else:
            out.append(item)
    return out


def run_async_prep_batch(
    paths: list[str],
    prep_func: Callable[[str], Any],
    *,
    concurrency: int | None = None,
) -> list[Any]:
    return asyncio.run(prep_paths_async(paths, prep_func, concurrency=concurrency))


class AsyncPrepLoopRunner:
    """Background asyncio loop scheduling prep on a thread pool."""

    def __init__(self, prep_func: Callable[[str], Any], *, concurrency: int | None = None):
        self._prep_func = prep_func
        self._concurrency = concurrency or ASYNC_IO_CONCURRENCY
        self._loop: asyncio.AbstractEventLoop | None = None
        self._thread: threading.Thread | None = None
        self._ready = threading.Event()
        self._futures: dict[str, asyncio.Future] = {}
        self._lock = threading.Lock()
        self._sem: asyncio.Semaphore | None = None

    def start(self) -> None:
        if self._thread and self._thread.is_alive():
            return
        self._ready.clear()
        self._thread = threading.Thread(target=self._thread_main, name="eml-async-prep", daemon=True)
        self._thread.start()
        if not self._ready.wait(timeout=30.0):
            raise RuntimeError("Async prep loop failed to start")

    def _thread_main(self) -> None:
        self._loop = asyncio.new_event_loop()
        asyncio.set_event_loop(self._loop)
        self._sem = asyncio.Semaphore(self._concurrency)
        self._ready.set()
        self._loop.run_forever()

    def submit(self, path: str) -> None:
        if not self._loop or not self._sem:
            return
        norm = os.path.normpath(os.path.abspath(path))

        async def _run() -> Any:
            assert self._sem is not None
            async with self._sem:
                return await self._loop.run_in_executor(None, self._prep_func, path)

        fut = asyncio.run_coroutine_threadsafe(_run(), self._loop)
        with self._lock:
            self._futures[norm] = fut

    def take(self, norm_path: str, *, timeout: float = 120.0) -> Any:
        import concurrent.futures

        with self._lock:
            fut = self._futures.pop(norm_path, None)
        if fut is None:
            return None
        try:
            return fut.result(timeout=timeout)
        except (asyncio.TimeoutError, concurrent.futures.TimeoutError):
            return None

    def shutdown(self) -> None:
        if self._loop and self._loop.is_running():
            self._loop.call_soon_threadsafe(self._loop.stop)
        if self._thread:
            self._thread.join(timeout=5.0)
        self._loop = None
        self._thread = None
        with self._lock:
            self._futures.clear()
