"""
LRU caches for repeated .eml metadata reads during large exports.
"""

from __future__ import annotations

import os
import time
from functools import lru_cache
from typing import Any


def file_cache_key(file_path: str, size: int | None = None) -> tuple[str, int, int]:
    """Stable cache key: path + size + mtime_ns."""
    path = os.path.normpath(os.path.abspath(file_path))
    try:
        st = os.stat(path)
        sz = int(size if size is not None else st.st_size)
        mtime = int(getattr(st, "st_mtime_ns", int(st.st_mtime * 1e9)))
    except OSError:
        sz = int(size or 0)
        mtime = 0
    return path, sz, mtime


class EmailMetadataCache:
    """TTL + LRU metadata cache for parallel prep / repeated header reads."""

    def __init__(self, max_size: int = 1000, ttl_seconds: int = 3600):
        self.max_size = max(1, max_size)
        self.ttl = max(60, ttl_seconds)
        self._cache: dict[str, dict[str, Any]] = {}
        self._timestamps: dict[str, float] = {}

    def get(self, cache_id: str) -> dict[str, Any] | None:
        if cache_id not in self._cache:
            return None
        if time.time() - self._timestamps[cache_id] >= self.ttl:
            self._cache.pop(cache_id, None)
            self._timestamps.pop(cache_id, None)
            return None
        return self._cache[cache_id]

    def set(self, cache_id: str, data: dict[str, Any]) -> None:
        if len(self._cache) >= self.max_size:
            oldest = min(self._timestamps, key=self._timestamps.get)
            self._cache.pop(oldest, None)
            self._timestamps.pop(oldest, None)
        self._cache[cache_id] = data
        self._timestamps[cache_id] = time.time()


@lru_cache(maxsize=4096)
def parse_headers_cached(path: str, size: int, mtime_ns: int) -> tuple[str, str, str]:
    """
    Read only the RFC822 header block (stops at first blank line).
    Returns (subject, from_addr, date_header).
    """
    subject = from_addr = date_header = ""
    try:
        with open(path, "rb") as handle:
            for raw in handle:
                line = raw.decode("utf-8", errors="replace")
                if line in ("\n", "\r\n"):
                    break
                if ":" not in line:
                    continue
                key, _, value = line.partition(":")
                key_l = key.strip().lower()
                value = value.strip()
                if key_l == "subject" and not subject:
                    subject = value
                elif key_l == "from" and not from_addr:
                    from_addr = value
                elif key_l == "date" and not date_header:
                    date_header = value
    except OSError:
        pass
    return subject, from_addr, date_header
