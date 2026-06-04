"""
Adaptive pacing for Outlook COM import (reduces RPC storms on failure).
"""

from __future__ import annotations

import os
import time
from collections import deque


def _env_float(name: str, default: float) -> float:
    try:
        raw = os.environ.get(name, "").strip()
        return float(raw) if raw else default
    except ValueError:
        return default


class AdaptiveRateLimiter:
    """Adjusts per-second pacing based on recent import success/failure."""

    def __init__(
        self,
        initial_rate: float = 10.0,
        max_rate: float = 100.0,
        min_rate: float = 1.0,
    ):
        self.current_rate = max(min_rate, min(initial_rate, max_rate))
        self.max_rate = max_rate
        self.min_rate = min_rate
        self._request_times: deque[float] = deque(maxlen=200)
        self._failure_count = 0
        self._success_count = 0

    def wait_if_needed(self) -> None:
        if not self._request_times:
            self._request_times.append(time.time())
            return
        if len(self._request_times) >= int(max(1, self.current_rate)):
            oldest = self._request_times[0]
            elapsed = time.time() - oldest
            if elapsed < 1.0:
                time.sleep(1.0 - elapsed)
        self._request_times.append(time.time())

    def record_success(self) -> None:
        self._success_count += 1
        if self._success_count >= 10 and self.current_rate < self.max_rate:
            self.current_rate = min(self.max_rate, self.current_rate * 1.1)
            self._success_count = 0

    def record_failure(self) -> None:
        self._failure_count += 1
        if self._failure_count >= 3:
            self.current_rate = max(self.min_rate, self.current_rate * 0.5)
            self._failure_count = 0
            time.sleep(min(2.0, _env_float("EML2PST_ADAPTIVE_BACKOFF", 2.0)))

    def get_stats(self) -> dict:
        return {
            "current_rate": round(self.current_rate, 2),
            "queue_length": len(self._request_times),
        }


ADAPTIVE_RATE_ENABLED = os.environ.get("EML2PST_ADAPTIVE_RATE", "").strip().lower() in (
    "1",
    "true",
    "yes",
)
