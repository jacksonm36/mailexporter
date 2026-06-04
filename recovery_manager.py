"""
JSON export checkpoints (complements export_results.csv and dedup_state.sqlite3).
"""

from __future__ import annotations

import glob
import json
import os
import time
from typing import Any


class RecoveryManager:
    def __init__(self, checkpoint_dir: str, *, keep_last: int = 10):
        self.checkpoint_dir = checkpoint_dir
        self.keep_last = max(1, keep_last)
        os.makedirs(self.checkpoint_dir, exist_ok=True)

    def save_checkpoint(self, state: dict[str, Any], checkpoint_id: str | None = None) -> str:
        if checkpoint_id is None:
            checkpoint_id = time.strftime("%Y%m%d_%H%M%S")
        path = os.path.join(self.checkpoint_dir, f"export_{checkpoint_id}.json")
        payload = {
            "timestamp": time.time(),
            "version": state.get("app_version", "1.0.2"),
            "processed_count": state.get("processed_count", 0),
            "converted": state.get("converted", 0),
            "skipped": state.get("skipped", 0),
            "errors": state.get("errors", 0),
            "last_file": state.get("last_file", ""),
            "cancelled": state.get("cancelled", False),
        }
        tmp = path + ".tmp"
        with open(tmp, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)
        os.replace(tmp, path)
        self._rotate_checkpoints()
        return checkpoint_id

    def load_latest_checkpoint(self) -> dict[str, Any] | None:
        paths = sorted(
            glob.glob(os.path.join(self.checkpoint_dir, "export_*.json")),
            key=os.path.getmtime,
        )
        if not paths:
            return None
        with open(paths[-1], encoding="utf-8") as handle:
            return json.load(handle)

    def _rotate_checkpoints(self) -> None:
        paths = sorted(
            glob.glob(os.path.join(self.checkpoint_dir, "export_*.json")),
            key=os.path.getmtime,
        )
        for old in paths[: -self.keep_last]:
            try:
                os.remove(old)
            except OSError:
                pass
