"""
Optional dedup strategies beyond content fingerprint / full SHA-256.
"""

from __future__ import annotations

import hashlib
import re

_MESSAGE_ID_RE = re.compile(br"^Message-ID:\s*<([^>\r\n]+)>", re.MULTILINE | re.IGNORECASE)
_IN_REPLY_RE = re.compile(br"^In-Reply-To:\s*<([^>\r\n]+)>", re.MULTILINE | re.IGNORECASE)
_SUBJECT_RE = re.compile(br"^Subject:\s*(.+)$", re.MULTILINE | re.IGNORECASE)
_PREFIX_RE = re.compile(r"^(re|fwd|fw)\s*:\s*", re.IGNORECASE)

STRATEGIES = frozenset({"content_hash", "message_id", "fuzzy_subject", "thread"})


def normalize_strategy(name: str) -> str:
    key = (name or "content_hash").strip().lower()
    return key if key in STRATEGIES else "content_hash"


def dedup_key_from_header_sample(data: bytes, strategy: str, file_size: int) -> str | None:
    """Return a dedup key from the first ~64 KiB of the message, or None to use fingerprint."""
    strategy = normalize_strategy(strategy)
    if strategy == "content_hash":
        return None
    sample = data[:65536]
    if strategy in ("message_id", "thread"):
        match = _MESSAGE_ID_RE.search(sample)
        if match:
            mid = match.group(1).decode("utf-8", errors="replace").strip()
            return f"mid:{mid}"
        match = _IN_REPLY_RE.search(sample)
        if match:
            irt = match.group(1).decode("utf-8", errors="replace").strip()
            return f"thread:{irt}"
        return None
    if strategy == "fuzzy_subject":
        match = _SUBJECT_RE.search(sample)
        if not match:
            return None
        subject = match.group(1).decode("utf-8", errors="replace").strip()
        subject = _PREFIX_RE.sub("", subject).strip().lower()
        if not subject:
            return None
        digest = hashlib.sha256(subject.encode("utf-8", errors="replace")).hexdigest()[:40]
        return f"subj:{digest}|sz:{file_size}"
    return None
