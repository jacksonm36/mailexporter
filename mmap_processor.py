"""
Memory-mapped I/O for large .eml files (reduces peak RAM vs read-all-then-parse).

Used by eml_to_pst_converter.py and mail_validation.py when file size >= threshold.
"""

from __future__ import annotations

import hashlib
import mmap
import os
import re
from email import policy
from email.parser import BytesParser

# Files at or above this size use mmap helpers (override: EML2PST_MMAP_THRESHOLD_MB).
def _env_int_mb(name: str, default_mb: int) -> int:
    raw = os.environ.get(name, "").strip()
    if not raw:
        return default_mb * 1024 * 1024
    try:
        return max(0, int(float(raw) * 1024 * 1024))
    except ValueError:
        return default_mb * 1024 * 1024


MMAP_READ_THRESHOLD = _env_int_mb("EML2PST_MMAP_THRESHOLD_MB", 4)
IO_CHUNK_SIZE = 65536
_RFC822_CRLF_NORM_RE = re.compile(rb"(?<!\r)\n|\r(?!\n)")


def should_use_mmap(file_size: int) -> bool:
    return MMAP_READ_THRESHOLD > 0 and file_size >= MMAP_READ_THRESHOLD


def file_size_or_zero(path: str) -> int:
    try:
        return os.path.getsize(path)
    except OSError:
        return 0


def read_file_bytes(path: str, *, max_bytes: int | None = None) -> bytes | None:
    """Read entire file; use mmap for large files (still returns bytes for callers that need them)."""
    try:
        size = os.path.getsize(path)
    except OSError:
        return None
    if max_bytes is not None and size > max_bytes:
        return None
    if size == 0:
        return b""
    if not should_use_mmap(size):
        try:
            with open(path, "rb") as handle:
                return handle.read()
        except OSError:
            return None
    try:
        with open(path, "rb") as handle:
            with mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                return mm[:]
    except OSError:
        return None


def hash_file_sha256(path: str) -> str | None:
    """Streaming SHA-256 via mmap chunks (no full-file bytes allocation)."""
    try:
        size = os.path.getsize(path)
    except OSError:
        return None
    digest = hashlib.sha256()
    try:
        with open(path, "rb") as handle:
            if size == 0:
                return digest.hexdigest()
            with mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                for offset in range(0, len(mm), IO_CHUNK_SIZE):
                    digest.update(mm[offset : offset + IO_CHUNK_SIZE])
        return digest.hexdigest()
    except OSError:
        return None


def read_head_tail_samples(
    path: str, size: int, head_len: int = 65536, tail_len: int = 65536
) -> tuple[bytes, bytes] | None:
    """Read head and optional tail through one mmap mapping."""
    if size <= 0:
        return b"", b""
    head_len = min(head_len, size)
    tail_len = min(tail_len, max(0, size - head_len))
    try:
        with open(path, "rb") as handle:
            with mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                head = bytes(mm[:head_len])
                if tail_len <= 0:
                    return head, b""
                start = size - tail_len
                if start < head_len:
                    return head, b""
                tail = bytes(mm[start : start + tail_len])
                return head, tail
    except OSError:
        return None


def parse_message_from_path(path: str):
    """Parse RFC822 message; mmap-backed for large files."""
    try:
        size = os.path.getsize(path)
    except OSError:
        return None
    parser = BytesParser(policy=policy.default)
    try:
        with open(path, "rb") as handle:
            if size == 0:
                return parser.parsebytes(b"")
            if should_use_mmap(size):
                with mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                    return parser.parsebytes(mm)
            return parser.parse(handle)
    except OSError:
        return None
    except Exception:
        return None


def copy_file_mmap(src: str, dest: str) -> None:
    """Copy file to dest using chunked mmap reads."""
    with open(src, "rb") as fin, open(dest, "wb") as fout:
        size = os.path.getsize(src)
        if size == 0:
            return
        with mmap.mmap(fin.fileno(), 0, access=mmap.ACCESS_READ) as mm:
            for offset in range(0, len(mm), IO_CHUNK_SIZE):
                fout.write(mm[offset : offset + IO_CHUNK_SIZE])
        fout.flush()
        os.fsync(fout.fileno())


def write_crlf_normalized_file(src: str, dest: str) -> None:
    """
    Stream CRLF normalization to dest without loading the whole source into RAM.
    """
    with open(src, "rb") as fin, open(dest, "wb") as fout:
        size = os.path.getsize(src)
        if size == 0:
            return
        with mmap.mmap(fin.fileno(), 0, access=mmap.ACCESS_READ) as mm:
            if not should_use_mmap(size):
                fout.write(_RFC822_CRLF_NORM_RE.sub(b"\r\n", mm[:]))
                fout.flush()
                os.fsync(fout.fileno())
                return
            i = 0
            n = len(mm)
            while i < n:
                b = mm[i]
                if b == 0x0D:
                    if i + 1 < n and mm[i + 1] == 0x0A:
                        fout.write(b"\r\n")
                        i += 2
                    else:
                        fout.write(b"\r\n")
                        i += 1
                elif b == 0x0A:
                    fout.write(b"\r\n")
                    i += 1
                else:
                    j = i
                    while j < n and mm[j] not in (0x0D, 0x0A):
                        j += 1
                    fout.write(mm[i:j])
                    i = j
        fout.flush()
        os.fsync(fout.fileno())


def needs_lf_to_crlf_conversion(path: str, sample: int = 8192) -> bool:
    """True when file appears LF-only (needs CRLF staging for Outlook)."""
    try:
        with open(path, "rb") as handle:
            if should_use_mmap(file_size_or_zero(path)):
                with mmap.mmap(handle.fileno(), 0, access=mmap.ACCESS_READ) as mm:
                    chunk = mm[: min(len(mm), sample)]
            else:
                chunk = handle.read(sample)
    except OSError:
        return False
    if not chunk:
        return False
    return b"\n" in chunk and b"\r\n" not in chunk
