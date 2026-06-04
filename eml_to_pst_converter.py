"""
Mail Exporter — EML/EMLX to Outlook PST or Exchange mailbox.
Portable Windows GUI; bundles Python/pywin32 when built as a single .exe.
Requires Microsoft Outlook (same bitness as this app) on the PC.
"""

from __future__ import annotations

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import os
import hashlib
from email import policy
from email.parser import BytesParser
from datetime import datetime
import threading
import subprocess
import sys
import time
import re
import gc
import logging
import tempfile
import uuid
import atexit
import struct
import shutil
import sqlite3
from pathlib import Path
from email.header import decode_header
from email.utils import parseaddr, parsedate_to_datetime

from env_config import load_mail_exporter_env

_ENV_FILE_LOADED = load_mail_exporter_env(log=False)

from i18n import (
    LANG_EN,
    LANG_HU,
    LANGUAGE_NAMES,
    detect_default_lang,
    lang_from_display,
    t,
)
APP_VERSION = "1.0.2"

from path_security import PathSecurityError, PathValidator
from cache_manager import file_cache_key, parse_headers_cached
from email_filter import EmailFilter
from rate_limiter import ADAPTIVE_RATE_ENABLED, AdaptiveRateLimiter
from recovery_manager import RecoveryManager
from smart_dedup import dedup_key_from_header_sample, normalize_strategy
from parallel_processor import (
    PARALLEL_PREP_WORKERS,
    create_eml_prep_prefetcher,
    EmlPrepResult,
)
from com_pipeline import (
    ComImportPipeline,
    ImportResult,
    com_pipeline_should_run,
)
from mmap_processor import (
    MMAP_READ_THRESHOLD,
    read_head_tail_samples,
    copy_file_mmap,
    file_size_or_zero,
    hash_file_sha256,
    needs_lf_to_crlf_conversion,
    parse_message_from_path,
    read_file_bytes,
    read_head_tail_samples,
    should_use_mmap,
    write_crlf_normalized_file,
)
from mail_validation import (
    ISSUE_WRONG_FOLDER,
    ISSUE_WRONG_PST_STORE,
    append_folder_placement_issues,
    body_looks_like_html,
    eml_parse_to_summary,
    format_inspection_errors,
    has_import_errors,
    read_outlook_mail_snapshot,
    validate_import_against_eml,
)


def _is_frozen() -> bool:
    return bool(getattr(sys, "frozen", False))


def _default_wlm_folder() -> str:
    return os.path.join(
        os.environ.get("LOCALAPPDATA", ""),
        "Microsoft",
        "Windows Live Mail",
    )


def _portable_staging_base() -> str:
    """Writable staging directory — prefer LOCALAPPDATA over project dir or %TEMP%."""
    candidates: list[str] = []
    local = os.environ.get("LOCALAPPDATA", "").strip()
    if local:
        candidates.append(os.path.join(local, "MailExporter", "Staging"))
    if _is_frozen():
        candidates.append(
            os.path.join(os.path.dirname(os.path.abspath(sys.executable)), "MailExporter_Staging")
        )
    else:
        candidates.append(
            os.path.join(os.path.dirname(os.path.abspath(__file__)), "MailExporter_Staging")
        )
    for preferred in candidates:
        try:
            os.makedirs(preferred, exist_ok=True)
            test = os.path.join(preferred, ".write_test")
            with open(test, "w", encoding="utf-8") as handle:
                handle.write("ok")
            os.remove(test)
            return preferred
        except OSError:
            continue
    fallback = os.path.join(local or tempfile.gettempdir(), "MailExporter_Staging")
    os.makedirs(fallback, exist_ok=True)
    return fallback

# Configure logging (fixed logger name so frozen .exe matches dev runs)
_LOG_FORMAT = "%(asctime)s - %(levelname)s - %(message)s"
if not logging.getLogger().handlers:
    logging.basicConfig(level=logging.INFO, format=_LOG_FORMAT)
logger = logging.getLogger("MailExporter")
logger.setLevel(logging.INFO)
if _ENV_FILE_LOADED:
    logger.info("Loaded settings from %s", _ENV_FILE_LOADED)
elif os.environ.get("EML2PST_PARALLEL_WORKERS") or os.environ.get("EML2PST_COM_PIPELINE"):
    logger.info("Using embedded or system EML2PST_* environment settings")


class _ImmediateFileHandler(logging.Handler):
    """Write each log record and flush immediately (reliable tail on Windows)."""

    def __init__(self, path: str, *, append: bool):
        super().__init__(logging.INFO)
        self.path = os.path.abspath(path)
        parent = os.path.dirname(self.path)
        if parent:
            os.makedirs(parent, exist_ok=True)
        mode = "a" if append else "w"
        self._stream = open(self.path, mode, encoding="utf-8", newline="\n")
        if not append:
            self._stream.write(
                "=== Mail Exporter export session "
                f"{datetime.now():%Y-%m-%d %H:%M:%S} ===\n"
            )
            self._stream.flush()
        self.setFormatter(logging.Formatter(_LOG_FORMAT))

    def emit(self, record: logging.LogRecord) -> None:
        try:
            self._stream.write(self.format(record) + "\n")
            self._stream.flush()
            try:
                os.fsync(self._stream.fileno())
            except OSError:
                pass
        except Exception:
            self.handleError(record)

    def flush(self) -> None:
        if self._stream and not self._stream.closed:
            self._stream.flush()

    def close(self) -> None:
        try:
            if self._stream and not self._stream.closed:
                self._stream.flush()
                self._stream.close()
        finally:
            super().close()


def _last_export_log_path() -> str:
    base = os.path.join(
        os.environ.get("LOCALAPPDATA", tempfile.gettempdir()),
        "MailExporter",
    )
    os.makedirs(base, exist_ok=True)
    return os.path.join(base, "last_export.log")

# Outlook constants
OL_MAIL_ITEM = 0
OL_DISCARD = 1
OL_STORE_UNICODE = 2  # Unicode PST format (Outlook 2003+)
OL_MAIL_CLASS = 43  # olMail — Items.Class for standard email messages
OL_SAVEAS_MSG = 3  # olMSG — Outlook .msg (OpenSharedItem-friendly)

# MAPI PT_SYSTIME properties Outlook uses for list columns (Received, Sent, sorting).
_OUTLOOK_MAPI_TIME_URLS = (
    "http://schemas.microsoft.com/mapi/proptag/0x0E060040",  # PR_MESSAGE_DELIVERY_TIME
    "http://schemas.microsoft.com/mapi/proptag/0x00390040",  # PR_CLIENT_SUBMIT_TIME
    "http://schemas.microsoft.com/mapi/proptag/0x00690040",  # PR_ORIGINAL_ARRIVAL_TIME
    "http://schemas.microsoft.com/mapi/proptag/0x30070040",  # PR_CREATION_TIME
    "http://schemas.microsoft.com/mapi/proptag/0x30080040",  # PR_LAST_MODIFICATION_TIME
)
_OUTLOOK_MAPI_MESSAGE_FLAGS = "http://schemas.microsoft.com/mapi/proptag/0x0E070003"
_OUTLOOK_MAPI_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001F"
_OUTLOOK_MAPI_SENDER_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0C1F001E"
_OUTLOOK_MAPI_SENDER_ADDRTYPE = "http://schemas.microsoft.com/mapi/proptag/0x0C1E001E"
_OUTLOOK_MAPI_SENT_REP_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0042001F"
_OUTLOOK_MAPI_SENT_REP_EMAIL = "http://schemas.microsoft.com/mapi/proptag/0x0065001F"
_MSGFLAG_UNSENT = 0x0008

# Folder names
INBOX_FOLDER_NAME = "Inbox"
OL_FOLDER_INBOX = 6
OL_FOLDER_SENT = 5
OL_FOLDER_DELETED = 3
OL_FOLDER_OUTBOX = 4
OL_FOLDER_DRAFTS = 16
OL_FOLDER_JUNK = 23
# Windows Live Mail / localized Outlook folder names -> olFolder* id
_STANDARD_FOLDER_BY_ALIAS: dict[str, int] = {
    "inbox": OL_FOLDER_INBOX,
    "beérkezett üzenetek": OL_FOLDER_INBOX,
    "sent items": OL_FOLDER_SENT,
    "sent messages": OL_FOLDER_SENT,
    "sent mail": OL_FOLDER_SENT,
    "sent": OL_FOLDER_SENT,
    "elküldött elemek": OL_FOLDER_SENT,
    "elküldött": OL_FOLDER_SENT,
    "deleted items": OL_FOLDER_DELETED,
    "deleted messages": OL_FOLDER_DELETED,
    "deleted": OL_FOLDER_DELETED,
    "trash": OL_FOLDER_DELETED,
    "lomtár": OL_FOLDER_DELETED,
    "törölt elemek": OL_FOLDER_DELETED,
    "törölt": OL_FOLDER_DELETED,
    "drafts": OL_FOLDER_DRAFTS,
    "piszkozatok": OL_FOLDER_DRAFTS,
    "outbox": OL_FOLDER_OUTBOX,
    "kimenő": OL_FOLDER_OUTBOX,
    "kimenő üzenetek": OL_FOLDER_OUTBOX,
    "junk e-mail": OL_FOLDER_JUNK,
    "junk email": OL_FOLDER_JUNK,
    "junk": OL_FOLDER_JUNK,
    "spam": OL_FOLDER_JUNK,
    "levélszemét": OL_FOLDER_JUNK,
}
# (olFolderId, name aliases for root scan, folder name when creating under PST root)
_STANDARD_PST_FOLDER_SPECS: dict[int, tuple[tuple[str, ...], str]] = {
    OL_FOLDER_INBOX: (("inbox", "beérkezett üzenetek"), INBOX_FOLDER_NAME),
    OL_FOLDER_SENT: (("sent items", "elküldött elemek", "sent"), "Sent Items"),
    OL_FOLDER_DELETED: (("deleted items", "törölt elemek"), "Deleted Items"),
    OL_FOLDER_OUTBOX: (("outbox", "kimenő", "kimenő üzenetek"), "Outbox"),
    OL_FOLDER_DRAFTS: (("drafts", "piszkozatok"), "Drafts"),
    OL_FOLDER_JUNK: (
        ("junk e-mail", "junk email", "levélszemét"),
        "Junk E-Mail",
    ),
}
# GetDefaultFolder(Inbox) on a damaged PST can return search folders — never import there.
_SEARCH_FOLDER_NAME_MARKERS = (
    "search folder",
    "spam search",
    "keresési mappa",
    "keresőmappa",
)


def standard_outlook_folder_id(folder_name: str) -> int | None:
    """Map a source folder label (WLM / localized Outlook) to olFolder* id."""
    return _STANDARD_FOLDER_BY_ALIAS.get((folder_name or "").lower().strip())


def find_standard_folder_in_parts(
    parts: list[str],
) -> tuple[int | None, int, list[str]]:
    """
    Locate the deepest standard Outlook folder segment in a relative path.

    WLM paths are usually Account\\Inbox\\... or Account\\Sent Items\\...;
    using the last match avoids mis-routing when a custom subfolder name
    appears earlier in the path.
    """
    if not parts:
        return None, -1, []
    match_idx = -1
    match_id: int | None = None
    for idx, segment in enumerate(parts):
        folder_id = standard_outlook_folder_id(segment)
        if folder_id is not None:
            match_idx = idx
            match_id = folder_id
    if match_idx < 0 or match_id is None:
        return None, -1, parts
    remainder = parts[:match_idx] + parts[match_idx + 1 :]
    return match_id, match_idx, remainder


def sent_state_for_folder_parts(parts: list[str]) -> bool | None:
    """
    Infer Sent flag from WLM/Outlook folder names in the source path.

    True = Sent Items; False = Drafts / Outbox; None = Inbox / Deleted / Junk / other.
    """
    folder_id, _, _ = find_standard_folder_in_parts(parts)
    if folder_id == OL_FOLDER_SENT:
        return True
    if folder_id in (OL_FOLDER_DRAFTS, OL_FOLDER_OUTBOX):
        return False
    return None

# Check for pywin32
OUTLOOK_AVAILABLE = False
WIN32COM = None
PYTHONCOM = None
PYWINTYPES = None
try:
    import win32com.client
    import pythoncom
    WIN32COM = win32com.client
    PYTHONCOM = pythoncom
    OUTLOOK_AVAILABLE = True
    try:
        import pywintypes as _pywintypes
        PYWINTYPES = _pywintypes
    except ImportError:
        pass
except ImportError:
    logger.warning("pywin32 not available - PST conversion will require installation")

def _env_int(name: str, default: int, *, lo: int = 1, hi: int = 500_000) -> int:
    try:
        v = int(os.environ.get(name, str(default)).strip())
        return max(lo, min(v, hi))
    except (TypeError, ValueError):
        return default


def _env_int_mb(name: str, default: int = 100, *, lo: int = 1, hi: int = 2048) -> int:
    try:
        v = int(os.environ.get(name, str(default)).strip())
        return max(lo, min(v, hi))
    except (TypeError, ValueError):
        return default


def _env_float(name: str, default: float, *, lo: float = 0.0, hi: float = 600.0) -> float:
    """Parse optional float env; clamp to [lo, hi] to avoid abuse or typos."""
    try:
        raw = os.environ.get(name, "")
        if not str(raw).strip():
            return default
        v = float(str(raw).strip())
        return max(lo, min(v, hi))
    except (TypeError, ValueError):
        return default


# Maximum files per export run (override: EML2PST_MAX_FILES)
MAX_FILES = _env_int("EML2PST_MAX_FILES", 1_000_000)
# Treeview preview cap — full list kept in memory for conversion; UI shows a sample only.
UI_PREVIEW_FILE_LIMIT = _env_int("EML2PST_UI_PREVIEW", 500, lo=50, hi=5000)
SCAN_PROGRESS_EVERY = _env_int("EML2PST_SCAN_PROGRESS_EVERY", 1000, lo=100, hi=10000)
CONVERSION_UI_EVERY = _env_int("EML2PST_UI_UPDATE_EVERY", 10, lo=1, hi=500)
UI_THROTTLE_SEC = _env_float("EML2PST_UI_THROTTLE_SEC", 0.25, lo=0.05, hi=2.0)
TREE_PREVIEW_CHUNK = _env_int("EML2PST_TREE_CHUNK", 50, lo=10, hi=500)
GC_EVERY_N_FILES = _env_int("EML2PST_GC_EVERY", 500, lo=0, hi=5000)
COM_RETRY_ATTEMPTS = _env_int("EML2PST_COM_RETRIES", 3, lo=1, hi=5)
CSV_FLUSH_EVERY = _env_int("EML2PST_CSV_FLUSH_EVERY", 50, lo=1, hi=1000)
CSV_UTF8_BOM = os.environ.get("EML2PST_CSV_BOM", "1").strip().lower() not in (
    "0",
    "false",
    "no",
)
SLOW_FILE_WARN_SEC = _env_float("EML2PST_SLOW_FILE_SEC", 120.0, lo=0.0, hi=3600.0)
PARALLEL_PARSE = os.environ.get("EML2PST_PARALLEL_PARSE", "").strip().lower() in (
    "1",
    "true",
    "yes",
)
USE_COM_PIPELINE = com_pipeline_should_run()
try:
    from async_processor import ASYNC_IO_ENABLED
except ImportError:
    ASYNC_IO_ENABLED = False
DEDUP_STRATEGY = normalize_strategy(os.environ.get("EML2PST_DEDUP_STRATEGY", "content_hash"))
CHECKPOINT_EVERY = _env_int("EML2PST_CHECKPOINT_EVERY", 0, lo=0, hi=50_000)


def _build_email_filter_from_env() -> EmailFilter | None:
    """Optional export filters from environment (all off by default)."""
    filt = EmailFilter()
    max_mb = os.environ.get("EML2PST_FILTER_MAX_MB", "").strip()
    if max_mb:
        try:
            filt.add_size_filter(int(float(max_mb) * 1024 * 1024))
        except ValueError:
            pass
    att = os.environ.get("EML2PST_FILTER_ATTACHMENTS", "").strip().lower()
    if att == "only":
        filt.add_attachment_filter(has_attachments=True)
    elif att == "none":
        filt.add_attachment_filter(has_attachments=False)
    return filt if filt.enabled else None
COM_RPC_COOLDOWN = _env_float("EML2PST_RPC_COOLDOWN", 5.0, lo=1.0, hi=120.0)
STAGING_WRITE_SETTLE = _env_float("EML2PST_STAGING_SETTLE", 0.35, lo=0.05, hi=5.0)
COM_PACE_SEC = _env_float("EML2PST_COM_PACE", 0.02, lo=0.0, hi=2.0)
MAX_ATTACHMENT_BYTES = _env_int_mb("EML2PST_MAX_ATTACHMENT_MB", 25) * 1024 * 1024
# OpenSharedItem often fails on large multipart/attachment .eml — skip native above this size.
NATIVE_IMPORT_MAX_BYTES = _env_int_mb("EML2PST_NATIVE_MAX_MB", 384) * 1024 * 1024
NATIVE_IMPORT_SNIFF_BYTES = _env_int("EML2PST_NATIVE_SNIFF", 262144, lo=8192, hi=1_048_576)
FULL_DEDUP_HASH = os.environ.get("EML2PST_FULL_DEDUP", "").strip().lower() in ("1", "true", "yes")
VALIDATE_IMPORT = os.environ.get("EML2PST_VALIDATE_IMPORT", "1").strip().lower() not in (
    "0",
    "false",
    "no",
    "off",
)
OUTLOOK_CRASH_WAIT_MAX = _env_int("EML2PST_OUTLOOK_WAIT_MAX", 3600, lo=60, hi=86_400)
OUTLOOK_CRASH_WAIT_POLL = _env_float("EML2PST_OUTLOOK_WAIT_POLL", 5.0, lo=2.0, hi=60.0)
PREFLIGHT_SIZE_SAMPLE = _env_int("EML2PST_PREFLIGHT_SAMPLE", 200, lo=50, hi=5000)
DEDUP_BACKEND_MODE = os.environ.get("EML2PST_DEDUP_BACKEND", "auto").strip().lower()
DEDUP_SQLITE_THRESHOLD = _env_int("EML2PST_DEDUP_SQLITE_THRESHOLD", 200_000, lo=10_000, hi=500_000)
DEDUP_SQLITE_COMMIT_EVERY = _env_int("EML2PST_DEDUP_SQLITE_COMMIT_EVERY", 500, lo=50, hi=5000)
DEDUP_DB_FILENAME = "dedup_state.sqlite3"
# Date: header in the first 64KB of each message (dedup keys must not ignore send time).
_DATE_HEADER_BYTES_RE = re.compile(br"^Date:\s*(.+)$", re.MULTILINE | re.IGNORECASE)
# Single-pass CRLF normalization (faster than split/join on large .eml bodies).
_RFC822_CRLF_NORM_RE = re.compile(rb"(?<!\r)\n|\r(?!\n)")
# Attachment sniff: HTML multipart uses Content-Disposition: inline without filename — not a file.
_ATTACHMENT_DISPOSITION_SNIFF_RE = re.compile(
    br"content-disposition\s*:\s*attachment\b",
    re.IGNORECASE,
)
_INLINE_NAMED_DISPOSITION_SNIFF_RE = re.compile(
    br"content-disposition\s*:\s*inline\b[^\r\n]{0,240}?filename\s*=",
    re.IGNORECASE,
)
_BINARY_PART_CONTENT_TYPE_SNIFF_RE = re.compile(
    rb"content-type\s*:\s*(?:application/(?:pdf|zip|x-zip|octet-stream)|image/(?:png|jpeg|jpg|gif))\b",
    re.IGNORECASE,
)
# HTML fragments without <html>/<body> (common in WLM / invoice templates).
_HTML_TAG_FRAGMENT_RE = re.compile(r"</?[a-zA-Z][^>]{0,240}>", re.I)
_BINARY_BODY_SIGNATURES = (
    b"%PDF-",
    b"\x89PNG\r\n\x1a\n",
    b"PK\x03\x04",
    b"GIF87a",
    b"GIF89a",
    b"\x1f\x8b\x08",
    b"Rar!",
    b"PK\x05\x06",
)
# Content-Type -> default extension when MIME parts omit filename=
_MIME_EXTENSION_MAP = {
    "application/pdf": ".pdf",
    "application/zip": ".zip",
    "application/x-zip-compressed": ".zip",
    "application/x-zip": ".zip",
    "application/x-rar-compressed": ".rar",
    "application/vnd.rar": ".rar",
    "application/x-7z-compressed": ".7z",
    "application/gzip": ".gz",
    "application/x-gzip": ".gz",
    "application/msword": ".doc",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.ms-excel": ".xls",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
    "application/vnd.ms-powerpoint": ".ppt",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation": ".pptx",
    "application/vnd.ms-outlook": ".msg",
    "application/octet-stream": ".bin",
    "message/rfc822": ".eml",
    "image/jpeg": ".jpg",
    "image/jpg": ".jpg",
    "image/png": ".png",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
    "image/webp": ".webp",
    "image/tiff": ".tif",
    "text/csv": ".csv",
    "text/plain": ".txt",
    "application/xml": ".xml",
    "text/xml": ".xml",
}
# New PST only: import in chunk PSTs of N emails, then merge into the final PST (0 = off).
PST_CHUNK_SIZE = _env_int("EML2PST_PST_CHUNK_SIZE", 100, lo=0, hi=10_000)


def _outlook_error_text(exc) -> str:
    if exc is None:
        return ""
    if isinstance(exc, str):
        return exc
    text = str(exc)
    if getattr(exc, "args", None):
        text = " ".join(str(a) for a in exc.args)
    return text


def _is_outlook_rpc_error(exc) -> bool:
    """True when Outlook COM/RPC is overloaded or the session died."""
    text = _outlook_error_text(exc)
    if not text:
        return False
    markers = (
        "RPC server is unavailable",
        "2147023174",
        "2147944122",
        "800706ba",
        "80080005",
        "80040154",
        "800706be",
        "Call was rejected by callee",
        "Server execution failed",
        "The message filter indicated",
        "Operation unavailable",
        "application is not running",
        "Invalid class string",
        "CoCreateInstance",
        "remote procedure call failed",
        "disconnected",
        "connection is invalid",
        "Outlook is not running",
    )
    return any(m.lower() in text.lower() for m in markers)


def _is_outlook_unavailable_error(exc) -> bool:
    """True when Outlook crashed, hung, or COM must be re-established."""
    return _is_outlook_rpc_error(exc)


def _outlook_process_running() -> bool:
    """True when OUTLOOK.EXE is present (Windows)."""
    if sys.platform != "win32":
        return True
    try:
        out = subprocess.run(
            [
                "tasklist",
                "/FI",
                "IMAGENAME eq OUTLOOK.EXE",
                "/NH",
            ],
            capture_output=True,
            text=True,
            timeout=15,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return "OUTLOOK.EXE" in (out.stdout or "").upper()
    except (OSError, subprocess.SubprocessError):
        return True


def _is_outlook_path_open_error(exc) -> bool:
    """True when OpenSharedItem failed because Outlook could not read the path/URI."""
    text = _outlook_error_text(exc).lower()
    if not text:
        return False
    markers = (
        "couldn't find",
        "cannot find this file",
        "2147024894",
        "invalid path or url",
        "moved or deleted",
    )
    return any(m in text for m in markers)


def _com_retry(label: str, operation):
    """Retry transient Outlook COM failures (RPC/server busy) on large batches."""
    last_err = None
    for attempt in range(COM_RETRY_ATTEMPTS):
        try:
            return operation()
        except Exception as e:
            last_err = e
            if attempt + 1 >= COM_RETRY_ATTEMPTS:
                break
            if _is_outlook_path_open_error(e):
                break
            if _is_outlook_rpc_error(e):
                time.sleep(COM_RPC_COOLDOWN * (attempt + 1))
            else:
                time.sleep(0.5 * (attempt + 1))
            logger.debug("%s retry %d/%d: %s", label, attempt + 2, COM_RETRY_ATTEMPTS, e)
    raise last_err


def _pattern_to_suffixes(pattern_input: str) -> set[str]:
    suffixes: set[str] = set()
    for part in pattern_input.lower().split(";"):
        part = part.strip()
        if part.startswith("*") and len(part) > 1:
            suffixes.add(part[1:])
        elif part.startswith("*."):
            suffixes.add(part[1:])
    return suffixes or {".eml"}


# Single-message read cap (mitigates huge / malicious .eml memory use)
MAX_EML_FILE_BYTES = _env_int_mb("EML2PST_MAX_FILE_MB", 100) * 1024 * 1024


def _safe_staging_subdir(raw: str | None, default: str = "EML2PST_Staging") -> str:
    """
    Single folder name only — env must not inject path traversal or drive letters.
    """
    s = (raw or default).strip()
    if len(s) > 80:
        s = s[:80]
    if not s or s in (".", ".."):
        return default
    for bad in (os.sep, ".."):
        if bad in s:
            return default
    if os.altsep and os.altsep in s:
        return default
    if ":" in s:  # Windows streams / drive-relative tricks
        return default
    if not re.match(r"^[\w.\- ]+$", s):
        return default
    return s


# After native import, Outlook may still read the staging .eml from disk asynchronously.
# Deleting it immediately causes "file may have been moved or deleted" (localized in HU/EN).
# Staging files are kept until the batch ends; optional extra delay after COM release.
_NATIVE_POST_BATCH_DELAY = _env_float("EML2PST_POST_BATCH_DELAY", 3.0)
# Not a dot-folder: Outlook often fails to open file:/// URLs under ".something" paths.
NATIVE_STAGING_SUBDIR = _safe_staging_subdir(os.environ.get("EML2PST_STAGING_SUBDIR"))

CSV_HEADERS = ("timestamp", "file_path", "status", "detail", "duration_sec", "target_folder")


def app_bitness() -> int:
    return struct.calcsize("P") * 8


_IMAGE_FILE_MACHINE_I386 = 0x014C
_IMAGE_FILE_MACHINE_AMD64 = 0x8664


def _parse_bitness_registry_value(raw) -> int | None:
    text = str(raw or "").strip().lower()
    if not text:
        return None
    if "64" in text or "x64" in text or "amd64" in text:
        return 64
    if "32" in text or "x86" in text:
        return 32
    return None


def find_outlook_exe_path() -> str | None:
    """Resolve installed OUTLOOK.EXE from App Paths (most reliable)."""
    try:
        import winreg
    except ImportError:
        return None
    for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
        for sub in (
            r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
            r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE",
        ):
            try:
                with winreg.OpenKey(hive, sub) as key:
                    path, _ = winreg.QueryValueEx(key, "")
                path = os.path.expandvars(str(path).strip().strip('"'))
                if path and os.path.isfile(path):
                    return os.path.normpath(path)
            except OSError:
                continue
    return None


def pe_image_bitness(exe_path: str) -> int | None:
    """Return 32 or 64 from the PE machine field of an executable."""
    try:
        with open(exe_path, "rb") as handle:
            handle.seek(0x3C)
            pe_offset = struct.unpack("<I", handle.read(4))[0]
            handle.seek(pe_offset + 4)
            machine = struct.unpack("<H", handle.read(2))[0]
        if machine == _IMAGE_FILE_MACHINE_AMD64:
            return 64
        if machine == _IMAGE_FILE_MACHINE_I386:
            return 32
    except OSError:
        pass
    return None


def detect_outlook_bitness() -> int | None:
    """Return 32 or 64 for the installed Outlook.exe, else None."""
    outlook_exe = find_outlook_exe_path()
    if outlook_exe:
        bits = pe_image_bitness(outlook_exe)
        if bits:
            return bits
    try:
        import winreg
    except ImportError:
        return None
    for ver in ("16.0", "15.0", "14.0", "12.0"):
        for hive in (winreg.HKEY_LOCAL_MACHINE, winreg.HKEY_CURRENT_USER):
            for suffix in ("", r"\WOW6432Node"):
                key_path = rf"SOFTWARE{suffix}\Microsoft\Office\{ver}\Outlook"
                try:
                    with winreg.OpenKey(hive, key_path) as key:
                        for value_name in ("OutlookBitness", "Bitness"):
                            try:
                                raw, _ = winreg.QueryValueEx(key, value_name)
                            except OSError:
                                continue
                            bits = _parse_bitness_registry_value(raw)
                            if bits:
                                return bits
                except OSError:
                    continue
    return None


def probe_outlook_com() -> tuple[bool, str]:
    """Try to start Outlook via COM; returns (ok, error_detail)."""
    if not OUTLOOK_AVAILABLE or not WIN32COM or not PYTHONCOM:
        return False, "Outlook automation libraries not available"
    com_initialized = False
    try:
        try:
            PYTHONCOM.CoInitialize()
            com_initialized = True
        except Exception:
            pass
        app = WIN32COM.Dispatch("Outlook.Application")
        _ = app.Version
        return True, ""
    except Exception as e:
        return False, str(e)
    finally:
        if com_initialized:
            try:
                PYTHONCOM.CoUninitialize()
            except Exception:
                pass


def recommended_exporter_exe_name(outlook_bits: int | None = None) -> str:
    """Portable exe file name that matches Outlook's bitness."""
    bits = outlook_bits if outlook_bits in (32, 64) else app_bitness()
    return f"MailExporter_x{bits}.exe"


def csv_sanitize(value: str) -> str:
    """Prevent Excel/CSV formula injection when opening export_results.csv."""
    if value is None:
        return ""
    s = str(value)
    if s and s[0] in ("=", "+", "-", "@", "\t", "\r"):
        return "'" + s
    return s


def log_sanitize(value) -> str:
    """Single-line, length-limited text safe for log files."""
    s = str(value or "").replace("\r", " ").replace("\n", " ")
    return s[:500] if len(s) > 500 else s


def format_export_eta(elapsed_sec: float, done: int, total: int) -> str:
    """Human-readable ETA from average rate (empty until at least two items finished)."""
    if done < 2 or total <= done or elapsed_sec <= 0:
        return ""
    remaining_sec = (total - done) * (elapsed_sec / done)
    if remaining_sec < 60:
        return "<1 min"
    minutes = int(remaining_sec // 60)
    if minutes < 90:
        return f"~{minutes} min"
    hours, mins = divmod(minutes, 60)
    if hours < 48:
        return f"~{hours}h {mins}m" if mins else f"~{hours}h"
    days, hours = divmod(hours, 24)
    return f"~{days}d {hours}h" if hours else f"~{days}d"


def release_com_object(obj) -> None:
    """Drop a COM reference so Outlook can reclaim message objects during long runs."""
    if obj is None:
        return
    try:
        del obj
    except Exception:
        pass


def safe_outlook_folder_name(name: str) -> str:
    s = (name or "Folder").strip() or "Folder"
    for ch in '\\/:*?"<>|':
        s = s.replace(ch, "_")
    return s[:255]


def relative_folder_parts(source_root: str, file_path: str) -> list[str]:
    """Subfolder path components relative to the scanned source root."""
    if not source_root:
        return []
    src = os.path.normpath(os.path.abspath(source_root))
    file_dir = os.path.normpath(os.path.dirname(os.path.abspath(file_path)))
    try:
        rel = os.path.relpath(file_dir, src)
    except ValueError:
        return []
    if rel in (".", ""):
        return []
    parts = []
    for part in rel.split(os.sep):
        if part in (".", "..") or not part:
            continue
        parts.append(safe_outlook_folder_name(part))
    return parts


def resolve_source_root(folder_path: str, file_paths: list[str]) -> str:
    """Best root for subfolder mapping: user folder or common ancestor."""
    if folder_path:
        root = os.path.normpath(os.path.abspath(folder_path.strip()))
        if os.path.isdir(root):
            return root
    if not file_paths:
        return ""
    dirs = [os.path.dirname(os.path.abspath(p)) for p in file_paths]
    try:
        return os.path.commonpath(dirs)
    except ValueError:
        return dirs[0] if dirs else ""


def export_output_dir(source_root: str, pst_path: str, export_mode: str) -> str:
    """Directory for export_results.csv, export.log, and dedup DB."""
    if export_mode != "mailbox" and pst_path:
        base = os.path.dirname(os.path.abspath(pst_path))
    elif source_root:
        base = os.path.abspath(source_root)
    else:
        base = _portable_staging_base()
    os.makedirs(base, exist_ok=True)
    return base


def export_log_paths(source_root: str, pst_path: str, export_mode: str) -> tuple[str, str]:
    """Return (csv_path, log_path) beside PST or source folder."""
    base = export_output_dir(source_root, pst_path, export_mode)
    return (
        os.path.join(base, "export_results.csv"),
        os.path.join(base, "export.log"),
    )


def normalize_pst_path(pst_path: str) -> str:
    """Absolute .pst path with parent directory created."""
    path = os.path.normpath(os.path.abspath((pst_path or "").strip()))
    if path and not path.lower().endswith(".pst"):
        path += ".pst"
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)
    return path


def validate_import_path(base_dir: str, user_path: str) -> str:
    """
    Resolve user_path and ensure it stays under base_dir (blocks .. traversal).
    Delegates to path_security.PathValidator.
    """
    try:
        return PathValidator.sanitize_path(user_path, base_dir, allow_absolute=True)
    except PathSecurityError as exc:
        raise ValueError(str(exc)) from exc


def pst_paths_equal(left: str, right: str) -> bool:
    """Compare PST paths on disk (case-insensitive on Windows)."""
    if not left or not right:
        return False
    a = os.path.normpath(os.path.abspath(left))
    b = os.path.normpath(os.path.abspath(right))
    if os.name == "nt":
        return os.path.normcase(a) == os.path.normcase(b)
    return a == b


def bytes_look_binary(data: bytes, *, sample: int = 8192) -> bool:
    """True when payload is likely not human-readable mail body text."""
    if not data:
        return False
    chunk = data[:sample]
    for sig in _BINARY_BODY_SIGNATURES:
        if chunk.startswith(sig):
            return True
    if b"\x00" in chunk:
        return True
    if len(chunk) < 32:
        return False
    printable = sum(1 for b in chunk if b in (9, 10, 13) or 32 <= b <= 126)
    return (printable / len(chunk)) < 0.72


def text_looks_binary(text: str) -> bool:
    """True when a decoded body string still looks like embedded binary."""
    if not text:
        return False
    sample = text[:8192]
    if sample.lstrip().startswith("%PDF-"):
        return True
    if sample.count("\\x") >= 8:
        return True
    try:
        encoded = sample.encode("utf-8", errors="ignore")
    except Exception:
        return True
    return bytes_look_binary(encoded)


def _legacy_export_log_path(log_path: str) -> str:
    """Older builds used export_log.txt beside the same folder."""
    return os.path.join(os.path.dirname(os.path.abspath(log_path)), "export_log.txt")


def load_resume_paths(csv_path: str, *, source_root: str = "") -> set[str]:
    """Paths successfully converted in a prior run."""
    done: set[str] = set()
    if not os.path.isfile(csv_path):
        return done
    try:
        with open(csv_path, newline="", encoding="utf-8") as handle:
            for row in csv.DictReader(handle):
                if (row.get("status") or "").lower() == "converted":
                    path = row.get("file_path", "").strip()
                    if not path:
                        continue
                    try:
                        if source_root:
                            path = validate_import_path(source_root, path)
                        else:
                            path = os.path.normpath(os.path.abspath(path))
                    except ValueError as exc:
                        logger.warning("Resume path skipped: %s", exc)
                        continue
                    done.add(path)
    except OSError as exc:
        logger.warning("Could not read resume log %s: %s", csv_path, exc)
    return done


def load_resume_paths_from_sqlite(db_path: str) -> set[str]:
    """Converted paths persisted in dedup_state.sqlite3 (survives CSV flush gaps)."""
    done: set[str] = set()
    if not db_path or not os.path.isfile(db_path):
        return done
    try:
        conn = sqlite3.connect(db_path)
        try:
            row = conn.execute(
                "SELECT 1 FROM sqlite_master WHERE type='table' AND name='converted_path'"
            ).fetchone()
            if not row:
                return done
            for (path,) in conn.execute(
                "SELECT path FROM converted_path WHERE status = 'converted'"
            ):
                if path:
                    done.add(os.path.normpath(os.path.abspath(path)))
        finally:
            conn.close()
    except sqlite3.Error as exc:
        logger.warning("Could not read resume state from %s: %s", db_path, exc)
    return done


def count_converted_in_csv(csv_path: str) -> int:
    """Count converted rows without loading paths (fast preflight)."""
    if not os.path.isfile(csv_path):
        return 0
    count = 0
    try:
        with open(csv_path, newline="", encoding="utf-8") as handle:
            for row in csv.DictReader(handle):
                if (row.get("status") or "").lower() == "converted":
                    count += 1
    except OSError as exc:
        logger.warning("Could not count resume log %s: %s", csv_path, exc)
    return count


class EmlToPstConverter:
    def __init__(self, root):
        self.root = root
        self.root.title(f"Mail Exporter {APP_VERSION}")
        # Default / min size: long HU/EN labels need room; bottom bar must stay visible.
        self.root.geometry("880x720")
        self.root.minsize(780, 640)
        self.root.resizable(True, True)
        self.root.configure(bg='#f0f0f0')
        
        # Variables
        self.folder_path = tk.StringVar()
        self.destination_path = tk.StringVar()
        self.pst_option = tk.StringVar(value="new")
        self.mailbox_store_var = tk.StringVar(value="")
        self.mailbox_folder_name = tk.StringVar(value="Imported EML")
        self._outlook_store_labels: list[str] = []
        self.remove_duplicates = tk.BooleanVar(value=True)
        self.preserve_subfolders = tk.BooleanVar(value=True)
        self.resume_from_log = tk.BooleanVar(value=True)
        # Off by default: use manual fallback + Windows file mtime for time when checked.
        self.strict_date_preservation = tk.BooleanVar(value=False)
        # Explorer "Date modified" / "Módosítás dátuma" on the original .eml file.
        # On by default for Windows Live Mail trees: Explorer "Date modified" matches user expectation.
        self.use_file_mtime_for_date = tk.BooleanVar(value=True)
        self.lang_var = tk.StringVar(value=LANGUAGE_NAMES[detect_default_lang()])
        self.file_pattern = tk.StringVar(value="*.eml")
        self.eml_files: list[str] = []
        self._scan_truncated = False
        self.processed_hashes = set()
        self._delivery_time_warn_count = 0
        self._wrong_store_warn_count = 0
        self._pst_layout_logged = False
        self._pst_direct_import_logged = False
        self._pst_standard_folder_cache: dict[tuple[str, int], object] = {}
        self._pst_std_folder_warn_logged: set[tuple[str, int]] = set()
        self._pst_std_folder_resolve_logged: set[tuple[str, int]] = set()
        self._folder_route_logged = False
        
        # Thread safety
        self._lock = threading.Lock()
        self._outlook_recovery_lock = threading.Lock()
        self._is_converting = False
        self._cancel_requested = False
        self._temp_files = []
        self._conversion_options = {}
        self._export_log_handlers: list[logging.Handler] = []
        self._csv_file = None
        self._csv_writer = None
        self._csv_rows_since_flush = 0
        self._conversion_thread = None
        self._dup_fingerprints: set[str] = set()
        self._dedup_backend = "memory"
        self._dedup_sqlite_conn = None
        self._dedup_sqlite_path = None
        self._dedup_sqlite_pending = 0
        self._dedup_sqlite_lock = threading.Lock()
        self._eml_file_sizes: dict[str, int] = {}
        # Native OpenSharedItem staging: paths kept until batch finishes (Outlook async read).
        self._native_staging_paths = []
        self._staging_dir = None  # set per run; folder next to target PST
        
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        atexit.register(self._cleanup_temp_files)
        
        self.create_widgets()
        self.apply_language()

    def _current_lang(self) -> str:
        return lang_from_display(self.lang_var.get())

    def apply_language(self, _event=None):
        """Refresh all UI strings (English / Hungarian)."""
        lang = self._current_lang()
        self.root.title(t(lang, "window_title", version=APP_VERSION))

        self._lf_folder.config(text=t(lang, "folder_section"))
        self._lf_pst.config(text=t(lang, "save_pst_section"))
        self._lf_list.config(text=t(lang, "eml_files_section"))
        self._btn_add.config(text=t(lang, "add_files"))
        self._lbl_pattern.config(text=t(lang, "file_pattern_label"))
        self._lbl_pattern_hint.config(text=t(lang, "file_pattern_hint"))
        self._rb_new.config(text=t(lang, "create_new_pst"))
        self._rb_existing.config(text=t(lang, "save_existing_pst"))
        self._rb_mailbox.config(text=t(lang, "export_to_mailbox"))
        self._lbl_mailbox_store.config(text=t(lang, "mailbox_store_label"))
        self._btn_refresh_stores.config(text=t(lang, "refresh_stores"))
        self._lbl_mailbox_folder.config(text=t(lang, "mailbox_folder_label"))
        self._btn_wlm.config(text=t(lang, "wlm_preset"))
        self._cb_dup.config(text=t(lang, "remove_duplicates"))
        self._cb_subfolders.config(text=t(lang, "preserve_subfolders"))
        self._cb_resume.config(text=t(lang, "resume_from_log"))
        self._cb_strict.config(text=t(lang, "strict_date_preservation"))
        explorer = t(lang, "explorer_date_modified")
        self._cb_mtime.config(text=t(lang, "use_file_mtime", explorer=explorer))
        self.file_tree.heading("name", text=t(lang, "col_name"))
        self.file_tree.heading("path", text=t(lang, "col_path"))
        self.file_tree.heading("size", text=t(lang, "col_size"))
        self.file_tree.heading("date", text=t(lang, "col_date"))
        self.update_file_count()
        self._lbl_dest.config(text=t(lang, "destination_label"))
        self._btn_browse_dest.config(text=t(lang, "browse_destination"))
        self._btn_exit.config(text=t(lang, "btn_exit"))
        self._btn_convert.config(text=t(lang, "btn_convert"))
        self._btn_cancel.config(text=t(lang, "btn_cancel"))
        self._lbl_lang.config(text=t(lang, "language_label"))
        self._on_export_target_changed()

        try:
            self.context_menu.entryconfig(0, label=t(lang, "context_remove"))
            self.context_menu.entryconfig(1, label=t(lang, "context_clear_all"))
        except tk.TclError:
            pass

        with self._lock:
            if not self._is_converting:
                self.status_label.config(text=t(lang, "status_ready"))

    def create_widgets(self):
        # Main frame: grid so the file list absorbs shrink/grow; dest + buttons stay visible.
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(0, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)
        # Row 3 = file list label frame — only this row expands vertically.
        main_frame.grid_rowconfigure(3, weight=1)

        # Language (Nyelv)
        lang_row = ttk.Frame(main_frame)
        lang_row.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        self._lbl_lang = ttk.Label(lang_row, text="")
        self._lbl_lang.pack(side=tk.LEFT)
        lang_combo = ttk.Combobox(
            lang_row,
            textvariable=self.lang_var,
            values=(LANGUAGE_NAMES[LANG_EN], LANGUAGE_NAMES[LANG_HU]),
            state="readonly",
            width=12,
        )
        lang_combo.pack(side=tk.LEFT, padx=(6, 0))
        lang_combo.bind("<<ComboboxSelected>>", self.apply_language)
        
        # === Add Folder Section ===
        self._lf_folder = ttk.LabelFrame(main_frame, text="", padding="10")
        self._lf_folder.grid(row=1, column=0, sticky="ew", pady=(0, 10))
        folder_frame = self._lf_folder
        
        # Folder path entry and browse
        path_frame = ttk.Frame(folder_frame)
        path_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.folder_entry = ttk.Entry(path_frame, textvariable=self.folder_path, width=70)
        self.folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        btn_col = ttk.Frame(path_frame)
        btn_col.pack(side=tk.RIGHT)
        self._btn_wlm = ttk.Button(btn_col, text="", command=self.use_wlm_preset)
        self._btn_wlm.pack(side=tk.RIGHT, padx=(0, 6))
        self._btn_add = ttk.Button(btn_col, text="", command=self.browse_folder)
        self._btn_add.pack(side=tk.RIGHT)
        
        # Wildcard pattern frame
        pattern_frame = ttk.Frame(folder_frame)
        pattern_frame.pack(fill=tk.X, pady=(5, 0))
        
        self._lbl_pattern = ttk.Label(pattern_frame, text="")
        self._lbl_pattern.pack(side=tk.LEFT)
        
        pattern_combo = ttk.Combobox(pattern_frame, textvariable=self.file_pattern, width=15)
        pattern_combo['values'] = ('*.eml', '*.emlx', '*.eml;*.emlx')
        pattern_combo.pack(side=tk.LEFT, padx=(5, 10))
        self._pattern_combo = pattern_combo
        
        self._lbl_pattern_hint = ttk.Label(pattern_frame, text="")
        self._lbl_pattern_hint.pack(side=tk.LEFT)
        
        # === Save in PST Section ===
        self._lf_pst = ttk.LabelFrame(main_frame, text="", padding="10")
        self._lf_pst.grid(row=2, column=0, sticky="ew", pady=(0, 10))
        pst_frame = self._lf_pst
        
        options_frame = ttk.Frame(pst_frame)
        options_frame.pack(fill=tk.X)
        
        self._rb_new = ttk.Radiobutton(
            options_frame, text="", variable=self.pst_option, value="new",
            command=self._on_export_target_changed,
        )
        self._rb_new.pack(side=tk.LEFT, padx=(0, 20))
        self._rb_existing = ttk.Radiobutton(
            options_frame, text="", variable=self.pst_option, value="existing",
            command=self._on_export_target_changed,
        )
        self._rb_existing.pack(side=tk.LEFT, padx=(0, 20))
        self._rb_mailbox = ttk.Radiobutton(
            options_frame, text="", variable=self.pst_option, value="mailbox",
            command=self._on_export_target_changed,
        )
        self._rb_mailbox.pack(side=tk.LEFT, padx=(0, 20))
        self._cb_dup = ttk.Checkbutton(
            options_frame, text="", variable=self.remove_duplicates
        )
        self._cb_dup.pack(side=tk.LEFT)
        self._cb_strict = ttk.Checkbutton(
            options_frame,
            text="",
            variable=self.strict_date_preservation,
        )
        self._cb_strict.pack(side=tk.LEFT, padx=(20, 0))

        pst_options_row2 = ttk.Frame(pst_frame)
        pst_options_row2.pack(fill=tk.X, pady=(6, 0))
        self._cb_subfolders = ttk.Checkbutton(
            pst_options_row2, text="", variable=self.preserve_subfolders
        )
        self._cb_subfolders.pack(side=tk.LEFT)
        self._cb_resume = ttk.Checkbutton(
            pst_options_row2, text="", variable=self.resume_from_log
        )
        self._cb_resume.pack(side=tk.LEFT, padx=(20, 0))
        self._cb_mtime = ttk.Checkbutton(
            pst_options_row2,
            text="",
            variable=self.use_file_mtime_for_date,
        )
        self._cb_mtime.pack(side=tk.LEFT, padx=(20, 0))

        mailbox_frame = ttk.Frame(pst_frame)
        mailbox_frame.pack(fill=tk.X, pady=(8, 0))
        self._lbl_mailbox_store = ttk.Label(mailbox_frame, text="")
        self._lbl_mailbox_store.pack(side=tk.LEFT)
        self._mailbox_store_combo = ttk.Combobox(
            mailbox_frame,
            textvariable=self.mailbox_store_var,
            state="readonly",
            width=48,
        )
        self._mailbox_store_combo.pack(side=tk.LEFT, padx=(6, 6))
        self._btn_refresh_stores = ttk.Button(
            mailbox_frame, text="", command=self._refresh_outlook_stores
        )
        self._btn_refresh_stores.pack(side=tk.LEFT)
        self._lbl_mailbox_folder = ttk.Label(mailbox_frame, text="")
        self._lbl_mailbox_folder.pack(side=tk.LEFT, padx=(16, 0))
        self._mailbox_folder_entry = ttk.Entry(
            mailbox_frame, textvariable=self.mailbox_folder_name, width=22
        )
        self._mailbox_folder_entry.pack(side=tk.LEFT, padx=(6, 0))
        
        # === File List Section ===
        self._lf_list = ttk.LabelFrame(main_frame, text="", padding="10")
        self._lf_list.grid(row=3, column=0, sticky="nsew", pady=(0, 10))
        list_frame = self._lf_list
        list_frame.grid_rowconfigure(0, weight=1)
        list_frame.grid_columnconfigure(0, weight=1)

        # Create Treeview with scrollbars
        tree_container = ttk.Frame(list_frame)
        tree_container.grid(row=0, column=0, sticky="nsew")
        
        # Scrollbars
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        y_scroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL)
        y_scroll.grid(row=0, column=1, sticky="ns")
        
        x_scroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL)
        x_scroll.grid(row=1, column=0, sticky="ew")
        
        # Treeview (height= rows; grid gives remaining space so list scales with window)
        self.file_tree = ttk.Treeview(tree_container, columns=("name", "path", "size", "date"),
                                       show="headings", height=8,
                                       yscrollcommand=y_scroll.set,
                                       xscrollcommand=x_scroll.set)
        
        self.file_tree.heading("name", text="")
        self.file_tree.heading("path", text="")
        self.file_tree.heading("size", text="")
        self.file_tree.heading("date", text="")
        
        self.file_tree.column("name", width=200, minwidth=150)
        self.file_tree.column("path", width=300, minwidth=200)
        self.file_tree.column("size", width=80, minwidth=60)
        self.file_tree.column("date", width=120, minwidth=100)
        
        self.file_tree.grid(row=0, column=0, sticky="nsew")
        
        y_scroll.config(command=self.file_tree.yview)
        x_scroll.config(command=self.file_tree.xview)
        
        # File count label
        self.file_count_label = ttk.Label(list_frame, text="")
        self.file_count_label.grid(row=1, column=0, sticky="w", pady=(5, 0))
        
        # Context menu for treeview
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="", command=self.remove_selected)
        self.context_menu.add_command(label="", command=self.clear_all)
        self.file_tree.bind("<Button-3>", self.show_context_menu)
        
        # === Destination Section ===
        dest_frame = ttk.Frame(main_frame)
        dest_frame.grid(row=4, column=0, sticky="ew", pady=(0, 10))
        
        self._lbl_dest = ttk.Label(dest_frame, text="")
        self._lbl_dest.pack(side=tk.LEFT)
        
        self.dest_entry = ttk.Entry(dest_frame, textvariable=self.destination_path, width=60)
        self.dest_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 10))
        
        self._btn_browse_dest = ttk.Button(
            dest_frame, text="", command=self.browse_destination
        )
        self._btn_browse_dest.pack(side=tk.RIGHT)
        
        # === Bottom Buttons ===
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, sticky="ew", pady=(10, 0))
        
        # Progress bar
        self.progress = ttk.Progressbar(button_frame, mode='determinate')
        self.progress.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(button_frame, text="")
        self.status_label.pack(side=tk.LEFT, padx=(0, 20))
        
        # Convert and Exit buttons
        self._btn_exit = ttk.Button(
            button_frame, text="", command=self._on_close, width=10
        )
        self._btn_exit.pack(side=tk.RIGHT, padx=(5, 0))
        
        self.convert_btn = ttk.Button(
            button_frame, text="", command=self.start_conversion, width=10
        )
        self.convert_btn.pack(side=tk.RIGHT)
        self._btn_convert = self.convert_btn
        self._btn_cancel = ttk.Button(
            button_frame, text="", command=self.request_cancel, width=10, state="disabled"
        )
        self._btn_cancel.pack(side=tk.RIGHT, padx=(0, 6))
        self._on_export_target_changed()

    def _on_export_target_changed(self, _event=None):
        """Show PST path or Exchange mailbox controls depending on export mode."""
        is_mailbox = self.pst_option.get() == "mailbox"
        dest_state = "disabled" if is_mailbox else "normal"
        self.dest_entry.config(state=dest_state)
        self._btn_browse_dest.config(state=dest_state)
        combo_state = "readonly" if is_mailbox else "disabled"
        self._mailbox_store_combo.config(state=combo_state)
        self._btn_refresh_stores.config(state="normal" if is_mailbox else "disabled")
        self._mailbox_folder_entry.config(state="normal" if is_mailbox else "disabled")
        lang = self._current_lang()
        self._lbl_dest.config(
            text=t(lang, "mailbox_store_label") if is_mailbox else t(lang, "destination_label")
        )
        if is_mailbox and not self._outlook_store_labels:
            self._refresh_outlook_stores_async(lang)

    def _refresh_outlook_stores_async(self, lang: str | None = None):
        """Load Outlook stores on a worker thread so the UI stays responsive."""
        if lang is None:
            lang = self._current_lang()
        with self._lock:
            if self._is_converting:
                return
        thread = threading.Thread(
            target=self._refresh_outlook_stores_worker,
            args=(lang,),
        )
        thread.daemon = True
        thread.start()

    def _refresh_outlook_stores_worker(self, lang: str):
        if not OUTLOOK_AVAILABLE:
            self.root.after(0, lambda: self._refresh_outlook_stores_ui_error(lang, missing_pywin32=True))
            return
        com_initialized = False
        labels = None
        error = None
        try:
            if PYTHONCOM:
                PYTHONCOM.CoInitialize()
                com_initialized = True
            outlook = WIN32COM.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            labels = []
            for i in range(1, namespace.Stores.Count + 1):
                store = namespace.Stores.Item(i)
                labels.append(self._store_label(store, i))
        except Exception as e:
            error = str(e)
            logger.error("Could not list Outlook stores: %s", e)
        finally:
            if com_initialized:
                PYTHONCOM.CoUninitialize()
        self.root.after(
            0,
            lambda lg=lang, lbls=labels, err=error: self._apply_outlook_store_labels(
                lg, lbls, err
            ),
        )

    def _refresh_outlook_stores_ui_error(self, lang, *, missing_pywin32=False):
        if missing_pywin32:
            if _is_frozen():
                messagebox.showerror(
                    t(lang, "title_error"),
                    t(lang, "msg_outlook_required_frozen"),
                )
            else:
                self.prompt_install_pywin32()
            return
        messagebox.showerror(
            t(lang, "title_error"),
            t(lang, "msg_stores_error", err="Unknown error"),
        )

    def _apply_outlook_store_labels(self, lang, labels, error):
        if error:
            messagebox.showerror(
                t(lang, "title_error"),
                t(lang, "msg_stores_error", err=error),
            )
            return
        labels = labels or []
        self._outlook_store_labels = labels
        self._mailbox_store_combo["values"] = labels
        if labels and not self.mailbox_store_var.get():
            self.mailbox_store_var.set(labels[0])
        elif self.mailbox_store_var.get() not in labels:
            self.mailbox_store_var.set(labels[0] if labels else "")
        if not labels:
            messagebox.showwarning(
                t(lang, "title_warning"),
                t(lang, "warn_no_stores"),
            )

    def _clear_file_list(self):
        """Remove all scanned files from the list (UI thread)."""
        self.file_tree.delete(*self.file_tree.get_children())
        with self._lock:
            self.eml_files.clear()
            self._scan_truncated = False
        self.update_file_count()

    def _start_folder_scan(self, folder: str, lang: str, *, replace: bool = True):
        """Scan a folder on a worker thread; optionally replace the current list."""
        if self._is_converting_locked():
            return
        if replace:
            self._clear_file_list()
        self.folder_path.set(folder)
        pattern_snapshot = self.file_pattern.get()
        thread = threading.Thread(
            target=self._scan_folder_thread,
            args=(folder, lang, pattern_snapshot),
        )
        thread.daemon = True
        thread.start()

    def use_wlm_preset(self):
        """Point source folder at the default Windows Live Mail storage location."""
        if self._is_converting_locked():
            return
        lang = self._current_lang()
        wlm = _default_wlm_folder()
        if not os.path.isdir(wlm):
            messagebox.showwarning(
                t(lang, "title_warning"),
                t(lang, "warn_wlm_not_found", path=wlm),
            )
            return
        self._start_folder_scan(wlm, lang)

    def _refresh_outlook_stores(self):
        """Refresh mailbox list (button handler)."""
        if self._is_converting_locked():
            return
        self._refresh_outlook_stores_async(self._current_lang())

    def _is_converting_locked(self) -> bool:
        with self._lock:
            return self._is_converting

    def browse_folder(self):
        """Browse for folder containing EML/EMLX files"""
        if self._is_converting_locked():
            return
        folder = filedialog.askdirectory(
            title=t(self._current_lang(), "dialog_select_folder")
        )
        if folder:
            self._start_folder_scan(folder, self._current_lang())
    
    def _scan_folder_thread(self, folder, lang, pattern_snapshot):
        """Background thread for folder scanning"""
        self.root.after(
            0,
            lambda: self.status_label.config(text=t(lang, "status_scanning")),
        )

        last_scan_ui = [0.0]

        def on_progress(count: int):
            now = time.monotonic()
            if (
                count < MAX_FILES
                and (now - last_scan_ui[0]) < UI_THROTTLE_SEC
            ):
                return
            last_scan_ui[0] = now
            self.root.after(
                0,
                lambda c=count, lg=lang: self.status_label.config(
                    text=t(lg, "status_scanning_progress", n=c)
                ),
            )

        try:
            files, truncated, size_cache = self._scan_folder_impl(
                folder, lang, pattern_snapshot, on_progress=on_progress
            )
            self.root.after(
                0,
                lambda f=files, tr=truncated, sc=size_cache, lg=lang: self._apply_scan_results(
                    f, lg, truncated=tr, file_sizes=sc
                ),
            )
        except Exception as e:
            logger.error("Error scanning folder: %s", e)
            self.root.after(
                0,
                lambda err=str(e), lg=lang: messagebox.showerror(
                    t(lg, "title_error"),
                    t(lg, "err_scan", err=err),
                ),
            )

    def _scan_folder_impl(self, folder, lang, pattern_input, on_progress=None):
        """Scan folder for EML/EMLX files using os.walk (scales to large mail stores)."""
        if not self._is_valid_pattern(pattern_input):
            raise ValueError(t(lang, "err_invalid_pattern"))

        suffixes = _pattern_to_suffixes(pattern_input)
        root = os.path.normpath(os.path.abspath(folder))
        skip_dir_names = {
            NATIVE_STAGING_SUBDIR,
            "MailExporter_Staging",
            "EML2PST_Staging",
        }
        found_files: list[str] = []
        seen: set[str] = set()
        truncated = False
        size_cache: dict[str, int] = {}

        for dirpath, dirnames, filenames in os.walk(root, topdown=True):
            dirnames[:] = [
                d
                for d in dirnames
                if d not in skip_dir_names and not d.startswith(".")
            ]
            for name in filenames:
                lower = name.lower()
                if not any(lower.endswith(ext) for ext in suffixes):
                    continue
                full = os.path.normpath(os.path.join(dirpath, name))
                if full in seen:
                    continue
                seen.add(full)
                found_files.append(full)
                try:
                    size_cache[full] = os.path.getsize(full)
                except OSError:
                    pass
                if on_progress and len(found_files) % SCAN_PROGRESS_EVERY == 0:
                    on_progress(len(found_files))
                if len(found_files) >= MAX_FILES:
                    truncated = True
                    logger.warning("File limit reached (%d)", MAX_FILES)
                    return found_files, truncated, size_cache

        return found_files, truncated, size_cache

    def _apply_scan_results(
        self,
        files: list[str],
        lang: str,
        *,
        truncated: bool = False,
        file_sizes: dict[str, int] | None = None,
    ):
        """Store scanned paths and refresh the preview list (full list used for export)."""
        with self._lock:
            self.eml_files = files
            self._scan_truncated = truncated
            self._eml_file_sizes = dict(file_sizes or {})
        self._populate_file_tree_preview(lang)
        total = len(self.eml_files)
        self.status_label.config(text=t(lang, "status_found", n=total))
        if truncated:
            messagebox.showwarning(
                t(lang, "title_warning"),
                t(lang, "warn_scan_limit", limit=MAX_FILES),
            )
        logger.info("Scan complete: %d files%s", total, " (limit reached)" if truncated else "")

    def _insert_tree_row(self, file_path: str):
        filename = os.path.basename(file_path)
        folder = os.path.dirname(file_path)
        try:
            size = os.path.getsize(file_path)
            size_str = self.format_size(size)
            mod_time = datetime.fromtimestamp(os.path.getmtime(file_path))
            date_str = mod_time.strftime("%Y-%m-%d %H:%M")
        except OSError:
            size_str = "N/A"
            date_str = "N/A"
        self.file_tree.insert("", tk.END, values=(filename, folder, size_str, date_str))

    def _populate_file_tree_preview(self, lang: str):
        """Show up to UI_PREVIEW_FILE_LIMIT rows; conversion uses the full eml_files list."""
        self.file_tree.delete(*self.file_tree.get_children())
        preview = self.eml_files[:UI_PREVIEW_FILE_LIMIT]
        if not preview:
            self._update_file_count_label(lang)
            return
        self._populate_file_tree_chunk(preview, 0, lang)

    def _populate_file_tree_chunk(
        self, preview: list[str], start: int, lang: str
    ):
        """Insert preview rows in chunks so the Tk event loop stays responsive."""
        end = min(start + TREE_PREVIEW_CHUNK, len(preview))
        for i in range(start, end):
            self._insert_tree_row(preview[i])
        if end < len(preview):
            self.root.after(
                1,
                lambda nxt=end, lg=lang, pv=preview: self._populate_file_tree_chunk(
                    pv, nxt, lg
                ),
            )
        else:
            self._update_file_count_label(lang)

    def _add_files_to_list(self, files, lang=None):
        """Legacy entry — bulk apply after scan."""
        if lang is None:
            lang = self._current_lang()
        self._apply_scan_results(files, lang, truncated=False)

    def _is_valid_pattern(self, pattern_input):
        """Validate file pattern to prevent dangerous patterns"""
        allowed_patterns = re.compile(r'^[\*\?a-zA-Z0-9_\-\.;]+$')
        if not allowed_patterns.match(pattern_input):
            return False

        patterns = pattern_input.lower().split(';')
        for p in patterns:
            p = p.strip()
            if p and not (p.endswith('.eml') or p.endswith('.emlx')):
                return False
        return True

    def add_file_to_list(self, file_path):
        """Add a single file to the in-memory list and preview when under the UI cap."""
        lang = self._current_lang()
        with self._lock:
            norm = os.path.normpath(os.path.abspath(file_path))
            if norm in self.eml_files:
                return
            self.eml_files.append(norm)
            try:
                self._eml_file_sizes[norm] = os.path.getsize(norm)
            except OSError:
                self._eml_file_sizes.pop(norm, None)
            total = len(self.eml_files)
            show_row = total <= UI_PREVIEW_FILE_LIMIT

        if show_row:
            self._insert_tree_row(norm)
        self._update_file_count_label(lang)

    def _update_file_count_label(self, lang: str | None = None):
        if lang is None:
            lang = self._current_lang()
        total = len(self.eml_files)
        shown = min(total, UI_PREVIEW_FILE_LIMIT)
        if total > shown:
            self.file_count_label.config(
                text=t(lang, "total_files_preview", total=total, shown=shown)
            )
        else:
            self.file_count_label.config(text=t(lang, "total_files", n=total))

    def update_file_count(self):
        """Update the file count label from the full in-memory list."""
        self._update_file_count_label()

    def format_size(self, size):
        """Format file size in human-readable format"""
        for unit in ['B', 'KB', 'MB', 'GB']:
            if size < 1024:
                return f"{size:.1f} {unit}"
            size /= 1024
        return f"{size:.1f} TB"

    def show_context_menu(self, event):
        """Show context menu on right-click"""
        if self._is_converting_locked():
            return
        self.context_menu.post(event.x_root, event.y_root)

    def remove_selected(self):
        """Remove selected items from the list"""
        if self._is_converting_locked():
            return
        total = len(self.eml_files)
        if total > UI_PREVIEW_FILE_LIMIT:
            lang = self._current_lang()
            messagebox.showinfo(
                t(lang, "title_warning"),
                t(lang, "warn_preview_remove", total=total, shown=UI_PREVIEW_FILE_LIMIT),
            )
            return
        selected = self.file_tree.selection()
        for item in selected:
            values = self.file_tree.item(item)['values']
            if values:
                file_path = os.path.normpath(os.path.join(str(values[1]), str(values[0])))
                with self._lock:
                    if file_path in self.eml_files:
                        self.eml_files.remove(file_path)
            self.file_tree.delete(item)
        lang = self._current_lang()
        self._populate_file_tree_preview(lang)

    def clear_all(self):
        """Clear all items from the list"""
        if self._is_converting_locked():
            return
        self._clear_file_list()
        
    def browse_destination(self):
        """Browse for destination PST file (not used for mailbox export)."""
        if self._is_converting_locked():
            return
        if self.pst_option.get() == "mailbox":
            return
        lang = self._current_lang()
        pst_type = (t(lang, "dialog_filetype_pst"), "*.pst")
        all_type = (t(lang, "dialog_all_files"), "*.*")
        initial = self.destination_path.get().strip()
        initialdir = os.path.dirname(initial) if initial else ""
        initialfile = os.path.basename(initial) if initial else ""

        # Existing PST: open dialog (append). New PST: save-as dialog (create path only).
        use_open = self.pst_option.get() == "existing"
        if not use_open and initial:
            initial_norm = normalize_pst_path(initial)
            if os.path.isfile(initial_norm):
                use_open = True

        if use_open:
            file_path = filedialog.askopenfilename(
                title=t(lang, "dialog_open_pst"),
                filetypes=[pst_type, all_type],
                initialdir=initialdir or None,
                initialfile=initialfile or None,
            )
        else:
            file_path = filedialog.asksaveasfilename(
                title=t(lang, "dialog_save_pst"),
                defaultextension=".pst",
                filetypes=[pst_type, all_type],
                initialdir=initialdir or None,
                initialfile=initialfile or None,
            )

        if file_path:
            path = normalize_pst_path(file_path)
            self.destination_path.set(path)
            if os.path.isfile(path):
                self.pst_option.set("existing")
            else:
                self.pst_option.set("new")
            
    def get_email_hash(self, file_path):
        """Full SHA-256 for duplicate detection (used when FULL_DEDUP is enabled)."""
        sz = file_size_or_zero(file_path)
        if should_use_mmap(sz):
            digest = hash_file_sha256(file_path)
            if digest is None:
                logger.warning(
                    "Could not hash file %s: %s",
                    log_sanitize(file_path),
                    "mmap read failed",
                )
            return digest
        try:
            sha256 = hashlib.sha256()
            with open(file_path, "rb") as f:
                for chunk in iter(lambda: f.read(65536), b""):
                    sha256.update(chunk)
            return sha256.hexdigest()
        except OSError as e:
            logger.warning("Could not hash file %s: %s", log_sanitize(file_path), log_sanitize(e))
            return None

    def _extract_date_from_rfc822_sample(self, data: bytes) -> str:
        """Parse Date: from the first chunk of an .eml for dedup (not Outlook display)."""
        if not data:
            return ""
        match = _DATE_HEADER_BYTES_RE.search(data[:65536])
        if not match:
            return ""
        try:
            raw = match.group(1).decode("utf-8", errors="replace").strip()
        except Exception:
            return ""
        if not raw:
            return ""
        try:
            dt = parsedate_to_datetime(raw)
            if dt is not None:
                if dt.tzinfo is not None:
                    dt = dt.astimezone().replace(tzinfo=None)
                return dt.strftime("%Y-%m-%d %H:%M:%S")
        except (TypeError, ValueError, OverflowError):
            pass
        return raw

    def _file_content_fingerprint(self, file_path: str, size: int) -> tuple[str, str] | None:
        """
        Fast duplicate fingerprint: size + Date header + head/tail sample.
        Returns (hash_hex, normalized_header_date) so SQLite/memory dedup
        does not treat different send times as the same message.
        """
        try:
            digest = hashlib.sha256()
            digest.update(str(size).encode("ascii"))
            header_date = ""
            if should_use_mmap(size):
                samples = read_head_tail_samples(file_path, size)
                if samples is None:
                    return None
                head, tail = samples
            else:
                with open(file_path, "rb") as handle:
                    head = handle.read(65536)
                    tail = b""
                    if size > 131072:
                        handle.seek(size - 65536)
                        tail = handle.read(65536)
            header_date = self._extract_date_from_rfc822_sample(head)
            digest.update(head)
            if header_date:
                digest.update(b"\x00hdr-date\x00")
                digest.update(header_date.encode("utf-8", errors="replace"))
            if tail:
                digest.update(tail)
            return digest.hexdigest(), header_date
        except OSError as e:
            logger.warning(
                "Could not fingerprint file %s: %s",
                log_sanitize(file_path),
                log_sanitize(e),
            )
            return None

    @staticmethod
    def _outlook_text_is_blank(value) -> bool:
        """True when an Outlook/COM field is empty or the literal string 'None'."""
        if value is None:
            return True
        text = str(value).strip()
        return not text or text.lower() == "none"

    @staticmethod
    def _body_looks_like_html(body: str) -> bool:
        return body_looks_like_html(body)

    @staticmethod
    def _normalize_html_body_for_outlook(body: str) -> str:
        """Wrap HTML fragments so Outlook renders them instead of showing raw tags."""
        body = body.strip()
        if not body:
            return body
        if "<html" not in body.lower():
            body = (
                "<!DOCTYPE html><html><head>"
                '<meta http-equiv="Content-Type" content="text/html; charset=utf-8">'
                f"</head><body>{body}</body></html>"
            )
        return body

    def _resolve_from_header(self, msg) -> str:
        """Best-effort From for WLM exports (From, Sender, Reply-To, Return-Path, …)."""
        if msg is None:
            return ""
        for key in (
            "From",
            "Sender",
            "Reply-To",
            "Return-Path",
            "X-Sender",
            "X-Original-From",
            "X-Authenticated-User",
        ):
            raw = msg.get(key)
            if not raw:
                continue
            decoded = self._decode_mime_header_field(raw)
            if decoded:
                display, addr = parseaddr(decoded)
                if addr or display:
                    return decoded
        return ""

    @staticmethod
    def _decode_mime_header_field(value) -> str:
        """Decode RFC 2047 Subject/From and repair common UTF-8/Latin-1 mojibake."""
        if value is None:
            return ""
        if isinstance(value, bytes):
            value = value.decode("utf-8", errors="replace")
        text = str(value).strip()
        if not text:
            return ""
        try:
            chunks: list[str] = []
            for fragment, charset in decode_header(text):
                if isinstance(fragment, bytes):
                    enc = charset or "utf-8"
                    try:
                        chunks.append(fragment.decode(enc, errors="replace"))
                    except LookupError:
                        chunks.append(fragment.decode("utf-8", errors="replace"))
                else:
                    chunks.append(str(fragment))
            text = "".join(chunks).strip()
        except Exception:
            pass
        if "Ã" in text or "â€" in text:
            try:
                repaired = text.encode("latin-1", errors="ignore").decode(
                    "utf-8", errors="replace"
                )
                if repaired.strip():
                    text = repaired.strip()
            except Exception:
                pass
        return text

    def _apply_sender_from_parsed_from(
        self, mail, from_raw: str, *, is_sent_message: bool
    ) -> None:
        """Set Outlook sender fields so the message list is not blank / 'None'."""
        from_raw = self._decode_mime_header_field(from_raw)
        if not from_raw:
            return
        display, addr = parseaddr(from_raw)
        display = self._decode_mime_header_field(display) or display.strip()
        addr = (addr or "").strip()
        if not addr and "@" in from_raw:
            m = re.search(r"[\w.\-+]+@[\w.\-]+\.\w+", from_raw)
            if m:
                addr = m.group(0)
        if not display and addr:
            display = addr
        if not display and not addr:
            return
        try:
            if is_sent_message:
                mail.SentOnBehalfOfName = display or from_raw
                if addr:
                    try:
                        mail.SentOnBehalfOfEmailAddress = addr
                    except Exception:
                        pass
            else:
                if self._outlook_text_is_blank(getattr(mail, "SenderName", None)):
                    mail.SenderName = display or addr or from_raw
                if addr and self._outlook_text_is_blank(
                    getattr(mail, "SenderEmailAddress", None)
                ):
                    mail.SenderEmailAddress = addr
                    try:
                        mail.SenderEmailType = "SMTP"
                    except Exception:
                        pass
        except (AttributeError, TypeError) as e:
            logger.debug("Could not set sender properties: %s", e)
        self._apply_sender_mapi_properties(
            mail, display or addr or from_raw, addr, is_sent_message=is_sent_message
        )

    def _apply_sender_mapi_properties(
        self,
        mail,
        display_name: str,
        email_addr: str,
        *,
        is_sent_message: bool,
    ) -> None:
        """
        Set sender via MAPI — required on many Outlook builds after Sent=True / Save.
        OM SenderName/SenderEmailAddress alone often stay blank in the message list.
        """
        display_name = (display_name or "").strip()
        email_addr = (email_addr or "").strip()
        if not display_name and not email_addr:
            return
        try:
            pa = mail.PropertyAccessor
        except Exception as e:
            logger.debug("Sender MAPI unavailable: %s", e)
            return
        if is_sent_message:
            props = (
                (_OUTLOOK_MAPI_SENT_REP_NAME, display_name),
                (_OUTLOOK_MAPI_SENT_REP_EMAIL, email_addr),
            )
        else:
            props = (
                (_OUTLOOK_MAPI_SENDER_NAME, display_name),
                (_OUTLOOK_MAPI_SENDER_EMAIL, email_addr),
            )
        for url, value in props:
            if not value:
                continue
            try:
                pa.SetProperty(url, value)
            except Exception as e:
                logger.debug("SetProperty sender %s: %s", url, e)
        if email_addr and not is_sent_message:
            try:
                pa.SetProperty(_OUTLOOK_MAPI_SENDER_ADDRTYPE, "SMTP")
            except Exception:
                pass

    def _reapply_sender_from_email_data(
        self,
        mail,
        email_data: dict | None,
        source_file_path: str | None,
        conversion_options: dict,
    ) -> None:
        """Re-apply sender after Received/Sent flags (Outlook may clear list fields)."""
        if not email_data:
            return
        from_raw = (email_data.get("from") or "").strip()
        if not from_raw and email_data.get("message") is not None:
            from_raw = self._resolve_from_header(email_data["message"])
        if not from_raw:
            return
        sent_state = None
        if source_file_path:
            sent_state = self._sent_state_for_source_path(
                source_file_path, conversion_options
            )
        self._apply_sender_from_parsed_from(
            mail, from_raw, is_sent_message=(sent_state is True)
        )

    def _copy_sender_from_mail_item(
        self, dest_mail, source_mail, *, is_sent_message: bool
    ) -> None:
        """Copy non-blank sender fields from an OpenSharedItem mail item."""
        if self._outlook_text_is_blank(getattr(dest_mail, "SenderName", None)):
            try:
                name = getattr(source_mail, "SenderName", None)
                if not self._outlook_text_is_blank(name):
                    if is_sent_message:
                        dest_mail.SentOnBehalfOfName = name
                    else:
                        dest_mail.SenderName = name
            except Exception:
                pass
        if not is_sent_message and self._outlook_text_is_blank(
            getattr(dest_mail, "SenderEmailAddress", None)
        ):
            try:
                addr = getattr(source_mail, "SenderEmailAddress", None)
                if not self._outlook_text_is_blank(addr):
                    dest_mail.SenderEmailAddress = addr
                    try:
                        dest_mail.SenderEmailType = "SMTP"
                    except Exception:
                        pass
            except Exception:
                pass

    def parse_eml(self, file_path):
        """
        Parse an EML file and return email data.
        Uses the stdlib parser so all MIME headers (Date, Received chain, etc.)
        stay available for transport-header and date preservation in Outlook.
        """
        try:
            msg = parse_message_from_path(file_path)
            if msg is None:
                return None

            subject = self._decode_mime_header_field(msg.get("Subject"))
            return {
                'subject': subject or "(No Subject)",
                'from': self._resolve_from_header(msg),
                'to': self._decode_mime_header_field(msg.get("To")),
                'cc': self._decode_mime_header_field(msg.get("Cc")),
                'date': msg.get('Date', ''),
                'body': self.get_email_body(msg),
                'attachments': self.get_attachments(msg),
                'message': msg
            }
        except Exception as e:
            logger.error("Error parsing EML file %s: %s", log_sanitize(file_path), log_sanitize(e))
            return None

    def _part_payload_bytes(self, part) -> bytes:
        ctype = part.get_content_type()
        if ctype == "message/rfc822":
            nested = part.get_payload()
            if isinstance(nested, list):
                nested = nested[0] if nested else None
            if nested is not None:
                try:
                    if hasattr(nested, "as_bytes"):
                        return nested.as_bytes()
                    if hasattr(nested, "as_string"):
                        return nested.as_string().encode("utf-8", errors="replace")
                except Exception as e:
                    logger.debug("Could not serialize message/rfc822 part: %s", e)
        try:
            payload = part.get_payload(decode=True)
        except Exception:
            payload = None
        if isinstance(payload, bytes):
            return payload
        if isinstance(payload, str):
            return payload.encode("utf-8", errors="replace")
        raw = part.get_payload(decode=False)
        if isinstance(raw, bytes):
            return raw
        if isinstance(raw, str):
            return raw.encode("utf-8", errors="replace")
        return b""

    def _decode_part_as_text(self, part) -> str | None:
        ctype = part.get_content_type()
        if not ctype.startswith("text/"):
            return None
        try:
            content = part.get_content()
            if isinstance(content, bytes):
                charset = part.get_content_charset() or "utf-8"
                try:
                    text = content.decode(charset, errors="replace")
                except LookupError:
                    text = content.decode("utf-8", errors="replace")
            elif content is None:
                raw = self._part_payload_bytes(part)
                if not raw:
                    return None
                charset = part.get_content_charset() or "utf-8"
                try:
                    text = raw.decode(charset, errors="replace")
                except LookupError:
                    text = raw.decode("utf-8", errors="replace")
            else:
                text = str(content)
        except (KeyError, LookupError, UnicodeDecodeError, TypeError, ValueError) as e:
            logger.debug("Could not decode %s part: %s", ctype, e)
            return None
        if text_looks_binary(text) or bytes_look_binary(self._part_payload_bytes(part)):
            return None
        return text

    def _is_attachment_like_part(self, part) -> bool:
        if part.get_content_maintype() == "multipart":
            return False
        ctype = part.get_content_type()
        if ctype == "message/rfc822":
            return True
        disp = (part.get_content_disposition() or "").lower()
        if disp in ("attachment", "inline"):
            if part.get_filename():
                return True
            if not ctype.startswith("text/"):
                return True
        if part.get_filename() and not ctype.startswith("text/"):
            return True
        if ctype.startswith(("application/", "image/", "audio/", "video/")):
            return True
        return False

    def _attachment_filename_for_part(self, part, *, index: int = 0) -> str:
        filename = part.get_filename()
        if filename:
            return self._sanitize_filename(filename)
        ctype = part.get_content_type()
        if ctype == "message/rfc822":
            return self._sanitize_filename(f"attached-message-{index or 1}.eml")
        ext = _MIME_EXTENSION_MAP.get(ctype)
        if not ext:
            main = part.get_content_maintype()
            sub = part.get_content_subtype()
            if main and sub:
                ext = f".{sub.split('.')[-1][:12]}"
            else:
                ext = ".bin"
        return self._sanitize_filename(f"attachment-{index or 1}{ext}")

    def _unique_attachment_filename(self, filename: str, used: set[str]) -> str:
        base = self._sanitize_filename(filename)
        if base.lower() not in used:
            used.add(base.lower())
            return base
        name, ext = os.path.splitext(base)
        for n in range(2, 1000):
            candidate = f"{name}_{n}{ext}"
            if candidate.lower() not in used:
                used.add(candidate.lower())
                return candidate
        fallback = f"attachment_{uuid.uuid4().hex[:8]}{ext or '.bin'}"
        used.add(fallback.lower())
        return fallback

    def _attachment_record(self, filename: str, data: bytes, content_type: str = "") -> dict:
        return {
            "filename": self._sanitize_filename(filename),
            "data": data,
            "content_type": content_type or "",
        }

    def _merge_attachment_records(self, *lists) -> list[dict]:
        """Merge attachment lists from MIME parse and Outlook COM; dedupe by name+size."""
        merged: list[dict] = []
        seen: set[tuple[str, int]] = set()
        for lst in lists:
            for att in lst or []:
                data = att.get("data")
                if not data:
                    continue
                if len(data) > MAX_ATTACHMENT_BYTES:
                    continue
                fname = self._sanitize_filename(att.get("filename") or "attachment")
                key = (fname.lower(), len(data))
                if key in seen:
                    continue
                seen.add(key)
                merged.append(
                    self._attachment_record(
                        fname,
                        data,
                        att.get("content_type") or "",
                    )
                )
        return merged

    def _collect_outlook_attachments(self, source_mail) -> list[dict]:
        """Read attachments from an Outlook MailItem via SaveAsFile."""
        results: list[dict] = []
        try:
            atts = source_mail.Attachments
            count = int(atts.Count)
        except Exception:
            return results
        if count <= 0:
            return results
        temp_dir = tempfile.mkdtemp(prefix="me_read_")
        try:
            used: set[str] = set()
            for i in range(1, count + 1):
                try:
                    att = atts.Item(i)
                    fname = self._unique_attachment_filename(
                        att.FileName or f"attach_{i}", used
                    )
                    path = os.path.join(temp_dir, fname)
                    att.SaveAsFile(path)
                    with open(path, "rb") as handle:
                        data = handle.read()
                    if data:
                        results.append(self._attachment_record(fname, data))
                except Exception as e:
                    logger.debug("Outlook attachment read failed: %s", e)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)
        return results

    def _sniff_eml_has_attachments(self, file_path: str) -> bool:
        """
        True when the .eml likely has a real file attachment (not HTML inline parts).

        WLM/HTML mail often has Content-Disposition: inline on text/html only — that
        must not trigger the 'attachments but none imported' warning.
        """
        try:
            with open(file_path, "rb") as handle:
                head = handle.read(min(os.path.getsize(file_path), NATIVE_IMPORT_SNIFF_BYTES))
        except OSError:
            return False
        if not head:
            return False
        if _ATTACHMENT_DISPOSITION_SNIFF_RE.search(head):
            return True
        if _INLINE_NAMED_DISPOSITION_SNIFF_RE.search(head):
            return True
        if _BINARY_PART_CONTENT_TYPE_SNIFF_RE.search(head):
            return True
        return False

    def _apply_attachment_chain(
        self,
        dest_mail,
        *,
        email_data: dict | None = None,
        outlook_mail=None,
        source_file_path: str | None = None,
    ) -> int:
        """
        Import attachments using a fallback chain:
          1) MIME parts parsed from the .eml (PDF/PNG/ZIP/inline/nested .eml)
          2) Outlook MailItem.Attachments from OpenSharedItem (fills gaps)
        """
        parsed_atts: list[dict] = []
        if email_data is None and source_file_path:
            email_data = self.parse_eml(source_file_path)
        if email_data:
            parsed_atts = list(email_data.get("attachments") or [])

        outlook_atts: list[dict] = []
        if outlook_mail is not None:
            outlook_atts = self._collect_outlook_attachments(outlook_mail)

        merged = self._merge_attachment_records(parsed_atts, outlook_atts)
        if merged:
            self._add_attachments(dest_mail, merged)
            return len(merged)

        if source_file_path and self._sniff_eml_has_attachments(source_file_path):
            logger.warning(
                "Message appears to contain attachments but none were imported: %s",
                log_sanitize(source_file_path),
            )
        return 0

    def _set_mail_body(self, mail, body: str) -> None:
        """Set Body/HTMLBody only when content is safe readable text."""
        if not body or text_looks_binary(body):
            return
        if self._body_looks_like_html(body):
            mail.HTMLBody = self._normalize_html_body_for_outlook(body)
        else:
            mail.Body = body

    def _apply_parsed_display_fields(
        self,
        mail,
        email_data: dict | None,
        source_file_path: str | None,
        conversion_options: dict,
        *,
        outlook_source_mail=None,
    ) -> None:
        """
        Apply Subject, recipients, HTML/plain body, and sender from parsed .eml.

        Staged OpenSharedItem imports often leave blank or literal 'None' list fields;
        always overlay parsed MIME metadata after the item exists in the PST.
        """
        if not email_data:
            return
        subject = (email_data.get("subject") or "").strip()
        if subject and subject != "(No Subject)":
            try:
                if self._outlook_text_is_blank(getattr(mail, "Subject", None)):
                    mail.Subject = subject
            except Exception as e:
                logger.debug("Could not set Subject: %s", e)
        elif self._outlook_text_is_blank(getattr(mail, "Subject", None)):
            try:
                mail.Subject = "(No Subject)"
            except Exception:
                pass

        for key, prop in (("to", "To"), ("cc", "CC")):
            val = (email_data.get(key) or "").strip()
            if not val:
                continue
            try:
                if self._outlook_text_is_blank(getattr(mail, prop, None)):
                    setattr(mail, prop, val)
            except Exception as e:
                logger.debug("Could not set %s: %s", prop, e)

        self._copy_body_to_mail_item(
            mail, outlook_source_mail or mail, email_data
        )

        sent_state = None
        if source_file_path:
            sent_state = self._sent_state_for_source_path(
                source_file_path, conversion_options
            )
        from_raw = (email_data.get("from") or "").strip()
        if not from_raw and email_data.get("message") is not None:
            from_raw = self._resolve_from_header(email_data["message"])
        if from_raw:
            self._apply_sender_from_parsed_from(
                mail,
                from_raw,
                is_sent_message=(sent_state is True),
            )
        elif outlook_source_mail is not None:
            self._copy_sender_from_mail_item(
                mail,
                outlook_source_mail,
                is_sent_message=(sent_state is True),
            )

    def _copy_body_to_mail_item(self, dest_mail, source_mail, parsed_email: dict | None) -> None:
        """
        Prefer parsed MIME text (correct text/* parts). Fall back to Outlook item
        only when it does not look like embedded binary (PDF bytes in Body).
        """
        parsed_body = (parsed_email or {}).get("body") or ""
        if parsed_body and not text_looks_binary(parsed_body):
            self._set_mail_body(dest_mail, parsed_body)
            return

        for prop, use_html in (("HTMLBody", True), ("Body", False)):
            try:
                content = getattr(source_mail, prop, None)
            except Exception:
                content = None
            if not content or text_looks_binary(str(content)):
                continue
            if use_html:
                dest_mail.HTMLBody = content
            else:
                dest_mail.Body = content
            return

    def get_email_body(self, msg):
        """Extract human-readable body; never return PDF/binary attachment bytes."""
        plain_parts: list[str] = []
        html_parts: list[str] = []
        for part in msg.walk():
            if part.get_content_maintype() == "multipart":
                continue
            if self._is_attachment_like_part(part):
                continue
            ctype = part.get_content_type()
            if ctype not in ("text/plain", "text/html"):
                continue
            text = self._decode_part_as_text(part)
            if not text:
                continue
            if ctype == "text/html":
                html_parts.append(text)
            else:
                plain_parts.append(text)
        if html_parts:
            return html_parts[0]
        if plain_parts:
            return plain_parts[0]
        return ""

    def get_attachments(self, msg):
        """Extract attachment and inline non-text parts (PDF, images, etc.)."""
        attachments = []
        seen: set[tuple[str, int]] = set()
        if not msg.is_multipart():
            if self._is_attachment_like_part(msg):
                data = self._part_payload_bytes(msg)
                filename = self._attachment_filename_for_part(msg)
                key = (filename, len(data))
                if data and key not in seen:
                    seen.add(key)
                    if len(data) <= MAX_ATTACHMENT_BYTES:
                        attachments.append({
                            'filename': filename,
                            'data': data,
                            'content_type': msg.get_content_type(),
                        })
            return attachments

        for part in msg.walk():
            if not self._is_attachment_like_part(part):
                continue
            data = self._part_payload_bytes(part)
            if not data:
                continue
            filename = self._attachment_filename_for_part(
                part, index=len(attachments) + 1
            )
            key = (filename, len(data))
            if key in seen:
                continue
            seen.add(key)
            if len(data) > MAX_ATTACHMENT_BYTES:
                logger.warning(
                    "Skipping oversized attachment %s (%d bytes)",
                    log_sanitize(filename),
                    len(data),
                )
                continue
            attachments.append({
                'filename': filename,
                'data': data,
                'content_type': part.get_content_type(),
            })
        return attachments
    
    def _sanitize_filename(self, filename):
        """Sanitize filename to prevent path traversal attacks."""
        return PathValidator.create_safe_filename(filename, max_length=200)
        
    def _on_close(self):
        """Handle window close with proper cleanup"""
        lang = self._current_lang()
        converting = False
        thread = None
        with self._lock:
            converting = self._is_converting
            thread = self._conversion_thread
        if converting:
            if not messagebox.askyesno(
                t(lang, "confirm_exit_title"),
                t(lang, "confirm_exit_msg"),
            ):
                return
            with self._lock:
                self._cancel_requested = True
        if converting and thread is not None and thread.is_alive():
            self.status_label.config(text=t(lang, "status_cancelling"))
            thread.join(timeout=45.0)
        self._cleanup_temp_files()
        self.root.destroy()
    
    def _cleanup_temp_files(self):
        """Remove any leftover temp files created during attachment handling"""
        self._cleanup_native_staging_files()
        with self._lock:
            paths = list(self._temp_files)
            self._temp_files.clear()
        for path in paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
            except OSError as e:
                logger.debug("Could not remove temp file %s: %s", path, e)

    def _ensure_native_staging_dir(self, pst_path: str) -> str:
        """Folder beside PST for staging (non-hidden name; Outlook-friendly)."""
        base = os.path.dirname(os.path.abspath(pst_path))
        if not base:
            base = _portable_staging_base()
        staging = os.path.join(base, NATIVE_STAGING_SUBDIR)
        try:
            os.makedirs(staging, exist_ok=True)
            test = os.path.join(staging, ".write_test")
            with open(test, "w", encoding="utf-8") as handle:
                handle.write("ok")
            os.remove(test)
            return staging
        except OSError:
            return self._ensure_staging_dir(_portable_staging_base())

    def _native_temp_staging_dirs(self, source_file_path: str) -> list:
        """
        Short LOCALAPPDATA paths first (best for OpenSharedItem), then beside PST.
        """
        dirs: list[str] = []
        portable = _portable_staging_base()
        try:
            os.makedirs(portable, exist_ok=True)
            dirs.append(portable)
        except OSError:
            pass
        if self._staging_dir and os.path.isdir(self._staging_dir):
            if self._staging_dir not in dirs:
                dirs.append(self._staging_dir)
        return dirs

    def _wait_for_staging_file(self, path: str) -> None:
        """Let the filesystem flush before Outlook OpenSharedItem reads the file."""
        deadline = time.perf_counter() + STAGING_WRITE_SETTLE
        last_size = -1
        while time.perf_counter() < deadline:
            try:
                size = os.path.getsize(path)
                if size > 0 and size == last_size:
                    return
                last_size = size
            except OSError:
                pass
            time.sleep(0.05)

    def _register_staging_file(self, path: str) -> None:
        with self._lock:
            if path not in self._native_staging_paths:
                self._native_staging_paths.append(path)

    def _purge_orphan_staging_eml(self, staging_dir: str) -> None:
        """Remove leftover staging .eml from interrupted prior runs."""
        try:
            for name in os.listdir(staging_dir):
                if not name.endswith(".eml"):
                    continue
                if name.startswith(("emlx_as_eml_", "import_")):
                    p = os.path.join(staging_dir, name)
                    try:
                        os.remove(p)
                    except OSError:
                        pass
        except OSError:
            pass

    def _cleanup_native_staging_files(self) -> None:
        """Delete deferred native-import staging files; drop from _temp_files."""
        with self._lock:
            paths = list(self._native_staging_paths)
            self._native_staging_paths.clear()
        for p in paths:
            try:
                if os.path.isfile(p):
                    os.remove(p)
            except OSError as e:
                logger.debug("Could not remove native staging file %s: %s", p, e)
            with self._lock:
                if p in self._temp_files:
                    self._temp_files.remove(p)

    def request_cancel(self):
        """Ask the worker thread to stop after the current message."""
        lang = self._current_lang()
        with self._lock:
            if not self._is_converting:
                return
            self._cancel_requested = True
        self.status_label.config(text=t(lang, "status_cancelling"))

    def _is_cancelled(self) -> bool:
        with self._lock:
            return self._cancel_requested

    def _reset_cancel_flag(self):
        with self._lock:
            self._cancel_requested = False

    def _set_conversion_ui_active(self, active: bool):
        def _apply():
            if active:
                self.convert_btn.config(state="disabled")
                self._btn_cancel.config(state="normal")
                for widget in (
                    self._btn_add,
                    self._btn_wlm,
                    self.folder_entry,
                    self._pattern_combo,
                    self._rb_new,
                    self._rb_existing,
                    self._rb_mailbox,
                    self._cb_dup,
                    self._cb_strict,
                    self._cb_subfolders,
                    self._cb_resume,
                    self._cb_mtime,
                    self._btn_browse_dest,
                    self._btn_refresh_stores,
                    self._mailbox_store_combo,
                    self._mailbox_folder_entry,
                    self.dest_entry,
                ):
                    try:
                        widget.config(state="disabled")
                    except tk.TclError:
                        pass
            else:
                self.convert_btn.config(state="normal")
                self._btn_cancel.config(state="disabled")
                self.folder_entry.config(state="normal")
                self._pattern_combo.config(state="normal")
                self._rb_new.config(state="normal")
                self._rb_existing.config(state="normal")
                self._rb_mailbox.config(state="normal")
                self._cb_dup.config(state="normal")
                self._cb_strict.config(state="normal")
                self._cb_subfolders.config(state="normal")
                self._cb_resume.config(state="normal")
                self._cb_mtime.config(state="normal")
                self._btn_add.config(state="normal")
                self._btn_wlm.config(state="normal")
                self.dest_entry.config(state="normal")
                self._on_export_target_changed()

        self.root.after(0, _apply)

    def _ui_pulse(self, message: str | None = None):
        """Keep the window responsive during main-thread preflight work."""
        if message is not None:
            self.status_label.config(text=message)
        try:
            self.root.update_idletasks()
        except tk.TclError:
            pass

    def _run_preflight(self, conversion_options: dict, lang: str) -> bool:
        """Warn about bitness mismatch and show summary before export."""
        self._ui_pulse(t(lang, "status_preflight"))
        lines = []
        file_count = len(self.eml_files)
        total_bytes = 0
        if file_count <= PREFLIGHT_SIZE_SAMPLE:
            for path in self.eml_files:
                try:
                    total_bytes += os.path.getsize(path)
                except OSError:
                    pass
            size_mb = total_bytes / (1024 * 1024)
        else:
            sample = self.eml_files[:PREFLIGHT_SIZE_SAMPLE]
            sample_bytes = 0
            sampled = 0
            for path in sample:
                try:
                    sample_bytes += os.path.getsize(path)
                    sampled += 1
                except OSError:
                    pass
            if sampled:
                avg = sample_bytes / sampled
                total_bytes = int(avg * file_count)
            size_mb = total_bytes / (1024 * 1024)
            lines.append(t(lang, "preflight_size_estimate", n=file_count))
        lines.append(
            t(
                lang,
                "preflight_summary",
                count=file_count,
                size=f"{size_mb:.1f}",
            )
        )

        app_bits = app_bitness()
        outlook_exe = find_outlook_exe_path()
        outlook_bits = detect_outlook_bitness()
        bitness_mismatch = (
            outlook_bits is not None and outlook_bits != app_bits
        )
        if outlook_exe:
            lines.append(
                t(
                    lang,
                    "preflight_outlook_exe",
                    path=outlook_exe,
                    bits=outlook_bits if outlook_bits is not None else "?",
                )
            )
        elif outlook_bits is None:
            lines.append(t(lang, "preflight_outlook_unknown"))

        com_ok, com_err = True, ""
        if not bitness_mismatch:
            self._ui_pulse(t(lang, "status_checking_outlook"))
            com_ok, com_err = probe_outlook_com()
        if bitness_mismatch:
            lines.append(
                t(
                    lang,
                    "preflight_bitness",
                    app=app_bits,
                    outlook=outlook_bits,
                    recommended=recommended_exporter_exe_name(outlook_bits),
                )
            )
            if not com_ok and com_err:
                lines.append(
                    t(lang, "preflight_com_failed", err=log_sanitize(com_err))
                )
        elif not com_ok and com_err:
            lines.append(
                t(lang, "preflight_com_failed", err=log_sanitize(com_err))
            )

        export_mode = conversion_options.get("pst_option")
        if export_mode != "mailbox":
            pst_path = normalize_pst_path(conversion_options.get("destination_path", ""))
            conversion_options["destination_path"] = pst_path
            dest_dir = os.path.dirname(pst_path)
            if not os.path.isdir(dest_dir):
                messagebox.showerror(
                    t(lang, "title_error"),
                    t(lang, "err_dest_dir", path=dest_dir),
                )
                return False
            if os.path.isfile(pst_path):
                lines.append(t(lang, "preflight_existing_pst"))

        csv_path, _ = export_log_paths(
            conversion_options.get("source_root", ""),
            conversion_options.get("destination_path", ""),
            export_mode,
        )
        if conversion_options.get("resume_from_log") and os.path.isfile(csv_path):
            resumed = count_converted_in_csv(csv_path)
            if resumed:
                lines.append(t(lang, "preflight_resume", n=resumed))

        if bitness_mismatch:
            messagebox.showerror(
                t(lang, "title_error"),
                "\n".join(lines),
            )
            return False

        lines.append(t(lang, "preflight_continue"))
        return messagebox.askyesno(
            t(lang, "preflight_title"),
            "\n".join(lines),
        )

    def _setup_export_log_handlers(self, log_path: str, *, append_log: bool) -> str:
        """
        Open export.log, export_log.txt, and %LOCALAPPDATA%\\MailExporter\\last_export.log.
        Returns the primary log path used beside the PST (or fallback if not writable).
        """
        log_path = os.path.abspath(log_path)
        log_targets = [log_path]
        legacy = _legacy_export_log_path(log_path)
        if os.path.normcase(legacy) != os.path.normcase(log_path):
            log_targets.append(legacy)
        log_targets.append(_last_export_log_path())

        self._export_log_handlers = []
        opened_paths: list[str] = []
        for target in log_targets:
            use_append = append_log and target != _last_export_log_path()
            try:
                handler = _ImmediateFileHandler(target, append=use_append)
            except OSError as exc:
                logger.warning("Could not open log file %s: %s", target, exc)
                continue
            logger.addHandler(handler)
            self._export_log_handlers.append(handler)
            if target != _last_export_log_path():
                opened_paths.append(handler.path)

        if not opened_paths:
            fallback = os.path.join(_portable_staging_base(), "export.log")
            handler = _ImmediateFileHandler(fallback, append=append_log)
            logger.addHandler(handler)
            self._export_log_handlers.append(handler)
            opened_paths = [handler.path]
            log_path = handler.path

        logger.info("Export started — log: %s", opened_paths[0])
        logger.info("Also mirrored to: %s", _last_export_log_path())
        self._flush_export_log()
        return log_path

    def _open_export_csv(self, csv_path: str, *, append_csv: bool) -> None:
        parent = os.path.dirname(os.path.abspath(csv_path))
        if parent:
            os.makedirs(parent, exist_ok=True)
        logger.info("CSV: %s", csv_path)
        write_header = not append_csv or not os.path.isfile(csv_path)
        mode = "a" if append_csv and os.path.isfile(csv_path) else "w"
        encoding = "utf-8"
        if mode == "w" and CSV_UTF8_BOM:
            encoding = "utf-8-sig"
        self._csv_file = open(csv_path, mode, newline="", encoding=encoding)
        self._csv_writer = csv.writer(self._csv_file)
        self._csv_rows_since_flush = 0
        if write_header:
            self._csv_writer.writerow(CSV_HEADERS)
            self._csv_file.flush()
        self._flush_export_log()

    def _teardown_export_logging(self):
        if self._export_log_handlers:
            logger.info("Export session finished")
            self._flush_export_log()
        if self._csv_file:
            try:
                self._csv_file.flush()
                self._csv_file.close()
            except OSError:
                pass
            self._csv_file = None
            self._csv_writer = None
        for handler in self._export_log_handlers:
            logger.removeHandler(handler)
            try:
                handler.close()
            except OSError:
                pass
        self._export_log_handlers = []

    def _flush_export_log(self):
        for handler in self._export_log_handlers:
            try:
                handler.flush()
            except OSError:
                pass

    def _write_csv_row(
        self,
        file_path: str,
        status: str,
        detail: str,
        duration_sec: float,
        target_folder: str,
    ):
        if not self._csv_writer:
            return
        self._csv_writer.writerow(
            [
                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                csv_sanitize(file_path),
                csv_sanitize(status),
                csv_sanitize(detail),
                csv_sanitize(f"{duration_sec:.2f}"),
                csv_sanitize(target_folder),
            ]
        )
        self._csv_rows_since_flush += 1
        if self._csv_rows_since_flush >= CSV_FLUSH_EVERY:
            self._csv_file.flush()
            self._csv_rows_since_flush = 0
            self._flush_export_log()
        self._record_converted_path_sqlite(file_path, status)

    def _record_converted_path_sqlite(self, file_path: str, status: str) -> None:
        """Persist per-file export status for resume after crash/cancel."""
        if self._dedup_backend != "sqlite" or not file_path:
            return
        conn = self._dedup_sqlite_conn
        if conn is None:
            return
        norm = os.path.normpath(os.path.abspath(file_path))
        st = (status or "").lower().strip()
        with self._dedup_sqlite_lock:
            try:
                conn.execute(
                    """
                    INSERT INTO converted_path(path, status, updated_at)
                    VALUES (?, ?, ?)
                    ON CONFLICT(path) DO UPDATE SET
                        status = excluded.status,
                        updated_at = excluded.updated_at
                    """,
                    (norm, st, datetime.now().isoformat(timespec="seconds")),
                )
                if st == "converted":
                    self._dedup_sqlite_pending += 1
                    if self._dedup_sqlite_pending >= DEDUP_SQLITE_COMMIT_EVERY:
                        conn.commit()
                        conn.execute("BEGIN")
                        self._dedup_sqlite_pending = 0
            except sqlite3.Error as e:
                logger.debug("SQLite converted_path update failed: %s", e)

    def _resolve_target_folder_for_file(
        self,
        base_folder,
        file_path: str,
        conversion_options: dict,
        folder_cache: dict,
    ):
        """
        Route each .eml to the matching Outlook folder (Inbox, Sent, Deleted, …).

        Standard WLM/Outlook folder names in the source path always map to the
        corresponding PST folder. Optional preserve_subfolders only controls
        whether account/custom segments (e.g. Account (user@domain)) are recreated under
        that folder.
        """
        source_root = conversion_options.get("source_root", "")
        parts = relative_folder_parts(source_root, file_path)
        if not parts:
            return base_folder, ""
        store = None
        try:
            store = base_folder.Store
        except Exception:
            pass
        preserve_nested = bool(conversion_options.get("preserve_subfolders"))
        mapped_folder, remainder, std_segment = self._map_source_folder_parts(
            store, parts
        )
        if mapped_folder is not None:
            if preserve_nested and remainder:
                prefix = (std_segment,) if std_segment else ()
                folder = self._get_or_create_folder_path(
                    mapped_folder,
                    remainder,
                    folder_cache,
                    cache_prefix=prefix,
                )
            else:
                folder = mapped_folder
        elif preserve_nested:
            folder = self._get_or_create_folder_path(base_folder, parts, folder_cache)
        else:
            folder = base_folder
        folder = self._coerce_import_folder(folder, conversion_options)
        if not self._folder_route_logged:
            nested_note = (
                "account/custom subfolders under each standard folder"
                if preserve_nested
                else "flat — only Inbox/Sent/Deleted/Outbox/Drafts/Junk"
            )
            logger.info(
                "Folder routing: path segments map to Outlook Inbox, Sent Items, "
                "Deleted Items, Outbox, Drafts, or Junk (%s)",
                nested_note,
            )
            self._folder_route_logged = True
        return folder, "\\".join(parts)

    def _map_source_folder_parts(self, store, parts: list[str]):
        """
        Map Inbox / Sent / Deleted / Outbox / etc. from the relative source path.

        Returns (outlook_folder, remainder_segments, matched_segment_name).
        """
        if not parts or store is None:
            return None, parts, ""
        folder_id, match_idx, remainder = find_standard_folder_in_parts(parts)
        if folder_id is None or match_idx < 0:
            return None, parts, ""
        folder = self._get_or_create_standard_folder(store, folder_id)
        if folder is None:
            return None, parts, ""
        return folder, remainder, parts[match_idx]

    def _get_or_create_folder_path(
        self,
        base_folder,
        parts: list[str],
        cache: dict,
        *,
        cache_prefix: tuple[str, ...] = (),
    ):
        folder = base_folder
        built: list[str] = list(cache_prefix)
        for part in parts:
            built.append(part)
            key = tuple(built)
            if key in cache:
                folder = cache[key]
                continue
            folder = self._get_or_create_named_folder(folder, part)
            cache[key] = folder
        return folder
    
    def start_conversion(self):
        """Start the conversion process in a separate thread"""
        lang = self._current_lang()
        with self._lock:
            if self._is_converting:
                messagebox.showwarning(
                    t(lang, "title_warning"), t(lang, "warn_in_progress")
                )
                return
            if not self.eml_files:
                messagebox.showwarning(
                    t(lang, "title_warning"), t(lang, "warn_no_files")
                )
                return

        # Snapshot UI state on the main thread; worker threads must not read Tk variables.
        dest_raw = self.destination_path.get().strip()
        dest_path = normalize_pst_path(dest_raw) if dest_raw else ""
        pst_option = self.pst_option.get()
        if pst_option != "mailbox" and dest_path and os.path.isfile(dest_path):
            pst_option = "existing"
        conversion_options = {
            "destination_path": dest_path,
            "remove_duplicates": bool(self.remove_duplicates.get()),
            "strict_date_preservation": bool(self.strict_date_preservation.get()),
            "use_file_mtime_for_date": bool(self.use_file_mtime_for_date.get()),
            "preserve_subfolders": bool(self.preserve_subfolders.get()),
            "resume_from_log": bool(self.resume_from_log.get()),
            "pst_option": pst_option,
            "mailbox_store_label": self.mailbox_store_var.get().strip(),
            "mailbox_folder_name": safe_outlook_folder_name(
                self.mailbox_folder_name.get().strip() or "Imported EML"
            ),
            "source_root": resolve_source_root(
                self.folder_path.get().strip(),
                list(self.eml_files),
            ),
            "lang": lang,
            "dedup_backend": self._pick_dedup_backend(
                len(self.eml_files),
                resume_from_log=bool(self.resume_from_log.get()),
            ),
            "pst_chunk_size": PST_CHUNK_SIZE,
            "validate_import": VALIDATE_IMPORT,
            "parallel_workers": PARALLEL_PREP_WORKERS,
            "parallel_parse": PARALLEL_PARSE,
            "use_com_pipeline": USE_COM_PIPELINE,
            "async_io": ASYNC_IO_ENABLED,
        }

        if conversion_options["pst_option"] == "mailbox":
            if not conversion_options["mailbox_store_label"]:
                messagebox.showwarning(
                    t(lang, "title_warning"), t(lang, "warn_no_mailbox")
                )
                return
        elif not conversion_options["destination_path"]:
            messagebox.showwarning(
                t(lang, "title_warning"), t(lang, "warn_no_destination")
            )
            return

        if not self._run_preflight(conversion_options, lang):
            return

        with self._lock:
            self._conversion_options = conversion_options
            self._is_converting = True
            self._cancel_requested = False
            self._delivery_time_warn_count = 0
            self._wrong_store_warn_count = 0
            self._pst_layout_logged = False
            self._pst_direct_import_logged = False
            self._pst_standard_folder_cache.clear()
            self._pst_std_folder_warn_logged.clear()
            self._pst_std_folder_resolve_logged.clear()
            self._folder_route_logged = False

        self._set_conversion_ui_active(True)

        thread = threading.Thread(
            target=self._convert_files_thread,
            args=(conversion_options.copy(),),
        )
        thread.daemon = True
        with self._lock:
            self._conversion_thread = thread
        thread.start()
    
    def _convert_files_thread(self, conversion_options):
        """Thread wrapper for conversion with COM initialization"""
        com_initialized = False
        lang = conversion_options.get("lang", LANG_EN)
        try:
            if PYTHONCOM:
                PYTHONCOM.CoInitialize()
                com_initialized = True
            self.convert_files(conversion_options)
        except Exception as exc:
            logger.exception("Export thread crashed: %s", exc)
            err = str(exc)

            def _show_thread_err(message=err, lg=lang):
                messagebox.showerror(
                    t(lg, "title_conversion_error"),
                    t(lg, "msg_conversion_error", err=message),
                )

            self.root.after(0, _show_thread_err)
        finally:
            if com_initialized:
                PYTHONCOM.CoUninitialize()
            with self._lock:
                self._is_converting = False
            self._set_conversion_ui_active(False)
        
    def convert_files(self, conversion_options):
        """Convert EML files to PST format"""
        lang = conversion_options.get("lang", LANG_EN)
        export_mode = conversion_options.get("pst_option", "new")
        csv_path, log_path = export_log_paths(
            conversion_options.get("source_root", ""),
            conversion_options.get("destination_path", ""),
            export_mode,
        )
        conversion_options["csv_path"] = csv_path

        append_log = bool(
            conversion_options.get("resume_from_log") and os.path.isfile(csv_path)
        )
        try:
            log_path = self._setup_export_log_handlers(
                log_path, append_log=append_log
            )
        except OSError as exc:
            logger.error("Could not open export log: %s", exc)
            self.root.after(
                0,
                lambda lg=lang, err=str(exc): messagebox.showerror(
                    t(lg, "title_error"),
                    t(lg, "err_export_log", err=err),
                ),
            )
            return
        conversion_options["log_path"] = log_path

        resume_paths: set[str] = set()
        if conversion_options.get("resume_from_log"):
            logger.info("Loading resume state from %s ...", csv_path)
            self._flush_export_log()
            source_root = conversion_options.get("source_root") or ""
            resume_paths = load_resume_paths(csv_path, source_root=source_root)
            db_path = self._dedup_db_path(conversion_options)
            sqlite_paths = load_resume_paths_from_sqlite(db_path)
            if sqlite_paths and source_root:
                safe_sqlite: set[str] = set()
                for p in sqlite_paths:
                    try:
                        safe_sqlite.add(validate_import_path(source_root, p))
                    except ValueError as exc:
                        logger.warning("Resume SQLite path skipped: %s", exc)
                sqlite_paths = safe_sqlite
            if sqlite_paths:
                extra = len(sqlite_paths - resume_paths)
                resume_paths |= sqlite_paths
                if extra:
                    logger.info(
                        "Resume: merged %d additional path(s) from SQLite %s",
                        extra,
                        db_path,
                    )
            logger.info("Resume: %d file(s) already marked converted", len(resume_paths))
            self._flush_export_log()
        conversion_options["resume_paths"] = resume_paths

        email_filter = _build_email_filter_from_env()
        if email_filter:
            conversion_options["email_filter"] = email_filter
            logger.info("Pre-import email filters: ON")
        if ADAPTIVE_RATE_ENABLED:
            conversion_options["_rate_limiter"] = AdaptiveRateLimiter()
            logger.info("Adaptive COM pacing: ON (EML2PST_ADAPTIVE_RATE)")
        if CHECKPOINT_EVERY > 0:
            ck_base = export_output_dir(
                conversion_options.get("source_root", ""),
                conversion_options.get("destination_path", ""),
                export_mode,
            )
            conversion_options["_recovery"] = RecoveryManager(
                os.path.join(ck_base, "checkpoints")
            )
            logger.info(
                "Export checkpoints every %d message(s) in %s/checkpoints",
                CHECKPOINT_EVERY,
                ck_base,
            )
        if DEDUP_STRATEGY != "content_hash":
            logger.info("Dedup strategy: %s (EML2PST_DEDUP_STRATEGY)", DEDUP_STRATEGY)

        append_csv = bool(resume_paths) and conversion_options.get(
            "resume_from_log"
        )
        try:
            self._open_export_csv(csv_path, append_csv=append_csv)
        except OSError as exc:
            logger.error("Could not open export CSV: %s", exc)
            self._teardown_export_logging()
            self.root.after(
                0,
                lambda lg=lang, err=str(exc): messagebox.showerror(
                    t(lg, "title_error"),
                    t(lg, "err_export_log", err=err),
                ),
            )
            return

        total_files = len(self.eml_files)
        self.root.after(
            0,
            lambda lg=lang, n=total_files, lp=log_path: self._ui_pulse(
                t(lg, "status_export_starting", n=n) + "\n" + t(lg, "status_log_path", path=lp)
            ),
        )
        if conversion_options.get("use_file_mtime_for_date", True):
            logger.info(
                "Message dates: Explorer 'Date modified' on each .eml (recommended for Windows Live Mail)"
            )
        else:
            logger.info(
                "Message dates: email Date / Received headers (Explorer date checkbox is OFF)"
            )

        with self._lock:
            total = len(self.eml_files)
            files_to_process = self.eml_files.copy()
            
        self.root.after(0, lambda: self.progress.config(maximum=total, value=0))
        with self._lock:
            self.processed_hashes.clear()
            self._dup_fingerprints.clear()
        self._init_dedup_backend(conversion_options)
        self._import_validation_repairs = 0
        self._import_validation_rebuilds = 0
        self._import_validation_failures = 0
        if MMAP_READ_THRESHOLD > 0:
            logger.info(
                "Large .eml I/O: memory-mapped reads from %d MB (EML2PST_MMAP_THRESHOLD_MB)",
                MMAP_READ_THRESHOLD // (1024 * 1024),
            )
        if conversion_options.get("validate_import", VALIDATE_IMPORT):
            logger.info(
                "Per-message import validation: ON (folder placement, sender/subject, "
                "HTML, attachments; set EML2PST_VALIDATE_IMPORT=0 to disable)"
            )
            logger.info(
                "Outlook crash recovery: wait up to %ds for Outlook.exe (EML2PST_OUTLOOK_WAIT_MAX)",
                OUTLOOK_CRASH_WAIT_MAX,
            )
        else:
            logger.info("Per-message import validation: OFF")
        
        cancelled = False
        try:
            if not OUTLOOK_AVAILABLE:
                logger.error("Export aborted — pywin32 / Outlook automation not available")
                self._flush_export_log()
                self.root.after(0, self.prompt_install_pywin32)
                return

            self.convert_with_outlook(files_to_process, conversion_options)
        except Exception as e:
            error_msg = str(e)
            logger.error("Conversion error: %s", log_sanitize(error_msg))
            lg = conversion_options.get("lang", LANG_EN)

            def _show_conv_err(err=error_msg, lang=lg):
                messagebox.showerror(
                    t(lang, "title_conversion_error"),
                    t(lang, "msg_conversion_error", err=err),
                )

            self.root.after(0, _show_conv_err)
        finally:
            self._close_dedup_backend(conversion_options)
            self._teardown_export_logging()

    def _dedup_db_path(self, conversion_options: dict) -> str:
        csv_path = conversion_options.get("csv_path", "")
        csv_dir = os.path.dirname(os.path.abspath(csv_path)) if csv_path else ""
        if not csv_dir:
            csv_dir = self._staging_dir or _portable_staging_base()
        return os.path.join(csv_dir, DEDUP_DB_FILENAME)

    def _pick_dedup_backend(
        self, total_files: int, *, resume_from_log: bool = False
    ) -> str:
        """Choose memory/sqlite dedup backend based on env preference and scale."""
        mode = DEDUP_BACKEND_MODE
        if mode in ("memory", "sqlite"):
            return mode
        if mode not in ("", "auto"):
            logger.warning(
                "Unknown EML2PST_DEDUP_BACKEND=%s; falling back to auto mode",
                mode,
            )
        if resume_from_log:
            return "sqlite"
        return "sqlite" if total_files >= DEDUP_SQLITE_THRESHOLD else "memory"

    def _init_dedup_backend(self, conversion_options: dict) -> None:
        self._dedup_backend = conversion_options.get("dedup_backend", "memory")
        self._dedup_sqlite_pending = 0
        if self._dedup_backend != "sqlite":
            return
        csv_dir = os.path.dirname(self._dedup_db_path(conversion_options))
        os.makedirs(csv_dir, exist_ok=True)
        db_path = self._dedup_db_path(conversion_options)
        resume = bool(conversion_options.get("resume_from_log"))
        if not resume:
            for suffix in ("", "-wal", "-shm"):
                p = db_path + suffix
                try:
                    if os.path.isfile(p):
                        os.remove(p)
                except OSError as e:
                    logger.debug("Could not remove old dedup db %s: %s", p, e)
        conn = sqlite3.connect(db_path, check_same_thread=False)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA temp_store=MEMORY")
        conn.execute("PRAGMA cache_size=-20000")
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS seen (
                hash TEXT PRIMARY KEY,
                header_date TEXT
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS converted_path (
                path TEXT PRIMARY KEY,
                status TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        try:
            conn.execute("ALTER TABLE seen ADD COLUMN header_date TEXT")
        except sqlite3.OperationalError:
            pass
        conn.execute("BEGIN")
        self._dedup_sqlite_conn = conn
        self._dedup_sqlite_path = db_path
        resume_n = len(conversion_options.get("resume_paths") or ())
        if resume and os.path.isfile(db_path):
            try:
                n_seen = int(conn.execute("SELECT COUNT(*) FROM seen").fetchone()[0])
            except sqlite3.Error:
                n_seen = 0
            if n_seen > 0 and resume_n == 0:
                logger.warning(
                    "Resume is on but export_results.csv has 0 converted rows; "
                    "clearing stale dedup database (%d keys) so messages are not skipped",
                    n_seen,
                )
                conn.execute("DELETE FROM seen")
                conn.execute("DELETE FROM converted_path")
                n_seen = 0
            logger.info(
                "Dedup backend: sqlite resume %s (%s keys, %d converted in CSV, commit every %d)",
                db_path,
                n_seen,
                resume_n,
                DEDUP_SQLITE_COMMIT_EVERY,
            )
        else:
            logger.info(
                "Dedup backend: sqlite (%s, Date-aware keys, commit every %d)",
                db_path,
                DEDUP_SQLITE_COMMIT_EVERY,
            )

    def _close_dedup_backend(self, conversion_options: dict | None = None) -> None:
        with self._dedup_sqlite_lock:
            conn = self._dedup_sqlite_conn
            self._dedup_sqlite_conn = None
            self._dedup_sqlite_path = None
            self._dedup_sqlite_pending = 0
            if conn is None:
                return
            try:
                conn.commit()
            except sqlite3.Error:
                pass
            try:
                conn.close()
            except sqlite3.Error:
                pass
        # Keep dedup_state.sqlite3 beside export_results.csv for resume runs.

    def _dedup_key_for_file(self, file_path: str, size: int) -> tuple[str, str] | None:
        """Return (dedup_hash, header_date) — hash includes Date: when fingerprinting."""
        if not FULL_DEDUP_HASH and DEDUP_STRATEGY != "content_hash":
            samples = read_head_tail_samples(file_path, size, head_len=65536, tail_len=0)
            if samples:
                head, _tail = samples
                alt = dedup_key_from_header_sample(head, DEDUP_STRATEGY, size)
                if alt:
                    header_date = self._extract_date_from_rfc822_sample(head)
                    return alt, header_date
        if FULL_DEDUP_HASH:
            full_hash = self.get_email_hash(file_path)
            if not full_hash:
                return None
            header_date = ""
            try:
                with open(file_path, "rb") as handle:
                    header_date = self._extract_date_from_rfc822_sample(
                        handle.read(65536)
                    )
            except OSError:
                pass
            # Full SHA-256 ignores Date:; fold header date into the key like fingerprint mode.
            if header_date:
                return f"{full_hash}\x00{header_date}", header_date
            return full_hash, header_date
        return self._file_content_fingerprint(file_path, size)

    def _is_duplicate_and_mark(self, key: str, header_date: str = "") -> bool:
        """Check and mark dedup key in active backend (header_date stored in SQLite)."""
        if self._dedup_backend == "sqlite":
            conn = self._dedup_sqlite_conn
            if conn is None:
                raise RuntimeError("SQLite dedup backend not initialized")
            with self._dedup_sqlite_lock:
                changes_before = conn.total_changes
                conn.execute(
                    "INSERT OR IGNORE INTO seen(hash, header_date) VALUES (?, ?)",
                    (key, header_date or None),
                )
                inserted = (conn.total_changes - changes_before) > 0
                if inserted:
                    self._dedup_sqlite_pending += 1
                    if self._dedup_sqlite_pending >= DEDUP_SQLITE_COMMIT_EVERY:
                        conn.commit()
                        conn.execute("BEGIN")
                        self._dedup_sqlite_pending = 0
                elif header_date:
                    logger.debug(
                        "Duplicate skipped (hash match, header Date was %s)",
                        header_date,
                    )
                return not inserted
        with self._lock:
            if FULL_DEDUP_HASH:
                if key in self.processed_hashes:
                    return True
                self.processed_hashes.add(key)
                return False
            if key in self._dup_fingerprints:
                return True
            self._dup_fingerprints.add(key)
            return False
    
    def prompt_install_pywin32(self):
        """Prompt user to install pywin32 (source runs only; frozen exe bundles it)."""
        lang = self._current_lang()
        if _is_frozen():
            messagebox.showerror(
                t(lang, "title_error"),
                t(lang, "msg_outlook_required_frozen"),
            )
            return
        result = messagebox.askyesno(
            t(lang, "title_missing_dep"),
            t(lang, "msg_missing_dep"),
        )
        
        if result:
            self.status_label.config(text=t(lang, "status_installing_pywin32"))
            self.root.update()
            
            try:
                subprocess.check_call(
                    [sys.executable, "-m", "pip", "install", "pywin32"],
                    creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == 'win32' else 0
                )
                messagebox.showinfo(
                    t(lang, "title_install_ok"),
                    t(lang, "msg_install_ok"),
                )
            except subprocess.CalledProcessError as e:
                logger.error("Failed to install pywin32: %s", e)
                messagebox.showerror(
                    t(lang, "title_install_fail"),
                    t(lang, "msg_install_fail"),
                )
            
            self.status_label.config(text=t(lang, "status_ready"))
            
    def convert_with_outlook(self, files_to_process, conversion_options):
        """Convert using Outlook COM interface (requires Microsoft Outlook)"""
        lang = conversion_options.get("lang", LANG_EN)
        export_mode = conversion_options.get("pst_option", "new")
        pst_path = conversion_options.get("destination_path", "")

        if export_mode == "mailbox":
            self._staging_dir = self._ensure_staging_dir(_portable_staging_base())
        else:
            self._staging_dir = self._ensure_native_staging_dir(pst_path)
        self._purge_orphan_staging_eml(self._staging_dir)
        seen_staging = {self._staging_dir}
        for fp in files_to_process:
            root = os.path.dirname(os.path.abspath(fp))
            sd = os.path.join(root, NATIVE_STAGING_SUBDIR)
            if sd in seen_staging:
                continue
            seen_staging.add(sd)
            if os.path.isdir(sd):
                self._purge_orphan_staging_eml(sd)
        with self._lock:
            self._native_staging_paths.clear()

        self._update_status(t(lang, "status_connecting"))
        logger.info("Connecting to Outlook (%d file(s) queued)", len(files_to_process))

        outlook = None
        namespace = None
        prev_display_alerts = None
        total = len(files_to_process)
        skip_main_outlook = (
            export_mode != "mailbox"
            and USE_COM_PIPELINE
            and not self._should_use_pst_chunks(conversion_options, total)
        )
        try:
            if not skip_main_outlook:
                try:
                    outlook = WIN32COM.Dispatch("Outlook.Application")
                except Exception as e:
                    raise RuntimeError(f"Could not connect to Outlook: {e}") from e

                try:
                    prev_display_alerts = outlook.DisplayAlerts
                    outlook.DisplayAlerts = False
                except Exception:
                    prev_display_alerts = None

                try:
                    namespace = outlook.GetNamespace("MAPI")
                except Exception as e:
                    raise RuntimeError(f"Could not access MAPI namespace: {e}") from e

                conversion_options["_session_outlook"] = outlook

            if export_mode != "mailbox":
                pst_path = normalize_pst_path(pst_path)
                conversion_options["destination_path"] = pst_path
            folder_cache: dict = {}
            if export_mode == "mailbox":
                store_label = conversion_options.get("mailbox_store_label", "")
                folder_name = conversion_options.get("mailbox_folder_name", "Imported EML")
                store = self._find_store_by_label(namespace, store_label)
                if not store:
                    raise RuntimeError(
                        f"Could not find Outlook mailbox: {store_label}"
                    )
                self._update_status(t(lang, "status_preparing_mailbox"))
                target_folder = self._resolve_mailbox_target_folder(store, folder_name)
                note = t(
                    lang,
                    "note_mailbox_saved",
                    store=store_label,
                    folder=folder_name,
                )
                converted, skipped, errors, error_messages, skipped_messages, cancelled = (
                    self._process_email_files(
                        outlook,
                        namespace,
                        target_folder,
                        files_to_process,
                        total,
                        conversion_options,
                        folder_cache,
                    )
                )
            elif self._should_use_pst_chunks(conversion_options, total):
                (
                    converted,
                    skipped,
                    errors,
                    error_messages,
                    skipped_messages,
                    cancelled,
                    note,
                ) = self._convert_with_pst_chunks(
                    outlook,
                    namespace,
                    files_to_process,
                    conversion_options,
                    lang,
                    pst_path,
                )
            elif USE_COM_PIPELINE:
                pst_path = normalize_pst_path(pst_path)
                conversion_options["destination_path"] = pst_path
                note = t(lang, "note_pst_saved", path=pst_path)
                logger.info(
                    "COM pipeline: dedicated Outlook STA worker (not parallel COM to one PST)"
                )
                converted, skipped, errors, error_messages, skipped_messages, cancelled = (
                    self._process_email_files(
                        None,
                        None,
                        None,
                        files_to_process,
                        total,
                        conversion_options,
                        folder_cache,
                    )
                )
            else:
                self._update_status(t(lang, "status_creating_pst"))
                pst_path = self._ensure_pst_store_attached(
                    outlook, namespace, pst_path, conversion_options
                )
                conversion_options["destination_path"] = pst_path
                pst_store = self._find_pst_store(namespace, pst_path)
                if not pst_store:
                    raise RuntimeError(
                        f"Could not access PST file after attach. Path: {pst_path}"
                    )
                self._log_attached_pst_store(pst_store, pst_path)
                target_folder = self._get_or_create_inbox(pst_store)
                note = t(lang, "note_pst_saved", path=pst_path)
                converted, skipped, errors, error_messages, skipped_messages, cancelled = (
                    self._process_email_files(
                        outlook,
                        namespace,
                        target_folder,
                        files_to_process,
                        total,
                        conversion_options,
                        folder_cache,
                    )
                )
            
            if error_messages:
                note += t(lang, "note_errors") + "\n".join(error_messages[:5])
                if len(error_messages) > 5:
                    note += t(
                        lang,
                        "note_more_errors",
                        n=len(error_messages) - 5,
                    )
            if skipped_messages:
                note += t(lang, "note_skipped") + "\n".join(skipped_messages[:5])
                if len(skipped_messages) > 5:
                    note += t(
                        lang,
                        "note_more_skipped",
                        n=len(skipped_messages) - 5,
                    )

            csv_path = conversion_options.get("csv_path", "")
            if csv_path:
                note += t(lang, "note_csv_saved", path=csv_path)
            log_path = conversion_options.get("log_path", "")
            if log_path:
                note += t(lang, "note_log_saved", path=log_path)
            repairs = getattr(self, "_import_validation_repairs", 0)
            rebuilds = getattr(self, "_import_validation_rebuilds", 0)
            val_fail = getattr(self, "_import_validation_failures", 0)
            if repairs or rebuilds or val_fail:
                note += (
                    f"\nImport validation: {repairs} repaired, "
                    f"{rebuilds} rebuilt, {val_fail} failed."
                )
                logger.info(
                    "Import validation summary: %d repaired, %d rebuilt, %d failed",
                    repairs,
                    rebuilds,
                    val_fail,
                )
                
            if cancelled:
                self.root.after(
                    0,
                    lambda c=converted, s=skipped, e=errors, n=note, lg=lang: self.show_completion(
                        c, s, e, n, lg, cancelled=True
                    ),
                )
            else:
                self.root.after(
                    0,
                    lambda c=converted, s=skipped, e=errors, n=note, lg=lang: self.show_completion(
                        c, s, e, n, lg
                    ),
                )
        finally:
            if outlook is not None and prev_display_alerts is not None:
                try:
                    outlook.DisplayAlerts = prev_display_alerts
                except Exception:
                    pass
            try:
                del namespace
            except Exception:
                pass
            try:
                del outlook
            except Exception:
                pass
            gc.collect()
            time.sleep(max(0.0, _NATIVE_POST_BATCH_DELAY))
            self._cleanup_native_staging_files()
    
    def _ensure_staging_dir(self, base_path: str) -> str:
        os.makedirs(base_path, exist_ok=True)
        return base_path

    def _store_label(self, store, index: int) -> str:
        name = str(store.DisplayName or f"Store {index}")
        path = str(store.FilePath or "").strip()
        if path.lower().endswith(".pst"):
            return f"{name} (PST)"
        return f"{name} (Mailbox)"

    def _find_store_by_label(self, namespace, label: str):
        for i in range(1, namespace.Stores.Count + 1):
            store = namespace.Stores.Item(i)
            if self._store_label(store, i) == label:
                return store
        return None
    
    def _update_status(self, text):
        """Thread-safe status update"""
        self.root.after(0, lambda t=text: self.status_label.config(text=t))
    
    def _cleanup_stale_store_for_path(self, namespace, pst_path):
        """Remove stale store reference only for the target PST path."""
        try:
            target = normalize_pst_path(pst_path)
            stale_stores = []
            for store in namespace.Stores:
                try:
                    file_path = store.FilePath
                    if (
                        file_path
                        and pst_paths_equal(file_path, target)
                        and not os.path.exists(file_path)
                    ):
                        stale_stores.append((store, file_path))
                except (AttributeError, OSError):
                    continue

            for store, file_path in stale_stores:
                try:
                    root = store.GetRootFolder()
                    namespace.RemoveStore(root)
                    logger.info("Removed stale store reference: %s", file_path)
                except Exception as e:
                    logger.warning("Could not remove stale store %s: %s", file_path, e)

        except Exception as e:
            logger.debug("Target stale store cleanup failed: %s", e)

    def _stores_for_pst_path(self, namespace, pst_path: str) -> list:
        """All Outlook stores currently attached to this PST path."""
        target = normalize_pst_path(pst_path)
        matches = []
        for store in namespace.Stores:
            try:
                file_path = store.FilePath
                if file_path and pst_paths_equal(file_path, target):
                    matches.append(store)
            except (AttributeError, OSError):
                continue
        return matches

    def _dedupe_attached_stores_for_path(self, namespace, pst_path: str) -> object | None:
        """Keep one attached store for a PST path; remove duplicate profile entries."""
        matches = self._stores_for_pst_path(namespace, pst_path)
        if not matches:
            return None
        primary = matches[0]
        for extra in matches[1:]:
            try:
                namespace.RemoveStore(extra.GetRootFolder())
                logger.info(
                    "Removed duplicate Outlook store for %s (display name %r)",
                    pst_path,
                    getattr(extra, "DisplayName", ""),
                )
            except Exception as e:
                logger.warning("Could not remove duplicate store: %s", e)
        return primary

    def _log_attached_pst_store(self, store, pst_path: str) -> None:
        try:
            display = str(store.DisplayName or "")
            file_path = str(store.FilePath or "")
        except Exception:
            display = file_path = ""
        logger.info(
            "Export target PST: %s (Outlook name %r, attached path %s)",
            normalize_pst_path(pst_path),
            display,
            file_path,
        )
        if file_path and not pst_paths_equal(file_path, pst_path):
            logger.warning(
                "Outlook store path mismatch — expected %s, got %s",
                pst_path,
                file_path,
            )
        if not self._pst_layout_logged:
            self._log_pst_store_layout(store, pst_path)
            self._pst_layout_logged = True

    def _outlook_folder_item_count(self, folder) -> int:
        try:
            return int(folder.Items.Count)
        except Exception:
            return -1

    def _nested_account_folder_count(self, parent_folder) -> tuple[str, int] | None:
        """First subfolder under a standard folder (usually Account (user@…))."""
        try:
            subfolders = parent_folder.Folders
            count = int(subfolders.Count)
        except Exception:
            return None
        for idx in range(1, count + 1):
            try:
                child = subfolders.Item(idx)
                name = str(child.Name or "").strip()
                if not name or self._folder_looks_like_search_container(child):
                    continue
                return name, self._outlook_folder_item_count(child)
            except Exception:
                continue
        return None

    def _log_pst_store_layout(self, store, pst_path: str) -> None:
        """Log where messages will appear (helps when Outlook shows an empty PST root)."""
        pst_path = normalize_pst_path(pst_path)
        pst_label = os.path.splitext(os.path.basename(pst_path))[0] or "PST"
        top_names: list[str] = []
        try:
            root = store.GetRootFolder()
            for i in range(1, int(root.Folders.Count) + 1):
                top_names.append(str(root.Folders.Item(i).Name))
        except Exception as e:
            logger.debug("Could not list PST root folders: %s", e)
        if top_names:
            logger.info(
                "PST %s top-level folders (store root — usually not where mail is): %s",
                pst_path,
                ", ".join(top_names),
            )
        else:
            logger.info("PST %s has no subfolders under the store root yet", pst_path)
        layout_lines: list[str] = []
        for folder_id, label in (
            (OL_FOLDER_INBOX, "Inbox"),
            (OL_FOLDER_SENT, "Sent Items"),
            (OL_FOLDER_DELETED, "Deleted Items"),
            (OL_FOLDER_OUTBOX, "Outbox"),
            (OL_FOLDER_DRAFTS, "Drafts"),
            (OL_FOLDER_JUNK, "Junk"),
        ):
            try:
                folder = self._get_or_create_pst_standard_folder(store, folder_id)
                if folder is None:
                    continue
                name = str(getattr(folder, "Name", "") or "").strip() or label
                direct = self._outlook_folder_item_count(folder)
                nested = self._nested_account_folder_count(folder)
                if nested:
                    sub_name, sub_count = nested
                    layout_lines.append(
                        f"{name} → {sub_name} (~{sub_count} items)"
                    )
                elif direct >= 0:
                    layout_lines.append(f"{name} (~{direct} items)")
            except Exception as e:
                logger.debug("Layout scan %s: %s", label, e)
        if layout_lines:
            logger.info(
                "Mail in %s is under standard Outlook folders (not the PST root): %s",
                pst_label,
                "; ".join(layout_lines),
            )
        logger.info(
            "In Outlook: expand %s → Inbox / Beérkezett üzenetek → account folder "
            "(e.g. Account (user@domain)) for imported Inbox .eml; same account folder "
            "under Deleted Items / Elküldött for other WLM folders. "
            "Uncheck 'preserve account subfolders' for a flat Inbox only.",
            pst_label,
        )

    def _ensure_pst_store_attached(
        self, outlook, namespace, pst_path, conversion_options
    ) -> str:
        """
        Attach exactly one Outlook store for pst_path without dialogs.

        - Reuses an already-attached store (no RemoveStore/AddStore churn).
        - Creates the file with AddStoreEx when missing.
        - Opens an existing file with AddStore when present.
        - Never deletes an existing PST on disk.
        """
        pst_path = normalize_pst_path(pst_path)
        self._cleanup_stale_store_for_path(namespace, pst_path)

        existing = self._dedupe_attached_stores_for_path(namespace, pst_path)
        if existing is not None:
            logger.info("Reusing PST already open in Outlook: %s", pst_path)
            self._prepare_pst_store_for_import(
                outlook, namespace, existing, pst_path
            )
            return pst_path

        file_exists = os.path.isfile(pst_path)

        if file_exists:
            namespace.AddStore(pst_path)
            if conversion_options.get("pst_option") == "new":
                logger.info(
                    "PST file already exists — attaching and appending (not replacing): %s",
                    pst_path,
                )
            else:
                logger.info("Attached existing PST: %s", pst_path)
        else:
            if conversion_options.get("pst_option") == "existing":
                raise FileNotFoundError(f"PST file not found: {pst_path}")
            self._create_new_pst(outlook, namespace, pst_path)

        time.sleep(0.5)
        store = self._find_pst_store(namespace, pst_path)
        if not store:
            raise RuntimeError(
                f"Could not attach PST in Outlook. Path: {pst_path}"
            )
        self._prepare_pst_store_for_import(outlook, namespace, store, pst_path)
        return pst_path

    def _remove_existing_store(self, namespace, pst_path):
        """Detach a PST store from the Outlook profile (used for chunk merge cleanup)."""
        pst_path = normalize_pst_path(pst_path)
        stores_to_remove = self._stores_for_pst_path(namespace, pst_path)

        for store in stores_to_remove:
            try:
                root = store.GetRootFolder()
                namespace.RemoveStore(root)
                logger.info("Removed store reference: %s", pst_path)
                time.sleep(0.5)
            except Exception as e:
                logger.warning("Could not remove store reference: %s", e)
    
    def _create_new_pst(self, outlook, namespace, pst_path):
        """Create a new PST file using the most reliable method"""
        # Method 1: Try AddStoreEx (preferred - creates Unicode PST)
        try:
            namespace.AddStoreEx(pst_path, OL_STORE_UNICODE)
            logger.info("Created PST using AddStoreEx: %s", pst_path)
            return
        except AttributeError:
            logger.debug("AddStoreEx not available (older Outlook version)")
        except Exception as e1:
            logger.debug("AddStoreEx failed: %s", e1)
        
        # Method 2: Try AddStore
        try:
            namespace.AddStore(pst_path)
            logger.info("Created PST using AddStore: %s", pst_path)
            return
        except Exception as e2:
            logger.debug("AddStore failed: %s", e2)
        
        # Method 3: Initialize Outlook with a temp item, then try AddStore
        try:
            temp_mail = outlook.CreateItem(OL_MAIL_ITEM)
            temp_mail.Subject = "Temp"
            temp_mail.Save()
            temp_mail.Delete()
            namespace.AddStore(pst_path)
            logger.info("Created PST using AddStore after init: %s", pst_path)
            return
        except Exception as e3:
            logger.debug("Method 3 failed: %s", e3)
        
        raise RuntimeError(
            "Could not create PST file. Please try:\n"
            "1. Close Outlook completely and restart the converter\n"
            "2. Choose a different filename\n"
            "3. Run Outlook as Administrator"
        )
    
    def _find_pst_store(self, namespace, pst_path, retries=5):
        """Find the PST store in Outlook with retry logic"""
        target = normalize_pst_path(pst_path)

        for attempt in range(retries):
            for store in namespace.Stores:
                try:
                    if store.FilePath and pst_paths_equal(store.FilePath, target):
                        return store
                except (AttributeError, OSError):
                    continue

            if attempt < retries - 1:
                time.sleep(1)
                logger.debug("PST store not found, retry %d/%d", attempt + 2, retries)

        return None

    def _folder_has_valid_name(self, folder) -> bool:
        try:
            return bool(str(folder.Name or "").strip())
        except Exception:
            return False

    def _store_file_path(self, store) -> str:
        try:
            return normalize_pst_path(str(store.FilePath or ""))
        except Exception:
            return ""

    def _store_id(self, store) -> str:
        try:
            return str(store.StoreID or "")
        except Exception:
            return ""

    def _stores_match(self, folder_store, store) -> bool:
        """True when two Store objects refer to the same PST/mailbox."""
        if folder_store is None or store is None:
            return False
        left = self._store_file_path(folder_store)
        right = self._store_file_path(store)
        if left and right and pst_paths_equal(left, right):
            return True
        left_id = self._store_id(folder_store)
        right_id = self._store_id(store)
        if left_id and right_id and left_id == right_id:
            return True
        try:
            return folder_store == store
        except Exception:
            return False

    def _folder_in_store(self, folder, store) -> bool:
        try:
            return self._stores_match(folder.Store, store)
        except Exception:
            return False

    def _folder_belongs_to_pst(self, folder, pst_path: str, *, store=None) -> bool:
        pst_path = normalize_pst_path(pst_path)
        try:
            f_store = folder.Store
        except Exception:
            return False
        store_path = self._store_file_path(f_store)
        if store_path and pst_paths_equal(store_path, pst_path):
            return True
        if store is not None and self._stores_match(f_store, store):
            return True
        return False

    def _folder_supports_mail_items(self, folder) -> bool:
        try:
            _ = folder.Items
            return True
        except Exception:
            return False

    def _folder_class_name(self, folder) -> str:
        try:
            return str(folder.FolderClass or "").lower()
        except Exception:
            return ""

    def _folder_name_looks_like_search(self, name: str) -> bool:
        lower = (name or "").lower().strip()
        if not lower:
            return False
        return any(marker in lower for marker in _SEARCH_FOLDER_NAME_MARKERS)

    def _folder_looks_like_search_container(self, folder) -> bool:
        if "search" in self._folder_class_name(folder):
            return True
        return self._folder_name_looks_like_search(str(getattr(folder, "Name", "") or ""))

    def _inbox_name_is_recognized(self, name: str) -> bool:
        lower = (name or "").lower().strip()
        if not lower:
            return False
        if lower in _STANDARD_PST_FOLDER_SPECS[OL_FOLDER_INBOX][0]:
            return True
        return "beérkezett" in lower and "üzenet" in lower

    def _is_usable_standard_folder(self, folder, folder_id: int, store) -> bool:
        """Reject search folders and other non-mail containers from GetDefaultFolder."""
        if folder is None or not self._folder_in_store(folder, store):
            return False
        if not self._folder_supports_mail_items(folder):
            return False
        if folder_id != OL_FOLDER_JUNK and self._folder_looks_like_search_container(
            folder
        ):
            return False
        if folder_id == OL_FOLDER_INBOX:
            name = str(getattr(folder, "Name", "") or "").strip()
            if not name:
                return False
            if not self._inbox_name_is_recognized(name):
                return False
        return True

    def _pst_folder_ready_for_import(self, folder, pst_path: str, store) -> bool:
        """Inbox (or equivalent) is in the export PST and accepts new mail items."""
        if folder is None:
            return False
        if not self._folder_in_store(folder, store):
            return False
        if not self._folder_supports_mail_items(folder):
            return False
        if self._folder_looks_like_search_container(folder):
            return False
        if self._store_file_path(store):
            return self._folder_belongs_to_pst(folder, pst_path, store=store)
        return True

    def _discover_pst_mail_inbox(self, store, pst_path: str):
        """Last-resort Inbox when GetDefaultFolder returns an empty display name."""
        spec = _STANDARD_PST_FOLDER_SPECS.get(OL_FOLDER_INBOX)
        aliases = spec[0] if spec else ("inbox",)
        root = store.GetRootFolder()
        found = self._find_folder_in_tree(root, aliases, max_depth=4)
        if found is not None and self._folder_in_store(found, store):
            return found
        try:
            root_folders = root.Folders
            count = int(root_folders.Count)
        except Exception:
            return None
        for idx in range(1, count + 1):
            try:
                child = root_folders.Item(idx)
            except Exception:
                continue
            if not self._folder_in_store(child, store):
                continue
            if not self._folder_supports_mail_items(child):
                continue
            if self._folder_looks_like_search_container(child):
                continue
            try:
                child_name = str(child.Name or "").strip()
            except Exception:
                child_name = ""
            if child_name and self._inbox_name_is_recognized(child_name):
                return child
            try:
                if int(child.DefaultItemType) == OL_MAIL_ITEM:
                    return child
            except Exception:
                return child
        return None

    def _folder_is_usable_import_target(
        self, folder, pst_path: str, *, store=None
    ) -> bool:
        if folder is None:
            return False
        if self._folder_has_valid_name(folder):
            store_path = self._folder_store_path(folder)
            return bool(store_path and pst_paths_equal(store_path, pst_path))
        if store is not None:
            return self._pst_folder_ready_for_import(folder, pst_path, store)
        return self._folder_belongs_to_pst(folder, pst_path) and self._folder_supports_mail_items(
            folder
        )

    def _find_folder_under_root(self, root_folder, names: tuple[str, ...]):
        return self._find_folder_in_tree(root_folder, names, max_depth=1)

    def _find_folder_in_tree(
        self, root_folder, names: tuple[str, ...], *, max_depth: int = 4
    ):
        """Find a subfolder by display name (case-insensitive), limited depth."""
        wanted = {n.lower().strip() for n in names if n}

        def walk(folder, depth: int):
            if depth > max_depth:
                return None
            try:
                subfolders = folder.Folders
            except Exception:
                return None
            try:
                count = int(subfolders.Count)
            except Exception:
                count = 0
            for idx in range(1, count + 1):
                try:
                    child = subfolders.Item(idx)
                except Exception:
                    continue
                try:
                    child_name = str(child.Name or "").lower().strip()
                except Exception:
                    child_name = ""
                if child_name in wanted:
                    return child
                found = walk(child, depth + 1)
                if found is not None:
                    return found
            return None

        try:
            return walk(root_folder, 0)
        except Exception as e:
            logger.debug("Scan folder tree: %s", e)
        return None

    def _label_pst_store_for_outlook(self, store, pst_path: str) -> None:
        """Show PST basename (e.g. export) in Outlook navigation instead of generic label."""
        pst_path = normalize_pst_path(pst_path)
        label = os.path.splitext(os.path.basename(pst_path))[0] or "MailExporter PST"
        try:
            current = str(store.DisplayName or "").strip()
            if current.lower() in ("outlook data file", ""):
                store.DisplayName = label
                logger.info(
                    "Outlook store display name set to %r for %s",
                    label,
                    pst_path,
                )
        except Exception as e:
            logger.debug("Could not set PST display name: %s", e)

    def _activate_pst_store_for_export(
        self, outlook, namespace, pst_store, pst_path: str
    ) -> None:
        """
        Prefer making the export PST the active store for Items.Add.

        DefaultStore is read-only on many Outlook 2019 builds; CurrentStore usually
        works. If both fail, imports use explicit PST folder Items.Add only.
        """
        pst_path = normalize_pst_path(pst_path)
        self._label_pst_store_for_outlook(pst_store, pst_path)
        label = str(pst_store.DisplayName or os.path.basename(pst_path))
        activated = False
        for prop_name, assign in (
            ("CurrentStore", lambda: setattr(namespace, "CurrentStore", pst_store)),
            (
                "DefaultStore",
                lambda: setattr(outlook.Session, "DefaultStore", pst_store),
            ),
        ):
            try:
                assign()
                logger.info(
                    "Outlook %s for export: %r (%s)",
                    prop_name,
                    label,
                    pst_path,
                )
                activated = True
                break
            except Exception as e:
                logger.debug("Could not set Session/Namespace %s: %s", prop_name, e)
        if not activated:
            logger.info(
                "Could not switch Outlook active store — creating mail via explicit "
                "PST folders on %s",
                pst_path,
            )

    def _prepare_pst_store_for_import(
        self, outlook, namespace, pst_store, pst_path: str
    ) -> None:
        """
        Ensure the PST has a writable Inbox and is the active export target.

        GetDefaultFolder(Inbox) can return a folder with an empty name; imports
        then miss the PST and land in the profile Drafts store instead.
        """
        pst_path = normalize_pst_path(pst_path)
        self._activate_pst_store_for_export(outlook, namespace, pst_store, pst_path)
        inbox = self._get_or_create_inbox(pst_store)
        if not self._pst_folder_ready_for_import(inbox, pst_path, pst_store):
            logger.warning(
                "Primary Inbox on %s not ready (name=%r) — searching PST for a mail folder",
                pst_path,
                str(getattr(inbox, "Name", "") if inbox else ""),
            )
            inbox_key = (pst_path, OL_FOLDER_INBOX)
            if inbox_key in self._pst_standard_folder_cache:
                del self._pst_standard_folder_cache[inbox_key]
            inbox = self._discover_pst_mail_inbox(pst_store, pst_path) or inbox
            if inbox is not None and self._pst_folder_ready_for_import(
                inbox, pst_path, pst_store
            ):
                self._pst_standard_folder_cache[inbox_key] = inbox
        if not self._pst_folder_ready_for_import(inbox, pst_path, pst_store):
            raise RuntimeError(
                f"Could not access a writable Inbox on PST {pst_path}. "
                "Close Outlook, remove the PST from the profile, and try again — "
                "or delete the export PST and export to a new file."
            )
        name = str(getattr(inbox, "Name", "") or "").strip() or "Inbox"
        logger.info("PST import Inbox ready: %r (%s)", name, pst_path)

    def _create_mail_in_pst_target_folder(self, dest_folder, pst_path: str):
        """
        Create a mail item in the target PST (not the profile Outlook.pst).

        Tries Items.Add on dest_folder first (works when DefaultStore cannot be
        set), then store Inbox + Move as fallback.
        """
        pst_path = normalize_pst_path(pst_path)
        store = dest_folder.Store
        if not self._folder_belongs_to_pst(dest_folder, pst_path, store=store):
            raise RuntimeError(
                f"Target folder is not in export PST {pst_path}"
            )
        last_err: Exception | None = None
        for label, factory in (
            ("dest_folder.Items.Add", lambda: dest_folder.Items.Add(OL_MAIL_ITEM)),
            (
                "store Inbox Items.Add + Move",
                lambda: self._add_mail_via_store_inbox(store, dest_folder),
            ),
        ):
            try:
                mail = factory()
            except Exception as e:
                last_err = e
                logger.debug("%s failed: %s", label, e)
                continue
            if self._message_in_expected_pst(
                mail, pst_path, expected_store=store
            ):
                return mail
            try:
                mail.Delete()
            except Exception:
                pass
            last_err = RuntimeError(
                f"{label} created item outside export PST {pst_path}"
            )
        raise RuntimeError(
            f"Could not create mail in {pst_path}: {last_err}"
        )

    def _add_mail_via_store_inbox(self, store, dest_folder):
        """Create on the PST Inbox, then Move into dest_folder."""
        inbox = self._get_or_create_inbox(store)
        mail = inbox.Items.Add(OL_MAIL_ITEM)
        if (
            getattr(dest_folder, "EntryID", None)
            and getattr(inbox, "EntryID", None)
            and dest_folder.EntryID != inbox.EntryID
        ):
            mail = mail.Move(dest_folder)
        return mail

    def _coerce_import_folder(self, folder, conversion_options: dict):
        """Replace broken/empty Outlook folders with a real PST Inbox."""
        pst_path = normalize_pst_path(conversion_options.get("destination_path") or "")
        try:
            store = folder.Store
        except Exception:
            store = None
        if not pst_path or self._folder_is_usable_import_target(
            folder, pst_path, store=store
        ):
            return folder
        if store is None:
            return folder
        bad_name = str(getattr(folder, "Name", "") or "")
        logger.warning(
            "Invalid import folder %r on %s — using PST Inbox instead",
            bad_name or "(no name)",
            pst_path,
        )
        return self._get_or_create_inbox(store)

    def _pst_std_folder_cache_key(
        self, store, folder_id: int
    ) -> tuple[str, int] | None:
        pst_path = normalize_pst_path(str(getattr(store, "FilePath", "") or ""))
        if not pst_path:
            return None
        return (pst_path, folder_id)

    def _get_or_create_pst_standard_folder(self, store, folder_id: int):
        """Writable PST folder for a standard olFolder* id (Inbox, Sent, Outbox, ...)."""
        cache_key = self._pst_std_folder_cache_key(store, folder_id)
        if cache_key is not None:
            cached = self._pst_standard_folder_cache.get(cache_key)
            if cached is not None and self._is_usable_standard_folder(
                cached, folder_id, store
            ):
                return cached
            if cache_key in self._pst_standard_folder_cache:
                del self._pst_standard_folder_cache[cache_key]

        spec = _STANDARD_PST_FOLDER_SPECS.get(folder_id)
        if spec is None:
            try:
                candidate = store.GetDefaultFolder(folder_id)
                if candidate is not None and self._is_usable_standard_folder(
                    candidate, folder_id, store
                ):
                    if cache_key is not None:
                        self._pst_standard_folder_cache[cache_key] = candidate
                    return candidate
            except Exception as e:
                logger.debug("GetDefaultFolder(%s): %s", folder_id, e)
            return None
        aliases, create_name = spec
        pst_path = str(getattr(store, "FilePath", "") or "")
        root_folder = store.GetRootFolder()
        try:
            candidate = store.GetDefaultFolder(folder_id)
            if candidate is not None and self._is_usable_standard_folder(
                candidate, folder_id, store
            ):
                if cache_key is not None:
                    self._pst_standard_folder_cache[cache_key] = candidate
                return candidate
            if candidate is not None and (
                cache_key is None
                or cache_key not in self._pst_std_folder_warn_logged
            ):
                bad_name = str(getattr(candidate, "Name", "") or "").strip() or (
                    "(empty name)"
                )
                logger.warning(
                    "GetDefaultFolder(%s) on %s returned unusable folder %r — "
                    "locating or creating %r (logged once per PST/folder)",
                    folder_id,
                    pst_path or "PST",
                    bad_name,
                    create_name,
                )
                if cache_key is not None:
                    self._pst_std_folder_warn_logged.add(cache_key)
        except Exception as e:
            logger.debug("GetDefaultFolder(%s): %s", folder_id, e)
        found = self._find_folder_in_tree(root_folder, aliases, max_depth=3)
        if found is not None and self._is_usable_standard_folder(
            found, folder_id, store
        ):
            if (
                folder_id == OL_FOLDER_INBOX
                and cache_key is not None
                and cache_key not in self._pst_std_folder_resolve_logged
            ):
                logger.info(
                    "Using existing Inbox folder %r on %s",
                    str(getattr(found, "Name", "") or create_name),
                    pst_path or "PST",
                )
                self._pst_std_folder_resolve_logged.add(cache_key)
            if cache_key is not None:
                self._pst_standard_folder_cache[cache_key] = found
            return found
        try:
            created = root_folder.Folders.Add(create_name)
            if (
                folder_id == OL_FOLDER_INBOX
                and cache_key is not None
                and cache_key not in self._pst_std_folder_resolve_logged
            ):
                logger.warning(
                    "Created new %r on %s because Outlook's default Inbox mapping "
                    "was corrupted (search folder). Inbox mail will appear under "
                    "%s → %r → account subfolders.",
                    create_name,
                    pst_path or "PST",
                    os.path.splitext(os.path.basename(pst_path))[0] or "PST",
                    create_name,
                )
                self._pst_std_folder_resolve_logged.add(cache_key)
            if cache_key is not None:
                self._pst_standard_folder_cache[cache_key] = created
            return created
        except Exception as e:
            raise RuntimeError(
                f"Could not create standard folder {create_name!r} on "
                f"{pst_path or 'PST'}: {e}"
            ) from e

    def _get_or_create_inbox(self, store):
        """Writable Inbox on the PST (localized name or English 'Inbox' under store root)."""
        return self._get_or_create_pst_standard_folder(store, OL_FOLDER_INBOX)

    def _get_or_create_standard_folder(self, store, folder_id: int):
        return self._get_or_create_pst_standard_folder(store, folder_id)

    def _sent_state_for_source_path(
        self, source_file_path: str | None, conversion_options: dict
    ) -> bool | None:
        if not source_file_path:
            return None
        parts = relative_folder_parts(
            conversion_options.get("source_root", ""),
            source_file_path,
        )
        return sent_state_for_folder_parts(parts)

    def _resolve_mailbox_target_folder(self, store, folder_name: str):
        """
        Pick a writable folder for mailbox import.
        Exchange/M365 often reject new folders on the store root; use Inbox when needed.
        """
        safe_name = safe_outlook_folder_name(folder_name or "Imported EML")
        store_path = str(getattr(store, "FilePath", "") or "").strip().lower()
        is_pst_store = store_path.endswith(".pst")
        if is_pst_store:
            return self._get_or_create_named_folder(store.GetRootFolder(), safe_name)
        try:
            inbox = store.GetDefaultFolder(6)  # olFolderInbox
            return self._get_or_create_named_folder(inbox, safe_name)
        except Exception as e:
            logger.debug("Default Inbox unavailable, using store root: %s", e)
            return self._get_or_create_named_folder(store.GetRootFolder(), safe_name)

    def _get_or_create_named_folder(self, parent_folder, folder_name):
        """Find or create a subfolder by name under a mailbox or PST root."""
        safe_name = safe_outlook_folder_name(folder_name)
        try:
            subfolders = parent_folder.Folders
        except Exception as e:
            raise RuntimeError(
                f"Folder {getattr(parent_folder, 'Name', '?')!r} cannot contain "
                f"subfolders ({e})"
            ) from e
        try:
            count = int(subfolders.Count)
        except Exception:
            count = 0
        for idx in range(1, count + 1):
            try:
                folder = subfolders.Item(idx)
            except Exception:
                continue
            try:
                if str(folder.Name or "").lower() == safe_name.lower():
                    return folder
            except Exception:
                continue
        try:
            return subfolders.Add(safe_name)
        except Exception as e:
            for idx in range(1, count + 1):
                try:
                    folder = subfolders.Item(idx)
                    if str(folder.Name or "").lower() == safe_name.lower():
                        return folder
                except Exception:
                    continue
            parent_label = str(getattr(parent_folder, "Name", "") or "?")
            raise RuntimeError(
                f"Could not create subfolder {safe_name!r} under {parent_label}: {e}"
            ) from e
    
    def _export_targets_pst(self, conversion_options: dict) -> bool:
        """True when exporting into a .pst file (not Exchange mailbox)."""
        return (
            conversion_options.get("pst_option") != "mailbox"
            and bool((conversion_options.get("destination_path") or "").strip())
        )

    def _verify_mail_in_target_pst(
        self, mail_item, conversion_options: dict, *, source_file_path: str | None = None
    ) -> bool:
        pst_path = normalize_pst_path(conversion_options.get("destination_path") or "")
        if not pst_path:
            return True
        mail_store = None
        try:
            mail_store = mail_item.Parent.Store
        except Exception:
            pass
        if self._message_in_expected_pst(
            mail_item, pst_path, expected_store=mail_store
        ):
            return True
        try:
            parent_name = str(mail_item.Parent.Name)
        except Exception:
            parent_name = "?"
        self._wrong_store_warn_count += 1
        n = self._wrong_store_warn_count
        if n <= 3 or n % 50 == 0:
            logger.warning(
                "After import, message is in %s (%s), not %s — %s",
                parent_name or "(no name)",
                self._mail_item_store_path(mail_item) or "(profile store)",
                pst_path,
                log_sanitize(source_file_path or ""),
            )
        return False

    def _import_validation_enabled(self, conversion_options: dict) -> bool:
        if conversion_options.get("_skip_import_validation"):
            return False
        return bool(conversion_options.get("validate_import", VALIDATE_IMPORT))

    def _validate_imported_mail_item(
        self,
        mail,
        file_path: str,
        conversion_options: dict,
        email_data: dict | None = None,
        *,
        expected_folder=None,
    ):
        if email_data is None:
            email_data = self.parse_eml(file_path)
        summary = eml_parse_to_summary(email_data, file_path)
        snapshot = read_outlook_mail_snapshot(mail)
        inspection = validate_import_against_eml(
            snapshot, summary, source=file_path
        )
        if expected_folder is not None:
            self._append_folder_placement_validation(
                inspection, mail, expected_folder, conversion_options
            )
        return inspection

    def _repair_imported_mail_item(
        self,
        mail,
        file_path: str,
        conversion_options: dict,
        email_data: dict | None = None,
        *,
        expected_folder=None,
    ) -> None:
        if email_data is None:
            email_data = self.parse_eml(file_path)
        if expected_folder is not None:
            moved = self._relocate_mail_to_dest_folder(mail, expected_folder)
            if moved is not None:
                mail = moved
        if not email_data:
            try:
                mail.Save()
            except Exception:
                pass
            return
        self._apply_parsed_display_fields(
            mail, email_data, file_path, conversion_options
        )
        self._reapply_sender_from_email_data(
            mail, email_data, file_path, conversion_options
        )
        self._apply_attachment_chain(
            mail,
            email_data=email_data,
            source_file_path=file_path,
        )
        if expected_folder is not None:
            moved = self._relocate_mail_to_dest_folder(mail, expected_folder)
            if moved is not None:
                mail = moved
        try:
            mail.Save()
        except Exception as e:
            logger.debug("Save after import repair: %s", e)

    def _ensure_import_quality(
        self,
        mail,
        file_path: str,
        outlook,
        target_folder,
        conversion_options: dict,
        *,
        email_data: dict | None = None,
        from_validation_rebuild: bool = False,
        expected_folder=None,
    ) -> str:
        """
        Validate PST/mailbox item right after import; repair or rebuild if misaligned.

        Returns 'converted' or an error detail string (bad item removed when possible).
        """
        if not self._import_validation_enabled(conversion_options):
            return "converted"

        if email_data is None:
            email_data = self.parse_eml(file_path)

        inspection = self._validate_imported_mail_item(
            mail,
            file_path,
            conversion_options,
            email_data,
            expected_folder=expected_folder,
        )
        if not has_import_errors(inspection):
            return "converted"

        detail = format_inspection_errors(inspection)
        logger.warning(
            "Import validation failed for %s — repairing (%s)",
            log_sanitize(file_path),
            detail,
        )
        self._repair_imported_mail_item(
            mail,
            file_path,
            conversion_options,
            email_data,
            expected_folder=expected_folder,
        )
        inspection = self._validate_imported_mail_item(
            mail,
            file_path,
            conversion_options,
            email_data,
            expected_folder=expected_folder,
        )
        if not has_import_errors(inspection):
            self._import_validation_repairs = (
                getattr(self, "_import_validation_repairs", 0) + 1
            )
            logger.info(
                "Import validation repaired: %s",
                log_sanitize(file_path),
            )
            return "converted"

        detail = format_inspection_errors(inspection)
        try:
            mail.Delete()
        except Exception as e:
            logger.debug("Delete invalid import before rebuild: %s", e)
        release_com_object(mail)
        if (
            not from_validation_rebuild
            and self._export_targets_pst(conversion_options)
            and outlook is not None
        ):
            rebuild_opts = dict(conversion_options)
            rebuild_opts["_skip_import_validation"] = False
            rebuild = self._process_single_email_manual_fallback(
                file_path,
                target_folder,
                outlook,
                require_date_preservation=bool(
                    conversion_options.get("strict_date_preservation")
                ),
                conversion_options=rebuild_opts,
                try_staged_first=True,
                from_validation_rebuild=True,
            )
            if rebuild == "converted":
                self._import_validation_rebuilds = (
                    getattr(self, "_import_validation_rebuilds", 0) + 1
                )
                logger.info(
                    "Import rebuilt after validation: %s",
                    log_sanitize(file_path),
                )
                return "converted"
            return f"Import validation failed ({detail}); rebuild: {rebuild}"

        self._import_validation_failures = (
            getattr(self, "_import_validation_failures", 0) + 1
        )
        return f"Import validation failed: {detail}"

    def _import_eml_direct_to_pst_folder(
        self,
        file_path: str,
        target_folder,
        outlook,
        conversion_options: dict,
    ) -> str:
        """
        Create the message with Items.Add in the target PST folder.

        OpenSharedItem always opens in the profile Drafts store on 32-bit Outlook;
        Copy/Move into an attached PST is unreliable. Direct creation avoids that.
        Returns 'converted' or an error detail string.
        """
        if not self._pst_direct_import_logged:
            logger.info(
                "PST export: staged OpenSharedItem first; Items.Add fallback if staging fails"
            )
            self._pst_direct_import_logged = True
        target_folder = self._coerce_import_folder(target_folder, conversion_options)
        namespace = outlook.GetNamespace("MAPI")
        staged = self._try_staged_eml_import_with_dates(
            namespace,
            target_folder,
            file_path,
            conversion_options,
            outlook=outlook,
        )
        if staged == "converted":
            return "converted"
        logger.debug(
            "Staged PST import failed for %s (%s); using Items.Add fallback",
            log_sanitize(file_path),
            log_sanitize(staged),
        )
        return self._process_single_email_manual_fallback(
            file_path,
            target_folder,
            outlook,
            require_date_preservation=bool(
                conversion_options.get("strict_date_preservation")
            ),
            conversion_options=conversion_options,
            try_staged_first=False,
        )

    def _resolve_export_target_folder(self, namespace, conversion_options):
        """Resolve the base import folder for PST or mailbox export."""
        export_mode = conversion_options.get("pst_option", "new")
        pst_path = conversion_options.get("destination_path", "")
        if export_mode == "mailbox":
            store_label = conversion_options.get("mailbox_store_label", "")
            folder_name = conversion_options.get("mailbox_folder_name", "Imported EML")
            store = self._find_store_by_label(namespace, store_label)
            if not store:
                raise RuntimeError(f"Could not find Outlook mailbox: {store_label}")
            return self._resolve_mailbox_target_folder(store, folder_name)
        pst_store = self._find_pst_store(namespace, pst_path)
        if not pst_store:
            raise RuntimeError(f"Could not access PST file. Path: {pst_path}")
        return self._get_or_create_inbox(pst_store)

    def _clear_pst_session_caches(self) -> None:
        """Drop cached Outlook folder COM objects after reconnect or store churn."""
        self._pst_standard_folder_cache.clear()
        self._pst_std_folder_warn_logged.clear()
        self._pst_std_folder_resolve_logged.clear()

    def _outlook_folder_display_path(self, folder) -> str:
        """Build a human-readable folder path (e.g. Inbox\\Account (user@domain))."""
        parts: list[str] = []
        current = folder
        for _ in range(40):
            try:
                name = str(current.Name or "").strip()
            except Exception:
                break
            if name:
                parts.append(name)
            try:
                current = current.Parent
            except Exception:
                break
            if current is None:
                break
        parts.reverse()
        return "\\".join(parts)

    def _mail_in_dest_folder(self, mail, dest_folder) -> tuple[bool, str, str]:
        """True when the item's parent folder is the intended import folder."""
        try:
            parent = mail.Parent
            if (
                getattr(parent, "EntryID", None)
                and getattr(dest_folder, "EntryID", None)
                and parent.EntryID == dest_folder.EntryID
            ):
                path = self._outlook_folder_display_path(parent)
                return True, path, path
        except Exception:
            pass
        try:
            actual = self._outlook_folder_display_path(mail.Parent)
        except Exception:
            actual = "?"
        try:
            expected = self._outlook_folder_display_path(dest_folder)
        except Exception:
            expected = "?"
        if actual == expected:
            return True, actual, expected
        return False, actual, expected

    def _append_folder_placement_validation(
        self,
        inspection,
        mail,
        dest_folder,
        conversion_options: dict,
    ) -> None:
        pst_path = normalize_pst_path(conversion_options.get("destination_path") or "")
        in_store = True
        actual_store = self._mail_item_store_path(mail)
        if pst_path and self._export_targets_pst(conversion_options):
            in_store = self._message_in_expected_pst(mail, pst_path)
        in_folder, actual_path, expected_path = self._mail_in_dest_folder(
            mail, dest_folder
        )
        append_folder_placement_issues(
            inspection,
            in_correct_folder=in_folder,
            actual_folder_path=actual_path,
            expected_folder_path=expected_path,
            in_correct_store=in_store,
            actual_store=actual_store or "?",
            expected_store=pst_path or "",
        )

    def _probe_outlook_application(self, outlook, conversion_options: dict | None = None) -> bool:
        """False when the Outlook COM server is not responding."""
        if conversion_options and conversion_options.get("_outlook_recovering"):
            return True
        try:
            _ = outlook.Application.Version
            namespace = outlook.GetNamespace("MAPI")
            _ = namespace.Stores.Count
            return True
        except Exception:
            return False

    def _dispatch_outlook_application(self):
        """Attach to a running Outlook or start one; raises on total failure."""
        last_err = None
        for attempt in range(3):
            try:
                try:
                    outlook = WIN32COM.GetActiveObject("Outlook.Application")
                except Exception:
                    outlook = WIN32COM.Dispatch("Outlook.Application")
                if self._probe_outlook_application(outlook):
                    return outlook
            except Exception as e:
                last_err = e
            time.sleep(COM_RPC_COOLDOWN * (attempt + 1))
        if last_err:
            raise last_err
        raise RuntimeError("Could not connect to Outlook")

    def _ui_set_status(self, text: str) -> None:
        try:
            self.root.after(0, lambda: self.status_label.config(text=text))
        except Exception:
            pass

    def _recover_outlook_session(self, conversion_options, *, lang: str | None = None):
        """
        Wait for Outlook to recover after crash/hang, then reconnect COM.

        Export pauses in a loop until Outlook responds or the user cancels.
        """
        with self._outlook_recovery_lock:
            return self._recover_outlook_session_locked(conversion_options, lang=lang)

    def _recover_outlook_session_locked(
        self, conversion_options, *, lang: str | None = None
    ):
        if lang is None:
            lang = conversion_options.get("lang", LANG_EN)
        conversion_options["_outlook_recovering"] = True
        try:
            return self._recover_outlook_session_impl(conversion_options, lang=lang)
        finally:
            conversion_options["_outlook_recovering"] = False

    def _recover_outlook_session_impl(self, conversion_options, *, lang: str):
        logger.warning(
            "Outlook unavailable — export paused until Outlook is healthy "
            "(max wait %ds, poll every %.0fs)",
            OUTLOOK_CRASH_WAIT_MAX,
            OUTLOOK_CRASH_WAIT_POLL,
        )
        self._ui_set_status(t(lang, "status_waiting_outlook"))
        self._flush_export_log()
        deadline = time.monotonic() + OUTLOOK_CRASH_WAIT_MAX
        last_log = 0.0
        while time.monotonic() < deadline:
            if self._is_cancelled():
                raise RuntimeError("Cancel requested while waiting for Outlook")
            if not _outlook_process_running():
                now = time.monotonic()
                if now - last_log >= 30.0:
                    logger.warning(
                        "Outlook.exe is not running — start or restart Outlook; "
                        "export will continue automatically when it is back"
                    )
                    last_log = now
            else:
                try:
                    gc.collect()
                    self._clear_pst_session_caches()
                    outlook = self._dispatch_outlook_application()
                    try:
                        outlook.DisplayAlerts = False
                    except Exception:
                        pass
                    namespace = outlook.GetNamespace("MAPI")
                    target_folder = self._resolve_export_target_folder(
                        namespace, conversion_options
                    )
                    logger.info("Outlook reconnected — resuming export")
                    self._ui_set_status(t(lang, "status_outlook_resumed"))
                    conversion_options["_session_outlook"] = outlook
                    return outlook, namespace, target_folder
                except Exception as e:
                    now = time.monotonic()
                    if now - last_log >= 30.0:
                        logger.warning(
                            "Outlook not ready yet (%s) — still waiting...",
                            log_sanitize(e),
                        )
                        last_log = now
            time.sleep(OUTLOOK_CRASH_WAIT_POLL)
        raise RuntimeError(
            f"Outlook did not recover within {OUTLOOK_CRASH_WAIT_MAX}s — "
            "restart Outlook and resume the export"
        )

    def _refresh_outlook_session_after_failure(
        self,
        session: dict,
        conversion_options: dict,
        folder_cache: dict,
        *,
        lang: str,
        consecutive_failures: int = 1,
    ) -> None:
        extra_pause = COM_RPC_COOLDOWN * min(consecutive_failures, 4)
        if extra_pause > COM_RPC_COOLDOWN:
            logger.warning(
                "Repeated Outlook failures (%d) — extra pause %.1fs",
                consecutive_failures,
                extra_pause,
            )
            time.sleep(extra_pause - COM_RPC_COOLDOWN)
        session["outlook"], session["namespace"], session["target_folder"] = (
            self._recover_outlook_session(conversion_options, lang=lang)
        )
        try:
            session["outlook"].DisplayAlerts = False
        except Exception:
            pass
        conversion_options["_session_outlook"] = session["outlook"]
        folder_cache.clear()

    def _should_use_pst_chunks(self, conversion_options: dict, file_count: int) -> bool:
        """True when importing into a PST in batches (direct import into final PST)."""
        if conversion_options.get("pst_option") == "mailbox":
            return False
        chunk_size = int(conversion_options.get("pst_chunk_size") or 0)
        return chunk_size > 0 and file_count > chunk_size

    def _convert_with_pst_chunks(
        self,
        outlook,
        namespace,
        files_to_process: list[str],
        conversion_options: dict,
        lang: str,
        final_pst_path: str,
    ):
        """
        Batched import directly into the final PST (preserves dates; no chunk merge).
        Chain: EML batch -> final PST, repeated every pst_chunk_size messages.
        """
        chunk_size = int(conversion_options["pst_chunk_size"])
        total = len(files_to_process)
        n_batches = (total + chunk_size - 1) // chunk_size

        converted = skipped = errors = 0
        error_messages: list[str] = []
        skipped_messages: list[str] = []
        cancelled = False
        completed_batches = 0
        folder_cache: dict = {}
        batch_opts = dict(conversion_options)
        batch_opts["destination_path"] = final_pst_path
        note = ""

        logger.info(
            "Batched PST export: %d files in %d batch(es) of up to %d into %s",
            total,
            n_batches,
            chunk_size,
            final_pst_path,
        )

        self._update_status(t(lang, "status_creating_pst"))
        self._cleanup_stale_store_for_path(namespace, final_pst_path)
        final_pst_path = self._ensure_pst_store_attached(
            outlook, namespace, final_pst_path, batch_opts
        )
        batch_opts["destination_path"] = final_pst_path
        pst_store = self._find_pst_store(namespace, final_pst_path)
        if not pst_store:
            raise RuntimeError(
                f"Could not access final PST for batched export: {final_pst_path}"
            )
        self._log_attached_pst_store(pst_store, final_pst_path)
        target_folder = self._get_or_create_inbox(pst_store)

        try:
            for batch_idx in range(n_batches):
                if self._is_cancelled():
                    cancelled = True
                    break

                batch_start = batch_idx * chunk_size
                batch = files_to_process[batch_start : batch_start + chunk_size]
                batch_num = batch_idx + 1

                self._update_status(
                    t(lang, "status_chain_import", cur=batch_num, total=n_batches)
                )

                c, s, e, em, sm, can = self._process_email_files(
                    outlook,
                    namespace,
                    target_folder,
                    batch,
                    len(batch),
                    batch_opts,
                    folder_cache,
                    progress_offset=batch_start,
                    progress_total=total,
                )
                converted += c
                skipped += s
                errors += e
                error_messages.extend(em)
                skipped_messages.extend(sm)
                completed_batches += 1

                if can:
                    cancelled = True
                    break

                gc.collect()
                time.sleep(max(0.0, _NATIVE_POST_BATCH_DELAY))

            note = ""
            if completed_batches:
                note += t(
                    lang,
                    "note_pst_chained",
                    links=completed_batches,
                    size=chunk_size,
                )
                if cancelled:
                    note += t(lang, "note_pst_partial_merge")
                note += t(lang, "note_pst_saved", path=final_pst_path)
            elif cancelled:
                note += t(
                    lang,
                    "note_chain_cancelled",
                    dir=os.path.dirname(os.path.abspath(final_pst_path)),
                )
        except Exception:
            logger.exception("Batched PST export failed for %s", final_pst_path)
            raise

        return (
            converted,
            skipped,
            errors,
            error_messages,
            skipped_messages,
            cancelled,
            note,
        )

    def _merge_chunk_pst_into_final(
        self, namespace, chunk_pst_path: str, final_pst_path: str
    ) -> None:
        """Attach a chunk PST and copy all mail items into the final PST."""
        self._remove_existing_store(namespace, chunk_pst_path)
        namespace.AddStore(chunk_pst_path)
        time.sleep(0.5)
        chunk_store = self._find_pst_store(namespace, chunk_pst_path)
        final_store = self._find_pst_store(namespace, final_pst_path)
        if not chunk_store or not final_store:
            raise RuntimeError(
                f"Cannot merge chunk PST (chunk={chunk_store}, final={final_store})"
            )
        try:
            chunk_inbox = self._get_or_create_inbox(chunk_store)
            final_inbox = self._get_or_create_inbox(final_store)
            copied, failed = self._copy_outlook_folder_tree(chunk_inbox, final_inbox)
            logger.info(
                "Merged %s into %s (%d items copied, %d failed)",
                chunk_pst_path,
                final_pst_path,
                copied,
                failed,
            )
        finally:
            self._remove_existing_store(namespace, chunk_pst_path)

    def _copy_outlook_folder_tree(self, source_folder, dest_folder) -> tuple[int, int]:
        """Move mail items and subfolders from source_folder into dest_folder."""
        copied = 0
        failed = 0
        try:
            items = source_folder.Items
        except Exception as e:
            logger.warning("Could not read folder items: %s", e)
            items = None

        if items is not None:
            try:
                count = int(items.Count)
            except Exception:
                count = 0
            for idx in range(count, 0, -1):
                try:
                    item = items.Item(idx)
                    if int(item.Class) != OL_MAIL_CLASS:
                        continue
                    try:
                        item.Move(dest_folder)
                    except Exception:
                        item.Copy(dest_folder)
                    copied += 1
                except Exception as e:
                    failed += 1
                    logger.debug("Could not move/copy mail item %d: %s", idx, e)

        try:
            subfolders = source_folder.Folders
            sub_count = int(subfolders.Count)
        except Exception:
            sub_count = 0

        for idx in range(1, sub_count + 1):
            try:
                sub = subfolders.Item(idx)
                name = str(sub.Name)
                dest_sub = self._get_or_create_named_folder(dest_folder, name)
                sub_copied, sub_failed = self._copy_outlook_folder_tree(sub, dest_sub)
                copied += sub_copied
                failed += sub_failed
            except Exception as e:
                failed += 1
                logger.warning("Could not merge subfolder index %d: %s", idx, e)

        return copied, failed

    def _cleanup_chunk_directory(self, chunk_dir: str, chunk_paths: list[str]) -> None:
        for path in chunk_paths:
            try:
                if os.path.isfile(path):
                    os.remove(path)
            except OSError as e:
                logger.debug("Could not remove chunk PST %s: %s", path, e)
        try:
            if os.path.isdir(chunk_dir) and not os.listdir(chunk_dir):
                os.rmdir(chunk_dir)
        except OSError:
            pass

    def _prep_eml_without_com(
        self, file_path: str, conversion_options: dict
    ) -> EmlPrepResult:
        """Disk/CPU prep for one .eml (no Outlook COM). Used by parallel prefetch workers."""
        norm_path = os.path.normpath(os.path.abspath(file_path))
        try:
            sz = self._eml_file_sizes.get(norm_path)
            if sz is None:
                sz = os.path.getsize(file_path)
        except OSError as e:
            return EmlPrepResult(
                file_path=file_path,
                norm_path=norm_path,
                size=0,
                prep_error=f"Cannot read file: {e}",
            )
        dedup_key = None
        header_date = ""
        if conversion_options.get("remove_duplicates"):
            dedup_pair = self._dedup_key_for_file(file_path, sz)
            if dedup_pair is None:
                if FULL_DEDUP_HASH:
                    return EmlPrepResult(
                        file_path=file_path,
                        norm_path=norm_path,
                        size=sz,
                        prep_error="Cannot compute duplicate hash",
                    )
                return EmlPrepResult(
                    file_path=file_path,
                    norm_path=norm_path,
                    size=sz,
                    prep_error="Cannot compute duplicate fingerprint",
                )
            dedup_key, header_date = dedup_pair
        email_data = None
        if conversion_options.get("parallel_parse"):
            email_data = self.parse_eml(file_path)
        return EmlPrepResult(
            file_path=file_path,
            norm_path=norm_path,
            size=sz,
            dedup_key=dedup_key,
            header_date=header_date,
            email_data=email_data,
        )

    def _com_pipeline_worker_start(self, conversion_options: dict, folder_cache: dict) -> dict:
        """Initialize Outlook COM on the pipeline worker thread (STA)."""
        import pythoncom

        pythoncom.CoInitialize()
        ctx: dict = {"com_initialized": True, "folder_cache": folder_cache}
        try:
            outlook = self._dispatch_outlook_application()
            try:
                outlook.DisplayAlerts = False
            except Exception:
                pass
            namespace = outlook.GetNamespace("MAPI")
            export_mode = conversion_options.get("pst_option", "new")
            if export_mode != "mailbox":
                pst_path = normalize_pst_path(
                    conversion_options.get("destination_path", "")
                )
                pst_path = self._ensure_pst_store_attached(
                    outlook, namespace, pst_path, conversion_options
                )
                conversion_options["destination_path"] = pst_path
                pst_store = self._find_pst_store(namespace, pst_path)
                if pst_store:
                    self._log_attached_pst_store(pst_store, pst_path)
            target_folder = self._resolve_export_target_folder(namespace, conversion_options)
            conversion_options["_session_outlook"] = outlook
            ctx["session"] = {
                "outlook": outlook,
                "namespace": namespace,
                "target_folder": target_folder,
            }
            return ctx
        except Exception:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            raise

    def _com_pipeline_worker_stop(self, ctx: dict | None) -> None:
        if not ctx or not ctx.get("com_initialized"):
            return
        import pythoncom

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    def _com_pipeline_import_one(
        self,
        ctx: dict,
        file_path: str,
        prep: EmlPrepResult | None,
        conversion_options: dict,
    ) -> ImportResult:
        """Run one import on the COM worker thread."""
        session = ctx["session"]
        folder_cache = ctx["folder_cache"]
        lang = conversion_options.get("lang", LANG_EN)
        conversion_options["_prep_result"] = prep
        conversion_options["_parsed_email_data"] = prep.email_data if prep else None
        consecutive_rpc_failures = 0
        try:
            if not self._probe_outlook_application(
                session["outlook"], conversion_options
            ):
                self._refresh_outlook_session_after_failure(
                    session,
                    conversion_options,
                    folder_cache,
                    lang=lang,
                    consecutive_failures=1,
                )
            dest_folder, target_label = self._resolve_target_folder_for_file(
                session["target_folder"], file_path, conversion_options, folder_cache
            )
            status, detail = self._process_single_email(
                session["namespace"],
                dest_folder,
                file_path,
                session["outlook"],
                conversion_options,
            )
            if status == "error" and _is_outlook_unavailable_error(detail):
                consecutive_rpc_failures += 1
                self._refresh_outlook_session_after_failure(
                    session,
                    conversion_options,
                    folder_cache,
                    lang=lang,
                    consecutive_failures=consecutive_rpc_failures,
                )
                dest_folder, target_label = self._resolve_target_folder_for_file(
                    session["target_folder"], file_path, conversion_options, folder_cache
                )
                status, detail = self._process_single_email(
                    session["namespace"],
                    dest_folder,
                    file_path,
                    session["outlook"],
                    conversion_options,
                )
            return ImportResult(
                path=file_path,
                success=status != "error",
                status=status,
                detail=detail or "",
                target_label=target_label,
            )
        except Exception as exc:
            return ImportResult(
                path=file_path,
                success=False,
                status="error",
                detail=str(exc),
                exception=exc,
            )
        finally:
            conversion_options.pop("_prep_result", None)
            conversion_options.pop("_parsed_email_data", None)

    def _process_email_files_com_pipeline(
        self,
        files_to_process,
        total,
        conversion_options,
        folder_cache,
        *,
        progress_offset: int = 0,
        progress_total: int | None = None,
    ):
        """Prep on this thread; Outlook import on a single dedicated COM worker."""
        if progress_total is None:
            progress_total = total
        converted = skipped = errors = 0
        error_messages: list[str] = []
        skipped_messages: list[str] = []
        cancelled = False
        lang = conversion_options.get("lang", LANG_EN)
        rate_limiter = conversion_options.get("_rate_limiter")
        recovery = conversion_options.get("_recovery")
        conversion_options.setdefault("_export_started_mono", time.monotonic())
        last_conv_ui = [0.0]
        parallel_workers = int(conversion_options.get("parallel_workers") or 0)
        prefetcher = None
        worker_ctx: dict = {}

        def _import_one(path: str, prep):
            return self._com_pipeline_import_one(
                worker_ctx, path, prep, conversion_options
            )

        def _on_worker_start():
            nonlocal worker_ctx
            worker_ctx = self._com_pipeline_worker_start(conversion_options, folder_cache)
            return worker_ctx

        pipeline = ComImportPipeline(
            import_one=_import_one,
            on_worker_start=_on_worker_start,
            on_worker_stop=self._com_pipeline_worker_stop,
        )
        pipeline.start()
        if pipeline.worker_error:
            raise RuntimeError(f"COM pipeline failed to start: {pipeline.worker_error}")

        if parallel_workers > 0:
            prefetcher = create_eml_prep_prefetcher(
                lambda p, opts=conversion_options: self._prep_eml_without_com(p, opts),
                max_workers=parallel_workers,
            )
            prefetcher.start(files_to_process)

        try:
            for i, file_path in enumerate(files_to_process):
                if self._is_cancelled():
                    cancelled = True
                    break
                norm_path = os.path.normpath(os.path.abspath(file_path))
                prep_result = None
                if prefetcher is not None:
                    prep_result = prefetcher.take(norm_path)
                pipeline.submit(file_path, prep_result)
                result = pipeline.get_result(timeout=600.0)
                if result is None:
                    errors += 1
                    error_messages.append(f"{os.path.basename(file_path)}: COM pipeline timeout")
                    continue
                current_file = os.path.basename(file_path)
                started = time.perf_counter()
                status = result.status or "error"
                detail = result.detail or ""
                target_label = result.target_label or ""
                if result.exception:
                    status = "error"
                    detail = str(result.exception)

                duration = time.perf_counter() - started
                if status == "converted":
                    converted += 1
                    if rate_limiter:
                        rate_limiter.record_success()
                elif status == "skipped":
                    skipped += 1
                    if detail:
                        skipped_messages.append(f"{current_file}: {detail}")
                else:
                    errors += 1
                    error_messages.append(f"{current_file}: {detail or 'Unknown error'}")
                    if rate_limiter:
                        rate_limiter.record_failure()
                self._write_csv_row(file_path, status, detail or "", duration, target_label)
                done = i + 1
                if done == 1 or done % CONVERSION_UI_EVERY == 0 or done == total:
                    overall_done = progress_offset + done
                    now = time.monotonic()
                    started_mono = conversion_options.get("_export_started_mono") or now
                    eta = format_export_eta(now - started_mono, overall_done, progress_total)
                    self.root.after(
                        0,
                        lambda f=current_file, cur=overall_done, tot=progress_total, lg=lang, e=eta: self.status_label.config(
                            text=(
                                t(lg, "status_converting_eta", name=f, cur=cur, total=tot, eta=e)
                                if e
                                else t(lg, "status_converting", name=f, cur=cur, total=tot)
                            )
                        ),
                    )
                    self.root.after(0, lambda v=overall_done: self.progress.config(value=v))
                if recovery and CHECKPOINT_EVERY > 0 and done % CHECKPOINT_EVERY == 0:
                    recovery.save_checkpoint(
                        {
                            "app_version": APP_VERSION,
                            "processed_count": progress_offset + done,
                            "converted": converted,
                            "skipped": skipped,
                            "errors": errors,
                            "last_file": file_path,
                            "cancelled": False,
                        }
                    )
                if rate_limiter:
                    rate_limiter.wait_if_needed()
                elif COM_PACE_SEC > 0:
                    time.sleep(COM_PACE_SEC)
        finally:
            pipeline.shutdown()
            if prefetcher is not None:
                prefetcher.shutdown(cancel_pending=cancelled)
        return converted, skipped, errors, error_messages, skipped_messages, cancelled

    def _process_email_files(
        self,
        outlook,
        namespace,
        target_folder,
        files_to_process,
        total,
        conversion_options,
        folder_cache,
        *,
        progress_offset: int = 0,
        progress_total: int | None = None,
    ):
        """Process all email files"""
        if conversion_options.get("use_com_pipeline"):
            return self._process_email_files_com_pipeline(
                files_to_process,
                total,
                conversion_options,
                folder_cache,
                progress_offset=progress_offset,
                progress_total=progress_total,
            )
        if progress_total is None:
            progress_total = total
        converted = 0
        skipped = 0
        errors = 0
        error_messages = []
        skipped_messages = []
        cancelled = False
        
        lang = conversion_options.get("lang", LANG_EN)
        session = {
            "outlook": outlook,
            "namespace": namespace,
            "target_folder": target_folder,
        }
        consecutive_rpc_failures = 0
        last_conv_ui = [0.0]
        rate_limiter = conversion_options.get("_rate_limiter")
        recovery = conversion_options.get("_recovery")
        conversion_options.setdefault("_export_started_mono", time.monotonic())
        logger.info(
            "Processing %d file(s); overall progress target %d",
            total,
            progress_total,
        )
        self._flush_export_log()
        lang = conversion_options.get("lang", LANG_EN)
        parallel_workers = int(conversion_options.get("parallel_workers") or 0)
        prefetcher = None
        if parallel_workers > 0:
            logger.info(
                "Parallel prep: %d worker(s) for fingerprints%s "
                "(Outlook COM import remains single-threaded)",
                parallel_workers,
                " + MIME parse" if conversion_options.get("parallel_parse") else "",
            )
            prefetcher = create_eml_prep_prefetcher(
                lambda p, opts=conversion_options: self._prep_eml_without_com(p, opts),
                max_workers=parallel_workers,
            )
            prefetcher.start(files_to_process)
        try:
            for i, file_path in enumerate(files_to_process):
                if self._is_cancelled():
                    cancelled = True
                    logger.info("Cancel requested — stopping after %d/%d", i, total)
                    break

                norm_path = os.path.normpath(os.path.abspath(file_path))
                prep_result = None
                if prefetcher is not None:
                    prep_result = prefetcher.take(norm_path)
                conversion_options["_prep_result"] = prep_result
                conversion_options["_parsed_email_data"] = (
                    prep_result.email_data if prep_result else None
                )

                if not self._probe_outlook_application(
                    session["outlook"], conversion_options
                ):
                    try:
                        self._refresh_outlook_session_after_failure(
                            session,
                            conversion_options,
                            folder_cache,
                            lang=lang,
                            consecutive_failures=consecutive_rpc_failures + 1,
                        )
                        consecutive_rpc_failures = 0
                    except Exception as wait_err:
                        errors += 1
                        error_messages.append(
                            f"{os.path.basename(file_path)}: {wait_err}"
                        )
                        logger.error("Outlook recovery failed: %s", wait_err)
                        break

                current_file = os.path.basename(file_path)
                started = time.perf_counter()
                target_label = ""
                try:
                    dest_folder, target_label = self._resolve_target_folder_for_file(
                        session["target_folder"], file_path, conversion_options, folder_cache
                    )
                    status, detail = self._process_single_email(
                        session["namespace"],
                        dest_folder,
                        file_path,
                        session["outlook"],
                        conversion_options,
                    )
                    if status == "error" and _is_outlook_unavailable_error(detail):
                        consecutive_rpc_failures += 1
                        try:
                            self._refresh_outlook_session_after_failure(
                                session,
                                conversion_options,
                                folder_cache,
                                lang=lang,
                                consecutive_failures=consecutive_rpc_failures,
                            )
                        except Exception as wait_err:
                            errors += 1
                            error_messages.append(f"{current_file}: {wait_err}")
                            logger.error("Outlook recovery failed: %s", wait_err)
                            break
                        dest_folder, target_label = self._resolve_target_folder_for_file(
                            session["target_folder"], file_path, conversion_options, folder_cache
                        )
                        status, detail = self._process_single_email(
                            session["namespace"],
                            dest_folder,
                            file_path,
                            session["outlook"],
                            conversion_options,
                        )
                    duration = time.perf_counter() - started
                    if SLOW_FILE_WARN_SEC > 0 and duration >= SLOW_FILE_WARN_SEC:
                        logger.warning(
                            "Slow message (%.1fs): %s",
                            duration,
                            log_sanitize(file_path),
                        )
                    if status == "converted":
                        consecutive_rpc_failures = 0
                        converted += 1
                        if rate_limiter:
                            rate_limiter.record_success()
                    elif status == "skipped":
                        skipped += 1
                        if detail:
                            skipped_messages.append(f"{current_file}: {detail}")
                    else:
                        errors += 1
                        error_messages.append(f"{current_file}: {detail or 'Unknown error'}")
                        if rate_limiter:
                            rate_limiter.record_failure()
                    self._write_csv_row(
                        file_path, status, detail or "", duration, target_label
                    )
                except Exception as e:
                    duration = time.perf_counter() - started
                    if _is_outlook_unavailable_error(e):
                        consecutive_rpc_failures += 1
                        try:
                            self._refresh_outlook_session_after_failure(
                                session,
                                conversion_options,
                                folder_cache,
                                lang=lang,
                                consecutive_failures=consecutive_rpc_failures,
                            )
                            dest_folder, target_label = self._resolve_target_folder_for_file(
                                session["target_folder"],
                                file_path,
                                conversion_options,
                                folder_cache,
                            )
                            status, detail = self._process_single_email(
                                session["namespace"],
                                dest_folder,
                                file_path,
                                session["outlook"],
                                conversion_options,
                            )
                            duration = time.perf_counter() - started
                            if status == "converted":
                                converted += 1
                            elif status == "skipped":
                                skipped += 1
                                if detail:
                                    skipped_messages.append(f"{current_file}: {detail}")
                            else:
                                errors += 1
                                error_messages.append(
                                    f"{current_file}: {detail or 'Unknown error'}"
                                )
                            self._write_csv_row(
                                file_path, status, detail or "", duration, target_label
                            )
                            continue
                        except Exception as retry_err:
                            e = retry_err
                    errors += 1
                    error_messages.append(f"{current_file}: {e}")
                    logger.error("Error processing %s: %s", log_sanitize(file_path), log_sanitize(e))
                    self._write_csv_row(file_path, "error", str(e), duration, target_label)
                finally:
                    conversion_options.pop("_prep_result", None)
                    conversion_options.pop("_parsed_email_data", None)

                done = i + 1
                report_progress = (
                    done == 1 or done % CONVERSION_UI_EVERY == 0 or done == total
                )
                if report_progress:
                    overall_done = progress_offset + done
                    logger.info(
                        "Progress %d/%d — %s",
                        overall_done,
                        progress_total,
                        log_sanitize(current_file),
                    )
                    self._flush_export_log()
                    force_ui = done == 1 or done == total
                    now = time.monotonic()
                    if force_ui or (now - last_conv_ui[0]) >= UI_THROTTLE_SEC:
                        last_conv_ui[0] = now
                        started_mono = conversion_options.get("_export_started_mono") or now
                        eta = format_export_eta(now - started_mono, overall_done, progress_total)
                        self.root.after(
                            0,
                            lambda f=current_file, cur=overall_done, tot=progress_total, lg=lang, e=eta: self.status_label.config(
                                text=(
                                    t(lg, "status_converting_eta", name=f, cur=cur, total=tot, eta=e)
                                    if e
                                    else t(lg, "status_converting", name=f, cur=cur, total=tot)
                                )
                            ),
                        )
                        self.root.after(
                            0, lambda v=overall_done: self.progress.config(value=v)
                        )

                if recovery and CHECKPOINT_EVERY > 0 and done % CHECKPOINT_EVERY == 0:
                    recovery.save_checkpoint(
                        {
                            "app_version": APP_VERSION,
                            "processed_count": progress_offset + done,
                            "converted": converted,
                            "skipped": skipped,
                            "errors": errors,
                            "last_file": file_path,
                            "cancelled": False,
                        }
                    )

                if GC_EVERY_N_FILES and done % GC_EVERY_N_FILES == 0:
                    gc.collect()

                if rate_limiter:
                    rate_limiter.wait_if_needed()
                elif COM_PACE_SEC > 0:
                    time.sleep(COM_PACE_SEC)
        finally:
            if prefetcher is not None:
                prefetcher.shutdown(cancel_pending=cancelled)

        return converted, skipped, errors, error_messages, skipped_messages, cancelled
    
    def _process_single_email(self, namespace, target_folder, file_path, outlook, conversion_options):
        """Process a single email file"""
        source_root = (conversion_options.get("source_root") or "").strip()
        if source_root:
            try:
                file_path = validate_import_path(source_root, file_path)
            except ValueError as exc:
                return "error", str(exc)
        norm_path = os.path.normpath(os.path.abspath(file_path))
        resume_paths = conversion_options.get("resume_paths") or set()
        if norm_path in resume_paths:
            return "skipped", t(
                conversion_options.get("lang", LANG_EN), "skip_resume"
            )

        prep: EmlPrepResult | None = conversion_options.get("_prep_result")
        try:
            if prep is not None:
                if prep.prep_error:
                    return "error", prep.prep_error
                sz = prep.size
            else:
                sz = self._eml_file_sizes.get(norm_path)
                if sz is None:
                    sz = os.path.getsize(file_path)
            if sz > MAX_EML_FILE_BYTES:
                mb = MAX_EML_FILE_BYTES // (1024 * 1024)
                return (
                    "error",
                    f"File too large ({sz // (1024 * 1024)} MB); max {mb} MB (set EML2PST_MAX_FILE_MB)",
                )
        except OSError as e:
            return "error", f"Cannot read file: {e}"

        email_filter = conversion_options.get("email_filter")
        if email_filter and email_filter.enabled:
            meta = conversion_options.get("_parsed_email_data")
            if meta is None:
                meta = self.parse_eml(file_path)
            if meta:
                att_count = len(meta.get("attachments") or [])
                meta_row = {
                    "subject": meta.get("subject"),
                    "from": meta.get("from"),
                    "date": meta.get("date"),
                    "size": sz,
                    "attachment_count": att_count,
                }
                if not email_filter.apply(meta_row):
                    return "skipped", "Filtered out by export rules"

        # Check for duplicates (fingerprint by default; optional full SHA-256 via EML2PST_FULL_DEDUP)
        if conversion_options["remove_duplicates"]:
            if prep is not None and prep.dedup_key is not None:
                dedup_key, header_date = prep.dedup_key, prep.header_date
            elif prep is not None and prep.prep_error:
                return "error", prep.prep_error
            else:
                dedup_pair = self._dedup_key_for_file(file_path, sz)
                if not dedup_pair:
                    if FULL_DEDUP_HASH:
                        return "error", "Cannot compute duplicate hash"
                    return "error", "Cannot compute duplicate fingerprint"
                dedup_key, header_date = dedup_pair
            if self._is_duplicate_and_mark(dedup_key, header_date):
                detail = "Duplicate content"
                if header_date:
                    detail = f"Duplicate content (Date: {header_date})"
                return "skipped", detail

        # PST: never OpenSharedItem (opens profile Drafts; Copy into target PST fails silently).
        if self._export_targets_pst(conversion_options):
            pst_result = self._import_eml_direct_to_pst_folder(
                file_path, target_folder, outlook, conversion_options
            )
            if pst_result == "converted":
                return "converted", ""
            if conversion_options.get("strict_date_preservation"):
                return "skipped", pst_result or "PST direct import failed"
            return "error", pst_result or "PST direct import failed"
        
        staged_already_tried = False
        if self._path_needs_native_staging(file_path):
            staged_already_tried = True
            staged_result = self._try_staged_eml_import_with_dates(
                namespace,
                target_folder,
                file_path,
                conversion_options,
                outlook=outlook,
            )
            if staged_result == "converted":
                return "converted", ""

        # Method 1: Native Outlook import (best MIME preservation: full RFC822 in PST)
        try_native = self._should_try_native_import(file_path, sz)
        # WLM paths with ()/@ already failed staged OpenSharedItem; native in-place fails too.
        if staged_already_tried and self._path_needs_native_staging(file_path):
            try_native = False
        native_result = None
        if try_native:
            native_result = self._import_with_outlook_native(
                namespace, target_folder, file_path, conversion_options
            )
            if native_result == "converted":
                return "converted", ""
        else:
            logger.debug(
                "Skipping native import for %s (%d bytes) — using manual import",
                log_sanitize(file_path),
                sz,
            )

        if conversion_options["strict_date_preservation"]:
            if native_result:
                logger.warning(
                    "Native import failed for %s; trying strict metadata fallback. Error: %s",
                    log_sanitize(file_path),
                    log_sanitize(native_result),
                )
            strict_result = self._process_single_email_manual_fallback(
                file_path,
                target_folder,
                outlook,
                require_date_preservation=True,
                conversion_options=conversion_options,
                try_staged_first=not staged_already_tried,
            )
            if strict_result == "converted":
                return "converted", ""
            return "skipped", (
                "Native import required to preserve original arrival date. "
                f"Reason: {native_result}; strict fallback failed: {strict_result}"
            )

        if _is_outlook_unavailable_error(native_result):
            logger.warning(
                "Native import failed for %s (Outlook COM unavailable): %s",
                log_sanitize(file_path),
                log_sanitize(native_result),
            )
            return "error", native_result
        if native_result:
            logger.warning(
                "Native import failed for %s; strict date preservation is OFF, using manual fallback. Error: %s",
                log_sanitize(file_path),
                log_sanitize(native_result),
            )
        fallback_result = self._process_single_email_manual_fallback(
            file_path,
            target_folder,
            outlook,
            require_date_preservation=False,
            conversion_options=conversion_options,
            try_staged_first=not staged_already_tried,
        )
        if fallback_result == "converted":
            return "converted", ""
        return "error", fallback_result

    def _process_single_email_manual_fallback(
        self,
        file_path,
        target_folder,
        outlook,
        require_date_preservation,
        conversion_options,
        *,
        try_staged_first: bool = True,
        from_validation_rebuild: bool = False,
    ):
        """
        Manual fallback import.
        Prefer staged OpenSharedItem (full MIME) with date stamping; if that fails,
        create the message in the PST folder via Items.Add (no Move from Drafts).
        """
        if try_staged_first and not from_validation_rebuild:
            staged_result = self._try_staged_eml_import_with_dates(
                outlook.GetNamespace("MAPI"),
                target_folder,
                file_path,
                conversion_options,
                outlook=outlook,
            )
            if staged_result == "converted":
                return "converted"
            if require_date_preservation:
                return (
                    f"Staged import required for date preservation failed: {staged_result}"
                )
            logger.debug(
                "Staged import failed for %s (%s); using Items.Add fallback",
                log_sanitize(file_path),
                staged_result,
            )

        try:
            email_data = self.parse_eml(file_path)
            if not email_data:
                return "Could not parse email"

            dt_local = self._resolve_outlook_datetime_for_source(
                file_path, email_data, conversion_options
            )
            if require_date_preservation and dt_local is None:
                return "Could not preserve original arrival date metadata"

            pst_path = normalize_pst_path(
                conversion_options.get("destination_path") or ""
            )
            target_store = None
            try:
                target_store = target_folder.Store
            except Exception:
                pass
            if pst_path and not self._folder_belongs_to_pst(
                target_folder, pst_path, store=target_store
            ):
                target_folder = self._coerce_import_folder(
                    target_folder, conversion_options
                )
            if pst_path:
                mail = self._create_mail_in_pst_target_folder(
                    target_folder, pst_path
                )
            else:
                mail = target_folder.Items.Add(OL_MAIL_ITEM)
            if dt_local is not None:
                self._stamp_delivery_on_new_item(mail, dt_local)

            self._apply_original_metadata(
                mail,
                email_data,
                source_file_path=file_path,
                apply_dates=False,
            )

            self._apply_parsed_display_fields(
                mail,
                email_data,
                file_path,
                conversion_options,
            )

            self._apply_attachment_chain(
                mail,
                email_data=email_data,
                source_file_path=file_path,
            )

            namespace = outlook.GetNamespace("MAPI")
            sent_state = self._sent_state_for_source_path(file_path, conversion_options)
            mark_as_received = sent_state is not False
            date_ok = self._finalize_message_datetime(
                mail,
                namespace,
                file_path,
                email_data,
                conversion_options,
                dest_folder=target_folder,
                pst_path=conversion_options.get("destination_path"),
                dt_local=dt_local,
                mark_as_received=mark_as_received,
            )
            if require_date_preservation and not date_ok:
                try:
                    mail.Delete()
                except Exception:
                    pass
                return "Could not preserve original arrival date metadata"
            if self._export_targets_pst(conversion_options):
                if not self._verify_mail_in_target_pst(
                    mail, conversion_options, source_file_path=file_path
                ):
                    try:
                        mail.Delete()
                    except Exception:
                        pass
                    return "Message was not saved in the target PST"
            self._reapply_sender_from_email_data(
                mail, email_data, file_path, conversion_options
            )
            try:
                mail.Save()
            except Exception as e:
                logger.debug("Save after sender re-apply: %s", e)
            return self._ensure_import_quality(
                mail,
                file_path,
                outlook,
                target_folder,
                conversion_options,
                email_data=email_data,
                from_validation_rebuild=from_validation_rebuild,
                expected_folder=target_folder,
            )
        except Exception as e:
            logger.debug("Manual fallback failed for %s: %s", file_path, e)
            return str(e)

    def _open_shared_item_candidates(self, file_path, *, allow_uri: bool = False):
        """
        Outlook OpenSharedItem often rejects plain paths and returns
        'Invalid path or URL'. Try absolute path and short path; file:/// URI
        only when explicitly allowed (often fails for async/lazy reads).
        """
        candidates = []
        try:
            abs_path = os.path.normpath(os.path.abspath(file_path))
        except OSError:
            abs_path = file_path

        if os.path.isfile(abs_path):
            # 1) Absolute path (most reliable for OpenSharedItem)
            candidates.append(abs_path)
            # 2) Forward slashes (some Outlook builds prefer this over backslashes)
            forward = abs_path.replace("\\", "/")
            if forward not in candidates:
                candidates.append(forward)
            # 3) Short path (helps with spaces / Unicode in path)
            if OUTLOOK_AVAILABLE:
                try:
                    import win32api
                    short = win32api.GetShortPathName(abs_path)
                    if short and os.path.isfile(short) and short not in candidates:
                        candidates.append(short)
                except Exception as e:
                    logger.debug("GetShortPathName skipped: %s", e)
            # 4) file:/// URI — last resort; skip for staging copies
            if allow_uri:
                try:
                    uri = Path(abs_path).resolve().as_uri()
                    if uri not in candidates:
                        candidates.append(uri)
                except Exception as e:
                    logger.debug("as_uri failed: %s", e)

        # De-dupe preserving order
        seen = set()
        out = []
        for c in candidates:
            if c and c not in seen:
                seen.add(c)
                out.append(c)
        return out

    def _open_shared_item(self, namespace, file_path, *, allow_uri: bool = False):
        """Try OpenSharedItem with path variants; return mail item or raise last error."""
        last_err = None
        for candidate in self._open_shared_item_candidates(file_path, allow_uri=allow_uri):
            try:
                return namespace.OpenSharedItem(candidate)
            except Exception as e:
                last_err = e
                logger.debug("OpenSharedItem failed for %r: %s", candidate, e)
        if last_err:
            raise last_err
        raise OSError(f"Not a file or unreadable: {file_path}")

    def _outlook_com_datetime(self, dt_naive: datetime):
        """
        Outlook often ignores Python datetime for MAPI / MailItem time fields;
        pywintypes.Time is required for list columns (Received, etc.).
        """
        if dt_naive.tzinfo is not None:
            dt_naive = dt_naive.astimezone().replace(tzinfo=None)
        if PYWINTYPES:
            try:
                return PYWINTYPES.Time(dt_naive)
            except Exception:
                pass
        return dt_naive

    def _naive_datetime_from_com(self, com_val) -> datetime | None:
        """Normalize Outlook/pywintypes COM times to naive local datetime."""
        if com_val is None:
            return None
        try:
            if hasattr(com_val, "year"):
                dt = com_val
                if getattr(dt, "tzinfo", None) is not None:
                    dt = dt.astimezone().replace(tzinfo=None)
                return dt.replace(microsecond=0)
        except Exception:
            pass
        return None

    def _is_windows_live_mail_path(self, source_file_path: str | None) -> bool:
        if not source_file_path:
            return False
        return "windows live mail" in source_file_path.replace("/", "\\").lower()

    def _file_mtime_datetime(self, source_file_path: str | None) -> datetime | None:
        if not source_file_path or not os.path.isfile(source_file_path):
            return None
        try:
            return datetime.fromtimestamp(os.path.getmtime(source_file_path)).replace(
                microsecond=0
            )
        except OSError:
            return None

    def _header_datetime_from_email(self, email_data: dict | None) -> datetime | None:
        if not email_data:
            return None
        date_raw = self._first_parseable_date_header(email_data)
        if not date_raw:
            return None
        try:
            dt = parsedate_to_datetime(date_raw)
            if dt is None:
                return None
            if dt.tzinfo is not None:
                return dt.astimezone().replace(tzinfo=None).replace(microsecond=0)
            return dt.replace(microsecond=0)
        except (TypeError, ValueError, OverflowError):
            return None

    def _resolve_outlook_datetime_for_source(
        self,
        source_file_path: str | None,
        email_data: dict | None,
        conversion_options: dict,
    ) -> datetime | None:
        """
        Pick the instant for Outlook Received / Sent columns.

        Default (checkbox on): Explorer Date modified on the .eml — matches Live Mail
        folder view. Windows Live Mail trees always use mtime. Otherwise Date:
        header, then mtime as fallback.
        """
        mtime_dt = self._file_mtime_datetime(source_file_path)
        use_mtime = conversion_options.get("use_file_mtime_for_date", True)
        if self._is_windows_live_mail_path(source_file_path):
            use_mtime = True
        if use_mtime and mtime_dt is not None:
            return mtime_dt
        header_dt = self._header_datetime_from_email(email_data)
        if header_dt is not None:
            return header_dt
        return mtime_dt

    def _datetime_source_label(
        self, source_file_path: str | None, conversion_options: dict
    ) -> str:
        use_mtime = conversion_options.get("use_file_mtime_for_date", True)
        if self._is_windows_live_mail_path(source_file_path):
            use_mtime = True
        return "file modification time" if use_mtime else "email header"

    def _verify_outlook_message_datetime(
        self, mail_item, expected: datetime, *, tolerance_sec: int = 120
    ) -> bool:
        """True when ReceivedTime / delivery MAPI time matches expected (not import time)."""
        expected = expected.replace(microsecond=0)
        now = datetime.now().replace(microsecond=0)
        expected_is_historical = abs((expected - now).total_seconds()) > 3600

        def _matches(got: datetime | None) -> bool:
            if got is None:
                return False
            if expected_is_historical and abs((got - now).total_seconds()) <= tolerance_sec:
                return False
            return abs((got - expected).total_seconds()) <= tolerance_sec

        try:
            if _matches(self._naive_datetime_from_com(mail_item.ReceivedTime)):
                return True
        except Exception:
            pass
        try:
            pa = mail_item.PropertyAccessor
            got = self._naive_datetime_from_com(
                pa.GetProperty(_OUTLOOK_MAPI_TIME_URLS[0])
            )
            if _matches(got):
                return True
        except Exception:
            pass
        return False

    def _apply_mapi_delivery_times(self, mail_item, com_val) -> None:
        """Write delivery/received MAPI times (works even when OM properties are read-only)."""
        if com_val is None:
            return
        try:
            pa = mail_item.PropertyAccessor
            for url in _OUTLOOK_MAPI_TIME_URLS:
                try:
                    pa.SetProperty(url, com_val)
                except Exception as e:
                    logger.debug("SetProperty %s: %s", url, e)
        except Exception as e:
            logger.debug("PropertyAccessor unavailable: %s", e)

    def _clear_unsent_mapi_flag(self, mail_item) -> None:
        """Clear MSGFLAG_UNSENT so Outlook does not treat imported mail as a draft."""
        try:
            pa = mail_item.PropertyAccessor
            flags = int(pa.GetProperty(_OUTLOOK_MAPI_MESSAGE_FLAGS))
            pa.SetProperty(_OUTLOOK_MAPI_MESSAGE_FLAGS, flags & ~_MSGFLAG_UNSENT)
        except Exception as e:
            logger.debug("Clear MSGFLAG_UNSENT: %s", e)

    def _mark_mail_imported_received(self, mail_item, dt_local: datetime | None = None) -> None:
        """
        Imported RFC822 is received mail, not a user draft.

        Items left with Sent=False are filed under Drafts (Piszkozatok) on save.
        Setting Sent=True can reset OM date fields — re-apply MAPI delivery times
        before and after marking received so Explorer dates stay intact.
        """
        com_val = self._outlook_com_datetime(dt_local) if dt_local is not None else None
        if com_val is not None:
            self._apply_mapi_delivery_times(mail_item, com_val)
        self._clear_unsent_mapi_flag(mail_item)
        try:
            mail_item.Sent = True
        except Exception:
            pass
        if com_val is not None:
            self._apply_mapi_delivery_times(mail_item, com_val)
            try:
                mail_item.ReceivedTime = com_val
            except Exception as e:
                logger.debug("ReceivedTime after Sent=True: %s", e)
            try:
                mail_item.SentOn = com_val
            except Exception as e:
                logger.debug("SentOn after Sent=True: %s", e)

    def _stamp_delivery_for_import(
        self, mail_item, dt_local: datetime, *, mark_received: bool = True
    ) -> None:
        """
        Stamp delivery times before Save.

        mark_received=True (Inbox, Sent, Deleted, ...): clear draft flag, Sent=True.
        mark_received=False (Drafts / Outbox): keep Sent=False in the target folder.
        """
        dt_local = dt_local.replace(microsecond=0)
        com_val = self._outlook_com_datetime(dt_local)
        if com_val is None:
            return
        self._apply_mapi_delivery_times(mail_item, com_val)
        try:
            mail_item.ReceivedTime = com_val
        except Exception as e:
            logger.debug("MailItem.ReceivedTime: %s", e)
        try:
            mail_item.SentOn = com_val
        except Exception as e:
            logger.debug("MailItem.SentOn: %s", e)
        for prop in ("CreationTime", "LastModificationTime"):
            try:
                setattr(mail_item, prop, com_val)
            except Exception:
                pass
        if mark_received:
            try:
                mail_item.Sent = False
            except Exception:
                pass
            self._mark_mail_imported_received(mail_item, dt_local)
        else:
            try:
                mail_item.Sent = False
            except Exception:
                pass

    def _save_imported_mail(
        self, mail_item, dt_local: datetime | None = None
    ) -> None:
        self._mark_mail_imported_received(mail_item, dt_local)
        try:
            mail_item.Save()
        except Exception as e:
            logger.debug("Save imported mail: %s", e)
        if dt_local is not None:
            com_val = self._outlook_com_datetime(dt_local.replace(microsecond=0))
            if com_val is not None:
                self._apply_mapi_delivery_times(mail_item, com_val)
                try:
                    mail_item.Save()
                except Exception as e:
                    logger.debug("Save after post-save MAPI stamp: %s", e)

    def _discard_open_shared_item(self, mail_item) -> None:
        """Remove temporary OpenSharedItem mail (usually opened in Drafts)."""
        try:
            mail_item.Delete()
            return
        except Exception as e:
            logger.debug("OpenSharedItem Delete failed: %s", e)
        try:
            mail_item.Close(OL_DISCARD)
        except Exception as e:
            logger.debug("OpenSharedItem Close(Discard) failed: %s", e)

    def _force_outlook_message_datetime(self, mail_item, dt_local: datetime) -> bool:
        """
        Set delivery/received times on a MailItem (MAPI + OM).

        Outlook often resets times to "now" on Move into a PST; ReceivedTime is
        read-only unless Sent is cleared briefly. Returns True only when verified.
        """
        if dt_local is None:
            return False
        self._stamp_delivery_for_import(mail_item, dt_local)
        return self._verify_outlook_message_datetime(mail_item, dt_local)

    def _stamp_outlook_item_with_explorer_mtime(self, mail_item, source_eml_path: str) -> bool:
        """Set message times from Explorer Date modified on the source .eml."""
        dt_local = self._resolve_outlook_datetime_for_source(
            source_eml_path, None, {"use_file_mtime_for_date": True}
        )
        if dt_local is None:
            return False
        ok = self._force_outlook_message_datetime(mail_item, dt_local)
        if not ok:
            logger.warning(
                "Explorer mtime stamp had no effect for %s (Outlook may block writes)",
                source_eml_path,
            )
        return ok

    def _log_datetime_outcome(
        self,
        ok: bool,
        source_file_path: str | None,
        conversion_options: dict,
        dt_local: datetime,
    ) -> None:
        label = self._datetime_source_label(source_file_path, conversion_options)
        path = log_sanitize(source_file_path or "")
        if ok:
            logger.info(
                "Applied delivery time from %s (%s): %s",
                label,
                dt_local.strftime("%Y-%m-%d %H:%M"),
                path,
            )
        else:
            self._delivery_time_warn_count += 1
            n = self._delivery_time_warn_count
            if n == 4:
                logger.warning(
                    "Delivery time did not stick for several messages (Outlook PST quirk). "
                    "Further per-file warnings are summarized every 50 messages. "
                    "Use a brand-new PST and MailExporter_x32.exe with Explorer date enabled."
                )
            if n <= 3 or n % 50 == 0:
                logger.warning(
                    "Delivery time from %s (%s) did not stick in Outlook — "
                    "leave 'Explorer date' checked and use a fresh PST: %s",
                    label,
                    dt_local.strftime("%Y-%m-%d %H:%M"),
                    path,
                )

    def _mail_item_store_path(self, mail_item) -> str:
        try:
            return str(mail_item.Parent.Store.FilePath or "")
        except Exception:
            return ""

    def _folder_store_path(self, folder) -> str:
        try:
            return normalize_pst_path(str(folder.Store.FilePath or ""))
        except Exception:
            return ""

    def _message_in_expected_pst(
        self, mail_item, expected_pst: str, *, expected_store=None
    ) -> bool:
        """True when the item's store is the target PST."""
        if not expected_pst:
            return True
        current = normalize_pst_path(self._mail_item_store_path(mail_item))
        if current and pst_paths_equal(current, expected_pst):
            return True
        if expected_store is None:
            return False
        try:
            return self._stores_match(mail_item.Parent.Store, expected_store)
        except Exception:
            return False

    def _import_open_shared_into_pst_folder(
        self, mail_item, dest_folder, conversion_options: dict
    ):
        """
        Move OpenSharedItem mail into the export PST before Save.

        OpenSharedItem opens in the profile Drafts store (often with an empty
        Store.FilePath). Saving there leaves mail outside the export PST; Copy first.
        """
        dest_folder = self._coerce_import_folder(dest_folder, conversion_options)
        pst_path = conversion_options.get("destination_path") or ""
        expected = normalize_pst_path(pst_path) if pst_path else ""
        dest_store = self._folder_store_path(dest_folder)
        if expected and dest_store and not pst_paths_equal(dest_store, expected):
            logger.warning(
                "Target folder is not in export PST %s (folder store %s)",
                expected,
                dest_store or "?",
            )
            return mail_item

        current = normalize_pst_path(self._mail_item_store_path(mail_item))
        if expected and current and pst_paths_equal(current, expected):
            moved = self._relocate_mail_to_dest_folder(mail_item, dest_folder)
            return moved if moved is not None else mail_item

        try:
            copied = mail_item.Copy(dest_folder)
            dest_store_obj = dest_folder.Store
            if expected and not self._message_in_expected_pst(
                copied, expected, expected_store=dest_store_obj
            ):
                logger.warning(
                    "Copy into PST folder %r did not stick — retrying via PST Inbox then Move",
                    str(getattr(dest_folder, "Name", "")),
                )
                try:
                    copied.Delete()
                except Exception:
                    pass
                store = dest_folder.Store
                inbox = self._get_or_create_inbox(store)
                copied = mail_item.Copy(inbox)
                try:
                    if (
                        getattr(dest_folder, "EntryID", None)
                        and getattr(inbox, "EntryID", None)
                        and dest_folder.EntryID != inbox.EntryID
                    ):
                        copied = copied.Move(dest_folder)
                except Exception as move_err:
                    logger.warning(
                        "Move from PST Inbox into %r failed: %s",
                        str(getattr(dest_folder, "Name", "")),
                        move_err,
                    )
            try:
                mail_item.Delete()
            except Exception as e:
                logger.debug("Delete temp OpenSharedItem after PST copy: %s", e)
            return copied
        except Exception as e:
            logger.warning(
                "Could not copy message into PST folder before import — "
                "mail may land in profile Drafts: %s",
                e,
            )
            return mail_item

    def _relocate_mail_to_dest_folder(self, mail_item, dest_folder):
        """Move mail into dest_folder when Outlook filed it elsewhere (e.g. default Drafts)."""
        try:
            parent = mail_item.Parent
            if (
                getattr(parent, "EntryID", None)
                and getattr(dest_folder, "EntryID", None)
                and parent.EntryID == dest_folder.EntryID
            ):
                return mail_item
        except Exception:
            pass
        try:
            return mail_item.Move(dest_folder)
        except Exception as e:
            logger.debug("Move into target PST folder: %s", e)
            try:
                copied = mail_item.Copy(dest_folder)
                try:
                    mail_item.Delete()
                except Exception as e_del:
                    logger.debug("Delete source after PST copy: %s", e_del)
                return copied
            except Exception as e2:
                logger.warning("Could not move/copy mail into PST folder: %s", e2)
                return mail_item

    def _commit_mail_to_pst_folder(
        self,
        mail_item,
        dest_folder,
        pst_path: str | None,
        dt_local: datetime | None,
        *,
        source_file_path: str | None = None,
        mark_as_received: bool = True,
    ):
        """
        After Save, ensure the item is under dest_folder in the target PST.

        OpenSharedItem and unsent Items.Add often land in the profile's Drafts
        store instead of the attached PST when that PST is not the default store.
        """
        expected = normalize_pst_path(pst_path) if pst_path else ""
        if dt_local is not None and mark_as_received:
            self._mark_mail_imported_received(mail_item, dt_local)
        elif dt_local is not None:
            com_val = self._outlook_com_datetime(dt_local.replace(microsecond=0))
            if com_val is not None:
                self._apply_mapi_delivery_times(mail_item, com_val)
            try:
                mail_item.Sent = False
            except Exception:
                pass
        moved = self._relocate_mail_to_dest_folder(mail_item, dest_folder)
        if moved is not None:
            mail_item = moved
        if dt_local is not None:
            com_val = self._outlook_com_datetime(dt_local.replace(microsecond=0))
            if com_val is not None:
                self._apply_mapi_delivery_times(mail_item, com_val)
        try:
            mail_item.Save()
        except Exception as e:
            logger.debug("Commit save to PST folder: %s", e)

        try:
            parent_name = str(mail_item.Parent.Name)
        except Exception:
            parent_name = "?"
        mail_store = None
        try:
            mail_store = mail_item.Parent.Store
        except Exception:
            pass
        if expected and not self._message_in_expected_pst(
            mail_item, expected, expected_store=mail_store
        ):
            current = self._mail_item_store_path(mail_item)
            self._wrong_store_warn_count += 1
            n = self._wrong_store_warn_count
            if n <= 3 or n % 50 == 0:
                logger.warning(
                    "Message in %s (%s), not target PST %s — %s",
                    parent_name or "(no name)",
                    current or "(no PST path — profile store)",
                    expected,
                    log_sanitize(source_file_path or ""),
                )
            moved = self._relocate_mail_to_dest_folder(mail_item, dest_folder)
            if moved is not None:
                mail_item = moved
                if dt_local is not None:
                    com_val = self._outlook_com_datetime(dt_local.replace(microsecond=0))
                    if com_val is not None:
                        self._apply_mapi_delivery_times(mail_item, com_val)
                try:
                    mail_item.Save()
                except Exception:
                    pass
        return mail_item

    def _stamp_delivery_on_new_item(self, mail_item, dt_local: datetime) -> None:
        """Set delivery/creation MAPI times on a new unsaved item (before body/sender)."""
        com_val = self._outlook_com_datetime(dt_local)
        if com_val is None:
            return
        self._apply_mapi_delivery_times(mail_item, com_val)
        for prop in ("ReceivedTime", "CreationTime", "LastModificationTime"):
            try:
                setattr(mail_item, prop, com_val)
            except Exception:
                pass

    def _copy_attachments_between_mail_items(self, source_mail, dest_mail) -> None:
        try:
            atts = source_mail.Attachments
            count = int(atts.Count)
        except Exception:
            return
        if count <= 0:
            return
        temp_dir = tempfile.mkdtemp(prefix="me_att_")
        try:
            for i in range(1, count + 1):
                try:
                    att = atts.Item(i)
                    fname = att.FileName or f"attach_{i}"
                    safe = re.sub(r'[<>:"/\\|?*]', "_", fname)[:200]
                    path = os.path.join(temp_dir, safe)
                    att.SaveAsFile(path)
                    dest_mail.Attachments.Add(path)
                except Exception as e:
                    logger.debug("Attachment transfer failed: %s", e)
        finally:
            shutil.rmtree(temp_dir, ignore_errors=True)

    def _transfer_opened_mail_to_folder(
        self,
        mail_item,
        target_folder,
        namespace,
        source_file_path: str | None,
        email_data: dict | None,
        conversion_options: dict,
        *,
        outlook=None,
    ):
        """
        Import OpenSharedItem mail into the target PST folder.

        OpenSharedItem opens in the profile Drafts store; copy into the PST
        before Save so mail is not left in the default profile (empty FilePath).
        """
        mail_item = self._import_open_shared_into_pst_folder(
            mail_item, target_folder, conversion_options
        )
        dt_local = self._resolve_outlook_datetime_for_source(
            source_file_path, email_data, conversion_options
        )
        if email_data is None and source_file_path:
            email_data = self.parse_eml(source_file_path)

        if email_data:
            self._apply_original_metadata(
                mail_item,
                email_data,
                source_file_path=source_file_path,
                apply_dates=False,
            )
            self._apply_parsed_display_fields(
                mail_item,
                email_data,
                source_file_path,
                conversion_options,
            )

        self._apply_attachment_chain(
            mail_item,
            email_data=email_data,
            outlook_mail=mail_item,
            source_file_path=source_file_path,
        )

        pst_path = conversion_options.get("destination_path")
        self._finalize_message_datetime(
            mail_item,
            namespace,
            source_file_path,
            email_data,
            conversion_options,
            dest_folder=target_folder,
            pst_path=pst_path,
            dt_local=dt_local,
        )
        if source_file_path and self._import_validation_enabled(conversion_options):
            outlook = outlook or conversion_options.get("_session_outlook")
            quality = self._ensure_import_quality(
                mail_item,
                source_file_path,
                outlook,
                target_folder,
                conversion_options,
                email_data=email_data,
                expected_folder=target_folder,
            )
            if quality != "converted":
                raise RuntimeError(quality)
        return mail_item

    def _finalize_message_datetime(
        self,
        mail_item,
        namespace,
        source_file_path: str | None,
        email_data: dict | None,
        conversion_options: dict,
        *,
        dest_folder=None,
        pst_path: str | None = None,
        dt_local: datetime | None = None,
        mark_as_received: bool = True,
    ) -> bool:
        """Force delivery time, Save, reload from store if needed, verify once."""
        if dt_local is None:
            dt_local = self._resolve_outlook_datetime_for_source(
                source_file_path, email_data, conversion_options
            )
        if dt_local is None:
            if dest_folder is not None:
                mail_item = self._commit_mail_to_pst_folder(
                    mail_item,
                    dest_folder,
                    pst_path or conversion_options.get("destination_path"),
                    None,
                    source_file_path=source_file_path,
                    mark_as_received=mark_as_received,
                )
            else:
                self._save_imported_mail(mail_item)
            return False

        ok = False
        item = mail_item
        for _ in range(3):
            self._stamp_delivery_for_import(
                item, dt_local, mark_received=mark_as_received
            )
            try:
                item.Save()
            except Exception as e:
                logger.debug("Save after date stamp: %s", e)
            com_val = self._outlook_com_datetime(dt_local.replace(microsecond=0))
            if com_val is not None:
                self._apply_mapi_delivery_times(item, com_val)
            if self._verify_outlook_message_datetime(item, dt_local):
                try:
                    item.Save()
                except Exception as e:
                    logger.debug("Final save after verified date stamp: %s", e)
                ok = True
                break
            try:
                entry_id = item.EntryID
                store = item.Parent.Store
                store_id = store.StoreID if store is not None else None
                if entry_id and store_id:
                    item = namespace.GetItemFromID(entry_id, store_id)
            except Exception as e:
                logger.debug("Could not reload item for date stamp retry: %s", e)
                break

        self._log_datetime_outcome(ok, source_file_path, conversion_options, dt_local)
        if dest_folder is not None:
            item = self._commit_mail_to_pst_folder(
                item,
                dest_folder,
                pst_path or conversion_options.get("destination_path"),
                dt_local,
                source_file_path=source_file_path,
                mark_as_received=mark_as_received,
            )
        return ok

    def _apply_message_datetime_to_item(
        self,
        mail_item,
        namespace,
        source_file_path: str | None,
        email_data: dict | None,
        conversion_options: dict,
    ) -> bool:
        """Resolve source time, stamp, save, verify — single log line."""
        return self._finalize_message_datetime(
            mail_item,
            namespace,
            source_file_path,
            email_data,
            conversion_options,
        )

    def _open_shared_item_from_staged_eml(
        self, namespace, staged_eml_path: str
    ):
        """OpenSharedItem on staged .eml; optional .msg round-trip for stubborn Outlook builds."""
        last_err = None
        for allow_uri in (False, True):
            try:
                return self._open_shared_item(
                    namespace, staged_eml_path, allow_uri=allow_uri
                )
            except Exception as e:
                last_err = e
        msg_path = None
        try:
            base, _ = os.path.splitext(staged_eml_path)
            msg_path = base + ".msg"
            if os.path.isfile(msg_path):
                os.remove(msg_path)
            probe = self._open_shared_item(namespace, staged_eml_path, allow_uri=False)
            probe.SaveAs(msg_path, OL_SAVEAS_MSG)
            del probe
            self._wait_for_staging_file(msg_path)
            with self._lock:
                self._temp_files.append(msg_path)
            return self._open_shared_item(namespace, msg_path, allow_uri=False)
        except Exception as e:
            last_err = e
            if msg_path:
                self._register_staging_file(msg_path)
        if last_err:
            raise last_err
        raise OSError("Could not open staged message for import")

    def _try_staged_eml_import_with_dates(
        self,
        namespace,
        target_folder,
        file_path: str,
        conversion_options: dict,
        *,
        outlook=None,
    ) -> str:
        """
        Import via staged CRLF .eml under LOCALAPPDATA + OpenSharedItem (+ optional .msg).
        Returns 'converted' or an error string (never raises).
        """
        try:
            staged_path = self._stage_rfc822_eml_for_import(file_path, prefix="e")
            if not staged_path:
                sz = file_size_or_zero(file_path)
                if sz > NATIVE_IMPORT_MAX_BYTES:
                    return "Message too large for staged import"
                return "Could not read source message content"
            email_data = self.parse_eml(file_path)

            mail_item = self._open_shared_item_from_staged_eml(namespace, staged_path)
            try:
                saved = self._transfer_opened_mail_to_folder(
                    mail_item,
                    target_folder,
                    namespace,
                    file_path,
                    email_data,
                    conversion_options,
                    outlook=outlook,
                )
                if self._export_targets_pst(conversion_options) and not self._verify_mail_in_target_pst(
                    saved,
                    conversion_options,
                    source_file_path=file_path,
                ):
                    try:
                        saved.Delete()
                    except Exception:
                        pass
                    return "Message was not saved in the target PST"
                return "converted"
            finally:
                release_com_object(mail_item)
        except Exception as e:
            return str(e)

    def _apply_original_metadata(
        self,
        mail,
        email_data,
        source_file_path=None,
        use_file_mtime=True,
        prefer_file_mtime_first=False,
        *,
        apply_dates=True,
    ):
        """
        Preserve MIME headers and message time in Outlook (fallback path).

        - PR_TRANSPORT_MESSAGE_HEADERS: full header block from the .eml.
        - PR_MESSAGE_DELIVERY_TIME / PR_CLIENT_SUBMIT_TIME:
          If prefer_file_mtime_first and use_file_mtime: Windows ``mtime`` on the
          original file first (Explorer "Date modified"); else Date/Received headers,
          then mtime as last resort.

        Returns True if a message date was applied (header-based or mtime).
        """
        date_applied = False
        try:
            pa = mail.PropertyAccessor
        except Exception as e:
            logger.debug("PropertyAccessor unavailable: %s", e)
            return False

        try:
            headers = self._build_transport_headers(email_data)
            if headers:
                try:
                    pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001E", headers)
                except Exception:
                    # Unicode transport headers (some Outlook builds)
                    pa.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x007D001F", headers)
        except Exception as e:
            logger.debug("Could not set transport headers: %s", e)

        return False

    def _first_parseable_date_header(self, email_data):
        """
        First usable instant for Outlook delivery/submit time.

        Order: ``Date:``, then **each** ``Received:`` line (common in real mail;
        msg.get('Received') only returns one), then Resent-Date / Delivery-Date.
        """
        msg = email_data.get("message")

        def _parse_stamp(raw: str):
            if not raw:
                return None
            raw = raw.strip()
            try:
                if parsedate_to_datetime(raw) is not None:
                    return raw
            except (TypeError, ValueError):
                pass
            return None

        # Date:
        raw = None
        if msg is not None:
            raw = msg.get("Date")
        if not raw:
            raw = email_data.get("date")
        found = _parse_stamp(raw) if raw else None
        if found:
            return found

        # Every Received: (hop chain); date is usually after the last ';'
        if msg is not None:
            received_vals = msg.get_all("Received") or []
            if received_vals:
                for rec in received_vals:
                    candidate = rec
                    if ";" in rec:
                        candidate = rec.rsplit(";", 1)[-1].strip()
                    found = _parse_stamp(candidate)
                    if found:
                        return found

        for key in ("Resent-Date", "Delivery-Date"):
            raw = msg.get(key) if msg is not None else None
            if raw:
                found = _parse_stamp(raw)
                if found:
                    return found

        return email_data.get("date")

    def _build_transport_headers(self, email_data):
        """
        Build RFC822 header block for PR_TRANSPORT_MESSAGE_HEADERS (MIME headers).

        Prefer ``raw_items()`` so order and **repeated** headers (e.g. multiple
        ``Received:``) match the source. Otherwise join ``get_all()`` per header
        name so Received chains are not dropped.
        """
        msg = email_data.get("message")
        if msg is None:
            return ""
        try:
            if hasattr(msg, "raw_items"):
                lines = [f"{k}: {v}" for (k, v) in msg.raw_items()]
            else:
                lines = []
                for name in msg.keys():
                    vals = msg.get_all(name)
                    if not vals:
                        continue
                    for value in vals:
                        lines.append(f"{name}: {value}")
            if lines:
                return "\r\n".join(lines) + "\r\n"
        except Exception as e:
            logger.debug("Could not build transport headers: %s", e)
        return ""

    def _load_rfc822_bytes(self, file_path: str) -> bytes | None:
        """Read raw RFC822 bytes from .eml or decoded payload from .emlx."""
        ext = os.path.splitext(file_path)[1].lower()
        if ext == ".emlx":
            return self._emlx_to_rfc822_bytes(file_path)
        return read_file_bytes(file_path)

    def _stage_rfc822_eml_for_import(
        self, file_path: str, *, prefix: str = "e"
    ) -> str | None:
        """
        Stage .eml for OpenSharedItem without loading multi-MB files into RAM when possible.
        Returns staged path or None on failure / oversize.
        """
        norm = os.path.normpath(os.path.abspath(file_path))
        sz = self._eml_file_sizes.get(norm)
        if sz is None:
            sz = file_size_or_zero(file_path)
        if sz > NATIVE_IMPORT_MAX_BYTES:
            return None
        if should_use_mmap(sz):
            last_err = None
            for staging_dir in self._native_temp_staging_dirs(file_path):
                temp_eml_path = None
                try:
                    fd, temp_eml_path = tempfile.mkstemp(
                        suffix=".eml",
                        prefix=prefix,
                        dir=staging_dir,
                    )
                    os.close(fd)
                    if needs_lf_to_crlf_conversion(file_path):
                        write_crlf_normalized_file(file_path, temp_eml_path)
                    else:
                        copy_file_mmap(file_path, temp_eml_path)
                    if not os.path.isfile(temp_eml_path):
                        raise OSError(f"staging file missing: {temp_eml_path}")
                    with self._lock:
                        self._temp_files.append(temp_eml_path)
                    self._register_staging_file(temp_eml_path)
                    self._wait_for_staging_file(temp_eml_path)
                    return temp_eml_path
                except Exception as e:
                    last_err = e
                    if temp_eml_path:
                        self._register_staging_file(temp_eml_path)
                    logger.debug(
                        "mmap staging failed for %s in %s: %s",
                        log_sanitize(file_path),
                        staging_dir,
                        e,
                    )
            logger.debug(
                "Could not mmap-stage %s: %s",
                log_sanitize(file_path),
                last_err,
            )
            return None

        rfc822_bytes = self._load_rfc822_bytes(file_path)
        if not rfc822_bytes:
            return None
        rfc822_bytes = self._normalize_rfc822_line_endings(rfc822_bytes)
        try:
            return self._stage_file_for_native_import(
                file_path, prefix=prefix, data=rfc822_bytes
            )
        except OSError:
            return None

    def _should_try_native_import(self, file_path: str, file_size: int) -> bool:
        """
        OpenSharedItem (often via staged CRLF copy) preserves attachments well.
        Only skip native import for oversized messages; attachment/multipart mail
        is handled by staged import + attachment chain, not manual body parsing.
        """
        if file_size > NATIVE_IMPORT_MAX_BYTES:
            return False
        try:
            with open(file_path, "rb") as handle:
                head = handle.read(min(file_size, 8192))
        except OSError:
            return False
        return bool(head)

    def _path_needs_native_staging(self, file_path: str) -> bool:
        """
        True when the source path is unsafe for Outlook OpenSharedItem in place.
        Live Mail trees often use (), @, spaces, and long paths that trigger
        'Invalid path or URL' unless we stage a short flat .eml copy first.
        """
        try:
            path = os.path.normpath(os.path.abspath(file_path))
        except OSError:
            return True
        if len(path) > 180:
            return True
        if re.search(r'[()@#%&!\']', path):
            return True
        return False

    def _eml_source_likely_outlook_native_ready(self, file_path: str) -> bool:
        """
        True when OpenSharedItem can usually open the source file in place.
        LF-only RFC822 and problematic paths need a CRLF staging copy first.
        """
        if self._path_needs_native_staging(file_path):
            return False
        try:
            with open(file_path, "rb") as handle:
                chunk = handle.read(8192)
        except OSError:
            return False
        if not chunk:
            return False
        if b"\n" in chunk and b"\r\n" not in chunk:
            return False
        return True

    def _stage_file_for_native_import(
        self,
        source_path: str,
        *,
        prefix: str = "import_",
        data: bytes | None = None,
    ) -> str:
        """
        Copy or write a message to a simple staging .eml beside the PST.
        Outlook OpenSharedItem is much more reliable on short flat paths than
        on deep source trees or file:/// URIs.
        """
        last_err = None
        for staging_dir in self._native_temp_staging_dirs(source_path):
            temp_eml_path = None
            try:
                fd, temp_eml_path = tempfile.mkstemp(
                    suffix=".eml",
                    prefix=prefix,
                    dir=staging_dir,
                )
                os.close(fd)
                if data is not None:
                    with open(temp_eml_path, "wb") as handle:
                        handle.write(data)
                        handle.flush()
                        os.fsync(handle.fileno())
                else:
                    shutil.copy2(source_path, temp_eml_path)
                if not os.path.isfile(temp_eml_path):
                    raise OSError(f"staging file missing after write: {temp_eml_path}")
                with self._lock:
                    self._temp_files.append(temp_eml_path)
                self._register_staging_file(temp_eml_path)
                self._wait_for_staging_file(temp_eml_path)
                return temp_eml_path
            except Exception as e:
                last_err = e
                if temp_eml_path:
                    self._register_staging_file(temp_eml_path)
                logger.debug(
                    "Could not stage %s in %s: %s",
                    log_sanitize(source_path),
                    staging_dir,
                    e,
                )
        raise OSError(f"Could not stage file for native import: {last_err}")

    def _native_import_to_folder(
        self,
        namespace,
        target_folder,
        import_path: str,
        mtime_source_path: str,
        conversion_options: dict,
        *,
        allow_uri: bool = False,
    ):
        """OpenSharedItem + Move with COM retries for large export runs."""

        def _do_import():
            email_data = None
            if mtime_source_path:
                email_data = self.parse_eml(mtime_source_path)
            dt_local = self._resolve_outlook_datetime_for_source(
                mtime_source_path, email_data, conversion_options
            )
            if import_path.lower().endswith(".eml"):
                mail_item = self._open_shared_item_from_staged_eml(
                    namespace, import_path
                )
            else:
                mail_item = self._open_shared_item(
                    namespace, import_path, allow_uri=allow_uri
                )
            try:
                self._transfer_opened_mail_to_folder(
                    mail_item,
                    target_folder,
                    namespace,
                    mtime_source_path,
                    email_data,
                    conversion_options,
                    outlook=conversion_options.get("_session_outlook"),
                )
            finally:
                release_com_object(mail_item)

        _com_retry(f"import {os.path.basename(import_path)}", _do_import)

    def _import_with_outlook_native(self, namespace, target_folder, file_path, conversion_options):
        """
        Import via Outlook OpenSharedItem (best path for intact MIME: Date, Received
        chain, Content-Type, etc. stay with the message as Outlook stored them).

        Outlook rejects many LF-only .eml files and deep paths with 'Invalid path or URL'.
        We stage a flat CRLF copy beside the PST when needed.
        """
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in (".eml", ".emlx"):
            return f"Unsupported extension for native import: {ext}"

        def _try_native(import_path: str, mtime_source: str, *, allow_uri: bool = False):
            self._native_import_to_folder(
                namespace,
                target_folder,
                import_path,
                mtime_source,
                conversion_options,
                allow_uri=allow_uri,
            )

        last_err = None
        needs_staging = self._path_needs_native_staging(file_path)

        # 1) Fast path — only safe short paths without ()/@ etc.
        if (
            ext == ".eml"
            and not needs_staging
            and self._eml_source_likely_outlook_native_ready(file_path)
        ):
            try:
                _try_native(file_path, file_path, allow_uri=False)
                return "converted"
            except Exception as e:
                last_err = e
                if _is_outlook_rpc_error(e):
                    return str(e)
                logger.debug(
                    "OpenSharedItem direct failed for %s: %s",
                    log_sanitize(file_path),
                    e,
                )

        # 2) Staged CRLF copy — fixes LF-only test/real mail and deep source paths
        try:
            prefix = "import_"
            if ext == ".emlx":
                prefix = f"emlx_as_eml_{uuid.uuid4().hex[:8]}_"
            staged_path = self._stage_rfc822_eml_for_import(file_path, prefix=prefix)
            if not staged_path:
                sz = file_size_or_zero(file_path)
                if sz > NATIVE_IMPORT_MAX_BYTES:
                    return "Message too large for native import"
                return "Could not read source message content"
            _try_native(staged_path, file_path, allow_uri=False)
            return "converted"
        except Exception as e:
            last_err = e
            if _is_outlook_rpc_error(e):
                return str(e)
            logger.debug(
                "OpenSharedItem staged import failed for %s: %s",
                log_sanitize(file_path),
                e,
            )

        return str(last_err or "Native import failed")

    def _normalize_rfc822_line_endings(self, data):
        """Normalize RFC822 data to CRLF endings for better Outlook compatibility."""
        if not data:
            return data
        return _RFC822_CRLF_NORM_RE.sub(b"\r\n", data)

    def _emlx_to_rfc822_bytes(self, file_path):
        """
        Convert Apple .emlx to RFC822 .eml bytes.
        .emlx is usually: <byte_count_line>\\n<rfc822_message><plist...>
        """
        try:
            with open(file_path, "rb") as f:
                data = f.read()
        except OSError as e:
            logger.debug("Could not read EMLX %s: %s", file_path, e)
            return None

        newline_idx = data.find(b"\n")
        if newline_idx <= 0:
            return data

        first_line = data[:newline_idx].strip()
        if not first_line.isdigit():
            return data

        try:
            expected_len = int(first_line)
        except ValueError:
            return data

        start = newline_idx + 1
        end = start + expected_len
        if end <= len(data):
            return data[start:end]
        return data[start:]
    
    def _add_attachments(self, mail, attachments):
        """Write attachment bytes to temp files and attach with original filenames."""
        if not attachments:
            return
        temp_dir = tempfile.mkdtemp(prefix="me_att_")
        used_names: set[str] = set()
        try:
            for att in attachments:
                data = att.get("data")
                if not data:
                    continue
                if len(data) > MAX_ATTACHMENT_BYTES:
                    logger.warning(
                        "Skipping oversized attachment %s (%d bytes)",
                        log_sanitize(att.get("filename", "attachment")),
                        len(data),
                    )
                    continue
                safe_filename = self._unique_attachment_filename(
                    att.get("filename", "attachment"), used_names
                )
                path = os.path.join(temp_dir, safe_filename)
                try:
                    with open(path, "wb") as handle:
                        handle.write(data)
                    mail.Attachments.Add(path)
                    with self._lock:
                        self._temp_files.append(path)
                except (OSError, IOError) as e:
                    logger.warning(
                        "Could not add attachment %s: %s",
                        att.get("filename"),
                        e,
                    )
        finally:
            try:
                os.rmdir(temp_dir)
            except OSError:
                pass
        
    def show_completion(self, converted, skipped, errors, note="", lang=None, cancelled=False):
        """Show completion message"""
        if lang is None:
            lang = self._current_lang()
        if cancelled:
            self.status_label.config(text=t(lang, "status_cancelled"))
        else:
            self.status_label.config(text=t(lang, "status_complete"))
        self.progress.config(value=self.progress['maximum'])
        
        if cancelled:
            msg = t(
                lang,
                "msg_cancelled",
                converted=converted,
                skipped=skipped,
                errors=errors,
            )
        else:
            msg = t(
                lang,
                "msg_complete",
                converted=converted,
                skipped=skipped,
                errors=errors,
            )
        msg += note
        
        title = t(lang, "title_cancelled") if cancelled else t(lang, "title_complete")
        messagebox.showinfo(title, msg)


def main():
    root = tk.Tk()
    
    # Set style
    style = ttk.Style()
    style.theme_use('clam')
    
    EmlToPstConverter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
