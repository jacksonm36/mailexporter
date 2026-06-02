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
from pathlib import Path
from email.utils import parsedate_to_datetime

from i18n import (
    LANG_EN,
    LANG_HU,
    LANGUAGE_NAMES,
    detect_default_lang,
    lang_from_display,
    t,
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

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)

# Outlook constants
OL_MAIL_ITEM = 0
OL_DISCARD = 1
OL_STORE_UNICODE = 2  # Unicode PST format (Outlook 2003+)

# Folder names
INBOX_FOLDER_NAME = "Inbox"

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
MAX_FILES = _env_int("EML2PST_MAX_FILES", 100_000)
# Treeview preview cap — full list kept in memory for conversion; UI shows a sample only.
UI_PREVIEW_FILE_LIMIT = _env_int("EML2PST_UI_PREVIEW", 500, lo=50, hi=5000)
SCAN_PROGRESS_EVERY = _env_int("EML2PST_SCAN_PROGRESS_EVERY", 1000, lo=100, hi=10000)
CONVERSION_UI_EVERY = _env_int("EML2PST_UI_UPDATE_EVERY", 25, lo=1, hi=500)
GC_EVERY_N_FILES = _env_int("EML2PST_GC_EVERY", 500, lo=0, hi=5000)
COM_RETRY_ATTEMPTS = _env_int("EML2PST_COM_RETRIES", 3, lo=1, hi=5)
CSV_FLUSH_EVERY = _env_int("EML2PST_CSV_FLUSH_EVERY", 50, lo=1, hi=1000)
COM_RPC_COOLDOWN = _env_float("EML2PST_RPC_COOLDOWN", 5.0, lo=1.0, hi=120.0)
STAGING_WRITE_SETTLE = _env_float("EML2PST_STAGING_SETTLE", 0.35, lo=0.05, hi=5.0)
COM_PACE_SEC = _env_float("EML2PST_COM_PACE", 0.02, lo=0.0, hi=2.0)
MAX_ATTACHMENT_BYTES = _env_int_mb("EML2PST_MAX_ATTACHMENT_MB", 25) * 1024 * 1024
# OpenSharedItem often fails on large multipart/attachment .eml — skip native above this size.
NATIVE_IMPORT_MAX_BYTES = _env_int_mb("EML2PST_NATIVE_MAX_MB", 384) * 1024 * 1024
NATIVE_IMPORT_SNIFF_BYTES = _env_int("EML2PST_NATIVE_SNIFF", 262144, lo=8192, hi=1_048_576)
FULL_DEDUP_HASH = os.environ.get("EML2PST_FULL_DEDUP", "").strip().lower() in ("1", "true", "yes")
PREFLIGHT_SIZE_SAMPLE = _env_int("EML2PST_PREFLIGHT_SAMPLE", 200, lo=50, hi=5000)


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
        "Call was rejected by callee",
        "Server execution failed",
        "The message filter indicated",
    )
    return any(m in text for m in markers)


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


def detect_outlook_bitness() -> int | None:
    """Return 32 or 64 if Outlook is installed, else None."""
    try:
        import winreg
    except ImportError:
        return None
    for ver in ("16.0", "15.0", "14.0", "12.0"):
        for bits, prefix in (
            (64, r"SOFTWARE\Microsoft\Office"),
            (32, r"SOFTWARE\WOW6432Node\Microsoft\Office"),
        ):
            try:
                winreg.OpenKey(
                    winreg.HKEY_LOCAL_MACHINE, rf"{prefix}\{ver}\Outlook"
                )
                return bits
            except OSError:
                continue
    return None


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


def export_log_paths(source_root: str, pst_path: str, export_mode: str) -> tuple[str, str]:
    """Return (csv_path, log_path) beside PST or source folder."""
    if export_mode != "mailbox" and pst_path:
        base = os.path.dirname(os.path.abspath(pst_path))
    elif source_root:
        base = source_root
    else:
        base = _portable_staging_base()
    return (
        os.path.join(base, "export_results.csv"),
        os.path.join(base, "export_log.txt"),
    )


def load_resume_paths(csv_path: str) -> set[str]:
    """Paths successfully converted in a prior run."""
    done: set[str] = set()
    if not os.path.isfile(csv_path):
        return done
    try:
        with open(csv_path, newline="", encoding="utf-8") as handle:
            for row in csv.DictReader(handle):
                if (row.get("status") or "").lower() == "converted":
                    path = row.get("file_path", "").strip()
                    if path:
                        done.add(os.path.normpath(os.path.abspath(path)))
    except OSError as exc:
        logger.warning("Could not read resume log %s: %s", csv_path, exc)
    return done


class EmlToPstConverter:
    def __init__(self, root):
        self.root = root
        self.root.title("Mail Exporter")
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
        self.use_file_mtime_for_date = tk.BooleanVar(value=True)
        self.lang_var = tk.StringVar(value=LANGUAGE_NAMES[detect_default_lang()])
        self.file_pattern = tk.StringVar(value="*.eml")
        self.eml_files: list[str] = []
        self._scan_truncated = False
        self.processed_hashes = set()
        
        # Thread safety
        self._lock = threading.Lock()
        self._is_converting = False
        self._cancel_requested = False
        self._temp_files = []
        self._conversion_options = {}
        self._export_log_handler = None
        self._csv_file = None
        self._csv_writer = None
        self._csv_rows_since_flush = 0
        self._conversion_thread = None
        self._dup_fingerprints: set[str] = set()
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
        self.root.title(t(lang, "window_title"))

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

        def on_progress(count: int):
            self.root.after(
                0,
                lambda c=count, lg=lang: self.status_label.config(
                    text=t(lg, "status_scanning_progress", n=c)
                ),
            )

        try:
            files, truncated = self._scan_folder_impl(
                folder, lang, pattern_snapshot, on_progress=on_progress
            )
            self.root.after(
                0,
                lambda f=files, tr=truncated, lg=lang: self._apply_scan_results(
                    f, lg, truncated=tr
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
                if on_progress and len(found_files) % SCAN_PROGRESS_EVERY == 0:
                    on_progress(len(found_files))
                if len(found_files) >= MAX_FILES:
                    truncated = True
                    logger.warning("File limit reached (%d)", MAX_FILES)
                    return found_files, truncated

        return found_files, truncated

    def _apply_scan_results(self, files: list[str], lang: str, *, truncated: bool = False):
        """Store scanned paths and refresh the preview list (full list used for export)."""
        with self._lock:
            self.eml_files = files
            self._scan_truncated = truncated
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
        total = len(self.eml_files)
        preview = self.eml_files[:UI_PREVIEW_FILE_LIMIT]
        for file_path in preview:
            self._insert_tree_row(file_path)
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
        if self.pst_option.get() == "new":
            file_path = filedialog.asksaveasfilename(
                title=t(lang, "dialog_save_pst"),
                defaultextension=".pst",
                filetypes=[pst_type, all_type],
            )
        else:
            file_path = filedialog.askopenfilename(
                title=t(lang, "dialog_open_pst"),
                filetypes=[pst_type, all_type],
            )
        
        if file_path:
            self.destination_path.set(file_path)
            
    def get_email_hash(self, file_path):
        """Full SHA-256 for duplicate detection (used when fingerprint collides)."""
        try:
            sha256 = hashlib.sha256()
            with open(file_path, 'rb') as f:
                for chunk in iter(lambda: f.read(65536), b''):
                    sha256.update(chunk)
            return sha256.hexdigest()
        except OSError as e:
            logger.warning("Could not hash file %s: %s", log_sanitize(file_path), log_sanitize(e))
            return None

    def _file_content_fingerprint(self, file_path: str, size: int) -> str | None:
        """Fast duplicate fingerprint: size + head/tail sample (avoids full read at scale)."""
        try:
            digest = hashlib.sha256()
            digest.update(str(size).encode("ascii"))
            with open(file_path, "rb") as handle:
                head = handle.read(65536)
                digest.update(head)
                if size > 131072:
                    handle.seek(size - 65536)
                    digest.update(handle.read(65536))
            return digest.hexdigest()
        except OSError as e:
            logger.warning(
                "Could not fingerprint file %s: %s",
                log_sanitize(file_path),
                log_sanitize(e),
            )
            return None
            
    def parse_eml(self, file_path):
        """
        Parse an EML file and return email data.
        Uses the stdlib parser so all MIME headers (Date, Received chain, etc.)
        stay available for transport-header and date preservation in Outlook.
        """
        try:
            with open(file_path, 'rb') as f:
                msg = BytesParser(policy=policy.default).parse(f)
            
            return {
                'subject': msg.get('Subject', '(No Subject)'),
                'from': msg.get('From', ''),
                'to': msg.get('To', ''),
                'cc': msg.get('Cc', ''),
                'date': msg.get('Date', ''),
                'body': self.get_email_body(msg),
                'attachments': self.get_attachments(msg),
                'message': msg
            }
        except Exception as e:
            logger.error("Error parsing EML file %s: %s", log_sanitize(file_path), log_sanitize(e))
            return None
            
    def get_email_body(self, msg):
        """Extract email body from message"""
        body = ""
        if msg.is_multipart():
            for part in msg.walk():
                content_type = part.get_content_type()
                if content_type == "text/plain":
                    try:
                        body = part.get_content()
                        break
                    except (KeyError, LookupError, UnicodeDecodeError) as e:
                        logger.debug("Could not decode text/plain part: %s", e)
                elif content_type == "text/html" and not body:
                    try:
                        body = part.get_content()
                    except (KeyError, LookupError, UnicodeDecodeError) as e:
                        logger.debug("Could not decode text/html part: %s", e)
        else:
            try:
                body = msg.get_content()
            except (KeyError, LookupError, UnicodeDecodeError):
                payload = msg.get_payload(decode=True)
                body = payload.decode('utf-8', errors='replace') if payload else ""
        return body
        
    def get_attachments(self, msg):
        """Extract attachments from message"""
        attachments = []
        if msg.is_multipart():
            for part in msg.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        safe_filename = self._sanitize_filename(filename)
                        data = part.get_payload(decode=True)
                        if data and len(data) > MAX_ATTACHMENT_BYTES:
                            logger.warning(
                                "Skipping oversized attachment %s (%d bytes)",
                                log_sanitize(safe_filename),
                                len(data),
                            )
                            continue
                        attachments.append({
                            'filename': safe_filename,
                            'data': data,
                            'content_type': part.get_content_type()
                        })
        return attachments
    
    def _sanitize_filename(self, filename):
        """Sanitize filename to prevent path traversal attacks"""
        # Remove any path components
        filename = os.path.basename(filename)
        # Remove potentially dangerous characters
        filename = re.sub(r'[<>:"/\\|?*\x00-\x1f]', '_', filename)
        # Ensure non-empty
        if not filename:
            filename = "attachment"
        # Limit length
        if len(filename) > 200:
            name, ext = os.path.splitext(filename)
            filename = name[:200-len(ext)] + ext
        return filename
        
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
        Staging beside the PST only — avoid %TEMP% (Outlook async reads fail there).
        """
        dirs: list[str] = []
        if self._staging_dir and os.path.isdir(self._staging_dir):
            dirs.append(self._staging_dir)
        fallback = _portable_staging_base()
        if fallback not in dirs:
            try:
                os.makedirs(fallback, exist_ok=True)
                dirs.append(fallback)
            except OSError:
                pass
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

    def _run_preflight(self, conversion_options: dict, lang: str) -> bool:
        """Warn about bitness mismatch and show summary before export."""
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
        outlook_bits = detect_outlook_bitness()
        if outlook_bits is None:
            lines.append(t(lang, "preflight_outlook_unknown"))
        elif outlook_bits != app_bits:
            lines.append(
                t(
                    lang,
                    "preflight_bitness",
                    app=app_bits,
                    outlook=outlook_bits,
                )
            )

        export_mode = conversion_options.get("pst_option")
        if export_mode != "mailbox":
            pst_path = conversion_options.get("destination_path", "")
            dest_dir = os.path.dirname(os.path.abspath(pst_path))
            if not os.path.isdir(dest_dir):
                messagebox.showerror(
                    t(lang, "title_error"),
                    t(lang, "err_dest_dir", path=dest_dir),
                )
                return False

        csv_path, _ = export_log_paths(
            conversion_options.get("source_root", ""),
            conversion_options.get("destination_path", ""),
            export_mode,
        )
        if conversion_options.get("resume_from_log") and os.path.isfile(csv_path):
            resumed = len(load_resume_paths(csv_path))
            if resumed:
                lines.append(t(lang, "preflight_resume", n=resumed))

        lines.append(t(lang, "preflight_continue"))
        return messagebox.askyesno(
            t(lang, "preflight_title"),
            "\n".join(lines),
        )

    def _setup_export_logging(self, log_path: str, csv_path: str, append_csv: bool):
        for path in (log_path, csv_path):
            parent = os.path.dirname(os.path.abspath(path))
            if parent:
                os.makedirs(parent, exist_ok=True)
        self._export_log_handler = logging.FileHandler(log_path, encoding="utf-8")
        self._export_log_handler.setFormatter(
            logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
        )
        logger.addHandler(self._export_log_handler)
        logger.info("Export started — CSV: %s", csv_path)

        write_header = not append_csv or not os.path.isfile(csv_path)
        mode = "a" if append_csv and os.path.isfile(csv_path) else "w"
        self._csv_file = open(csv_path, mode, newline="", encoding="utf-8")
        self._csv_writer = csv.writer(self._csv_file)
        self._csv_rows_since_flush = 0
        if write_header:
            self._csv_writer.writerow(CSV_HEADERS)
            self._csv_file.flush()

    def _teardown_export_logging(self):
        if self._csv_file:
            try:
                self._csv_file.flush()
                self._csv_file.close()
            except OSError:
                pass
            self._csv_file = None
            self._csv_writer = None
        if self._export_log_handler:
            logger.removeHandler(self._export_log_handler)
            try:
                self._export_log_handler.close()
            except OSError:
                pass
            self._export_log_handler = None

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

    def _resolve_target_folder_for_file(
        self,
        base_folder,
        file_path: str,
        conversion_options: dict,
        folder_cache: dict,
    ):
        if not conversion_options.get("preserve_subfolders"):
            return base_folder, ""
        source_root = conversion_options.get("source_root", "")
        parts = relative_folder_parts(source_root, file_path)
        if not parts:
            return base_folder, ""
        folder = self._get_or_create_folder_path(base_folder, parts, folder_cache)
        return folder, "\\".join(parts)

    def _get_or_create_folder_path(self, base_folder, parts: list[str], cache: dict):
        folder = base_folder
        built: list[str] = []
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
            self._conversion_options = {
                "destination_path": os.path.abspath(self.destination_path.get().strip()),
                "remove_duplicates": bool(self.remove_duplicates.get()),
                "strict_date_preservation": bool(self.strict_date_preservation.get()),
                "use_file_mtime_for_date": bool(self.use_file_mtime_for_date.get()),
                "preserve_subfolders": bool(self.preserve_subfolders.get()),
                "resume_from_log": bool(self.resume_from_log.get()),
                "pst_option": self.pst_option.get(),
                "mailbox_store_label": self.mailbox_store_var.get().strip(),
                "mailbox_folder_name": safe_outlook_folder_name(
                    self.mailbox_folder_name.get().strip() or "Imported EML"
                ),
                "source_root": resolve_source_root(
                    self.folder_path.get().strip(),
                    list(self.eml_files),
                ),
                "lang": lang,
            }
            self._is_converting = True
            self._cancel_requested = False
            
        if self._conversion_options["pst_option"] == "mailbox":
            if not self._conversion_options["mailbox_store_label"]:
                with self._lock:
                    self._is_converting = False
                messagebox.showwarning(
                    t(lang, "title_warning"), t(lang, "warn_no_mailbox")
                )
                return
        elif not self._conversion_options["destination_path"]:
            with self._lock:
                self._is_converting = False
            messagebox.showwarning(
                t(lang, "title_warning"), t(lang, "warn_no_destination")
            )
            return

        if not self._run_preflight(self._conversion_options, lang):
            with self._lock:
                self._is_converting = False
            return

        self._set_conversion_ui_active(True)

        thread = threading.Thread(
            target=self._convert_files_thread,
            args=(self._conversion_options.copy(),),
        )
        thread.daemon = True
        with self._lock:
            self._conversion_thread = thread
        thread.start()
    
    def _convert_files_thread(self, conversion_options):
        """Thread wrapper for conversion with COM initialization"""
        com_initialized = False
        try:
            if PYTHONCOM:
                PYTHONCOM.CoInitialize()
                com_initialized = True
            self.convert_files(conversion_options)
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
        conversion_options["log_path"] = log_path

        resume_paths: set[str] = set()
        if conversion_options.get("resume_from_log"):
            resume_paths = load_resume_paths(csv_path)
        conversion_options["resume_paths"] = resume_paths

        append_csv = bool(resume_paths) and conversion_options.get("resume_from_log")
        self._setup_export_logging(log_path, csv_path, append_csv=append_csv)

        with self._lock:
            total = len(self.eml_files)
            files_to_process = self.eml_files.copy()
            
        self.root.after(0, lambda: self.progress.config(maximum=total, value=0))
        with self._lock:
            self.processed_hashes.clear()
            self._dup_fingerprints.clear()
        
        cancelled = False
        try:
            if not OUTLOOK_AVAILABLE:
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
            self._teardown_export_logging()
    
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
        
        outlook = None
        namespace = None
        try:
            try:
                outlook = WIN32COM.Dispatch("Outlook.Application")
            except Exception as e:
                raise RuntimeError(f"Could not connect to Outlook: {e}") from e
            
            try:
                namespace = outlook.GetNamespace("MAPI")
            except Exception as e:
                raise RuntimeError(f"Could not access MAPI namespace: {e}") from e

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
            else:
                self._cleanup_stale_store_for_path(namespace, pst_path)
                self._update_status(t(lang, "status_creating_pst"))
                self._setup_pst_store(outlook, namespace, pst_path, conversion_options)
                time.sleep(1)
                pst_store = self._find_pst_store(namespace, pst_path)
                if not pst_store:
                    raise RuntimeError(
                        f"Could not access PST file after creation. Path: {pst_path}"
                    )
                root_folder = pst_store.GetRootFolder()
                target_folder = self._get_or_create_inbox(root_folder)
                note = t(lang, "note_pst_saved", path=pst_path)
                
            total = len(files_to_process)
            folder_cache: dict = {}
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
            target = pst_path.lower()
            stale_stores = []
            for store in namespace.Stores:
                try:
                    file_path = store.FilePath
                    if file_path and file_path.lower() == target and not os.path.exists(file_path):
                        stale_stores.append((store, file_path))
                except (AttributeError, OSError):
                    continue
            
            # Remove stale stores
            for store, file_path in stale_stores:
                try:
                    root = store.GetRootFolder()
                    namespace.RemoveStore(root)
                    logger.info("Removed stale store reference: %s", file_path)
                except Exception as e:
                    logger.warning("Could not remove stale store %s: %s", file_path, e)
                    
        except Exception as e:
            logger.debug("Target stale store cleanup failed: %s", e)
    
    def _setup_pst_store(self, outlook, namespace, pst_path, conversion_options):
        """Create or open PST store"""
        try:
            # First, remove any existing store reference with this path
            self._remove_existing_store(namespace, pst_path)
            
            if conversion_options.get("pst_option") == "new":
                # Remove existing file if present
                if os.path.exists(pst_path):
                    try:
                        os.remove(pst_path)
                    except OSError as e:
                        logger.warning("Could not remove existing PST: %s", e)
                
                # Create PST using AddStoreEx (creates Unicode PST)
                self._create_new_pst(outlook, namespace, pst_path)
            else:
                # Open existing PST
                if not os.path.exists(pst_path):
                    raise FileNotFoundError(f"PST file not found: {pst_path}")
                namespace.AddStore(pst_path)
        except FileNotFoundError:
            raise
        except Exception as e:
            raise RuntimeError(f"Could not create/open PST file: {e}") from e
    
    def _remove_existing_store(self, namespace, pst_path):
        """Remove any existing store reference with the given path"""
        pst_path_lower = pst_path.lower()
        stores_to_remove = []
        
        # Find stores matching this path
        for store in namespace.Stores:
            try:
                if store.FilePath and store.FilePath.lower() == pst_path_lower:
                    stores_to_remove.append(store)
            except (AttributeError, OSError):
                continue
        
        # Remove found stores
        for store in stores_to_remove:
            try:
                root = store.GetRootFolder()
                namespace.RemoveStore(root)
                logger.info("Removed existing store reference: %s", pst_path)
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
        pst_path_lower = pst_path.lower()
        
        for attempt in range(retries):
            for store in namespace.Stores:
                try:
                    if store.FilePath and store.FilePath.lower() == pst_path_lower:
                        return store
                except (AttributeError, OSError):
                    continue
            
            # Wait and retry if not found
            if attempt < retries - 1:
                time.sleep(1)
                logger.debug("PST store not found, retry %d/%d", attempt + 2, retries)
        
        return None
    
    def _get_or_create_inbox(self, root_folder):
        """Find or create Inbox folder in PST"""
        for folder in root_folder.Folders:
            if folder.Name.lower() == INBOX_FOLDER_NAME.lower():
                return folder
        return root_folder.Folders.Add(INBOX_FOLDER_NAME)

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
        for folder in parent_folder.Folders:
            if str(folder.Name).lower() == safe_name.lower():
                return folder
        return parent_folder.Folders.Add(safe_name)
    
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
        return self._get_or_create_inbox(pst_store.GetRootFolder())

    def _recover_outlook_session(self, conversion_options):
        """Reconnect Outlook COM after RPC/session failure."""
        logger.warning("Outlook COM unavailable — pausing and reconnecting...")
        time.sleep(COM_RPC_COOLDOWN)
        gc.collect()
        outlook = WIN32COM.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        target_folder = self._resolve_export_target_folder(namespace, conversion_options)
        return outlook, namespace, target_folder

    def _process_email_files(
        self,
        outlook,
        namespace,
        target_folder,
        files_to_process,
        total,
        conversion_options,
        folder_cache,
    ):
        """Process all email files"""
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
        for i, file_path in enumerate(files_to_process):
            if self._is_cancelled():
                cancelled = True
                logger.info("Cancel requested — stopping after %d/%d", i, total)
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
                if status == "error" and _is_outlook_rpc_error(detail):
                    consecutive_rpc_failures += 1
                    extra_pause = COM_RPC_COOLDOWN * min(consecutive_rpc_failures, 4)
                    if extra_pause > COM_RPC_COOLDOWN:
                        logger.warning(
                            "Repeated Outlook RPC failures (%d) — pausing %.1fs before reconnect",
                            consecutive_rpc_failures,
                            extra_pause,
                        )
                        time.sleep(extra_pause - COM_RPC_COOLDOWN)
                    session["outlook"], session["namespace"], session["target_folder"] = (
                        self._recover_outlook_session(conversion_options)
                    )
                    folder_cache.clear()
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
                if status == "converted":
                    consecutive_rpc_failures = 0
                    converted += 1
                elif status == "skipped":
                    skipped += 1
                    if detail:
                        skipped_messages.append(f"{current_file}: {detail}")
                else:
                    errors += 1
                    error_messages.append(f"{current_file}: {detail or 'Unknown error'}")
                self._write_csv_row(
                    file_path, status, detail or "", duration, target_label
                )
            except Exception as e:
                duration = time.perf_counter() - started
                if _is_outlook_rpc_error(e):
                    try:
                        session["outlook"], session["namespace"], session["target_folder"] = (
                            self._recover_outlook_session(conversion_options)
                        )
                        folder_cache.clear()
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
                        if status == "converted":
                            converted += 1
                        elif status == "skipped":
                            skipped += 1
                            if detail:
                                skipped_messages.append(f"{current_file}: {detail}")
                        else:
                            errors += 1
                            error_messages.append(f"{current_file}: {detail or 'Unknown error'}")
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

            done = i + 1
            if done % CONVERSION_UI_EVERY == 0 or done == total:
                self.root.after(
                    0,
                    lambda f=current_file, idx=i, tot=total, lg=lang: self.status_label.config(
                        text=t(lg, "status_converting", name=f, cur=idx + 1, total=tot)
                    ),
                )
                self.root.after(0, lambda v=done: self.progress.config(value=v))

            if GC_EVERY_N_FILES and done % GC_EVERY_N_FILES == 0:
                gc.collect()

            if COM_PACE_SEC > 0:
                time.sleep(COM_PACE_SEC)
        
        return converted, skipped, errors, error_messages, skipped_messages, cancelled
    
    def _process_single_email(self, namespace, target_folder, file_path, outlook, conversion_options):
        """Process a single email file"""
        norm_path = os.path.normpath(os.path.abspath(file_path))
        resume_paths = conversion_options.get("resume_paths") or set()
        if norm_path in resume_paths:
            return "skipped", t(
                conversion_options.get("lang", LANG_EN), "skip_resume"
            )

        try:
            sz = os.path.getsize(file_path)
            if sz > MAX_EML_FILE_BYTES:
                mb = MAX_EML_FILE_BYTES // (1024 * 1024)
                return (
                    "error",
                    f"File too large ({sz // (1024 * 1024)} MB); max {mb} MB (set EML2PST_MAX_FILE_MB)",
                )
        except OSError as e:
            return "error", f"Cannot read file: {e}"

        # Check for duplicates (fingerprint by default; optional full SHA-256 via EML2PST_FULL_DEDUP)
        if conversion_options["remove_duplicates"]:
            fingerprint = self._file_content_fingerprint(file_path, sz)
            if fingerprint is None:
                return "error", "Cannot compute duplicate fingerprint"
            with self._lock:
                if FULL_DEDUP_HASH:
                    full_hash = self.get_email_hash(file_path)
                    if not full_hash:
                        return "error", "Cannot compute duplicate hash"
                    if full_hash in self.processed_hashes:
                        return "skipped", "Duplicate content"
                    self.processed_hashes.add(full_hash)
                else:
                    if fingerprint in self._dup_fingerprints:
                        return "skipped", "Duplicate content"
                    self._dup_fingerprints.add(fingerprint)
        
        # Method 1: Native Outlook import (best MIME preservation: full RFC822 in PST)
        try_native = self._should_try_native_import(file_path, sz)
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
            )
            if strict_result == "converted":
                return "converted", ""
            return "skipped", (
                "Native import required to preserve original arrival date. "
                f"Reason: {native_result}; strict fallback failed: {strict_result}"
            )

        if _is_outlook_rpc_error(native_result):
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
        )
        if fallback_result == "converted":
            return "converted", ""
        return "error", fallback_result

    def _process_single_email_manual_fallback(
        self, file_path, target_folder, outlook, require_date_preservation, conversion_options
    ):
        """
        Manual fallback import.
        Writes PR_TRANSPORT_MESSAGE_HEADERS from the parsed message (MIME headers).
        If strict (require_date_preservation): delivery time from Date/Received only;
        if that fails, the message is skipped.
        If not strict and "file mtime" is on: Windows modified time on the original
        .eml is applied first for delivery/submit time; otherwise Date/Received, then mtime.
        """
        try:
            email_data = self.parse_eml(file_path)
            if not email_data:
                return "Could not parse email"

            mail = outlook.CreateItem(OL_MAIL_ITEM)
            mail.Subject = str(email_data['subject'] or "(No Subject)")
            if email_data.get('to'):
                mail.To = str(email_data['to'])
            if email_data.get('cc'):
                mail.CC = str(email_data['cc'])

            body = email_data['body']
            if isinstance(body, str):
                if '<html' in body.lower() or '<body' in body.lower():
                    mail.HTMLBody = body
                else:
                    mail.Body = body
            else:
                mail.Body = str(body) if body else ""

            try:
                if email_data.get('from'):
                    mail.SentOnBehalfOfName = str(email_data['from'])
            except (AttributeError, TypeError) as e:
                logger.debug("Could not set sender: %s", e)

            use_mtime = conversion_options.get("use_file_mtime_for_date", True)
            prefer_mtime_first = not require_date_preservation and use_mtime
            # Dates are re-applied after Move(); set headers only before Move to avoid duplicate work.
            date_ok = self._apply_original_metadata(
                mail,
                email_data,
                source_file_path=file_path,
                use_file_mtime=use_mtime,
                prefer_file_mtime_first=prefer_mtime_first,
                apply_dates=not prefer_mtime_first,
            )
            if require_date_preservation and not date_ok:
                return "Could not preserve original arrival date metadata"

            self._add_attachments(mail, email_data.get('attachments', []))

            mail.Save()
            try:
                moved_mail = mail.Move(target_folder)
            except Exception as move_err:
                try:
                    mail.Delete()
                except Exception as del_err:
                    logger.debug("Could not delete orphan draft after Move failure: %s", del_err)
                raise move_err
            # Move into PST often resets Received/list times to "now"; re-apply on the stored item.
            self._apply_original_metadata(
                moved_mail,
                email_data,
                source_file_path=file_path,
                use_file_mtime=use_mtime,
                prefer_file_mtime_first=prefer_mtime_first,
            )
            moved_mail.Save()
            return "converted"
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

    def _stamp_outlook_item_with_explorer_mtime(self, mail_item, source_eml_path: str) -> bool:
        """
        Set delivery/submit/creation-style MAPI times and MailItem.ReceivedTime /
        SentOn from the source file's Windows last-write time (Explorer Date modified).
        """
        if not source_eml_path or not os.path.isfile(source_eml_path):
            return False
        try:
            ts = os.path.getmtime(source_eml_path)
            dt_local = datetime.fromtimestamp(ts)
        except OSError as e:
            logger.warning("Could not read mtime for %s: %s", source_eml_path, e)
            return False

        com_val = self._outlook_com_datetime(dt_local)
        ok = False
        # MAPI PT_SYSTIME tags the message list / sorting often use
        mapi_time_urls = (
            "http://schemas.microsoft.com/mapi/proptag/0x0E060040",  # PR_MESSAGE_DELIVERY_TIME
            "http://schemas.microsoft.com/mapi/proptag/0x00390040",  # PR_CLIENT_SUBMIT_TIME
            "http://schemas.microsoft.com/mapi/proptag/0x30070040",  # PR_CREATION_TIME
            "http://schemas.microsoft.com/mapi/proptag/0x30080040",  # PR_LAST_MODIFICATION_TIME
        )
        try:
            pa = mail_item.PropertyAccessor
            for url in mapi_time_urls:
                try:
                    pa.SetProperty(url, com_val)
                    ok = True
                except Exception as e:
                    logger.debug("SetProperty %s: %s", url, e)
        except Exception as e:
            logger.warning("PropertyAccessor unavailable for mtime stamp: %s", e)

        try:
            mail_item.ReceivedTime = com_val
            ok = True
        except Exception as e:
            logger.debug("MailItem.ReceivedTime: %s", e)
        try:
            mail_item.SentOn = com_val
            ok = True
        except Exception as e:
            logger.debug("MailItem.SentOn: %s", e)

        if not ok:
            logger.warning(
                "Explorer mtime stamp had no effect for %s (Outlook may block writes)",
                source_eml_path,
            )
        return ok

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

        def _apply_mtime_from_source():
            nonlocal date_applied
            if not use_file_mtime or not source_file_path:
                return
            try:
                if os.path.isfile(source_file_path):
                    ts = os.path.getmtime(source_file_path)
                    dt_local = datetime.fromtimestamp(ts)
                    com_val = self._outlook_com_datetime(dt_local)
                    for url in (
                        "http://schemas.microsoft.com/mapi/proptag/0x0E060040",
                        "http://schemas.microsoft.com/mapi/proptag/0x00390040",
                        "http://schemas.microsoft.com/mapi/proptag/0x30070040",
                        "http://schemas.microsoft.com/mapi/proptag/0x30080040",
                    ):
                        try:
                            pa.SetProperty(url, com_val)
                        except Exception:
                            pass
                    try:
                        mail.ReceivedTime = com_val
                    except Exception:
                        pass
                    try:
                        mail.SentOn = com_val
                    except Exception:
                        pass
                    date_applied = True
                    logger.info(
                        "Applied delivery time from file modification time: %s",
                        source_file_path,
                    )
            except Exception as e:
                logger.debug("Could not apply file mtime as date: %s", e)

        if not apply_dates:
            return False

        if prefer_file_mtime_first:
            _apply_mtime_from_source()

        if not date_applied:
            try:
                date_raw = self._first_parseable_date_header(email_data)
                if date_raw:
                    dt = parsedate_to_datetime(date_raw)
                    if dt is not None:
                        if dt.tzinfo is None:
                            dt_local = dt
                        else:
                            dt_local = dt.astimezone().replace(tzinfo=None)
                        com_val = self._outlook_com_datetime(dt_local)
                        for url in (
                            "http://schemas.microsoft.com/mapi/proptag/0x0E060040",
                            "http://schemas.microsoft.com/mapi/proptag/0x00390040",
                        ):
                            try:
                                pa.SetProperty(url, com_val)
                            except Exception:
                                pass
                        try:
                            mail.ReceivedTime = com_val
                        except Exception:
                            pass
                        try:
                            mail.SentOn = com_val
                        except Exception:
                            pass
                        date_applied = True
            except Exception as e:
                logger.debug("Could not set original date properties: %s", e)

        if not date_applied:
            _apply_mtime_from_source()

        return date_applied

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
        try:
            with open(file_path, "rb") as handle:
                return handle.read()
        except OSError as e:
            logger.debug("Could not read %s: %s", file_path, e)
            return None

    def _should_try_native_import(self, file_path: str, file_size: int) -> bool:
        """
        OpenSharedItem is fast but unreliable for large files and messages with
        attachments. Skip native import for those and use manual import directly.
        """
        if file_size > NATIVE_IMPORT_MAX_BYTES:
            return False
        try:
            with open(file_path, "rb") as handle:
                head = handle.read(min(file_size, NATIVE_IMPORT_SNIFF_BYTES)).lower()
        except OSError:
            return False
        if not head:
            return False
        if b"content-disposition:" in head and b"attachment" in head:
            return False
        if b"multipart/" in head[:4096] and file_size > 128 * 1024:
            return False
        return True

    def _eml_source_likely_outlook_native_ready(self, file_path: str) -> bool:
        """
        True when OpenSharedItem can usually open the source file in place.
        LF-only RFC822 and very long paths need a CRLF staging copy first.
        """
        try:
            with open(file_path, "rb") as handle:
                chunk = handle.read(8192)
        except OSError:
            return False
        if not chunk:
            return False
        if b"\n" in chunk and b"\r\n" not in chunk:
            return False
        if len(os.path.abspath(file_path)) > 200:
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
        use_explorer_mtime: bool,
        *,
        allow_uri: bool = False,
    ):
        """OpenSharedItem + Move with COM retries for large export runs."""

        def _do_import():
            mail_item = self._open_shared_item(namespace, import_path, allow_uri=allow_uri)
            try:
                moved_item = mail_item.Move(target_folder)
                moved_item.Save()
                if use_explorer_mtime:
                    self._stamp_outlook_item_with_explorer_mtime(moved_item, mtime_source_path)
                    moved_item.Save()
            finally:
                del mail_item

        _com_retry(f"import {os.path.basename(import_path)}", _do_import)

    def _import_with_outlook_native(self, namespace, target_folder, file_path, conversion_options):
        """
        Import via Outlook OpenSharedItem (best path for intact MIME: Date, Received
        chain, Content-Type, etc. stay with the message as Outlook stored them).

        Outlook rejects many LF-only .eml files and deep paths with 'Invalid path or URL'.
        We stage a flat CRLF copy beside the PST when needed.
        """
        use_explorer_mtime = conversion_options.get("use_file_mtime_for_date", True)
        ext = os.path.splitext(file_path)[1].lower()
        if ext not in (".eml", ".emlx"):
            return f"Unsupported extension for native import: {ext}"

        def _try_native(import_path: str, mtime_source: str, *, allow_uri: bool = False):
            self._native_import_to_folder(
                namespace,
                target_folder,
                import_path,
                mtime_source,
                use_explorer_mtime,
                allow_uri=allow_uri,
            )

        last_err = None

        # 1) Fast path — CRLF .eml already at a short-ish absolute path
        if ext == ".eml" and self._eml_source_likely_outlook_native_ready(file_path):
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
            rfc822_bytes = self._load_rfc822_bytes(file_path)
            if not rfc822_bytes:
                return "Could not read source message content"
            if len(rfc822_bytes) > NATIVE_IMPORT_MAX_BYTES:
                return "Message too large for native import"
            rfc822_bytes = self._normalize_rfc822_line_endings(rfc822_bytes)
            prefix = "import_"
            if ext == ".emlx":
                prefix = f"emlx_as_eml_{uuid.uuid4().hex[:8]}_"
            staged_path = self._stage_file_for_native_import(
                file_path,
                prefix=prefix,
                data=rfc822_bytes,
            )
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
        normalized = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
        return b"\r\n".join(normalized.split(b"\n"))

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
        """Add attachments to mail item"""
        for att in attachments:
            if not att.get('data'):
                continue
            
            temp_path = None
            try:
                filename = att.get('filename', 'attachment')
                safe_filename = self._sanitize_filename(filename)
                _, ext = os.path.splitext(safe_filename)
                
                fd, temp_path = tempfile.mkstemp(suffix=ext, prefix=f"eml_att_{uuid.uuid4().hex[:8]}_")
                with self._lock:
                    self._temp_files.append(temp_path)
                try:
                    os.write(fd, att['data'])
                finally:
                    os.close(fd)
                
                mail.Attachments.Add(temp_path)
            except (OSError, IOError) as e:
                logger.warning("Could not add attachment %s: %s", att.get('filename'), e)
            finally:
                if temp_path and os.path.exists(temp_path):
                    try:
                        os.remove(temp_path)
                        with self._lock:
                            if temp_path in self._temp_files:
                                self._temp_files.remove(temp_path)
                    except OSError as e:
                        logger.debug("Could not remove temp file %s: %s", temp_path, e)
        
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
