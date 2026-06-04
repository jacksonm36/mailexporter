# Changelog

## v1.0.2

### Bug fixes

- **CSV Unicode** — `export_results.csv` is written as UTF-8 with BOM on new files (Excel opens Hungarian/UTF-8 subjects correctly). Disable BOM with `EML2PST_CSV_BOM=0`.
- **COM memory** — explicit `release_com_object()` after staged imports; periodic `gc.collect()` during long runs (configurable via `EML2PST_GC_EVERY`).
- **Crash recovery race** — serialized Outlook reconnect with `_outlook_recovery_lock` and `_outlook_recovering` so health probes do not start overlapping recovery.
- **HTML validation** — `body_looks_like_html()` only flags known HTML tags (avoids false positives from `<user@domain>` or URL angle brackets).
- **Import quality** — blank `None` sender/subject, raw HTML previews, wrong folder/PST, attachment sniffing (see README).

### Improvements

- **Checkpoint / resume** — `export_results.csv` (per-file status, flushed every 50 rows) plus `dedup_state.sqlite3` (converted paths even if CSV lags). Enable **Resume from prior export log** in the UI.
- **Progress ETA** — status bar shows estimated time remaining after the first few messages.
- **Path traversal** — staging subfolder names and attachment filenames sanitized (`_safe_staging_subdir`, `_sanitize_filename`).
- **Outlook retry** — `EML2PST_COM_RETRIES`, RPC cooldown, crash wait (`EML2PST_OUTLOOK_WAIT_MAX`), session refresh after failures.

### Performance

- **`path_security.py`**: `PathValidator` / `PathSecurityError` for traversal, extension checks, safe filenames, and optional `scan_directory_safe`.
- **Parallel prep** (`parallel_processor.py`, `EML2PST_PARALLEL_WORKERS`): thread pool prefetches dedup fingerprints while COM import runs on the main thread. Do not parallelize Outlook COM (unstable).
- **Async prep** (`async_processor.py`, `EML2PST_ASYNC_IO=1`): asyncio scheduler for prep; optional `aiofiles`.
- **COM pipeline** (`com_pipeline.py`, `EML2PST_COM_PIPELINE=1` or `EML2PST_COM_WORKERS=1`): one dedicated Outlook STA worker thread. `EML2PST_COM_WORKERS>1` is not supported for a single PST.
- **Memory-mapped I/O** for large `.eml` files (`mmap_processor.py`, default from 4 MB via `EML2PST_MMAP_THRESHOLD_MB`): parsing, SHA-256 dedup, fingerprint head/tail, and streamed CRLF staging without loading the full file into RAM.
- SQLite dedup for large mailboxes (default from 200k files).
- File sizes cached at scan time (avoids repeated `stat()` during export).
- Faster CRLF normalization for native import staging.
- Slow-file warnings (`EML2PST_SLOW_FILE_SEC`, default 120s).

### Documentation

- README: troubleshooting and Unicode/CSV sections.

## v1.0.1 and earlier

See git history on [github.com/jacksonm36/mailexporter](https://github.com/jacksonm36/mailexporter).
