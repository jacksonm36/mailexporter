# Performance & feature matrix

Mail Exporter implements the safe parts of common optimization proposals. **Outlook COM import stays single-threaded.**

| # | Proposal | Status | Module / notes |
|---|----------|--------|----------------|
| 1 | Parallel thread pool | **Done** | `parallel_processor.py` — prep only (`EML2PST_PARALLEL_WORKERS`). Not parallel COM. |
| 2 | Memory-mapped EML | **Done** | `mmap_processor.py` — `EML2PST_MMAP_THRESHOLD_MB` (default 4) |
| 3 | LRU caching | **Done** | `cache_manager.py` — header LRU + metadata TTL |
| 4 | Async prep I/O | **Done** | `async_processor.py` — `EML2PST_ASYNC_IO=1` (optional `aiofiles`; asyncio prep scheduler) |
| 4b | COM pipeline | **Done** | `com_pipeline.py` — `EML2PST_COM_PIPELINE=1` or `EML2PST_COM_WORKERS=1` (one STA worker; **not** multi-thread PST writes) |
| 5 | SQLite state DB | **Done** | `dedup_state.sqlite3` + `export_results.csv` (WAL, resume) |
| 6 | Smart dedup strategies | **Done** | `smart_dedup.py` — `EML2PST_DEDUP_STRATEGY=message_id\|fuzzy_subject\|thread` |
| 7 | Email filters | **Done** | `email_filter.py` — optional pre-import skip |
| 8 | Web dashboard | **Skipped** | Use Tk progress + ETA; no Flask |
| 9 | Checkpoint rollback | **Done** | `recovery_manager.py` — JSON beside PST (`EML2PST_CHECKPOINT_EVERY`) |
| 10 | Adaptive rate limit | **Done** | `rate_limiter.py` — `EML2PST_ADAPTIVE_RATE=1` |
| 11 | Attachment compression | **Partial** | Size cap `EML2PST_MAX_ATTACHMENT_MB`; no zip-in-PST (Outlook risk) |
| 12 | HTML preview service | **Skipped** | Optional; not needed for batch export |

## Security

- `path_security.py` + `test_security.py`

## Env quick reference

```text
EML2PST_PARALLEL_WORKERS=4
EML2PST_ASYNC_IO=1
EML2PST_COM_PIPELINE=1
EML2PST_MMAP_THRESHOLD_MB=4
EML2PST_DEDUP_STRATEGY=content_hash
EML2PST_ADAPTIVE_RATE=1
EML2PST_CHECKPOINT_EVERY=500
```

**Do not** set `EML2PST_COM_WORKERS` > 1 for one PST — Outlook MAPI is single-threaded; values > 1 are forced to one pipeline worker.
