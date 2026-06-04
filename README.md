# Mail Exporter

Portable Windows GUI to import **EML/EMLX** files (including Windows Live Mail folders) into:

- a new or existing **Outlook PST** file, or
- an **Outlook mailbox** (Exchange / Microsoft 365 / any account configured in Outlook)

No Python or pip is required on the target PC when you use the pre-built executable.

## Download / portable exe

After building, use one of these outputs in `dist/`:

| Output | Use when |
|--------|----------|
| **`MailExporter_x32.exe`** (single file) | Copy one file to USB / another PC |
| **`MailExporter_x32/MailExporter_x32.exe`** (folder) | Antivirus blocks single-file builds — copy the whole folder |

Build commands:

```bash
python build_exe.py           # single portable .exe
python build_exe.py --onedir  # folder bundle (fewer antivirus false positives)
```

**Requirements on the PC where you run the exe:**

- Windows 10/11
- **Microsoft Outlook** installed and working
- **Same bitness** as the exe (64-bit Outlook needs `MailExporter_x64.exe` from a 64-bit Python build)

Nothing else to install — the exe bundles Python and pywin32.

### If the exe will not run ("Access is denied" or disappears)

PyInstaller apps are often flagged by **Windows Defender** or other antivirus. The file may appear in `dist/` but be locked or quarantined.

1. Open **Windows Security** → **Protection history**
2. Find **MailExporter_x32.exe** → **Allow** / **Restore**
3. Add an exclusion for `Documents\mailexporter\dist`
4. Rebuild: `python build_exe.py --onedir` and run the exe inside `dist\MailExporter_x32\`
5. Right-click the exe → **Properties** → if you see **Unblock**, check it → OK

## Quick start

1. Double-click `MailExporter_x64.exe`.
2. Click **Live Mail Folder** (or **Add Files**) and pick the folder with `.eml` files.
3. Choose export mode:
   - **Create New PST File** / **Save to Existing PST File**, or
   - **Import to Outlook Mailbox (Exchange / M365)** — pick mailbox and target folder name.
4. Click **Convert**. Keep Outlook open during the run.

### Options

- **Preserve subfolder structure** — mirrors Live Mail / source folders in the PST or mailbox target
- **Resume from prior export log** — skips files already marked `converted` in `export_results.csv`
- **Cancel** — stops after the current message; partial results are saved to the CSV and log

After export, check **`export_results.csv`** and **`export.log`** next to the PST (or source folder for mailbox mode). A copy is also written to **`export_log.txt`** for older workflows.

### Validate export quality (checker)

Use **`export_checker.py`** (or build **`ExportChecker_x32.exe`** with `python build_exe.py --checker`) on the same PC as Outlook to find blank **None** rows, raw HTML in the preview, missing attachments, and mismatches vs source `.eml` files:

```bash
python export_checker.py --pst C:\path\to\export.pst
python export_checker.py --pst export.pst --folder "Account (user@domain" --issues-only
python export_checker.py --pst export.pst --csv export_results.csv --compare-eml
```

Writes **`validation_report_<pst>_<timestamp>.csv`** next to the PST. Exit code **1** if any message has **error**-level issues (re-import recommended after updating Mail Exporter).

**During export**, each message is validated immediately after it lands in the PST (default **ON**): correct **folder** (Inbox/Sent/… under the right account path), correct **PST store**, sender/subject must not be blank/`None`, HTML must not show as raw tags, attachments must match the `.eml`. Failed items are repaired automatically (including **Move** into the right folder); if still bad, the app deletes them and retries via `Items.Add`, or records an error in `export_results.csv`. Disable with `EML2PST_VALIDATE_IMPORT=0`.

If **Outlook crashes or hangs**, export **pauses** and polls until `OUTLOOK.EXE` is healthy again (default wait up to 1 hour, `EML2PST_OUTLOOK_WAIT_MAX`). Restart Outlook manually if needed; the run resumes on the current file when COM reconnects.

## Build from source (developers only)

```bash
pip install -r requirements.txt
python build_exe.py
```

Output: `dist/MailExporter_x64.exe`

Run from source (dev):

```bash
python eml_to_pst_converter.py
```

## Windows Live Mail location

Default storage:

`%LOCALAPPDATA%\Microsoft\Windows Live Mail`

Use the **Live Mail Folder** button to scan it automatically.

## Notes

- Data never leaves your PC — all processing is local via Outlook COM.
- Large mailboxes take time; speed is limited by Outlook, not this app.
- For Exchange import, messages appear in the folder you specify (default: `Imported EML`) and sync per your Outlook/Exchange settings.

## Troubleshooting

| Symptom | What to do |
|--------|------------|
| Export stops with “waiting for Outlook” | Restart Outlook; export resumes when COM reconnects (up to `EML2PST_OUTLOOK_WAIT_MAX`, default 1 hour). |
| **None** sender/subject or raw `<br/>` in PST | Built with v1.0.2+; delete bad PST items (or new PST) and re-import — resume does not rewrite old items. |
| Blank rows after resume | Keep `export_results.csv` and `dedup_state.sqlite3` in the same folder as the PST; check **Resume from prior export log**. |
| “Access is denied” / exe quarantined | See [If the exe will not run](#if-the-exe-will-not-run-access-is-denied-or-disappears) above. |
| 32-bit vs 64-bit mismatch | Use `MailExporter_x32.exe` with 32-bit Outlook (typical on older VMs). |
| Very slow single messages | Log warns after `EML2PST_SLOW_FILE_SEC` (default 120s); check huge attachments or damaged `.eml`. |
| Antivirus blocks build | `python build_exe.py --onedir` and run from the folder bundle. |

## Unicode and CSV export

- Message text uses UTF-8 end-to-end where possible (RFC 2047 decoding, HTML `charset=utf-8`).
- **`export_results.csv`** is UTF-8 with BOM on a **new** export so Excel shows Hungarian characters correctly.
- To write plain UTF-8 without BOM: set `EML2PST_CSV_BOM=0` before starting.
- Paths and detail fields are sanitized for Excel formula injection (`csv_sanitize`).

## Environment variables (common)

| Variable | Default | Purpose |
|----------|---------|---------|
| `EML2PST_VALIDATE_IMPORT` | `1` | Per-message PST quality checks |
| `EML2PST_OUTLOOK_WAIT_MAX` | `3600` | Seconds to wait for Outlook after crash |
| `EML2PST_PST_CHUNK_SIZE` | `100` | Batch size for PST import |
| `EML2PST_DEDUP_BACKEND` | `auto` | `sqlite` for large jobs |
| `EML2PST_CSV_BOM` | `1` | UTF-8 BOM on new CSV |
| `EML2PST_SLOW_FILE_SEC` | `120` | Log warning if one file exceeds this |
| `EML2PST_MMAP_THRESHOLD_MB` | `4` | Use memory-mapped reads for larger `.eml` files (`0` = off) |
| `EML2PST_PARALLEL_WORKERS` | `0` | Parallel prep threads for dedup fingerprints (`0` = off). **Outlook import stays single-threaded.** |
| `EML2PST_PARALLEL_PARSE` | `0` | Also parse MIME in parallel (high RAM on huge jobs) |
| `EML2PST_DEDUP_STRATEGY` | `content_hash` | `message_id`, `fuzzy_subject`, or `thread` |
| `EML2PST_ADAPTIVE_RATE` | `0` | Adaptive COM pacing on failures |
| `EML2PST_CHECKPOINT_EVERY` | `0` | JSON checkpoint every N messages (off=0) |
| `EML2PST_FILTER_MAX_MB` | — | Skip .eml larger than N MB |
| `EML2PST_FILTER_ATTACHMENTS` | — | `only` or `none` |

See [PERFORMANCE_FEATURES.md](PERFORMANCE_FEATURES.md) for the full feature matrix (items 1–12).

See `CHANGELOG.md` for the full v1.0.2 list.

## Security tests

```bash
python test_security.py
python bug_check.py
```

## License

Open source — internal company use.
