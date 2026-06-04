# Security notes

## Threat model

This is a **local GUI** that reads user-chosen folders, writes a PST path the user selects, and drives **Microsoft Outlook** via COM. It does not expose a network service.

## Mitigations in code

- **Glob / patterns**: Only `*.eml` / `*.emlx`-style patterns are allowed (`_is_valid_pattern`); recursive search is capped at `MAX_FILES` (50 000).
- **Path traversal**: Attachment filenames are sanitized (`_sanitize_filename`). Staging folder name from `EML2PST_STAGING_SUBDIR` is restricted to a **single path segment** (no `..`, separators, or `:`).
- **Subprocess**: `pip install pywin32` uses `subprocess.check_call` with a **fixed argument list** (no `shell=True`).
- **Large files**: Each message file is rejected if over `EML2PST_MAX_FILE_MB` (default **100** MB, clamped 1–2048 MB) to limit memory exhaustion from a single `.eml`.
- **Env parsing**: `EML2PST_POST_BATCH_DELAY` is parsed as float and clamped (0–600 s).

## Residual risks

- **Malicious `.eml`**: Stdlib email parsing and MIME bodies can still use significant CPU/memory within the per-file cap. Only open mail you trust.
- **Outlook / COM**: Outlook runs with the user’s privileges; the app does not sandbox COM.
- **“New PST”**: Overwrites an existing file at the chosen path if present (by design).

## Reporting

If you find a security issue, report it privately to the repository maintainer.
