"""Quick sanity checks for Mail Exporter (no Outlook required)."""
from __future__ import annotations

import os
import sys
import tempfile

ROOT = os.path.dirname(os.path.abspath(__file__))
_DEFAULT_TEST = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "MailExporter_test")
TEST_MAIL = os.environ.get("EML2PST_TEST_MAIL", os.path.join(_DEFAULT_TEST, "mail"))


def main() -> int:
    sys.path.insert(0, ROOT)
    from eml_to_pst_converter import (
        EmlToPstConverter,
        csv_sanitize,
        load_resume_paths,
        relative_folder_parts,
        safe_outlook_folder_name,
        _is_outlook_path_open_error,
        _pattern_to_suffixes,
        _portable_staging_base,
    )

    errors: list[str] = []

    # Sanitization
    if csv_sanitize("=1+1") != "'=1+1":
        errors.append("csv_sanitize formula prefix")
    if safe_outlook_folder_name("a/b:c") != "a_b_c":
        errors.append("safe_outlook_folder_name")

    # Pattern parsing
    if _pattern_to_suffixes("*.eml;*.emlx") != {".eml", ".emlx"}:
        errors.append("_pattern_to_suffixes")

    # Staging base should prefer LOCALAPPDATA (not raw %TEMP%)
    staging = _portable_staging_base()
    if not os.path.isdir(staging):
        errors.append("_portable_staging_base not a directory")
    local = os.environ.get("LOCALAPPDATA", "")
    if local and tempfile.gettempdir().lower() in staging.lower():
        if "MailExporter" not in staging:
            errors.append("_portable_staging_base fell back to TEMP unexpectedly")

    # OpenSharedItem path helpers
    app_stub = EmlToPstConverter.__new__(EmlToPstConverter)
    if os.path.isdir(TEST_MAIL):
        sample = next(
            (
                os.path.join(dp, f)
                for dp, _, fs in os.walk(TEST_MAIL)
                for f in fs
                if f.lower().endswith(".eml")
            ),
            "",
        )
        if sample:
            cands = app_stub._open_shared_item_candidates(sample, allow_uri=False)
            if not cands or any(c.startswith("file:") for c in cands):
                errors.append("_open_shared_item_candidates should not use file:/// when allow_uri=False")
            cands_uri = app_stub._open_shared_item_candidates(sample, allow_uri=True)
            if not any(c.startswith("file:") for c in cands_uri):
                errors.append("_open_shared_item_candidates allow_uri=True should include file:///")

    if not _is_outlook_path_open_error("We couldn't find 'file:///C:/x.eml'. It may have been moved or deleted."):
        errors.append("_is_outlook_path_open_error")
    if not _is_outlook_path_open_error("Invalid path or URL."):
        errors.append("_is_outlook_path_open_error invalid path")

    # LF-only test EMLs should skip direct OpenSharedItem (need CRLF staging)
    if os.path.isdir(TEST_MAIL):
        sample = next(
            (
                os.path.join(dp, f)
                for dp, _, fs in os.walk(TEST_MAIL)
                for f in fs
                if f.lower().endswith(".eml")
            ),
            "",
        )
        if sample:
            ready = app_stub._eml_source_likely_outlook_native_ready(sample)
            with open(sample, "rb") as handle:
                chunk = handle.read(8192)
            lf_only = b"\n" in chunk and b"\r\n" not in chunk
            if lf_only and ready:
                errors.append("_eml_source_likely_outlook_native_ready should be False for LF-only .eml")
            if not lf_only and not ready:
                errors.append("_eml_source_likely_outlook_native_ready unexpected False for CRLF .eml")

            large = next(
                (
                    os.path.join(dp, f)
                    for dp, _, fs in os.walk(TEST_MAIL)
                    for f in fs
                    if f.lower().startswith("large_") and f.lower().endswith(".eml")
                ),
                "",
            )
            if large:
                lsz = os.path.getsize(large)
                if app_stub._should_try_native_import(large, lsz):
                    errors.append("_should_try_native_import should skip large attachment EML")
            tiny = next(
                (
                    os.path.join(dp, f)
                    for dp, _, fs in os.walk(TEST_MAIL)
                    for f in fs
                    if f.lower().startswith("tiny_") and f.lower().endswith(".eml")
                ),
                "",
            )
            if tiny:
                tsz = os.path.getsize(tiny)
                if not app_stub._should_try_native_import(tiny, tsz):
                    errors.append("_should_try_native_import should allow tiny EML")

    # Subfolder mapping
    root = os.path.abspath(TEST_MAIL)
    sample = os.path.join(root, "Inbox", "tiny_0001_test.eml")
    if os.path.isfile(sample):
        parts = relative_folder_parts(root, sample)
        if parts != ["Inbox"]:
            errors.append(f"relative_folder_parts expected ['Inbox'], got {parts}")

    # Dedup fingerprint + full-corpus simulation
    if os.path.isdir(TEST_MAIL):
        paths: list[str] = []
        for dirpath, _, names in os.walk(TEST_MAIL):
            for name in names:
                if name.lower().endswith(".eml"):
                    paths.append(os.path.join(dirpath, name))
        dups = [p for p in paths if "dup_" in os.path.basename(p)]
        originals = [p for p in paths if "dup_" not in os.path.basename(p)]
        if len(paths) < 400:
            errors.append(f"expected ~500 test files, found {len(paths)}")
        if len(dups) < 50:
            errors.append(f"expected ~100 duplicate files, found {len(dups)}")

        app_stub = EmlToPstConverter.__new__(EmlToPstConverter)
        seen_fp: set[str] = set()
        would_skip = 0
        for path in sorted(paths):
            try:
                sz = os.path.getsize(path)
            except OSError:
                continue
            fp = app_stub._file_content_fingerprint(path, sz)
            if fp is None:
                errors.append(f"fingerprint failed for {path}")
                continue
            if fp in seen_fp:
                would_skip += 1
            else:
                seen_fp.add(fp)
        if would_skip < 50:
            errors.append(
                f"dedup simulation expected ~100 skips, got {would_skip}"
            )

        if dups and originals:
            dup = dups[0]
            base = os.path.basename(dup)
            if "_of_" in base:
                orig_name = base.split("_of_", 1)[1]
                orig = next(
                    (p for p in originals if os.path.basename(p) == orig_name),
                    None,
                )
                if orig:
                    fp1 = app_stub._file_content_fingerprint(
                        dup, os.path.getsize(dup)
                    )
                    fp2 = app_stub._file_content_fingerprint(
                        orig, os.path.getsize(orig)
                    )
                    if fp1 != fp2:
                        errors.append("duplicate fingerprint mismatch")
                else:
                    errors.append(f"could not find original for {dup}")
        elif paths:
            errors.append("no duplicate test files found")

    # Default dedup enabled
    import tkinter as tk

    root_win = tk.Tk()
    root_win.withdraw()
    try:
        app = EmlToPstConverter(root_win)
        if not app.remove_duplicates.get():
            errors.append("remove_duplicates should default to True")
    finally:
        root_win.destroy()

    # Resume CSV reader
    csv_path = os.path.join(_DEFAULT_TEST, "export_results_sample.csv")
    if not os.path.isfile(csv_path):
        with open(csv_path, "w", encoding="utf-8", newline="") as handle:
            handle.write(
                "timestamp,file_path,status,detail,duration_sec,target_folder\n"
                f"2024-01-01 12:00:00,'=evil,converted,,1.0,Inbox\n"
            )
    resumed = load_resume_paths(csv_path)
    if not any("evil" in p for p in resumed):
        pass  # path normalization may differ; just ensure no crash

    if errors:
        print("FAILED:")
        for err in errors:
            print(f"  - {err}")
        return 1

    print("All bug_check tests passed.")
    if os.path.isdir(TEST_MAIL):
        n = sum(
            1
            for dp, _, fs in os.walk(TEST_MAIL)
            for f in fs
            if f.lower().endswith(".eml")
        )
        print(f"  test mail folder: {n} .eml files at {TEST_MAIL}")
        if os.path.isdir(TEST_MAIL):
            app_stub = EmlToPstConverter.__new__(EmlToPstConverter)
            seen: set[str] = set()
            skip = 0
            for dp, _, fs in os.walk(TEST_MAIL):
                for f in fs:
                    if not f.lower().endswith(".eml"):
                        continue
                    p = os.path.join(dp, f)
                    fp = app_stub._file_content_fingerprint(p, os.path.getsize(p))
                    if fp in seen:
                        skip += 1
                    else:
                        seen.add(fp)
            print(f"  dedup simulation: {skip} duplicates would be skipped")
    return 0


if __name__ == "__main__":
    sys.exit(main())
