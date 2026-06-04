#!/usr/bin/env python3
"""
Validate Mail Exporter output: detect blank "None" rows, raw HTML in body,
missing attachments, and mismatches vs source .eml files.

Requires Microsoft Outlook (same bitness as this Python / built exe).

Examples:
  python export_checker.py --pst C:\\exports\\pq2.pst
  python export_checker.py --pst pq2.pst --csv export_results.csv --compare-eml
  python export_checker.py --pst pq2.pst --folder "PQ (pqinfo@ 252" --report issues.csv
"""

from __future__ import annotations

import argparse
import csv
import logging
import os
import sys
import time
from collections import Counter
from datetime import datetime

from mail_validation import (
    ISSUE_AMBIGUOUS_IN_PST,
    ISSUE_EML_MISSING_FILE,
    ISSUE_MISSING_ATTACHMENT,
    ISSUE_NOT_FOUND_IN_PST,
    SEVERITY_ERROR,
    SEVERITY_WARNING,
    MailInspection,
    MailIssue,
    compare_eml_to_outlook,
    inspect_outlook_fields,
    inspect_parsed_eml,
    load_converted_paths_from_csv,
    outlook_text_is_blank,
    parse_eml_summary,
    read_outlook_mail_snapshot,
    sniff_eml_has_attachments,
    validate_import_against_eml,
    _normalize_match_key,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger("ExportChecker")

OL_MAIL_ITEM = 43

try:
    import pythoncom
    import win32com.client as win32com_client
except ImportError:
    pythoncom = None
    win32com_client = None


def normalize_pst_path(pst_path: str) -> str:
    path = os.path.normpath(os.path.abspath((pst_path or "").strip()))
    if path and not path.lower().endswith(".pst"):
        path += ".pst"
    return path


def pst_paths_equal(left: str, right: str) -> bool:
    if not left or not right:
        return False
    a = os.path.normpath(os.path.abspath(left))
    b = os.path.normpath(os.path.abspath(right))
    if os.name == "nt":
        return os.path.normcase(a) == os.path.normcase(b)
    return a == b


def _safe_com_str(value) -> str:
    if value is None:
        return ""
    try:
        return str(value).strip()
    except Exception:
        return ""


def _read_outlook_mail_fields(item) -> dict:
    fields = read_outlook_mail_snapshot(item)
    received = ""
    try:
        rt = item.ReceivedTime
        if rt is not None:
            received = str(rt)[:16]
    except Exception:
        pass
    folder_name = ""
    try:
        folder_name = _safe_com_str(item.Parent.Name)
    except Exception:
        pass
    fields["received"] = received
    fields["folder"] = folder_name
    fields["entry_id"] = _safe_com_str(getattr(item, "EntryID", ""))
    return fields


def _folder_name_matches(folder_name: str, filter_text: str) -> bool:
    if not filter_text:
        return True
    return filter_text.lower() in (folder_name or "").lower()


def iter_mail_items(folder, *, folder_filter: str = "", recursive: bool = True):
    """Yield (MailItem, folder_path_str) under folder."""
    try:
        folder_name = _safe_com_str(folder.Name)
    except Exception:
        folder_name = ""
    path = folder_name
    try:
        parent = folder.Parent
        if parent is not None and getattr(parent, "Name", None):
            parent_name = _safe_com_str(parent.Name)
            if parent_name and parent_name != folder_name:
                path = f"{parent_name}/{folder_name}"
    except Exception:
        pass

    if _folder_name_matches(folder_name, folder_filter) or _folder_name_matches(
        path, folder_filter
    ):
        try:
            items = folder.Items
            count = int(items.Count)
        except Exception:
            count = 0
        for idx in range(1, count + 1):
            try:
                item = items.Item(idx)
            except Exception:
                continue
            try:
                if int(item.Class) != OL_MAIL_ITEM:
                    continue
            except Exception:
                continue
            yield item, path

    if not recursive:
        return
    try:
        subfolders = folder.Folders
        sub_count = int(subfolders.Count)
    except Exception:
        return
    for idx in range(1, sub_count + 1):
        try:
            sub = subfolders.Item(idx)
        except Exception:
            continue
        yield from iter_mail_items(
            sub, folder_filter=folder_filter, recursive=True
        )


class OutlookPstSession:
    def __init__(self, pst_path: str):
        self.pst_path = normalize_pst_path(pst_path)
        self.outlook = None
        self.namespace = None
        self.store = None

    def __enter__(self):
        if win32com_client is None:
            raise RuntimeError("pywin32 is required (pip install pywin32)")
        if pythoncom is not None:
            pythoncom.CoInitialize()
        self.outlook = win32com_client.Dispatch("Outlook.Application")
        self.namespace = self.outlook.GetNamespace("MAPI")
        if os.path.isfile(self.pst_path):
            try:
                self.namespace.AddStore(self.pst_path)
                time.sleep(0.4)
            except Exception as e:
                logger.debug("AddStore: %s", e)
        self.store = self._find_store()
        if self.store is None:
            raise RuntimeError(f"PST not attached in Outlook: {self.pst_path}")
        return self

    def __exit__(self, *args):
        self.store = None
        self.namespace = None
        self.outlook = None
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    def _find_store(self):
        expected = self.pst_path
        for idx in range(1, int(self.namespace.Stores.Count) + 1):
            try:
                store = self.namespace.Stores.Item(idx)
                fp = ""
                try:
                    fp = normalize_pst_path(str(store.FilePath or ""))
                except Exception:
                    pass
                if fp and pst_paths_equal(fp, expected):
                    return store
            except Exception:
                continue
        return None

    def root_folder(self):
        return self.store.GetRootFolder()

    def iter_messages(self, *, folder_filter: str = "", recursive: bool = True):
        root = self.root_folder()
        yield from iter_mail_items(
            root, folder_filter=folder_filter, recursive=recursive
        )


def inspect_outlook_mail_item(
    item,
    *,
    folder_path: str = "",
    source_eml: str = "",
    expected_sniff_attachments: bool = False,
) -> MailInspection:
    fields = _read_outlook_mail_fields(item)
    source = source_eml or f"{folder_path}|{fields.get('entry_id', '')[:24]}"
    summary = None
    if source_eml and os.path.isfile(source_eml):
        summary = parse_eml_summary(source_eml)
    inspection = validate_import_against_eml(
        fields,
        summary,
        source=source,
    )
    if expected_sniff_attachments and not summary:
        extra = inspect_outlook_fields(
            expected_sniff_attachments=True,
            attachment_count=fields["attachment_count"],
            source=source,
        )
        for issue in extra.issues:
            if issue.code == ISSUE_MISSING_ATTACHMENT:
                inspection.issues.append(issue)
    return inspection


def scan_pst(
    pst_path: str,
    *,
    folder_filter: str = "",
    max_items: int = 0,
    progress_every: int = 500,
) -> tuple[list[MailInspection], int]:
    results: list[MailInspection] = []
    scanned = 0
    with OutlookPstSession(pst_path) as session:
        for item, folder_path in session.iter_messages(
            folder_filter=folder_filter, recursive=True
        ):
            scanned += 1
            insp = inspect_outlook_mail_item(item, folder_path=folder_path)
            results.append(insp)
            if progress_every and scanned % progress_every == 0:
                logger.info("Scanned %d mail item(s) in PST...", scanned)
            if max_items and scanned >= max_items:
                break
    return results, scanned


def _build_pst_index(
    pst_path: str,
    *,
    folder_filter: str = "",
    max_items: int = 0,
) -> tuple[dict[tuple[str, str], list[dict]], int]:
    """Index PST messages by (subject_key, sender_key) for .eml matching."""
    index: dict[tuple[str, str], list[dict]] = {}
    scanned = 0
    with OutlookPstSession(pst_path) as session:
        for item, folder_path in session.iter_messages(
            folder_filter=folder_filter, recursive=True
        ):
            scanned += 1
            fields = _read_outlook_mail_fields(item)
            subj_k, snd_k = _normalize_match_key(
                fields["subject"],
                fields["sender_name"] or fields["sender_email"],
            )
            key = (subj_k, snd_k)
            index.setdefault(key, []).append({**fields, "folder": folder_path})
            if max_items and scanned >= max_items:
                break
    return index, scanned


def compare_csv_to_pst(
    csv_path: str,
    pst_path: str,
    *,
    folder_filter: str = "",
    max_compare: int = 0,
    progress_every: int = 200,
) -> tuple[list[MailInspection], int, int]:
    rows = load_converted_paths_from_csv(csv_path)
    if max_compare and len(rows) > max_compare:
        rows = rows[:max_compare]
    logger.info(
        "Building PST index for %d converted path(s) (folder filter=%r)...",
        len(rows),
        folder_filter or "(all)",
    )
    index, pst_scanned = _build_pst_index(
        pst_path, folder_filter=folder_filter
    )
    logger.info("PST index: %d item(s) scanned", pst_scanned)

    results: list[MailInspection] = []
    for i, row in enumerate(rows, start=1):
        eml_path = row["file_path"]
        if not os.path.isfile(eml_path):
            insp = MailInspection(source=eml_path)
            insp.add(ISSUE_EML_MISSING_FILE, SEVERITY_ERROR, "File not found")
            results.append(insp)
            continue
        summary = parse_eml_summary(eml_path)
        eml_insp = inspect_parsed_eml(summary, source=eml_path)
        subj_k, snd_k = _normalize_match_key(
            (summary or {}).get("subject", ""),
            (summary or {}).get("from", ""),
        )
        matches = index.get((subj_k, snd_k), [])
        if not subj_k and not snd_k:
            matches = []
        if not matches and subj_k:
            matches = [
                entry
                for (sk, _), entries in index.items()
                if sk == subj_k
                for entry in entries
            ]
        if not matches:
            eml_insp.add(
                ISSUE_NOT_FOUND_IN_PST,
                SEVERITY_WARNING,
                "No PST item with same subject/sender key",
            )
            results.append(eml_insp)
        elif len(matches) > 3:
            eml_insp.add(
                ISSUE_AMBIGUOUS_IN_PST,
                SEVERITY_WARNING,
                f"{len(matches)} PST candidates",
            )
            results.append(eml_insp)
        else:
            match = matches[0]
            out_insp = inspect_outlook_fields(
                subject=match["subject"],
                sender_name=match["sender_name"],
                sender_email=match["sender_email"],
                body=match["body"],
                html_body=match["html_body"],
                attachment_count=match["attachment_count"],
                source=eml_path,
                expected_sniff_attachments=bool(
                    summary and summary.get("sniff_attachments")
                ),
            )
            for issue in eml_insp.issues:
                out_insp.issues.append(issue)
            for issue in compare_eml_to_outlook(summary, out_insp):
                out_insp.issues.append(issue)
            results.append(out_insp)
        if progress_every and i % progress_every == 0:
            logger.info("Compared %d / %d .eml file(s)...", i, len(rows))
    return results, len(rows), pst_scanned


def summarize_inspections(inspections: list[MailInspection]) -> dict:
    issue_counts: Counter[str] = Counter()
    error_items = 0
    warning_items = 0
    ok_items = 0
    for insp in inspections:
        has_err = any(i.severity == SEVERITY_ERROR for i in insp.issues)
        has_warn = any(i.severity == SEVERITY_WARNING for i in insp.issues)
        if has_err:
            error_items += 1
        elif has_warn:
            warning_items += 1
        else:
            ok_items += 1
        for issue in insp.issues:
            issue_counts[issue.code] += 1
    return {
        "total": len(inspections),
        "ok": ok_items,
        "with_errors": error_items,
        "warnings_only": warning_items,
        "issue_counts": dict(issue_counts),
    }


def write_report_csv(path: str, inspections: list[MailInspection]) -> None:
    parent = os.path.dirname(os.path.abspath(path))
    if parent:
        os.makedirs(parent, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(
            (
                "source",
                "subject",
                "sender",
                "severity",
                "issue_codes",
                "details",
            )
        )
        for insp in inspections:
            if not insp.issues:
                writer.writerow(
                    (
                        insp.source,
                        insp.subject,
                        insp.sender,
                        "ok",
                        "",
                        "",
                    )
                )
                continue
            worst = SEVERITY_ERROR if any(
                i.severity == SEVERITY_ERROR for i in insp.issues
            ) else SEVERITY_WARNING
            codes = ";".join(i.code for i in insp.issues)
            details = " | ".join(
                f"{i.code}: {i.detail}" if i.detail else i.code for i in insp.issues
            )
            writer.writerow(
                (
                    insp.source,
                    insp.subject,
                    insp.sender,
                    worst,
                    codes,
                    details,
                )
            )


def write_issues_only_csv(path: str, inspections: list[MailInspection]) -> None:
    bad = [i for i in inspections if i.issues]
    write_report_csv(path, bad)


def scan_eml_paths(paths: list[str]) -> list[MailInspection]:
    results: list[MailInspection] = []
    for path in paths:
        if not os.path.isfile(path):
            insp = MailInspection(source=path)
            insp.add(ISSUE_EML_MISSING_FILE, SEVERITY_ERROR, "File not found")
            results.append(insp)
            continue
        summary = parse_eml_summary(path)
        results.append(inspect_parsed_eml(summary, source=path))
    return results


def default_report_path(pst_path: str) -> str:
    base = os.path.splitext(os.path.basename(pst_path))[0]
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    return os.path.join(
        os.path.dirname(normalize_pst_path(pst_path)),
        f"validation_report_{base}_{stamp}.csv",
    )


def default_csv_beside_pst(pst_path: str) -> str:
    return os.path.join(
        os.path.dirname(normalize_pst_path(pst_path)),
        "export_results.csv",
    )


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Check Mail Exporter PST output for blank/corrupt/misaligned messages.",
    )
    parser.add_argument(
        "--pst",
        required=True,
        help="Path to the export .pst file (must be openable in Outlook)",
    )
    parser.add_argument(
        "--csv",
        default="",
        help="export_results.csv path (default: beside PST)",
    )
    parser.add_argument(
        "--compare-eml",
        action="store_true",
        help="Compare each converted .eml in CSV to matching PST items",
    )
    parser.add_argument(
        "--folder",
        default="",
        help="Only scan/compare under folders whose name contains this text",
    )
    parser.add_argument(
        "--report",
        default="",
        help="Write full validation report CSV (default: validation_report_<pst>_<time>.csv)",
    )
    parser.add_argument(
        "--issues-only",
        action="store_true",
        help="Report CSV lists only rows with problems",
    )
    parser.add_argument(
        "--max",
        type=int,
        default=0,
        help="Stop after N messages (0 = no limit)",
    )
    parser.add_argument(
        "--eml",
        action="append",
        default=[],
        help="Also validate specific .eml file(s) without PST (repeatable)",
    )
    args = parser.parse_args(argv)

    pst_path = normalize_pst_path(args.pst)
    if not os.path.isfile(pst_path):
        logger.error("PST not found: %s", pst_path)
        return 2

    inspections: list[MailInspection] = []

    if args.eml:
        logger.info("Checking %d .eml file(s) (parse only)...", len(args.eml))
        inspections.extend(scan_eml_paths(args.eml))

    csv_path = args.csv.strip() or default_csv_beside_pst(pst_path)

    if args.compare_eml:
        if not os.path.isfile(csv_path):
            logger.error("CSV not found for --compare-eml: %s", csv_path)
            return 2
        logger.info("Comparing converted rows in %s to %s", csv_path, pst_path)
        compared, n_eml, n_pst = compare_csv_to_pst(
            csv_path,
            pst_path,
            folder_filter=args.folder,
            max_compare=args.max or 0,
        )
        inspections.extend(compared)
        logger.info(
            "Compared %d .eml row(s); PST index used %d item(s)",
            n_eml,
            n_pst,
        )
    else:
        logger.info("Scanning all mail in PST: %s", pst_path)
        scanned, n = scan_pst(
            pst_path,
            folder_filter=args.folder,
            max_items=args.max or 0,
        )
        inspections.extend(scanned)
        logger.info("Scanned %d mail item(s)", n)

    summary = summarize_inspections(inspections)
    logger.info(
        "Validation summary: %d checked, %d OK, %d with errors, %d warnings only",
        summary["total"],
        summary["ok"],
        summary["with_errors"],
        summary["warnings_only"],
    )
    if summary["issue_counts"]:
        logger.info("Issue breakdown:")
        for code, count in sorted(
            summary["issue_counts"].items(), key=lambda x: (-x[1], x[0])
        ):
            logger.info("  %s: %d", code, count)

    report_path = args.report.strip() or default_report_path(pst_path)
    if args.issues_only:
        write_issues_only_csv(report_path, inspections)
    else:
        write_report_csv(report_path, inspections)
    logger.info("Report written: %s", report_path)

    if summary["with_errors"] > 0:
        logger.error(
            "FAILED: %d message(s) have errors — fix export or re-import affected mail",
            summary["with_errors"],
        )
        return 1
    if summary["warnings_only"] > 0:
        logger.warning(
            "PASSED with warnings: %d message(s)",
            summary["warnings_only"],
        )
        return 0
    logger.info("PASSED: no issues detected")
    return 0


if __name__ == "__main__":
    sys.exit(main())
