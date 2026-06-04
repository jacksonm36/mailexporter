"""
Mail quality checks for Mail Exporter (PST items and source .eml files).

Used by export_checker.py and bug_check.py. No tkinter / Outlook dependency.
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from email import policy
from email.header import decode_header
from email.parser import BytesParser

from mmap_processor import parse_message_from_path
from email.utils import parseaddr

# Issue codes (stable for reports and CI)
ISSUE_BLANK_SENDER = "blank_sender"
ISSUE_LITERAL_NONE_SENDER = "literal_none_sender"
ISSUE_BLANK_SUBJECT = "blank_subject"
ISSUE_LITERAL_NONE_SUBJECT = "literal_none_subject"
ISSUE_WEAK_SUBJECT = "weak_subject"
ISSUE_BLANK_BODY = "blank_body"
ISSUE_CORRUPTED_BODY = "corrupted_body"
ISSUE_HTML_AS_PLAIN = "html_as_plain"
ISSUE_MISSING_ATTACHMENT = "missing_attachment"
ISSUE_UNEXPECTED_ATTACHMENT = "unexpected_attachment"
ISSUE_EML_PARSE_FAILED = "eml_parse_failed"
ISSUE_EML_MISSING_FILE = "eml_missing_file"
ISSUE_SUBJECT_MISMATCH = "subject_mismatch"
ISSUE_SENDER_MISMATCH = "sender_mismatch"
ISSUE_NOT_FOUND_IN_PST = "not_found_in_pst"
ISSUE_AMBIGUOUS_IN_PST = "ambiguous_in_pst"
ISSUE_WRONG_FOLDER = "wrong_folder"
ISSUE_WRONG_PST_STORE = "wrong_pst_store"

SEVERITY_ERROR = "error"
SEVERITY_WARNING = "warning"

_HTML_TAG_FRAGMENT_RE = re.compile(r"</?[a-zA-Z][^>]{0,240}>", re.I)
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
_BINARY_BODY_SIGNATURES = (
    b"%PDF-",
    b"\x89PNG\r\n\x1a\n",
    b"PK\x03\x04",
)

NATIVE_IMPORT_SNIFF_BYTES = 262144


@dataclass
class MailIssue:
    code: str
    severity: str
    detail: str = ""


@dataclass
class MailInspection:
    """Result of checking one message (PST item fields and/or parsed .eml)."""

    source: str = ""
    subject: str = ""
    sender: str = ""
    issues: list[MailIssue] = field(default_factory=list)

    @property
    def ok(self) -> bool:
        return not any(i.severity == SEVERITY_ERROR for i in self.issues)

    def add(self, code: str, severity: str, detail: str = "") -> None:
        self.issues.append(MailIssue(code=code, severity=severity, detail=detail))


def outlook_text_is_blank(value) -> bool:
    if value is None:
        return True
    text = str(value).strip()
    return not text or text.lower() == "none"


def bytes_look_binary(data: bytes, *, sample: int = 8192) -> bool:
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
    if not text:
        return False
    sample = text[:8192]
    if sample.lstrip().startswith("%PDF-"):
        return True
    try:
        encoded = sample.encode("utf-8", errors="ignore")
    except Exception:
        return True
    return bytes_look_binary(encoded)


_KNOWN_HTML_TAGS = frozenset(
    {
        "html",
        "body",
        "head",
        "meta",
        "div",
        "span",
        "p",
        "br",
        "table",
        "tr",
        "td",
        "th",
        "a",
        "img",
        "font",
        "style",
        "script",
        "link",
        "ul",
        "ol",
        "li",
        "h1",
        "h2",
        "h3",
        "h4",
        "h5",
        "h6",
        "center",
        "blockquote",
        "pre",
        "hr",
        "strong",
        "em",
        "b",
        "i",
        "u",
    }
)


def _html_tag_name_from_match(tag_text: str) -> str | None:
    inner = tag_text.strip("<>/ \t").split(None, 1)[0].lower()
    if not inner or inner in ("http", "https", "mailto", "ftp"):
        return None
    return inner.split(":", 1)[0]


def body_looks_like_html(body: str) -> bool:
    if not body or text_looks_binary(body):
        return False
    sample = body.lstrip()[:8192].lower()
    if sample.startswith("<!doctype") or "<html" in sample or "<body" in sample:
        return True
    for match in _HTML_TAG_FRAGMENT_RE.finditer(body[:8192]):
        name = _html_tag_name_from_match(match.group(0))
        if name in _KNOWN_HTML_TAGS:
            return True
    return False


def html_stored_as_plain_text(plain_body: str, html_body: str) -> bool:
    """True when HTML tags appear in plain Body but HTMLBody is empty or identical."""
    plain = (plain_body or "").strip()
    html = (html_body or "").strip()
    if not plain or not body_looks_like_html(plain):
        return False
    if not html:
        return True
    if html == plain:
        return True
    if len(html) < 32 and body_looks_like_html(plain):
        return True
    return False


def sniff_eml_has_attachments(file_path: str) -> bool:
    try:
        with open(file_path, "rb") as handle:
            head = handle.read(
                min(os.path.getsize(file_path), NATIVE_IMPORT_SNIFF_BYTES)
            )
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


def decode_mime_header_field(value) -> str:
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


def resolve_from_header(msg) -> str:
    if msg is None:
        return ""
    for key in (
        "From",
        "Sender",
        "Reply-To",
        "Return-Path",
        "X-Sender",
        "X-Original-From",
    ):
        raw = msg.get(key)
        if not raw:
            continue
        decoded = decode_mime_header_field(raw)
        if decoded:
            display, addr = parseaddr(decoded)
            if addr or display:
                return decoded
    return ""


def _is_attachment_like_part(part) -> bool:
    disp = (part.get_content_disposition() or "").lower()
    if disp == "attachment":
        return True
    if disp == "inline":
        filename = part.get_filename()
        if filename:
            return True
        ctype = (part.get_content_type() or "").lower()
        if ctype.startswith("image/") or ctype in (
            "application/pdf",
            "application/zip",
            "application/octet-stream",
        ):
            return True
    ctype = (part.get_content_type() or "").lower()
    if ctype == "message/rfc822":
        return True
    return False


def _decode_part_as_text(part) -> str | None:
    try:
        content = part.get_content()
    except Exception:
        return None
    if isinstance(content, bytes):
        charset = part.get_content_charset() or "utf-8"
        try:
            return content.decode(charset, errors="replace")
        except (LookupError, UnicodeDecodeError):
            return content.decode("utf-8", errors="replace")
    if isinstance(content, str):
        return content
    return None


def extract_eml_body_and_attachments(msg) -> tuple[str, int, bool]:
    """Return (body text, attachment count, body_is_html)."""
    plain_parts: list[str] = []
    html_parts: list[str] = []
    attachments = 0
    for part in msg.walk():
        if part.get_content_maintype() == "multipart":
            continue
        if _is_attachment_like_part(part):
            attachments += 1
            continue
        ctype = part.get_content_type()
        if ctype not in ("text/plain", "text/html"):
            continue
        text = _decode_part_as_text(part)
        if not text:
            continue
        if ctype == "text/html":
            html_parts.append(text)
        else:
            plain_parts.append(text)
    if html_parts:
        return html_parts[0], attachments, True
    if plain_parts:
        return plain_parts[0], attachments, False
    return "", attachments, False


def parse_eml_summary(file_path: str) -> dict | None:
    """Lightweight .eml parse for validation (no Outlook)."""
    msg = parse_message_from_path(file_path)
    if msg is None:
        return None
    subject = decode_mime_header_field(msg.get("Subject"))
    body, att_count, is_html = extract_eml_body_and_attachments(msg)
    return {
        "file_path": file_path,
        "subject": subject,
        "from": resolve_from_header(msg),
        "body": body,
        "body_is_html": is_html,
        "attachment_count": att_count,
        "sniff_attachments": sniff_eml_has_attachments(file_path),
        "message": msg,
    }


def load_converted_paths_from_csv(csv_path: str) -> list[dict]:
    """Rows with status converted from export_results.csv."""
    import csv as csv_mod

    rows: list[dict] = []
    if not os.path.isfile(csv_path):
        return rows
    try:
        with open(csv_path, newline="", encoding="utf-8") as handle:
            for row in csv_mod.DictReader(handle):
                if (row.get("status") or "").lower() != "converted":
                    continue
                path = (row.get("file_path") or "").strip()
                if path:
                    rows.append(
                        {
                            "file_path": os.path.normpath(os.path.abspath(path)),
                            "target_folder": (row.get("target_folder") or "").strip(),
                        }
                    )
    except OSError:
        pass
    return rows


def _normalize_match_key(subject: str, sender: str) -> tuple[str, str]:
    subj = (subject or "").strip().lower()
    if subj in ("", "none", "(no subject)"):
        subj = ""
    snd = (sender or "").strip().lower()
    if snd in ("", "none"):
        snd = ""
    return subj, snd


def inspect_parsed_eml(summary: dict | None, *, source: str = "") -> MailInspection:
    result = MailInspection(source=source or (summary or {}).get("file_path", ""))
    if not summary:
        result.add(ISSUE_EML_PARSE_FAILED, SEVERITY_ERROR, "Could not parse .eml")
        return result
    result.subject = (summary.get("subject") or "").strip()
    result.sender = (summary.get("from") or "").strip()
    if not result.sender:
        result.add(ISSUE_BLANK_SENDER, SEVERITY_WARNING, "No From/Sender in .eml")
    body = (summary.get("body") or "").strip()
    if text_looks_binary(body):
        result.add(ISSUE_CORRUPTED_BODY, SEVERITY_ERROR, "Body looks binary in .eml")
    elif not body and not summary.get("attachment_count") and not summary.get(
        "sniff_attachments"
    ):
        result.add(ISSUE_BLANK_BODY, SEVERITY_WARNING, "No body or attachments in .eml")
    if not result.subject:
        result.add(ISSUE_BLANK_SUBJECT, SEVERITY_WARNING, "Empty Subject in .eml")
    return result


def inspect_outlook_fields(
    *,
    subject: str = "",
    sender_name: str = "",
    sender_email: str = "",
    body: str = "",
    html_body: str = "",
    attachment_count: int = 0,
    source: str = "",
    expected_sniff_attachments: bool = False,
) -> MailInspection:
    """Check Outlook MailItem fields (already read via COM)."""
    result = MailInspection(source=source, subject=(subject or "").strip())
    display_sender = (sender_name or "").strip()
    if outlook_text_is_blank(display_sender):
        if (sender_email or "").strip():
            display_sender = (sender_email or "").strip()
    result.sender = display_sender

    if outlook_text_is_blank(sender_name):
        if str(sender_name or "").strip().lower() == "none":
            result.add(
                ISSUE_LITERAL_NONE_SENDER,
                SEVERITY_ERROR,
                "SenderName is literal 'None'",
            )
        else:
            result.add(
                ISSUE_BLANK_SENDER,
                SEVERITY_ERROR,
                "Missing sender (blank list row / None group)",
            )

    if outlook_text_is_blank(subject):
        if str(subject or "").strip().lower() == "none":
            result.add(
                ISSUE_LITERAL_NONE_SUBJECT,
                SEVERITY_ERROR,
                "Subject is literal 'None'",
            )
        else:
            result.add(ISSUE_BLANK_SUBJECT, SEVERITY_ERROR, "Empty subject")
    elif subject.strip() in ("(No Subject)",):
        result.add(ISSUE_WEAK_SUBJECT, SEVERITY_WARNING, "Generic (No Subject)")

    plain = (body or "").strip()
    html = (html_body or "").strip()
    if text_looks_binary(plain) or text_looks_binary(html):
        result.add(ISSUE_CORRUPTED_BODY, SEVERITY_ERROR, "Body contains binary data")
    elif not plain and not html and attachment_count <= 0:
        result.add(
            ISSUE_BLANK_BODY,
            SEVERITY_ERROR,
            "No body text and no attachments",
        )
    elif html_stored_as_plain_text(plain, html):
        result.add(
            ISSUE_HTML_AS_PLAIN,
            SEVERITY_ERROR,
            "HTML tags visible in plain body (not rendered)",
        )

    if expected_sniff_attachments and attachment_count <= 0:
        result.add(
            ISSUE_MISSING_ATTACHMENT,
            SEVERITY_ERROR,
            "Source .eml likely has files but PST item has no attachments",
        )
    return result


def compare_eml_to_outlook(
    eml_summary: dict | None,
    outlook_inspection: MailInspection,
) -> list[MailIssue]:
    """Extra issues when a converted .eml is matched to a PST item."""
    extra: list[MailIssue] = []
    if not eml_summary:
        return extra
    eml_subj = (eml_summary.get("subject") or "").strip()
    if eml_subj and eml_subj not in ("(No Subject)",):
        out_subj = (outlook_inspection.subject or "").strip()
        if outlook_text_is_blank(out_subj) or out_subj.lower() == "none":
            extra.append(
                MailIssue(
                    ISSUE_SUBJECT_MISMATCH,
                    SEVERITY_ERROR,
                    f".eml subject {eml_subj[:80]!r} not on PST item",
                )
            )
        elif _normalize_match_key(eml_subj, "")[0] != _normalize_match_key(out_subj, "")[0]:
            if eml_subj.lower() not in out_subj.lower() and out_subj.lower() not in eml_subj.lower():
                extra.append(
                    MailIssue(
                        ISSUE_SUBJECT_MISMATCH,
                        SEVERITY_WARNING,
                        f".eml {eml_subj[:60]!r} vs PST {out_subj[:60]!r}",
                    )
                )
    eml_from = (eml_summary.get("from") or "").strip()
    if eml_from:
        _, eml_addr = parseaddr(eml_from)
        out_sender = (outlook_inspection.sender or "").strip()
        if outlook_text_is_blank(out_sender):
            extra.append(
                MailIssue(
                    ISSUE_SENDER_MISMATCH,
                    SEVERITY_ERROR,
                    f".eml From present; PST sender blank ({eml_addr or eml_from[:60]})",
                )
            )
    if eml_summary.get("sniff_attachments") and not any(
        i.code == ISSUE_MISSING_ATTACHMENT for i in outlook_inspection.issues
    ):
        pass  # inspect_outlook_fields already flags with expected_sniff_attachments
    return extra


def inspection_issue_codes(inspection: MailInspection) -> list[str]:
    return [i.code for i in inspection.issues]


def inspection_error_codes(inspection: MailInspection) -> list[str]:
    return [i.code for i in inspection.issues if i.severity == SEVERITY_ERROR]


def has_import_errors(inspection: MailInspection) -> bool:
    return any(i.severity == SEVERITY_ERROR for i in inspection.issues)


def eml_parse_to_summary(email_data: dict | None, file_path: str = "") -> dict | None:
    """Map Mail Exporter parse_eml() dict to validation summary."""
    if not email_data:
        return None
    body = (email_data.get("body") or "").strip()
    atts = email_data.get("attachments") or []
    path = file_path or ""
    return {
        "file_path": path,
        "subject": (email_data.get("subject") or "").strip(),
        "from": (email_data.get("from") or "").strip(),
        "body": body,
        "body_is_html": body_looks_like_html(body),
        "attachment_count": len(atts),
        "sniff_attachments": sniff_eml_has_attachments(path) if path else bool(atts),
    }


def read_outlook_mail_snapshot(mail_item) -> dict:
    """Read MailItem properties via COM for mid-import validation."""

    def _safe(prop: str) -> str:
        try:
            value = getattr(mail_item, prop, None)
            if value is None:
                return ""
            return str(value).strip()
        except Exception:
            return ""

    att_count = 0
    try:
        att_count = int(mail_item.Attachments.Count)
    except Exception:
        pass
    return {
        "subject": _safe("Subject"),
        "sender_name": _safe("SenderName"),
        "sender_email": _safe("SenderEmailAddress"),
        "body": _safe("Body"),
        "html_body": _safe("HTMLBody"),
        "attachment_count": att_count,
    }


def validate_import_against_eml(
    snapshot: dict,
    eml_summary: dict | None,
    *,
    source: str = "",
) -> MailInspection:
    """Validate a just-imported PST item against the source .eml."""
    path = (eml_summary or {}).get("file_path") or source
    sniff = bool((eml_summary or {}).get("sniff_attachments"))
    if not sniff and path:
        sniff = sniff_eml_has_attachments(path)
    att_n = int((eml_summary or {}).get("attachment_count") or 0)
    inspection = inspect_outlook_fields(
        subject=snapshot.get("subject", ""),
        sender_name=snapshot.get("sender_name", ""),
        sender_email=snapshot.get("sender_email", ""),
        body=snapshot.get("body", ""),
        html_body=snapshot.get("html_body", ""),
        attachment_count=int(snapshot.get("attachment_count") or att_n),
        source=source or path,
        expected_sniff_attachments=sniff,
    )
    if eml_summary:
        for issue in compare_eml_to_outlook(eml_summary, inspection):
            inspection.issues.append(issue)
    return inspection


def append_folder_placement_issues(
    inspection: MailInspection,
    *,
    in_correct_folder: bool,
    actual_folder_path: str,
    expected_folder_path: str,
    in_correct_store: bool = True,
    actual_store: str = "",
    expected_store: str = "",
) -> None:
    """Record PST store / subfolder placement problems on an inspection."""
    if not in_correct_store:
        inspection.add(
            ISSUE_WRONG_PST_STORE,
            SEVERITY_ERROR,
            f"message in {actual_store or '?'}, expected PST {expected_store or '?'}",
        )
    if not in_correct_folder:
        inspection.add(
            ISSUE_WRONG_FOLDER,
            SEVERITY_ERROR,
            f"in folder {actual_folder_path or '?'}, expected {expected_folder_path or '?'}",
        )


def format_inspection_errors(inspection: MailInspection) -> str:
    parts: list[str] = []
    for issue in inspection.issues:
        if issue.severity != SEVERITY_ERROR:
            continue
        if issue.detail:
            parts.append(f"{issue.code}: {issue.detail}")
        else:
            parts.append(issue.code)
    return "; ".join(parts)
