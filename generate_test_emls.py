"""Generate sample .eml files for Mail Exporter conversion and dedup testing."""
from __future__ import annotations

import os
import random
import shutil
import uuid
from datetime import datetime, timedelta
from email.message import EmailMessage
from email.utils import format_datetime, make_msgid

ROOT = os.path.dirname(os.path.abspath(__file__))
_DEFAULT_TEST = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "MailExporter_test")
OUT = os.environ.get("EML2PST_TEST_MAIL", os.path.join(_DEFAULT_TEST, "mail"))
SMALL_OUT = os.environ.get("EML2PST_TEST_QUICK", os.path.join(_DEFAULT_TEST, "quick"))

SUBJECTS = [
    "Quarterly budget review",
    "Meeting notes from Monday",
    "Invoice attached",
    "Welcome to the team",
    "Server maintenance window",
    "Re: Project timeline update",
    "Holiday schedule 2026",
    "Password reset confirmation",
    "FW: Customer feedback summary",
    "Action required: timesheet",
    "Lunch order for Friday",
    "Backup completed successfully",
    "New policy document",
    "Reminder: conference call at 3pm",
    "Your order has shipped",
    "Weekly status report",
    "Contract renewal notice",
    "Training session invite",
    "Expense report due",
    "System alert: disk usage",
]

SENDERS = [
    ("Alice Morgan", "alice.morgan@example.com"),
    ("Bob Chen", "bob.chen@contoso.com"),
    ("Carol Weiss", "carol.weiss@fabrikam.com"),
    ("IT Helpdesk", "helpdesk@company.local"),
    ("Notifications", "noreply@mailer.example.com"),
    ("David Park", "david.park@example.com"),
    ("Eva Novak", "eva.novak@contoso.com"),
    ("Finance Team", "finance@company.local"),
]

FOLDERS = ("Inbox", "Sent", "Archive", "Drafts")

BODIES = [
    "Hi,\n\nPlease review the attached notes and reply by end of week.\n\nThanks,\n",
    "Hello team,\n\nQuick update: we moved the deadline to next Friday.\n\nBest regards,\n",
    "Dear colleague,\n\nAutomated test message for Mail Exporter. Ref: {ref}\n\nRegards,\n",
    "Good morning,\n\nCan you confirm receipt? Random ref: {ref}\n\nCheers,\n",
    "All,\n\nReminder to complete the survey before COB today.\n\nThank you,\n",
]


def _build_message(
    index: int,
    *,
    body_extra: str = "",
    attachment: tuple[str, bytes] | None = None,
    base_date: datetime,
) -> EmailMessage:
    sender_name, sender_addr = random.choice(SENDERS)
    subject = random.choice(SUBJECTS)
    msg_date = base_date + timedelta(
        minutes=index,
        seconds=random.randint(0, 59),
    )
    ref = uuid.uuid4().hex[:8].upper()
    body_text = random.choice(BODIES).format(ref=ref) + sender_name + body_extra

    msg = EmailMessage()
    msg["From"] = f"{sender_name} <{sender_addr}>"
    msg["To"] = "you@company.local"
    msg["Subject"] = f"{subject} [{ref}]"
    msg["Date"] = format_datetime(msg_date)
    msg["Message-ID"] = make_msgid(domain="test.mailexporter.local")
    msg["X-Test-Index"] = str(index)
    msg.set_content(body_text)

    if attachment:
        name, data = attachment
        maintype, subtype = ("application", "octet-stream")
        if name.endswith(".txt"):
            maintype, subtype = "text", "plain"
        elif name.endswith(".pdf"):
            maintype, subtype = "application", "pdf"
        msg.add_attachment(data, maintype=maintype, subtype=subtype, filename=name)

    return msg


def _write_eml(path: str, msg: EmailMessage) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data = msg.as_bytes()
    # Outlook OpenSharedItem expects CRLF; match real Windows Live Mail exports.
    data = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n")
    data = b"\r\n".join(data.split(b"\n"))
    with open(path, "wb") as handle:
        handle.write(data)


def _size_tier(index: int) -> str:
    if index <= 100:
        return "tiny"
    if index <= 250:
        return "small"
    if index <= 350:
        return "medium"
    return "large"


def generate_batch(out_dir: str, *, total: int = 500, duplicate_count: int = 100) -> list[str]:
    """Create unique EML files plus byte-identical duplicates with different names."""
    if os.path.isdir(out_dir):
        for name in os.listdir(out_dir):
            path = os.path.join(out_dir, name)
            if os.path.isdir(path):
                shutil.rmtree(path, ignore_errors=True)
            else:
                try:
                    os.remove(path)
                except OSError:
                    pass
    os.makedirs(out_dir, exist_ok=True)

    base_date = datetime(2024, 1, 10, 8, 0, 0)
    unique_count = total - duplicate_count
    written: list[str] = []

    for i in range(1, unique_count + 1):
        tier = _size_tier(i)
        folder = FOLDERS[i % len(FOLDERS)]
        body_extra = ""
        attachment = None

        if tier == "tiny":
            body_extra = ""
        elif tier == "small":
            body_extra = "\n" + ("Line of detail. " * 120)
        elif tier == "medium":
            body_extra = "\n" + ("Medium body padding. " * 400)
            att_size = random.randint(10_000, 50_000)
            attachment = (
                f"doc_{i:04d}.bin",
                os.urandom(att_size),
            )
        else:
            body_extra = "\n" + ("Large message body block. " * 2000)
            att_size = random.randint(100_000, 400_000)
            attachment = (
                f"large_{i:04d}.dat",
                os.urandom(att_size),
            )

        msg = _build_message(
            i,
            body_extra=body_extra,
            attachment=attachment,
            base_date=base_date,
        )
        ref = uuid.uuid4().hex[:6]
        fname = f"{tier}_{i:04d}_{ref}.eml"
        path = os.path.join(out_dir, folder, fname)
        _write_eml(path, msg)
        written.append(path)
        if i % 100 == 0:
            print(f"  unique {i}/{unique_count}")

    dup_sources = random.sample(written, min(duplicate_count, len(written)))
    for j, src in enumerate(dup_sources, start=1):
        folder = os.path.basename(os.path.dirname(src))
        dup_name = f"dup_{j:04d}_of_{os.path.basename(src)}"
        dup_path = os.path.join(out_dir, folder, dup_name)
        shutil.copy2(src, dup_path)
        written.append(dup_path)

    # Pad duplicate count if sample was smaller than duplicate_count
    while len(written) < total and written:
        src = random.choice(written[:unique_count])
        folder = os.path.basename(os.path.dirname(src))
        dup_name = f"dup_extra_{len(written):04d}_{os.path.basename(src)}"
        dup_path = os.path.join(out_dir, folder, dup_name)
        shutil.copy2(src, dup_path)
        written.append(dup_path)

    return written


def generate_small_set(out_dir: str) -> None:
    """Keep a small 15-file set in test_emls/ for quick manual checks."""
    os.makedirs(out_dir, exist_ok=True)
    base_date = datetime(2024, 3, 15, 9, 0, 0)
    for i in range(15):
        msg = _build_message(i + 1, base_date=base_date)
        if (i + 1) % 5 == 0:
            msg.add_attachment(
                f"Test attachment #{i + 1}\n".encode("utf-8"),
                maintype="text",
                subtype="plain",
                filename=f"attachment_{i + 1}.txt",
            )
        path = os.path.join(out_dir, f"test_{i + 1:02d}.eml")
        _write_eml(path, msg)


def main() -> None:
    print(f"Generating 500 test EMLs (with duplicates) in:\n  {OUT}\n")
    paths = generate_batch(OUT, total=500, duplicate_count=100)
    unique = len({os.path.getsize(p): p for p in paths})  # rough; report sizes
    total_bytes = sum(os.path.getsize(p) for p in paths)
    print(f"\nDone: {len(paths)} files, ~{total_bytes / (1024 * 1024):.1f} MB total")
    print(f"  Unique sources: {500 - 100}, duplicates: 100")
    print(f"\nAlso refreshing 15 quick-test files in:\n  {SMALL_OUT}\n")
    generate_small_set(SMALL_OUT)
    print("Small set ready.")


if __name__ == "__main__":
    main()
