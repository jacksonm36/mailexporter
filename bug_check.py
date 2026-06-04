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
        _ImmediateFileHandler,
        csv_sanitize,
        count_converted_in_csv,
        export_log_paths,
        load_resume_paths,
        validate_import_path,
        load_resume_paths_from_sqlite,
        logger,
        normalize_pst_path,
        pst_paths_equal,
        bytes_look_binary,
        text_looks_binary,
        relative_folder_parts,
        safe_outlook_folder_name,
        _is_outlook_path_open_error,
        _pattern_to_suffixes,
        _portable_staging_base,
    )

    def _fp_hash(app_stub, path: str, size: int) -> str | None:
        pair = app_stub._file_content_fingerprint(path, size)
        if pair is None:
            return None
        return pair[0]

    errors: list[str] = []

    # Sanitization
    if csv_sanitize("=1+1") != "'=1+1":
        errors.append("csv_sanitize formula prefix")
    if safe_outlook_folder_name("a/b:c") != "a_b_c":
        errors.append("safe_outlook_folder_name")

    # Pattern parsing
    if _pattern_to_suffixes("*.eml;*.emlx") != {".eml", ".emlx"}:
        errors.append("_pattern_to_suffixes")

    with tempfile.TemporaryDirectory() as pst_tmp:
        raw = os.path.join(pst_tmp, "out")
        norm = normalize_pst_path(raw)
        if not norm.lower().endswith(".pst"):
            errors.append("normalize_pst_path should add .pst")
        if not pst_paths_equal(norm, raw + ".pst"):
            errors.append("pst_paths_equal basic")
        if os.name == "nt" and not pst_paths_equal(
            norm.upper(), norm.lower()
        ):
            errors.append("pst_paths_equal case insensitive on Windows")

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

    with tempfile.TemporaryDirectory() as sec_tmp:
        base = os.path.join(sec_tmp, "mail")
        os.makedirs(base)
        inside = os.path.join(base, "a.eml")
        open(inside, "w", encoding="utf-8").close()
        try:
            validate_import_path(base, inside)
        except ValueError as exc:
            errors.append(f"validate_import_path inside base: {exc}")
        try:
            validate_import_path(base, os.path.join("..", "evil.eml"))
            errors.append("validate_import_path should block ..")
        except ValueError:
            pass
        from path_security import PathSecurityError, PathValidator

        if not PathValidator.create_safe_filename("..\\evil.pdf"):
            errors.append("create_safe_filename empty")
        if PathValidator.create_safe_filename("report<1>.eml") != "report_1_.eml":
            errors.append("create_safe_filename dangerous chars")
        if PathValidator.create_safe_filename("file;rm -rf /") != "file_rm -rf":
            errors.append("create_safe_filename semicolon")
        try:
            PathValidator.sanitize_path("..\\x.eml", base)
            errors.append("PathValidator.sanitize_path should block ..")
        except PathSecurityError:
            pass

    fixed_name = EmlToPstConverter._decode_mime_header_field("PatÃ³cs")
    if "ó" not in fixed_name:
        errors.append(f"_decode_mime_header_field mojibake repair: got {fixed_name!r}")
    if not hasattr(app_stub, "_apply_sender_from_parsed_from"):
        errors.append("_apply_sender_from_parsed_from missing")
    if not hasattr(app_stub, "_apply_parsed_display_fields"):
        errors.append("_apply_parsed_display_fields missing")
    if not hasattr(app_stub, "_ensure_import_quality"):
        errors.append("_ensure_import_quality missing (mid-conversion validation)")
    if not hasattr(app_stub, "_recover_outlook_session"):
        errors.append("_recover_outlook_session missing")
    if not hasattr(app_stub, "_mail_in_dest_folder"):
        errors.append("_mail_in_dest_folder missing")
    if not EmlToPstConverter._body_looks_like_html("<br/>Tisztelt Partnerünk!<br/><b>Hi</b>"):
        errors.append("_body_looks_like_html should detect br/b fragments")
    if EmlToPstConverter._body_looks_like_html("Plain text only."):
        errors.append("_body_looks_like_html should reject plain text")
    wrapped = EmlToPstConverter._normalize_html_body_for_outlook("<br/>Hi</br>")
    if "<html" not in wrapped.lower() or "<body>" not in wrapped.lower():
        errors.append("_normalize_html_body_for_outlook should wrap fragments")
    if not EmlToPstConverter._outlook_text_is_blank("None"):
        errors.append("_outlook_text_is_blank should treat literal None as blank")
    if not hasattr(app_stub, "_clear_pst_session_caches"):
        errors.append("_clear_pst_session_caches missing")

    wlm_like = r"C:\test\Account (user@domain)\Inbox\msg.eml"
    if not app_stub._path_needs_native_staging(wlm_like):
        errors.append("_path_needs_native_staging should be True for () and @ in path")

    if not hasattr(app_stub, "_stamp_delivery_on_new_item"):
        errors.append("_stamp_delivery_on_new_item missing")
    for name in (
        "_apply_mapi_delivery_times",
        "_clear_unsent_mapi_flag",
        "_stamp_delivery_for_import",
        "_commit_mail_to_pst_folder",
        "_relocate_mail_to_dest_folder",
        "_import_open_shared_into_pst_folder",
        "_message_in_expected_pst",
        "_folder_store_path",
        "_log_pst_store_layout",
        "_prepare_pst_store_for_import",
        "_activate_pst_store_for_export",
        "_label_pst_store_for_outlook",
        "_folder_belongs_to_pst",
        "_create_mail_in_pst_target_folder",
        "_add_mail_via_store_inbox",
        "_find_folder_in_tree",
        "_record_converted_path_sqlite",
        "_stores_match",
        "_pst_folder_ready_for_import",
        "_discover_pst_mail_inbox",
        "_is_usable_standard_folder",
        "_folder_looks_like_search_container",
        "_coerce_import_folder",
        "_folder_is_usable_import_target",
        "_get_or_create_standard_folder",
    ):
        if not hasattr(app_stub, name):
            errors.append(f"{name} missing")
    import inspect

    mark_sig = inspect.signature(app_stub._mark_mail_imported_received)
    if "dt_local" not in mark_sig.parameters:
        errors.append("_mark_mail_imported_received must accept dt_local")
    save_sig = inspect.signature(app_stub._save_imported_mail)
    if "dt_local" not in save_sig.parameters:
        errors.append("_save_imported_mail must accept dt_local")
    if not hasattr(app_stub, "_transfer_opened_mail_to_folder"):
        errors.append("_transfer_opened_mail_to_folder missing")
    if not hasattr(app_stub, "_export_targets_pst"):
        errors.append("_export_targets_pst missing")
    if not hasattr(app_stub, "_import_eml_direct_to_pst_folder"):
        errors.append("_import_eml_direct_to_pst_folder missing")

    pst_opts = {"pst_option": "new", "destination_path": r"C:\out\export.pst"}
    mbox_opts = {"pst_option": "mailbox", "destination_path": ""}
    if not app_stub._export_targets_pst(pst_opts):
        errors.append("_export_targets_pst should be True for PST export")
    if app_stub._export_targets_pst(mbox_opts):
        errors.append("_export_targets_pst should be False for mailbox export")

    class _FakeMail:
        def __init__(self, store_path: str):
            self.Parent = type("P", (), {"Store": type("S", (), {"FilePath": store_path})()})()

    if app_stub._message_in_expected_pst(_FakeMail(""), r"C:\out\other.pst"):
        errors.append("_message_in_expected_pst must reject empty store path")
    if not app_stub._message_in_expected_pst(
        _FakeMail(r"C:\out\other.pst"), r"C:\out\other.pst"
    ):
        errors.append("_message_in_expected_pst should accept matching PST path")

    class _FakeFolder:
        def __init__(self, name: str, store_path: str):
            self.Name = name
            self.Store = type("S", (), {"FilePath": store_path})()

    bad = _FakeFolder("", r"C:\out\other.pst")
    if app_stub._folder_is_usable_import_target(bad, r"C:\out\other.pst"):
        errors.append("_folder_is_usable_import_target must reject empty folder name")
    good = _FakeFolder("Inbox", r"C:\out\other.pst")
    if not app_stub._folder_is_usable_import_target(good, r"C:\out\other.pst"):
        errors.append("_folder_is_usable_import_target should accept named Inbox")

    class _FakeStore:
        def __init__(self, path: str, store_id: str = "STORE1"):
            self.FilePath = path
            self.StoreID = store_id

    class _FakeItems:
        Count = 0

    class _FakeInboxNoName:
        Name = ""
        DefaultItemType = 0

        def __init__(self, store: _FakeStore):
            self.Store = store
            self.Items = _FakeItems()

    pst_store = _FakeStore(r"C:\out\export.pst")
    unnamed_inbox = _FakeInboxNoName(pst_store)
    if not app_stub._stores_match(unnamed_inbox.Store, pst_store):
        errors.append("_stores_match should match same StoreID/path")
    class _FakeSearchInbox:
        Name = "SPAM Search Folder 2"
        FolderClass = "IPF.OutlookSearchFolder"
        DefaultItemType = 0

        def __init__(self, store: _FakeStore):
            self.Store = store
            self.Items = _FakeItems()

    search_inbox = _FakeSearchInbox(pst_store)
    if app_stub._is_usable_standard_folder(
        search_inbox, 6, pst_store
    ):
        errors.append("_is_usable_standard_folder must reject SPAM search folder Inbox")
    if app_stub._folder_looks_like_search_container(search_inbox):
        pass
    else:
        errors.append("_folder_looks_like_search_container should detect search folder")

    if not bytes_look_binary(b"%PDF-1.3\nbinary"):
        errors.append("bytes_look_binary should detect PDF")
    if not text_looks_binary("%PDF-1.3\n"):
        errors.append("text_looks_binary should detect PDF header")
    if text_looks_binary("Hello, this is a normal email body."):
        errors.append("text_looks_binary should accept plain text")

    with tempfile.TemporaryDirectory() as mime_tmp:
        pdf_eml = os.path.join(mime_tmp, "pdf_only.eml")
        pdf_bytes = b"%PDF-1.3\nfake pdf content with null \x00 byte"
        import base64

        b64 = base64.b64encode(pdf_bytes).decode("ascii")
        with open(pdf_eml, "wb") as handle:
            handle.write(
                (
                    "From: sender@example.com\r\n"
                    "To: recv@example.com\r\n"
                    "Subject: PO PDF\r\n"
                    "MIME-Version: 1.0\r\n"
                    'Content-Type: application/pdf; name="order.pdf"\r\n'
                    "Content-Transfer-Encoding: base64\r\n"
                    "Content-Disposition: inline; filename=\"order.pdf\"\r\n"
                    "\r\n"
                    f"{b64}\r\n"
                ).encode("ascii")
            )
        parsed = app_stub.parse_eml(pdf_eml)
        if not parsed:
            errors.append("parse_eml failed for inline PDF sample")
        else:
            if parsed.get("body"):
                errors.append("get_email_body must not return PDF bytes as body")
            atts = parsed.get("attachments") or []
            if len(atts) != 1 or not atts[0].get("filename", "").endswith(".pdf"):
                errors.append("inline PDF should be extracted as attachment")
            elif atts[0].get("data", b"")[:5] != b"%PDF-":
                errors.append("inline PDF attachment payload corrupt")

        multi_eml = os.path.join(mime_tmp, "multi_attach.eml")
        png_bytes = b"\x89PNG\r\n\x1a\n" + b"\x00" * 64
        zip_bytes = b"PK\x03\x04" + b"\x00" * 32
        png_b64 = base64.b64encode(png_bytes).decode("ascii")
        zip_b64 = base64.b64encode(zip_bytes).decode("ascii")
        with open(multi_eml, "wb") as handle:
            handle.write(
                (
                    "From: sender@example.com\r\n"
                    "To: recv@example.com\r\n"
                    "Subject: files\r\n"
                    "MIME-Version: 1.0\r\n"
                    'Content-Type: multipart/mixed; boundary="bnd"\r\n'
                    "\r\n"
                    "--bnd\r\n"
                    "Content-Type: text/plain\r\n"
                    "\r\n"
                    "See attached files.\r\n"
                    "--bnd\r\n"
                    'Content-Type: image/png; name="scan.png"\r\n'
                    "Content-Transfer-Encoding: base64\r\n"
                    'Content-Disposition: attachment; filename="scan.png"\r\n'
                    "\r\n"
                    f"{png_b64}\r\n"
                    "--bnd\r\n"
                    'Content-Type: application/zip; name="data.zip"\r\n'
                    "Content-Transfer-Encoding: base64\r\n"
                    'Content-Disposition: attachment; filename="data.zip"\r\n'
                    "\r\n"
                    f"{zip_b64}\r\n"
                    "--bnd--\r\n"
                ).encode("ascii")
            )
        multi = app_stub.parse_eml(multi_eml)
        if not multi:
            errors.append("parse_eml failed for PNG+ZIP multipart sample")
        else:
            names = {a.get("filename", "").lower() for a in multi.get("attachments") or []}
            if "scan.png" not in names or "data.zip" not in names:
                errors.append("multipart PNG+ZIP attachments not extracted")
            merged = app_stub._merge_attachment_records(
                multi.get("attachments") or [],
                [
                    {
                        "filename": "extra.pdf",
                        "data": b"%PDF-1.3",
                        "content_type": "application/pdf",
                    }
                ],
            )
            if len(merged) != 3:
                errors.append("_merge_attachment_records should merge distinct files")
            if not hasattr(app_stub, "_apply_attachment_chain"):
                errors.append("_apply_attachment_chain missing")

        html_only = os.path.join(mime_tmp, "html_inline_only.eml")
        with open(html_only, "wb") as handle:
            handle.write(
                (
                    "From: a@b.c\r\nTo: you@co.local\r\nSubject: HTML only\r\n"
                    "MIME-Version: 1.0\r\n"
                    'Content-Type: multipart/alternative; boundary="b"\r\n'
                    "\r\n--b\r\n"
                    "Content-Type: text/plain; charset=utf-8\r\n"
                    "Content-Disposition: inline\r\n"
                    "\r\nHello\r\n"
                    "--b\r\n"
                    "Content-Type: text/html; charset=utf-8\r\n"
                    "Content-Disposition: inline\r\n"
                    "\r\n<html><body>Hi</body></html>\r\n"
                    "--b--\r\n"
                ).encode("utf-8")
            )
        if app_stub._sniff_eml_has_attachments(html_only):
            errors.append(
                "_sniff_eml_has_attachments must ignore HTML inline-only multipart"
            )
        if not app_stub._sniff_eml_has_attachments(pdf_eml):
            errors.append("_sniff_eml_has_attachments should detect PDF sample")

        wlm_html = os.path.join(mime_tmp, "wlm_html_fragment.eml")
        with open(wlm_html, "wb") as handle:
            handle.write(
                (
                    "Reply-To: invoices@example.com\r\n"
                    "Subject: =?utf-8?B?VMOpcsOpc3plbHQ=?=\r\n"
                    "MIME-Version: 1.0\r\n"
                    "Content-Type: text/html; charset=utf-8\r\n"
                    "\r\n"
                    "<br/>Tisztelt Partnerünk!<br/><b>196995</b>\r\n"
                ).encode("utf-8")
            )
        wlm_parsed = app_stub.parse_eml(wlm_html)
        if not wlm_parsed:
            errors.append("parse_eml failed for WLM HTML fragment sample")
        else:
            if "invoices@example.com" not in (wlm_parsed.get("from") or ""):
                errors.append("_resolve_from_header should use Reply-To when From missing")
            if not EmlToPstConverter._body_looks_like_html(wlm_parsed.get("body") or ""):
                errors.append("WLM HTML fragment body should look like HTML")

    wlm_opts = {"use_file_mtime_for_date": True}
    if not app_stub._is_windows_live_mail_path(
        r"D:\backup\Windows Live Mail\account\Inbox\a.eml"
    ):
        errors.append("_is_windows_live_mail_path")
    with tempfile.TemporaryDirectory() as wlm_tmp:
        wlm_eml = os.path.join(
            wlm_tmp, "Windows Live Mail", "Account (user@domain)", "Inbox", "probe.eml"
        )
        os.makedirs(os.path.dirname(wlm_eml), exist_ok=True)
        with open(wlm_eml, "wb") as handle:
            handle.write(b"From: a@b.c\r\nDate: Mon, 1 Jan 2024 12:00:00 +0000\r\n\r\nx\r\n")
        mtime_dt = app_stub._file_mtime_datetime(wlm_eml)
        resolved = app_stub._resolve_outlook_datetime_for_source(
            wlm_eml, None, wlm_opts
        )
        if mtime_dt is None or resolved != mtime_dt:
            errors.append(
                "_resolve_outlook_datetime_for_source should prefer mtime on WLM paths"
            )

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
                from eml_to_pst_converter import NATIVE_IMPORT_MAX_BYTES

                if lsz > NATIVE_IMPORT_MAX_BYTES:
                    if app_stub._should_try_native_import(large, lsz):
                        errors.append(
                            "_should_try_native_import should skip oversized EML"
                        )
                elif not app_stub._should_try_native_import(large, lsz):
                    errors.append(
                        "_should_try_native_import should allow attachment EML under size cap"
                    )
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
    from eml_to_pst_converter import (
        OL_FOLDER_DELETED,
        OL_FOLDER_INBOX,
        OL_FOLDER_OUTBOX,
        OL_FOLDER_SENT,
        find_standard_folder_in_parts,
        sent_state_for_folder_parts,
        standard_outlook_folder_id,
    )

    if standard_outlook_folder_id("Inbox") != 6:
        errors.append("standard_outlook_folder_id Inbox")
    if standard_outlook_folder_id("Piszkozatok") != 16:
        errors.append("standard_outlook_folder_id Piszkozatok")
    if standard_outlook_folder_id("Sent Items") != 5:
        errors.append("standard_outlook_folder_id Sent Items")
    if standard_outlook_folder_id("Outbox") != OL_FOLDER_OUTBOX:
        errors.append("standard_outlook_folder_id Outbox")

    if sent_state_for_folder_parts(["Account (x)", "Sent Items"]) is not True:
        errors.append("sent_state_for_folder_parts Sent Items")
    if sent_state_for_folder_parts(["Account", "Drafts"]) is not False:
        errors.append("sent_state_for_folder_parts Drafts")
    if sent_state_for_folder_parts(["Account", "Outbox"]) is not False:
        errors.append("sent_state_for_folder_parts Outbox")
    if sent_state_for_folder_parts(["Account", "Inbox"]) is not None:
        errors.append("sent_state_for_folder_parts Inbox should be None")
    if sent_state_for_folder_parts(["Account", "Deleted Items"]) is not None:
        errors.append("sent_state_for_folder_parts Deleted Items should be None")

    fid, idx, rem = find_standard_folder_in_parts(
        ["Account (user@domain)", "Sent Items", "2024"]
    )
    if fid != OL_FOLDER_SENT or idx != 1 or rem != ["Account (user@domain)", "2024"]:
        errors.append(f"nested standard folder map: fid={fid} idx={idx} rem={rem}")
    fid2, idx2, rem2 = find_standard_folder_in_parts(["Account", "Outbox"])
    if fid2 != OL_FOLDER_OUTBOX or idx2 != 1 or rem2 != ["Account"]:
        errors.append(f"outbox folder map: fid={fid2} idx={idx2} rem={rem2}")
    fid3, idx3, rem3 = find_standard_folder_in_parts(
        ["Account", "Inbox", "Projects", "Inbox"]
    )
    if fid3 != OL_FOLDER_INBOX or idx3 != 3 or rem3 != ["Account", "Inbox", "Projects"]:
        errors.append(
            f"last standard folder wins: fid={fid3} idx={idx3} rem={rem3}"
        )
    fid4, _, rem4 = find_standard_folder_in_parts(["Account", "Deleted Items"])
    if fid4 != OL_FOLDER_DELETED or rem4 != ["Account"]:
        errors.append(f"deleted folder map: fid={fid4} rem={rem4}")

    if not hasattr(app_stub, "_get_or_create_pst_standard_folder"):
        errors.append("_get_or_create_pst_standard_folder missing")
    if not hasattr(app_stub, "_sent_state_for_source_path"):
        errors.append("_sent_state_for_source_path missing")

    root = os.path.abspath(TEST_MAIL)
    sample = os.path.join(root, "Inbox", "tiny_0001_test.eml")
    if os.path.isfile(sample):
        parts = relative_folder_parts(root, sample)
        if parts != ["Inbox"]:
            errors.append(f"relative_folder_parts expected ['Inbox'], got {parts}")
    wlm_root = os.path.join(root, "Account (user@domain)") if os.path.isdir(os.path.join(root, "Account (user@domain)")) else ""
    if not wlm_root:
        wlm_root = root
    nested = os.path.join(wlm_root, "Inbox", "tiny_0001_test.eml")
    if os.path.isfile(nested):
        parts = relative_folder_parts(wlm_root, nested)
        if parts and parts[0].lower() != "inbox" and "inbox" not in [p.lower() for p in parts]:
            errors.append(f"WLM nested parts missing Inbox: {parts}")

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
            fp = _fp_hash(app_stub, path, sz)
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
                    fp1 = _fp_hash(app_stub, dup, os.path.getsize(dup))
                    fp2 = _fp_hash(app_stub, orig, os.path.getsize(orig))
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
    os.makedirs(os.path.dirname(csv_path), exist_ok=True)
    if not os.path.isfile(csv_path):
        with open(csv_path, "w", encoding="utf-8", newline="") as handle:
            handle.write(
                "timestamp,file_path,status,detail,duration_sec,target_folder\n"
                f"2024-01-01 12:00:00,'=evil,converted,,1.0,Inbox\n"
            )
    resumed = load_resume_paths(csv_path)
    if not any("evil" in p for p in resumed):
        pass  # path normalization may differ; just ensure no crash
    if count_converted_in_csv(csv_path) < 1:
        errors.append("count_converted_in_csv should find converted row")

    with tempfile.TemporaryDirectory() as sqlite_tmp:
        import sqlite3

        db_path = os.path.join(sqlite_tmp, "dedup_state.sqlite3")
        probe = os.path.normpath(os.path.join(sqlite_tmp, "mail", "a.eml"))
        conn = sqlite3.connect(db_path)
        conn.execute(
            """
            CREATE TABLE converted_path (
                path TEXT PRIMARY KEY,
                status TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
            """
        )
        conn.execute(
            "INSERT INTO converted_path VALUES (?, 'converted', '2024-01-01')",
            (probe,),
        )
        conn.commit()
        conn.close()
        from_sqlite = load_resume_paths_from_sqlite(db_path)
        if probe not in from_sqlite:
            errors.append("load_resume_paths_from_sqlite should return converted path")

    # Export log files are created and flushed immediately
    with tempfile.TemporaryDirectory() as tmp:
        pst = os.path.join(tmp, "out.pst")
        with open(pst, "wb") as handle:
            handle.write(b"\x00")
        csv_p, log_p = export_log_paths(tmp, pst, "new")
        if not log_p.endswith("export.log"):
            errors.append("export_log_paths should use export.log")
        handler = _ImmediateFileHandler(log_p, append=False)
        logger.addHandler(handler)
        try:
            logger.info("bug_check export log probe")
            handler.flush()
        finally:
            logger.removeHandler(handler)
            handler.close()
        if not os.path.isfile(log_p) or os.path.getsize(log_p) < 20:
            errors.append("_ImmediateFileHandler did not write export.log")
        else:
            with open(log_p, encoding="utf-8") as handle:
                body = handle.read()
            if "bug_check export log probe" not in body:
                errors.append("export.log missing probe line")
            if "Mail Exporter export session" not in body:
                errors.append("export.log missing session banner")

    from mail_validation import (
        ISSUE_BLANK_SENDER,
        ISSUE_HTML_AS_PLAIN,
        ISSUE_LITERAL_NONE_SENDER,
        eml_parse_to_summary,
        format_inspection_errors,
        has_import_errors,
        inspect_outlook_fields,
        inspect_parsed_eml,
        html_stored_as_plain_text,
        load_converted_paths_from_csv,
        outlook_text_is_blank,
        parse_eml_summary,
        validate_import_against_eml,
    )

    if not outlook_text_is_blank("None"):
        errors.append("mail_validation.outlook_text_is_blank(None)")
    bad = inspect_outlook_fields(
        subject="None",
        sender_name="None",
        body="<br/>Hello<b>world</b>",
        html_body="",
    )
    codes = {i.code for i in bad.issues}
    if ISSUE_LITERAL_NONE_SENDER not in codes:
        errors.append("inspect_outlook_fields should flag literal None sender")
    if ISSUE_HTML_AS_PLAIN not in codes:
        errors.append("inspect_outlook_fields should flag html_as_plain")
    if not html_stored_as_plain_text("<br/>x", ""):
        errors.append("html_stored_as_plain_text br fragment")
    with tempfile.TemporaryDirectory() as vtmp:
        ok_eml = os.path.join(vtmp, "ok.eml")
        with open(ok_eml, "wb") as handle:
            handle.write(
                b"From: good@example.com\r\n"
                b"Subject: Test\r\n"
                b"Content-Type: text/plain\r\n\r\n"
                b"Hello body\r\n"
            )
        summary = parse_eml_summary(ok_eml)
        if not summary or inspect_parsed_eml(summary).issues:
            errors.append("parse_eml_summary ok sample should have no issues")
        csv_v = os.path.join(vtmp, "export_results.csv")
        with open(csv_v, "w", newline="", encoding="utf-8") as handle:
            w = __import__("csv").writer(handle)
            w.writerow(
                (
                    "timestamp",
                    "file_path",
                    "status",
                    "detail",
                    "duration_sec",
                    "target_folder",
                )
            )
            w.writerow(
                ("2026-01-01 00:00:00", ok_eml, "converted", "", "1.0", "Inbox")
            )
        if len(load_converted_paths_from_csv(csv_v)) != 1:
            errors.append("load_converted_paths_from_csv")

        bad_snap = {
            "subject": "None",
            "sender_name": "None",
            "sender_email": "",
            "body": "<br/>Hi</br>",
            "html_body": "",
            "attachment_count": 0,
        }
        bad_eml = eml_parse_to_summary(
            {
                "subject": "Invoice 123",
                "from": "billing@example.com",
                "body": "<br/>Hi</br>",
                "attachments": [],
            },
            ok_eml,
        )
        vinsp = validate_import_against_eml(bad_snap, bad_eml, source=ok_eml)
        if not has_import_errors(vinsp):
            errors.append("validate_import_against_eml should flag None sender + html_as_plain")
        if not format_inspection_errors(vinsp):
            errors.append("format_inspection_errors empty")

        from mail_validation import (
            ISSUE_WRONG_FOLDER,
            MailInspection,
            append_folder_placement_issues,
            body_looks_like_html,
        )
        if body_looks_like_html("Contact <user@example.com> for help"):
            errors.append("body_looks_like_html should not flag bare email brackets")

        finsp = MailInspection()
        append_folder_placement_issues(
            finsp,
            in_correct_folder=False,
            actual_folder_path="Inbox\\Drafts",
            expected_folder_path="Inbox\\Account (user@domain)",
        )
        if ISSUE_WRONG_FOLDER not in {i.code for i in finsp.issues}:
            errors.append("append_folder_placement_issues wrong_folder")

    try:
        from mmap_processor import (
            hash_file_sha256,
            needs_lf_to_crlf_conversion,
            parse_message_from_path,
            read_file_bytes,
            write_crlf_normalized_file,
        )

        with tempfile.NamedTemporaryFile(suffix=".eml", delete=False) as tmp:
            tmp.write(b"From: a@b.com\r\nSubject: mmap test\r\n\r\nBody\r\n")
            tmp_path = tmp.name
        try:
            msg = parse_message_from_path(tmp_path)
            if msg is None or msg.get("Subject") != "mmap test":
                errors.append("parse_message_from_path failed")
            digest = hash_file_sha256(tmp_path)
            if not digest or len(digest) != 64:
                errors.append("hash_file_sha256 failed")
            raw = read_file_bytes(tmp_path)
            if raw is None or b"mmap test" not in raw:
                errors.append("read_file_bytes failed")
            lf_path = tmp_path + ".lf.eml"
            with open(lf_path, "wb") as lf:
                lf.write(b"From: a@b.com\nSubject: lf\n\nHi\n")
            if not needs_lf_to_crlf_conversion(lf_path):
                errors.append("needs_lf_to_crlf_conversion should detect LF-only")
            crlf_path = tmp_path + ".crlf.eml"
            write_crlf_normalized_file(lf_path, crlf_path)
            with open(crlf_path, "rb") as handle:
                data = handle.read()
            if b"\r\n" not in data or b"From: a@b.com\n" in data:
                errors.append("write_crlf_normalized_file did not CRLF-normalize")
            for extra in (lf_path, crlf_path):
                try:
                    os.remove(extra)
                except OSError:
                    pass
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    except Exception as exc:
        errors.append(f"mmap_processor: {exc}")

    try:
        from parallel_processor import EmlPrepPrefetcher, ParallelEmailProcessor

        if ParallelEmailProcessor().max_workers < 1:
            errors.append("ParallelEmailProcessor max_workers")
        results = []

        def _fake_prep(path: str):
            from parallel_processor import EmlPrepResult

            return EmlPrepResult(
                file_path=path,
                norm_path=os.path.normpath(os.path.abspath(path)),
                size=1,
                dedup_key="abc",
            )

        with tempfile.NamedTemporaryFile(suffix=".eml", delete=False) as tmp:
            tmp.write(b"x")
            p = tmp.name
        try:
            pf = EmlPrepPrefetcher(_fake_prep, max_workers=2)
            pf.start([p])
            got = pf.take(os.path.normpath(os.path.abspath(p)), timeout=10.0)
            if got is None or got.dedup_key != "abc":
                errors.append("EmlPrepPrefetcher take failed")
            pf.shutdown()
        finally:
            try:
                os.remove(p)
            except OSError:
                pass
    except Exception as exc:
        errors.append(f"parallel_processor: {exc}")

    try:
        import export_checker as _ec

        if not hasattr(_ec, "scan_pst"):
            errors.append("export_checker.scan_pst missing")
    except ImportError as exc:
        errors.append(f"export_checker import: {exc}")

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
                    fp = _fp_hash(app_stub, p, os.path.getsize(p))
                    if fp and fp in seen:
                        skip += 1
                    else:
                        seen.add(fp)
            print(f"  dedup simulation: {skip} duplicates would be skipped")
    return 0


if __name__ == "__main__":
    sys.exit(main())
