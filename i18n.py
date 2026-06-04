"""
English / Hungarian UI strings for EML to PST Converter.
"""
from __future__ import annotations

import locale
import os

LANG_EN = "en"
LANG_HU = "hu"

# Combobox display names
LANGUAGE_NAMES = {
    LANG_EN: "English",
    LANG_HU: "Magyar",
}


def detect_default_lang() -> str:
    """Pick Hungarian if Windows/user locale is Hungarian."""
    try:
        loc = locale.getdefaultlocale()[0]
        if loc and str(loc).lower().startswith("hu"):
            return LANG_HU
    except Exception:
        pass
    env = os.environ.get("EML2PST_LANG", "").strip().lower()
    if env in ("hu", "hungarian", "magyar"):
        return LANG_HU
    if env in ("en", "english"):
        return LANG_EN
    return LANG_EN


def lang_from_display(name: str) -> str:
    if name == LANGUAGE_NAMES[LANG_HU]:
        return LANG_HU
    return LANG_EN


STRINGS: dict[str, dict[str, str]] = {
    LANG_EN: {
        "window_title": "Mail Exporter",
        "folder_section": "Add Folder Having *.eml / *.emlx Files",
        "add_files": "Add Files",
        "wlm_preset": "Live Mail Folder",
        "file_pattern_label": "File Pattern:",
        "file_pattern_hint": "(Wildcards: * for any characters, ? for single character)",
        "save_pst_section": "Export Destination",
        "create_new_pst": "Create New PST File",
        "save_existing_pst": "Save to Existing PST File",
        "export_to_mailbox": "Import to Outlook Mailbox (Exchange / M365)",
        "mailbox_store_label": "Mailbox:",
        "mailbox_folder_label": "Target folder:",
        "refresh_stores": "Refresh",
        "warn_wlm_not_found": (
            "Windows Live Mail folder not found at:\n{path}\n\n"
            "Use Add Files to pick your .eml folder manually."
        ),
        "warn_no_mailbox": "Please select an Outlook mailbox.",
        "warn_no_stores": (
            "No Outlook mailboxes found. Open Outlook, sign in, then click Refresh."
        ),
        "msg_stores_error": "Could not list Outlook mailboxes:\n{err}",
        "msg_outlook_required_frozen": (
            "Microsoft Outlook must be installed and running on this PC.\n\n"
            "This portable app does not require Python or pip — only Outlook "
            "(64-bit Outlook for this 64-bit app)."
        ),
        "status_preparing_mailbox": "Preparing mailbox folder...",
        "note_mailbox_saved": (
            "\n\nImported to mailbox: {store}\nFolder: {folder}"
        ),
        "remove_duplicates": "Remove Duplicate Content",
        "strict_date_preservation": (
            "Strict header dates (skip if Date/Received cannot be applied — optional)"
        ),
        # Explorer column: English "Date modified" vs Hungarian "Módosítás dátuma"
        "explorer_date_modified": "Date modified",
        "use_file_mtime": (
            "Use Explorer {explorer} for message date (on by default; matches .eml folder)"
        ),
        "eml_files_section": "EML/EMLX Files",
        "col_name": "EML/EMLX Name",
        "col_path": "Path",
        "col_size": "Size",
        "col_date": "Date Modified",
        "total_files": "Total Files: {n}",
        "total_files_preview": "Total Files: {total} (showing first {shown} in list)",
        "destination_label": "Destination:",
        "browse_destination": "Browse Destination",
        "status_ready": "Ready",
        "btn_exit": "Exit",
        "btn_convert": "Convert",
        "language_label": "Language:",
        "context_remove": "Remove Selected",
        "context_clear_all": "Clear All",
        "dialog_select_folder": "Select Folder Containing EML/EMLX Files",
        "dialog_save_pst": "Save PST File",
        "dialog_open_pst": "Select Existing PST File",
        "dialog_filetype_pst": "Outlook PST Files",
        "dialog_all_files": "All Files",
        "confirm_exit_title": "Confirm",
        "confirm_exit_msg": "Conversion in progress. Exit anyway?",
        "warn_in_progress": "Conversion already in progress!",
        "warn_no_files": "No EML/EMLX files to convert!",
        "warn_no_destination": "Please select a destination path!",
        "err_dest_dir": "Destination directory does not exist: {path}",
        "err_scan": "Error scanning folder: {err}",
        "err_invalid_pattern": "Invalid file pattern. Only *.eml, *.emlx patterns are allowed.",
        "title_error": "Error",
        "title_warning": "Warning",
        "title_conversion_error": "Conversion Error",
        "msg_conversion_error": (
            "Error during conversion:\n{err}\n\n"
            "Make sure Microsoft Outlook is installed and working properly."
        ),
        "title_missing_dep": "Missing Dependency",
        "msg_missing_dep": (
            "The 'pywin32' library is required for PST conversion.\n\n"
            "Would you like to install it now?\n\n"
            "(This requires an internet connection)"
        ),
        "status_installing_pywin32": "Installing pywin32...",
        "title_install_ok": "Installation Complete",
        "msg_install_ok": (
            "pywin32 has been installed successfully!\n\n"
            "Please restart the application to use PST conversion."
        ),
        "title_install_fail": "Installation Failed",
        "msg_install_fail": (
            "Failed to install pywin32.\n\n"
            "Please run manually in command prompt:\npip install pywin32"
        ),
        "status_scanning": "Scanning folder...",
        "status_scanning_progress": "Scanning folder... {n} files found so far",
        "warn_scan_limit": (
            "Scan stopped at the file limit ({limit}). "
            "Increase EML2PST_MAX_FILES if you need to process more in one run."
        ),
        "status_found": "Found {n} files",
        "status_preflight": "Preparing export...",
        "status_checking_outlook": "Checking Outlook (COM)...",
        "status_export_starting": "Starting export ({n} files)...",
        "status_log_path": "Log: {path}  (mirror: %LOCALAPPDATA%\\MailExporter\\last_export.log)",
        "err_export_log": "Could not create export log or CSV:\n{err}",
        "status_connecting": "Connecting to Outlook...",
        "status_creating_pst": "Creating PST file...",
        "status_chunk_pst": "Creating chunk PST {cur}/{total}...",
        "status_merging_pst": "Merging chunk PSTs into final file...",
        "status_merging_chunk": "Merging chunk {cur}/{total}...",
        "status_chain_import": "Chain {cur}/{total}: importing batch to chunk PST...",
        "status_chain_merge": "Chain {cur}/{total}: merging chunk into final PST...",
        "status_converting": "Converting: {name} ({cur}/{total})",
        "status_complete": "Conversion Complete",
        "title_complete": "Complete",
        "msg_complete": (
            "Conversion Complete!\n\n"
            "Converted: {converted} files\n"
            "Skipped: {skipped} files\n"
            "Errors: {errors} files"
        ),
        "note_pst_saved": "\n\nPST saved to: {path}",
        "note_pst_chunked": (
            "\n\nImported in {chunks} temporary PST chunk(s) of up to {size} emails, "
            "then merged into the final PST."
        ),
        "note_pst_chained": (
            "\n\nBatched export: {links} batch(es) of up to {size} emails "
            "imported directly into the final PST."
        ),
        "note_pst_partial_merge": (
            "\n\nExport was cancelled; completed batch(es) are already in the PST."
        ),
        "note_chunks_left": "\n\nTemporary chunk folder (not merged): {dir}",
        "note_chain_cancelled": "\n\nExport cancelled before any batch was merged.",
        "note_errors": "\n\nErrors:\n",
        "note_skipped": "\n\nSkipped:\n",
        "note_more_errors": "\n... and {n} more errors",
        "note_more_skipped": "\n... and {n} more skipped",
        "preserve_subfolders": "Preserve account subfolders (Inbox/Sent/Deleted always matched)",
        "resume_from_log": "Resume from prior export log",
        "btn_cancel": "Cancel",
        "skip_resume": "Already imported (resume)",
        "preflight_title": "Ready to export",
        "preflight_summary": "Files: {count}  (~{size} MB)",
        "preflight_size_estimate": "(Size estimated from a sample of {n} files)",
        "warn_preview_remove": (
            "The list shows only the first {shown} of {total} files.\n\n"
            "Use Clear All to remove the entire batch, or rescan the folder to change the selection."
        ),
        "preflight_outlook_exe": "Outlook: {path} ({bits}-bit)",
        "preflight_bitness": (
            "This app is {app}-bit but Outlook is {outlook}-bit.\n"
            "Outlook COM test failed. Close this program and run {recommended} instead."
        ),
        "preflight_com_failed": "Outlook COM test failed: {err}",
        "preflight_outlook_unknown": "Could not locate OUTLOOK.EXE or read its bitness.",
        "preflight_resume": "Resume will skip {n} file(s) already marked converted in export_results.csv.",
        "preflight_continue": "Start export now?",
        "preflight_existing_pst": "Target PST already exists — new messages will be appended (file is not replaced).",
        "status_cancelling": "Cancelling...",
        "status_cancelled": "Cancelled",
        "title_cancelled": "Export cancelled",
        "msg_cancelled": (
            "Export cancelled.\n\n"
            "Converted: {converted} files\n"
            "Skipped: {skipped} files\n"
            "Errors: {errors} files"
        ),
        "note_csv_saved": "\n\nResults CSV: {path}",
        "note_log_saved": "\n\nLog file: {path}",
    },
    LANG_HU: {
        "window_title": "Mail Exporter",
        "folder_section": "Mappa hozzáadása *.eml / *.emlx fájlokkal",
        "add_files": "Fájlok hozzáadása",
        "wlm_preset": "Live Mail mappa",
        "file_pattern_label": "Fájlminta:",
        "file_pattern_hint": "(Helyettesítők: * több karakter, ? egy karakter)",
        "save_pst_section": "Export célja",
        "create_new_pst": "Új PST fájl létrehozása",
        "save_existing_pst": "Meglévő PST fájlba mentés",
        "export_to_mailbox": "Import Outlook postafiókba (Exchange / M365)",
        "mailbox_store_label": "Postafiók:",
        "mailbox_folder_label": "Célmappa:",
        "refresh_stores": "Frissítés",
        "warn_wlm_not_found": (
            "A Windows Live Mail mappa nem található:\n{path}\n\n"
            "Használja a Fájlok hozzáadása gombot a .eml mappa kiválasztásához."
        ),
        "warn_no_mailbox": "Válasszon Outlook postafiókot.",
        "warn_no_stores": (
            "Nem található Outlook postafiók. Indítsa el az Outlookot, jelentkezzen be, majd Frissítés."
        ),
        "msg_stores_error": "Az Outlook postafiókok listázása sikertelen:\n{err}",
        "msg_outlook_required_frozen": (
            "A Microsoft Outlook telepítve és futó legyen ezen a gépen.\n\n"
            "Ehhez a hordozható apphoz nem kell Python vagy pip — csak Outlook "
            "(64 bites Outlook ehhez a 64 bites apphoz)."
        ),
        "status_preparing_mailbox": "Postafiók mappa előkészítése...",
        "note_mailbox_saved": (
            "\n\nImportálva postafiókba: {store}\nMappa: {folder}"
        ),
        "remove_duplicates": "Duplikátumok eltávolítása",
        "strict_date_preservation": (
            "Szigorú fejléc dátumok (kihagyás, ha Date/Received nem alkalmazható — opcionális)"
        ),
        "explorer_date_modified": "Módosítás dátuma",
        "use_file_mtime": (
            "Tallózó {explorer} az üzenet dátumához (alapból be; egyezik az .eml mappával)"
        ),
        "eml_files_section": "EML/EMLX fájlok",
        "col_name": "EML/EMLX név",
        "col_path": "Útvonal",
        "col_size": "Méret",
        "col_date": "Módosítás dátuma",
        "total_files": "Összes fájl: {n}",
        "total_files_preview": "Összes fájl: {total} (listában az első {shown})",
        "destination_label": "Cél:",
        "browse_destination": "Cél tallózása",
        "status_ready": "Kész",
        "btn_exit": "Kilépés",
        "btn_convert": "Konvertálás",
        "language_label": "Nyelv:",
        "context_remove": "Kijelöltek eltávolítása",
        "context_clear_all": "Összes törlése",
        "dialog_select_folder": "Mappa kiválasztása EML/EMLX fájlokkal",
        "dialog_save_pst": "PST fájl mentése",
        "dialog_open_pst": "Meglévő PST fájl kiválasztása",
        "dialog_filetype_pst": "Outlook PST fájlok",
        "dialog_all_files": "Minden fájl",
        "confirm_exit_title": "Megerősítés",
        "confirm_exit_msg": "Konvertálás folyamatban. Biztosan kilép?",
        "warn_in_progress": "A konvertálás már folyamatban van!",
        "warn_no_files": "Nincs konvertálandó EML/EMLX fájl!",
        "warn_no_destination": "Válasszon célútvonalat!",
        "err_dest_dir": "A célmappa nem létezik: {path}",
        "err_scan": "Hiba a mappa beolvasásakor: {err}",
        "err_invalid_pattern": "Érvénytelen minta. Csak *.eml és *.emlx minták engedélyezettek.",
        "title_error": "Hiba",
        "title_warning": "Figyelmeztetés",
        "title_conversion_error": "Konvertálási hiba",
        "msg_conversion_error": (
            "Hiba a konvertálás során:\n{err}\n\n"
            "Ellenőrizze, hogy a Microsoft Outlook telepítve és működőképes-e."
        ),
        "title_missing_dep": "Hiányzó összetevő",
        "msg_missing_dep": (
            "A PST konvertáláshoz a 'pywin32' könyvtár szükséges.\n\n"
            "Telepítse most?\n\n"
            "(Internetkapcsolat szükséges)"
        ),
        "status_installing_pywin32": "pywin32 telepítése...",
        "title_install_ok": "Telepítés kész",
        "msg_install_ok": (
            "A pywin32 sikeresen települt.\n\n"
            "Indítsa újra az alkalmazást a PST konvertáláshoz."
        ),
        "title_install_fail": "Telepítés sikertelen",
        "msg_install_fail": (
            "A pywin32 telepítése nem sikerült.\n\n"
            "Parancssorból futtassa: pip install pywin32"
        ),
        "status_scanning": "Mappa beolvasása...",
        "status_scanning_progress": "Mappa beolvasása... eddig {n} fájl",
        "warn_scan_limit": (
            "A beolvasás elérte a fájllimitet ({limit}). "
            "Növelje az EML2PST_MAX_FILES értékét, ha több fájlt kell egyszerre feldolgozni."
        ),
        "status_found": "{n} fájl találva",
        "status_preflight": "Export előkészítése...",
        "status_checking_outlook": "Outlook ellenőrzése (COM)...",
        "status_export_starting": "Export indítása ({n} fájl)...",
        "status_log_path": "Napló: {path}  (másolat: %LOCALAPPDATA%\\MailExporter\\last_export.log)",
        "err_export_log": "Az export napló vagy CSV nem hozható létre:\n{err}",
        "status_connecting": "Kapcsolódás az Outlookhoz...",
        "status_creating_pst": "PST fájl létrehozása...",
        "status_chunk_pst": "Ideiglenes PST létrehozása {cur}/{total}...",
        "status_merging_pst": "Ideiglenes PST-ek egyesítése a végső fájlba...",
        "status_merging_chunk": "Darab egyesítése {cur}/{total}...",
        "status_chain_import": "Lánc {cur}/{total}: köteg importálása ideiglenes PST-be...",
        "status_chain_merge": "Lánc {cur}/{total}: ideiglenes PST egyesítése a végsőbe...",
        "status_converting": "Konvertálás: {name} ({cur}/{total})",
        "status_complete": "Konvertálás kész",
        "title_complete": "Kész",
        "msg_complete": (
            "Konvertálás kész!\n\n"
            "Konvertálva: {converted} fájl\n"
            "Kihagyva: {skipped} fájl\n"
            "Hibák: {errors} fájl"
        ),
        "note_pst_saved": "\n\nPST mentve ide: {path}",
        "note_pst_chunked": (
            "\n\nImport {chunks} ideiglenes PST darabban (max. {size} e-mail), "
            "majd egyesítés a végső PST-be."
        ),
        "note_pst_chained": (
            "\n\nKötegelt export: {links} köteg (max. {size} e-mail) "
            "közvetlenül a végső PST-be importálva."
        ),
        "note_pst_partial_merge": (
            "\n\nAz export megszakadt; a kész kötegek már a PST-ben vannak."
        ),
        "note_chunks_left": "\n\nIdeiglenes chunk mappa (nem egyesítve): {dir}",
        "note_chain_cancelled": "\n\nExport megszakítva, mielőtt bármely köteg egyesítve lett volna.",
        "note_errors": "\n\nHibák:\n",
        "note_skipped": "\n\nKihagyva:\n",
        "note_more_errors": "\n... és még {n} hiba",
        "note_more_skipped": "\n... és még {n} kihagyott",
        "preserve_subfolders": "Fiók almappák megőrzése (Beérkezett/Elküldött/Törölt mindig egyezik)",
        "resume_from_log": "Folytatás korábbi export naplóból",
        "btn_cancel": "Mégse",
        "skip_resume": "Már importálva (folytatás)",
        "preflight_title": "Export indítása",
        "preflight_summary": "Fájlok: {count}  (~{size} MB)",
        "preflight_size_estimate": "(Méret becslés {n} fájl mintájából)",
        "warn_preview_remove": (
            "A listában csak az első {shown} / {total} fájl látható.\n\n"
            "Az egész listához használja az Összes törlése gombot, vagy olvassa be újra a mappát."
        ),
        "preflight_outlook_exe": "Outlook: {path} ({bits} bites)",
        "preflight_bitness": (
            "Az app {app} bites, az Outlook {outlook} bites.\n"
            "Az Outlook COM teszt sikertelen. Zárja be ezt a programot, és futtassa: {recommended}"
        ),
        "preflight_com_failed": "Outlook COM teszt sikertelen: {err}",
        "preflight_outlook_unknown": "Az OUTLOOK.EXE nem található, vagy a bitness nem olvasható.",
        "preflight_resume": "Folytatáskor {n} fájl kihagyása (már converted az export_results.csv-ben).",
        "preflight_continue": "Indítja az exportot?",
        "preflight_existing_pst": "A cél PST már létezik — az új levelek hozzáadódnak (a fájl nem lesz felülírva).",
        "status_cancelling": "Megszakítás...",
        "status_cancelled": "Megszakítva",
        "title_cancelled": "Export megszakítva",
        "msg_cancelled": (
            "Export megszakítva.\n\n"
            "Konvertálva: {converted} fájl\n"
            "Kihagyva: {skipped} fájl\n"
            "Hibák: {errors} fájl"
        ),
        "note_csv_saved": "\n\nEredmény CSV: {path}",
        "note_log_saved": "\n\nNapló fájl: {path}",
    },
}


def t(lang: str, key: str, **kwargs) -> str:
    """Translate a string; falls back to English if key missing."""
    table = STRINGS.get(lang) or STRINGS[LANG_EN]
    s = table.get(key)
    if s is None:
        s = STRINGS[LANG_EN].get(key, key)
    if kwargs:
        return s.format(**kwargs)
    return s
