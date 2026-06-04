# PyInstaller --version-file resource (Windows PE version metadata).
# Sync filevers/prodvers with APP_VERSION in eml_to_pst_converter.py.

VSVersionInfo(
    ffi=FixedFileInfo(
        filevers=(1, 0, 2, 0),
        prodvers=(1, 0, 2, 0),
        mask=0x3F,
        flags=0x0,
        OS=0x40004,
        fileType=0x1,
        subtype=0x0,
        date=(0, 0),
    ),
    kids=[
        StringFileInfo(
            [
                StringTable(
                    "040904B0",
                    [
                        StringStruct("CompanyName", "Mail Exporter"),
                        StringStruct("FileDescription", "EML to PST Converter"),
                        StringStruct("FileVersion", "1.0.2"),
                        StringStruct("InternalName", "MailExporter"),
                        StringStruct("LegalCopyright", "Copyright (c) 2026"),
                        StringStruct("OriginalFilename", "MailExporter.exe"),
                        StringStruct("ProductName", "Mail Exporter"),
                        StringStruct("ProductVersion", "1.0.2"),
                    ],
                )
            ]
        ),
        VarFileInfo([VarStruct("Translation", [1033, 1200])]),
    ],
)
