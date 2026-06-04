"""
v1.0.2 patch verifier — checks that security/build fixes are present in the tree.

The naive string-replace patch script is superseded by integrated changes in:
  - path_security.py         (PathValidator, PathSecurityError)
  - eml_to_pst_converter.py  (validate_import_path → PathValidator, …)
  - build_exe.py             (--version-file=version_info.py on Windows)
  - version_info.py          (PE version 1.0.2)

Run:  python v1.0.2_patch.py
Then: python build_exe.py --arch 32
"""

from __future__ import annotations

import os
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))


def patch_path_traversal() -> bool:
    """Verify path traversal protection exists."""
    path = os.path.join(ROOT, "eml_to_pst_converter.py")
    with open(path, encoding="utf-8") as handle:
        content = handle.read()
    sec_path = os.path.join(ROOT, "path_security.py")
    ok = os.path.isfile(sec_path) and "class PathValidator" in open(
        sec_path, encoding="utf-8"
    ).read()
    with open(path, encoding="utf-8") as handle:
        conv = handle.read()
    ok = ok and "PathValidator" in conv and "def validate_import_path" in conv
    if ok:
        print("OK  Path security module (path_security.py + validate_import_path)")
    else:
        print("FAIL  path_security.py or PathValidator integration missing")
    return ok


def add_version_info() -> bool:
    """Verify Windows build embeds version metadata."""
    build_path = os.path.join(ROOT, "build_exe.py")
    version_path = os.path.join(ROOT, "version_info.py")
    with open(build_path, encoding="utf-8") as handle:
        build_src = handle.read()
    ok = (
        os.path.isfile(version_path)
        and "--version-file" in build_src
        and "1.0.2" in open(version_path, encoding="utf-8").read()
    )
    if ok:
        print("OK  Version info (version_info.py + build_exe --version-file)")
    else:
        print("FAIL  version_info.py or build_exe.py --version-file")
    return ok


def main() -> int:
    sys.path.insert(0, ROOT)
    checks = [patch_path_traversal(), add_version_info()]
    try:
        from eml_to_pst_converter import validate_import_path
        import tempfile

        with tempfile.TemporaryDirectory() as tmp:
            base = os.path.join(tmp, "mail")
            os.makedirs(base)
            good = os.path.join(base, "inbox", "a.eml")
            os.makedirs(os.path.dirname(good), exist_ok=True)
            open(good, "w", encoding="utf-8").close()
            validate_import_path(base, good)
            try:
                validate_import_path(base, os.path.join("..", "outside.eml"))
                print("FAIL  validate_import_path should block ..")
                checks.append(False)
            except ValueError:
                print("OK  validate_import_path blocks traversal")
                checks.append(True)
    except Exception as exc:
        print(f"FAIL  validate_import_path runtime test: {exc}")
        checks.append(False)

    if all(checks):
        print("\nAll v1.0.2 patch checks passed. Rebuild with: python build_exe.py")
        return 0
    print("\nSome checks failed.")
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
