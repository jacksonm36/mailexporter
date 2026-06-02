"""
Build a portable Mail Exporter executable (no Python install required on target PCs).

Requirements (build machine only):
    pip install -r requirements.txt

Usage:
    python build_exe.py              # single-file exe (portable)
    python build_exe.py --onedir     # folder bundle (more AV-friendly)
"""

from __future__ import annotations

import argparse
import os
import shutil
import struct
import subprocess
import sys
import tempfile

# win32com test/demos bloat the bundle and often trigger antivirus false positives.
EXCLUDE_MODULES = (
    "win32com.test",
    "win32com.demos",
    "win32com.makegw",
    "win32com.axdebug",
    "win32com.axscript",
    "win32com.directsound",
    "win32com.ifilter",
    "win32com.internet",
    "win32com.mapi",
    "win32com.propsys",
    "win32com.taskscheduler",
    "win32com.authorization",
    "win32com.bits",
    "win32com.adsi",
    "win32com.axcontrol",
)


def get_python_arch() -> int:
    return struct.calcsize("P") * 8


def check_pyinstaller() -> bool:
    try:
        import PyInstaller  # noqa: F401
        return True
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        return True


def _unblock_windows_file(path: str) -> None:
    if sys.platform != "win32":
        return
    try:
        subprocess.run(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                f"Unblock-File -LiteralPath '{path}'",
            ],
            check=False,
            capture_output=True,
            text=True,
        )
    except OSError:
        pass


def verify_executable(exe_path: str) -> bool:
    """Confirm the output is a readable PE file (not quarantined/corrupt)."""
    try:
        with open(exe_path, "rb") as handle:
            magic = handle.read(2)
    except OSError as exc:
        print(f"VERIFY FAILED: cannot read {exe_path}: {exc}")
        print(
            "Windows Defender or another antivirus may have quarantined the build.\n"
            "Add an exclusion for this project folder, then rebuild."
        )
        return False
    if magic != b"MZ":
        print(f"VERIFY FAILED: {exe_path} is not a valid Windows executable (missing MZ header)")
        return False
    print(f"Verified: {exe_path} is readable and has a valid PE header.")
    return True


def build_executable(*, onedir: bool = False, dist_dir: str | None = None) -> bool:
    arch = get_python_arch()
    script_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(script_dir, "eml_to_pst_converter.py")
    build_dir = os.path.join(script_dir, "build")
    if dist_dir is None:
        dist_dir = os.path.join(script_dir, "dist")
    # Fall back to LOCALAPPDATA if the project folder is not writable (e.g. synced Documents).
    try:
        probe = os.path.join(script_dir, ".build_write_probe")
        with open(probe, "w", encoding="utf-8") as handle:
            handle.write("ok")
        os.remove(probe)
    except OSError:
        local_base = os.path.join(
            os.environ.get("LOCALAPPDATA", tempfile.gettempdir()),
            "MailExporter_build",
        )
        build_dir = os.path.join(local_base, "build")
        if dist_dir == os.path.join(script_dir, "dist"):
            dist_dir = os.path.join(local_base, "dist")
        os.makedirs(build_dir, exist_ok=True)
        os.makedirs(dist_dir, exist_ok=True)
        print(f"Project folder not writable — using {local_base}")
    output_name = f"MailExporter_x{arch}"

    if not os.path.exists(main_script):
        print(f"Error: {main_script} not found")
        return False

    cmd = [
        sys.executable,
        "-m",
        "PyInstaller",
        "--windowed",
        f"--name={output_name}",
        f"--distpath={dist_dir}",
        f"--workpath={build_dir}",
        f"--specpath={build_dir}",
        "--clean",
        "--noconfirm",
        "--hidden-import=win32timezone",
        "--hidden-import=win32com.client",
        "--hidden-import=pythoncom",
        "--hidden-import=pywintypes",
    ]
    if onedir:
        cmd.append("--onedir")
    else:
        cmd.append("--onefile")
    for mod in EXCLUDE_MODULES:
        cmd.append(f"--exclude-module={mod}")
    cmd.append(main_script)

    mode = "folder bundle" if onedir else "single-file"
    print(f"Building portable {arch}-bit {mode} executable...\n")
    print(" ".join(cmd), "\n")

    try:
        subprocess.check_call(cmd)
    except subprocess.CalledProcessError as exc:
        print(f"Build failed: {exc}")
        return False

    if onedir:
        exe_path = os.path.join(dist_dir, output_name, f"{output_name}.exe")
    else:
        exe_path = os.path.join(dist_dir, f"{output_name}.exe")

    if not os.path.exists(exe_path):
        print("Error: executable was not created")
        return False

    _unblock_windows_file(exe_path)
    if not verify_executable(exe_path):
        return False

    size_mb = os.path.getsize(exe_path) / (1024 * 1024)
    print(f"\nSUCCESS: {exe_path} ({size_mb:.1f} MB)")
    print(f"Built for {arch}-bit Outlook — match exe bitness to your Outlook install.")
    if onedir:
        folder = os.path.dirname(exe_path)
        print(f"Portable folder: copy the entire '{folder}' directory to use elsewhere.")
    else:
        print("Portable: copy the single .exe anywhere; only Microsoft Outlook is required.")
    print(
        "\nIf Windows SmartScreen or antivirus blocks the file:\n"
        "  1. Windows Security -> Protection history -> allow/restored file\n"
        "  2. Or add an exclusion for this project's dist folder\n"
        "  3. Rebuild with: python build_exe.py --onedir  (often fewer false positives)"
    )
    return True


def clean_build_artifacts() -> None:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    build_dir = os.path.join(script_dir, "build")
    if os.path.exists(build_dir):
        shutil.rmtree(build_dir)
    for name in os.listdir(script_dir):
        if name.endswith(".spec"):
            os.remove(os.path.join(script_dir, name))


def main() -> int:
    parser = argparse.ArgumentParser(description="Build portable Mail Exporter exe")
    parser.add_argument(
        "--onedir",
        action="store_true",
        help="Build a folder bundle instead of a single file (more reliable with antivirus)",
    )
    parser.add_argument(
        "--dist-dir",
        default=None,
        help="Output directory (default: ./dist). Example: %%LOCALAPPDATA%%\\MailExporter",
    )
    args = parser.parse_args()

    dist_dir = args.dist_dir
    if dist_dir and "%" in dist_dir:
        dist_dir = os.path.expandvars(dist_dir)
    if dist_dir is None:
        dist_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist")

    if not check_pyinstaller():
        return 1
    ok = build_executable(onedir=args.onedir, dist_dir=dist_dir)
    if ok:
        clean_build_artifacts()
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
