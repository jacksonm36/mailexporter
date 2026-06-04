"""
Build a portable Mail Exporter executable (no Python install required on target PCs).

Requirements (build machine only):
    pip install -r requirements.txt

Usage:
    python build_exe.py              # both MailExporter_x32.exe and MailExporter_x64.exe
    python build_exe.py --onedir     # folder bundles (more AV-friendly)
    python build_exe.py --arch 32    # build only 32-bit (match 32-bit Outlook)
    python build_exe.py --arch 64    # build only 64-bit (match 64-bit Outlook)
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


def get_python_arch_for(python_exe: str) -> int:
    try:
        out = subprocess.run(
            [python_exe, "-c", "import struct; print(struct.calcsize('P') * 8)"],
            capture_output=True,
            text=True,
            check=True,
            timeout=30,
        )
        return int(out.stdout.strip())
    except (OSError, subprocess.CalledProcessError, ValueError) as exc:
        raise RuntimeError(f"Could not determine bitness of {python_exe}: {exc}") from exc


def resolve_python_for_arch(target_arch: int) -> str | None:
    """Find a Python executable matching target_arch (32 or 64)."""
    if get_python_arch() == target_arch:
        return sys.executable
    if sys.platform != "win32":
        return None
    candidates: list[str] = []
    if target_arch == 64:
        candidates.extend(
            [
                os.path.join(
                    os.environ.get("LOCALAPPDATA", ""),
                    "Python",
                    "pythoncore-3.14-64",
                    "python.exe",
                ),
                os.path.join(
                    os.environ.get("LOCALAPPDATA", ""),
                    "Programs",
                    "Python",
                    "Python314",
                    "python.exe",
                ),
                os.path.join(
                    os.environ.get("LOCALAPPDATA", ""),
                    "Programs",
                    "Python",
                    "Python312",
                    "python.exe",
                ),
            ]
        )
        launcher_tag = "3.14-64"
    else:
        candidates.extend(
            [
                os.path.join(
                    os.environ.get("LOCALAPPDATA", ""),
                    "Programs",
                    "Python",
                    "Python38-32",
                    "python.exe",
                ),
                os.path.join(
                    os.environ.get("LOCALAPPDATA", ""),
                    "Programs",
                    "Python",
                    "Python313-32",
                    "python.exe",
                ),
            ]
        )
        launcher_tag = "3.8-32"
    for path in candidates:
        if path and os.path.isfile(path):
            try:
                out = subprocess.run(
                    [path, "-c", "import struct; print(struct.calcsize('P') * 8)"],
                    capture_output=True,
                    text=True,
                    check=True,
                    timeout=30,
                )
                if int(out.stdout.strip()) == target_arch:
                    return os.path.normpath(path)
            except (OSError, subprocess.CalledProcessError, ValueError):
                continue
    py_exe = shutil.which("py") or os.path.join(os.environ.get("WINDIR", "C:\\Windows"), "py.exe")
    if os.path.isfile(py_exe):
        try:
            out = subprocess.run(
                [py_exe, f"-{launcher_tag}", "-c", "import sys; print(sys.executable)"],
                capture_output=True,
                text=True,
                check=True,
                timeout=30,
            )
            resolved = out.stdout.strip()
            if resolved and os.path.isfile(resolved):
                return os.path.normpath(resolved)
        except (OSError, subprocess.CalledProcessError):
            pass
    return None


def check_pyinstaller(python_exe: str | None = None) -> bool:
    python_exe = python_exe or sys.executable
    try:
        subprocess.run(
            [python_exe, "-c", "import PyInstaller"],
            check=True,
            capture_output=True,
            timeout=30,
        )
        return True
    except (subprocess.CalledProcessError, OSError):
        print("Installing PyInstaller...")
        subprocess.check_call(
            [python_exe, "-m", "pip", "install", "pyinstaller", "pywin32"]
        )
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


def _resolve_env_file(script_dir: str, env_file: str | None) -> str | None:
    if env_file:
        path = os.path.normpath(os.path.expandvars(env_file))
        return path if os.path.isfile(path) else None
    for name in (".env", ".env.example"):
        path = os.path.join(script_dir, name)
        if os.path.isfile(path):
            return path
    return None


def prepare_build_env(script_dir: str, env_file: str | None, *, embed: bool) -> str | None:
    """Generate embedded_env.py from .env and return the source path used."""
    from env_config import generate_embedded_env_py, parse_dotenv

    out = os.path.join(script_dir, "embedded_env.py")
    if env_file is None:
        return None
    if embed:
        if generate_embedded_env_py(env_file, out, source_label=os.path.basename(env_file)):
            n = len(parse_dotenv(env_file))
            print(f"Embedded {n} EML2PST_* setting(s) from {env_file} -> embedded_env.py")
    return env_file


def _deploy_env_file(exe_path: str, env_source: str | None) -> None:
    if not env_source or not os.path.isfile(exe_path):
        return
    from env_config import copy_env_beside_exe

    dest = copy_env_beside_exe(exe_path, env_source)
    if dest:
        print(f"Copied {env_source} -> {dest} (edit beside exe to override embedded defaults)")


def build_executable(
    *,
    onedir: bool = False,
    dist_dir: str | None = None,
    python_exe: str | None = None,
    checker: bool = False,
    env_file: str | None = None,
    embed_env: bool = True,
) -> bool:
    python_exe = python_exe or sys.executable
    arch = get_python_arch_for(python_exe)
    output_name = f"ExportChecker_x{arch}" if checker else f"MailExporter_x{arch}"
    script_dir = os.path.dirname(os.path.abspath(__file__))
    main_script = os.path.join(
        script_dir, "export_checker.py" if checker else "eml_to_pst_converter.py"
    )
    build_dir = os.path.join(script_dir, "build", output_name)
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
        build_dir = os.path.join(local_base, "build", output_name)
        if dist_dir == os.path.join(script_dir, "dist"):
            dist_dir = os.path.join(local_base, "dist")
        os.makedirs(build_dir, exist_ok=True)
        os.makedirs(dist_dir, exist_ok=True)
        print(f"Project folder not writable — using {local_base}")

    if not os.path.exists(main_script):
        print(f"Error: {main_script} not found")
        return False

    env_source = None
    if not checker:
        resolved_env = _resolve_env_file(script_dir, env_file)
        env_source = prepare_build_env(script_dir, resolved_env, embed=embed_env)

    window_flag = "--console" if checker else "--windowed"
    cmd = [
        python_exe,
        "-m",
        "PyInstaller",
        window_flag,
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
        "--hidden-import=env_config",
        "--hidden-import=embedded_env",
    ]
    if onedir:
        cmd.append("--onedir")
    else:
        cmd.append("--onefile")
    for mod in EXCLUDE_MODULES:
        cmd.append(f"--exclude-module={mod}")
    if sys.platform == "win32" and not checker:
        version_file = os.path.join(script_dir, "version_info.py")
        if os.path.isfile(version_file):
            cmd.append(f"--version-file={version_file}")
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
        bundle_dir = os.path.join(dist_dir, output_name)
        exe_path = os.path.join(bundle_dir, f"{output_name}.exe")
        if not os.path.isfile(exe_path):
            try:
                for name in os.listdir(bundle_dir):
                    if name.lower() == f"{output_name.lower()}.exe":
                        exe_path = os.path.join(bundle_dir, name)
                        break
            except OSError:
                pass
    else:
        exe_path = os.path.join(dist_dir, f"{output_name}.exe")

    if not os.path.isfile(exe_path):
        quarantined = []
        if onedir:
            try:
                for name in os.listdir(os.path.join(dist_dir, output_name)):
                    if name.lower().endswith((".cynet", ".blocked", ".quarantine")):
                        quarantined.append(name)
            except OSError:
                pass
        if quarantined:
            print(
                "Error: executable was quarantined by antivirus "
                f"({', '.join(quarantined)}). Allow the file or exclude the dist "
                "folder, then rebuild."
            )
        else:
            print("Error: executable was not created")
        return False

    _unblock_windows_file(exe_path)
    if not verify_executable(exe_path):
        return False

    if not checker and env_source:
        _deploy_env_file(exe_path, env_source)

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


def build_for_arch(
    target_arch: int,
    *,
    onedir: bool,
    dist_dir: str,
    python_exe: str | None = None,
    checker: bool = False,
    env_file: str | None = None,
    embed_env: bool = True,
) -> bool:
    if python_exe is None:
        python_exe = resolve_python_for_arch(target_arch)
        if not python_exe:
            print(
                f"Error: no {target_arch}-bit Python found. Install Python {target_arch}-bit "
                f"or pass --python PATH_TO_PYTHON.exe with --arch {target_arch}"
            )
            return False
        print(f"Using {target_arch}-bit Python: {python_exe}")
    if not check_pyinstaller(python_exe):
        return False
    if get_python_arch_for(python_exe) != target_arch:
        print(
            f"Error: {python_exe} is {get_python_arch_for(python_exe)}-bit, "
            f"expected {target_arch}-bit"
        )
        return False
    return build_executable(
        onedir=onedir,
        dist_dir=dist_dir,
        python_exe=python_exe,
        checker=checker,
        env_file=env_file,
        embed_env=embed_env,
    )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Build portable Mail Exporter exe (x32 + x64 by default)"
    )
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
    parser.add_argument(
        "--arch",
        type=int,
        choices=(32, 64),
        default=None,
        help="Build only this bitness (default: build both 32 and 64)",
    )
    parser.add_argument(
        "--python",
        default=None,
        help="Explicit Python executable (only with --arch; overrides auto-detect)",
    )
    parser.add_argument(
        "--checker",
        action="store_true",
        help="Build ExportChecker_x32/x64.exe (PST validation CLI) instead of MailExporter",
    )
    parser.add_argument(
        "--env-file",
        default=None,
        help="Path to .env to embed and copy beside exe (default: ./.env or ./.env.example)",
    )
    parser.add_argument(
        "--no-embed-env",
        action="store_true",
        help="Do not bake .env into embedded_env.py (still copies .env next to exe if present)",
    )
    args = parser.parse_args()

    if args.python and args.arch is None:
        print("Error: --python requires --arch 32 or --arch 64")
        return 1

    dist_dir = args.dist_dir
    if dist_dir and "%" in dist_dir:
        dist_dir = os.path.expandvars(dist_dir)
    if dist_dir is None:
        dist_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "dist")

    archs = [args.arch] if args.arch is not None else [32, 64]
    product = "ExportChecker" if args.checker else "MailExporter"
    if len(archs) == 2:
        print(f"Building both 32-bit and 64-bit {product} executables...\n")

    python_exe = None
    if args.python:
        python_exe = os.path.normpath(os.path.expandvars(args.python))
        if not os.path.isfile(python_exe):
            print(f"Error: Python not found: {python_exe}")
            return 1

    all_ok = True
    for target_arch in archs:
        if len(archs) > 1:
            print(f"\n{'=' * 60}\n{target_arch}-bit build\n{'=' * 60}\n")
        ok = build_for_arch(
            target_arch,
            onedir=args.onedir,
            dist_dir=dist_dir,
            python_exe=python_exe,
            checker=args.checker,
            env_file=args.env_file,
            embed_env=not args.no_embed_env,
        )
        if not ok:
            all_ok = False

    if all_ok:
        clean_build_artifacts()
        if len(archs) == 2:
            n32 = f"{product}_x32.exe"
            n64 = f"{product}_x64.exe"
            print(
                "\nBoth builds succeeded:\n"
                f"  {os.path.join(dist_dir, n32)}\n"
                f"  {os.path.join(dist_dir, n64)}"
            )
    return 0 if all_ok else 1


if __name__ == "__main__":
    sys.exit(main())
