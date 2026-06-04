"""
Path Security Module for Mail Exporter
Prevents directory traversal and path injection attacks
"""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import List, Optional, Union

class PathSecurityError(Exception):
    """Custom exception for path security violations"""

    pass


class PathValidator:
    """Validates and sanitizes file paths to prevent traversal attacks"""

    # Dangerous patterns to detect (relative / untrusted segments)
    DANGEROUS_PATTERNS = [
        r"\.\.[/\\]",  # Parent directory traversal
        r"~[/\\]",  # Home directory
        r"%[0-9A-Fa-f]{2}",  # URL encoded characters
        r"\$[A-Za-z0-9_]+",  # Environment variables
        r"[;|&`$<>]",  # Command injection chars
        r"\\\\[^\\]+\\[^\\]+",  # UNC paths (\\server\share)
    ]

    # Allowed extensions
    ALLOWED_EXTENSIONS = {
        ".eml",
        ".emlx",
        ".pst",
        ".csv",
        ".log",
        ".txt",
        ".db",
        ".sqlite3",
    }

    @classmethod
    def _path_within_base(cls, base_real: str, target_real: str) -> bool:
        if os.path.normcase(target_real) == os.path.normcase(base_real):
            return True
        prefix = base_real.rstrip(os.sep) + os.sep
        if os.name == "nt":
            return os.path.normcase(target_real).startswith(os.path.normcase(prefix))
        return target_real.startswith(prefix)

    @classmethod
    def _check_dangerous_patterns(cls, user_input: str) -> None:
        for pattern in cls.DANGEROUS_PATTERNS:
            if re.search(pattern, user_input, re.IGNORECASE):
                raise PathSecurityError(f"Dangerous pattern detected: {pattern}")

    @classmethod
    def sanitize_path(
        cls,
        user_input: str,
        base_dir: Union[str, Path],
        *,
        allow_absolute: bool = True,
    ) -> str:
        """
        Sanitize and validate a user-provided path.

        Args:
            user_input: Raw path input from user
            base_dir: Base directory that path must stay under
            allow_absolute: When True, absolute paths are allowed if under base_dir

        Returns:
            Sanitized absolute path

        Raises:
            PathSecurityError: If path is dangerous or escapes base_dir
        """
        user_input = str(user_input or "").strip()
        base_dir = str(base_dir or "").strip()
        if not base_dir:
            raise PathSecurityError("Missing base directory")
        if not user_input:
            raise PathSecurityError("Missing path")

        if "\0" in user_input:
            raise PathSecurityError("Null byte in path")
        user_input = user_input.replace("\0", "")
        cls._check_dangerous_patterns(user_input)

        base_real = os.path.realpath(base_dir)

        if os.path.isabs(user_input):
            if not allow_absolute:
                raise PathSecurityError("Absolute paths are not allowed")
            full_path = os.path.realpath(user_input)
        else:
            rel = user_input.replace("/", os.sep).replace("\\", os.sep)
            rel = re.sub(r"^[A-Za-z]:[/\\]", "", rel)
            rel = rel.lstrip(os.sep)
            full_path = os.path.realpath(os.path.join(base_real, rel))

        if not cls._path_within_base(base_real, full_path):
            raise PathSecurityError(f"Path traversal attempt: {user_input}")
        return full_path

    @classmethod
    def validate_file_extension(
        cls, filepath: str, allowed_extensions: Optional[set] = None
    ) -> bool:
        if allowed_extensions is None:
            allowed_extensions = cls.ALLOWED_EXTENSIONS
        ext = os.path.splitext(filepath)[1].lower()
        return ext in allowed_extensions

    @classmethod
    def scan_directory_safe(
        cls, base_dir: Union[str, Path], pattern: str = "*.eml"
    ) -> List[str]:
        base_dir = str(base_dir)
        if not os.path.isdir(base_dir):
            raise PathSecurityError(f"Not a directory: {base_dir}")

        safe_files: List[str] = []
        base_path = Path(base_dir).resolve()

        try:
            for file_path in base_path.rglob(pattern):
                try:
                    resolved = file_path.resolve()
                    if cls._path_within_base(str(base_path), str(resolved)):
                        safe_files.append(str(resolved))
                except (OSError, RuntimeError):
                    continue
            return safe_files
        except Exception as e:
            raise PathSecurityError(f"Error scanning directory: {e}") from e

    @classmethod
    def create_safe_filename(cls, original_name: str, max_length: int = 255) -> str:
        raw = str(original_name or "")
        parts = [p for p in re.split(r"[/\\]+", raw) if p]
        name = parts[-1] if parts else raw
        name = re.sub(r"[/\\:*?\"<>|;]", "_", name)
        name = re.sub(r"[\x00-\x1f\x7f]", "", name)
        name = name.strip(". ")
        if len(name) > max_length:
            name_base, ext = os.path.splitext(name)
            keep = max(1, max_length - len(ext) - 1)
            name = name_base[:keep] + "_" + ext
        return name if name else "unnamed_file"


def secure_file_operation(func):
    """Decorator to wrap file operations with path validation."""

    def wrapper(self, filepath, *args, **kwargs):
        if hasattr(self, "base_dir"):
            try:
                validated_path = PathValidator.sanitize_path(filepath, self.base_dir)
                return func(self, validated_path, *args, **kwargs)
            except PathSecurityError as e:
                if hasattr(self, "log_security_error"):
                    self.log_security_error(f"Blocked file access: {e}")
                raise
        return func(self, filepath, *args, **kwargs)

    return wrapper
