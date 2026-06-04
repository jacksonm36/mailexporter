#!/usr/bin/env python3
"""
Security test script for Mail Exporter
Tests path traversal protection
"""

from __future__ import annotations

import os
import shutil
import tempfile
import unittest

from path_security import PathSecurityError, PathValidator


class TestPathSecurity(unittest.TestCase):
    def setUp(self):
        """Create temporary test directories"""
        self.temp_dir = tempfile.mkdtemp()
        self.validator = PathValidator()

        self.test_file = os.path.join(self.temp_dir, "test.eml")
        with open(self.test_file, "w", encoding="utf-8") as handle:
            handle.write("Test content")

    def test_normal_path(self):
        """Test normal path validation"""
        path = "test.eml"
        result = self.validator.sanitize_path(path, self.temp_dir)
        self.assertTrue(os.path.samefile(result, self.test_file))

    def test_path_traversal_detection(self):
        """Test that path traversal is blocked"""
        attacks = [
            "../etc/passwd",
            "..\\windows\\system32",
            ".../.../.../etc/shadow",
            "%2e%2e%2f",  # URL encoded ../
            "..;/etc/passwd",
            "....//....//etc/passwd",
            "\\..\\..\\windows",
            "folder/../../etc/passwd",
        ]

        for attack in attacks:
            with self.subTest(attack=attack):
                with self.assertRaises(PathSecurityError):
                    self.validator.sanitize_path(attack, self.temp_dir)

    def test_null_byte_injection(self):
        """Test null byte injection protection"""
        attack = "test\0.exe"
        result = self.validator.sanitize_path(attack, self.temp_dir)
        self.assertNotIn("\0", result)
        self.assertTrue(result.endswith(".exe"))
        self.assertTrue(
            self.validator._path_within_base(
                os.path.realpath(self.temp_dir), result
            )
        )

    def test_extension_validation(self):
        """Test file extension validation"""
        self.assertTrue(self.validator.validate_file_extension("test.eml"))

        dangerous_exts = [".exe", ".bat", ".ps1", ".vbs", ".js", ".scr"]
        for ext in dangerous_exts:
            with self.subTest(ext=ext):
                self.assertFalse(
                    self.validator.validate_file_extension(f"malicious{ext}")
                )

    def test_safe_filename_creation(self):
        """Test filename sanitization"""
        cases = [
            ("../../../config", "config"),
            ("file;rm -rf /", "file_rm -rf"),
            ("malicious.bat", "malicious.bat"),
            ("report<1>.eml", "report_1_.eml"),
        ]
        for dangerous, expected in cases:
            with self.subTest(dangerous=dangerous):
                safe = self.validator.create_safe_filename(dangerous)
                self.assertEqual(safe, expected)

        long_name = "a" * 300 + ".eml"
        safe = self.validator.create_safe_filename(long_name, max_length=255)
        self.assertLessEqual(len(safe), 255)
        self.assertTrue(safe.endswith(".eml"))

    def test_directory_scan_safe(self):
        """Test safe directory scanning"""
        nested_dir = os.path.join(self.temp_dir, "level1", "level2")
        os.makedirs(nested_dir)

        safe_file = os.path.join(nested_dir, "safe.eml")
        with open(safe_file, "w", encoding="utf-8") as handle:
            handle.write("test")

        files = self.validator.scan_directory_safe(self.temp_dir, "*.eml")
        self.assertIn(os.path.realpath(safe_file), [os.path.realpath(p) for p in files])
        self.assertIn(os.path.realpath(self.test_file), [os.path.realpath(p) for p in files])

        if hasattr(os, "symlink"):
            try:
                link_path = os.path.join(self.temp_dir, "escape.link")
                target = "/etc/passwd" if os.name != "nt" else "C:\\Windows\\System32"
                os.symlink(target, link_path)
                files = self.validator.scan_directory_safe(self.temp_dir)
                resolved_outside = [
                    p
                    for p in files
                    if not self.validator._path_within_base(
                        os.path.realpath(self.temp_dir), os.path.realpath(p)
                    )
                ]
                self.assertEqual(resolved_outside, [])
            except (OSError, NotImplementedError):
                pass

    def test_absolute_path_under_base(self):
        """Absolute paths are allowed when still under base_dir"""
        result = self.validator.sanitize_path(self.test_file, self.temp_dir)
        self.assertTrue(os.path.samefile(result, self.test_file))

    def tearDown(self):
        shutil.rmtree(self.temp_dir, ignore_errors=True)


def run_security_tests() -> int:
    """Run all security tests"""
    suite = unittest.TestLoader().loadTestsFromTestCase(TestPathSecurity)
    runner = unittest.TextTestRunner(verbosity=2)
    result = runner.run(suite)

    if result.wasSuccessful():
        print("\nAll security tests passed.")
        return 0
    print("\nSecurity tests failed.")
    return 1


if __name__ == "__main__":
    raise SystemExit(run_security_tests())
