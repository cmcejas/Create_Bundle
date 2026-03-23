"""Tests for Word COM path normalization (no Word required)."""
import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bundle_script import _word_com_path


class WordComPathTests(unittest.TestCase):
    def test_forward_slashes_become_native(self):
        if sys.platform != 'win32':
            self.skipTest("Windows-specific path shape")
        # Same issue as Word error C:\\//Users/... when given C:/Users/...
        p = _word_com_path("C:/Users/Someone/Documents/file.docx")
        self.assertTrue(os.path.isabs(p))
        self.assertNotIn("//", p.replace("\\\\", ""))  # no accidental UNC
        self.assertTrue(p.startswith("C:\\") or p.startswith("c:\\"))

    def test_abspath_expands_relative(self):
        base = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        p = _word_com_path(os.path.join(base, "bundle_script.py"))
        self.assertTrue(os.path.isfile(p))


if __name__ == "__main__":
    unittest.main()
