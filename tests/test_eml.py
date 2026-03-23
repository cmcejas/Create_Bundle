"""Unit tests for .eml parsing (no Word/Outlook required)."""
import os
import tempfile
import unittest

# Import from project root
import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from bundle_script import _eml_to_html, _eml_body_to_html_fragment
from email import policy
from email.parser import BytesParser


SAMPLE_PLAIN_EML = b"""\
From: Alice Example <alice@example.com>
To: bob@example.com
Subject: Plain test subject
Date: Mon, 15 Jan 2024 14:30:00 +0000
MIME-Version: 1.0
Content-Type: text/plain; charset=utf-8

Line one
Line two with <tags> & ampersands
"""

SAMPLE_HTML_EML = b"""\
From: =?utf-8?B?0J3QsNC30LDQvdC+0LI=?= <html-sender@example.org>
To: recipient@example.org
Subject: =?utf-8?B?SFRNTCBib2R5IHRlc3Q=?=
Date: Wed, 06 Mar 2024 09:15:22 -0500
MIME-Version: 1.0
Content-Type: multipart/alternative; boundary="bound123"

--bound123
Content-Type: text/plain; charset=utf-8

Plain fallback
--bound123
Content-Type: text/html; charset=utf-8

<html><body><p>Hello <b>HTML</b></p><style>div { margin: 1px; }</style></body></html>
--bound123--
"""


class EmlToHtmlTests(unittest.TestCase):
    def test_plain_eml_writes_expected_fields(self):
        with tempfile.TemporaryDirectory() as td:
            path = os.path.join(td, "test.eml")
            out = os.path.join(td, "out.html")
            with open(path, "wb") as f:
                f.write(SAMPLE_PLAIN_EML)
            _eml_to_html(path, out)
            with open(out, encoding="utf-8") as f:
                html = f.read()
        self.assertIn("From:", html)
        self.assertIn("alice@example.com", html)
        self.assertIn("To:", html)
        self.assertIn("bob@example.com", html)
        self.assertIn("Subject:", html)
        self.assertIn("Plain test subject", html)
        self.assertIn("Date:", html)
        self.assertIn("2024", html)
        self.assertIn("Line one", html)
        self.assertIn("&lt;tags&gt;", html)

    def test_html_eml_prefers_html_and_brace_css_safe(self):
        """Curly braces in embedded CSS must not break template formatting."""
        with tempfile.TemporaryDirectory() as td:
            path = os.path.join(td, "test.eml")
            out = os.path.join(td, "out.html")
            with open(path, "wb") as f:
                f.write(SAMPLE_HTML_EML)
            _eml_to_html(path, out)
            with open(out, encoding="utf-8") as f:
                html = f.read()
        self.assertIn("Hello", html)
        self.assertIn("HTML", html)
        self.assertIn("margin: 1px", html)

    def test_eml_body_fallback_walk(self):
        msg = BytesParser(policy=policy.default).parsebytes(SAMPLE_HTML_EML)
        frag = _eml_body_to_html_fragment(msg)
        self.assertIn("Hello", frag)
        self.assertNotIn("<html>", frag.lower() or "")


if __name__ == "__main__":
    unittest.main()
