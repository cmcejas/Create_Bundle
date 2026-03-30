"""
Create_Bundle – PDF Bundler (GUI)
Merges PDFs, Word docs, and Outlook .msg emails into a single PDF.
"""

import os
import sys
import re
import html as html_mod
import shutil
import tempfile
import datetime
import time
import traceback
import threading
import queue as queue_mod

import customtkinter as ctk
from tkinter import filedialog, messagebox

# ── Constants ────────────────────────────────────────────────────────────────

RETRY_ATTEMPTS = 3
RETRY_DELAY_BASE = 2
DOC_READY_POLL = 0.3            # interval when polling for doc ready
DOC_READY_TIMEOUT = 8.0         # max seconds to wait for doc to be ready
FILE_STABLE_POLL = 0.2          # interval when checking output file
FILE_STABLE_CHECKS = 2          # unchanged size checks before "stable"

SUPPORTED_EXT = {'.pdf', '.doc', '.docx', '.msg'}

EMAIL_HTML_TEMPLATE = """\
<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
  body {{
    font-family: Calibri, Arial, Helvetica, sans-serif;
    font-size: 11pt;
    color: #222;
    margin: 50px 60px;
    line-height: 1.5;
  }}
  .email-header {{
    border-bottom: 2px solid #b0b0b0;
    padding-bottom: 12px;
    margin-bottom: 20px;
  }}
  .header-row {{
    margin: 3px 0;
  }}
  .header-label {{
    font-weight: bold;
    color: #444;
    display: inline-block;
    width: 75px;
  }}
  .email-body {{
    margin-top: 10px;
  }}
</style>
</head>
<body>
<div class="email-header">
  <div class="header-row"><span class="header-label">From:</span> {from_field}</div>
  <div class="header-row"><span class="header-label">To:</span> {to_field}</div>
  <div class="header-row"><span class="header-label">Date:</span> {date_field}</div>
  <div class="header-row"><span class="header-label">Subject:</span> {subject_field}</div>
</div>
<div class="email-body">
{body_html}
</div>
</body>
</html>
"""


def _base_dir():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


# ── System checks ────────────────────────────────────────────────────────────

def _check_word_available():
    """
    Check if Microsoft Word is installed and available via COM.
    Returns True if available, False otherwise.
    """
    try:
        import comtypes.client
        app = comtypes.client.CreateObject('Word.Application')
        app.Quit()
        return True
    except Exception:
        return False


def _show_word_missing_dialog():
    """Show a user-friendly dialog explaining Word is required."""
    # Delay import to avoid early tkinter initialization
    from tkinter import messagebox as msg
    msg.showwarning(
        "Microsoft Word Required",
        "Microsoft Word is required to convert Word documents and emails to PDF.\n\n"
        "Please install Microsoft Word from:\n"
        "  https://www.microsoft.com/office\n\n"
        "You can still work with PDF files without Word.\n"
        "The app will continue, but Word-dependent files will fail to process."
    )


def _check_system_requirements():
    """
    Perform startup checks and warn user if requirements are missing.
    Run this after Tkinter is initialized but before showing main window.
    """
    if sys.platform != 'win32':
        from tkinter import messagebox as msg
        msg.showerror(
            "Windows Only",
            "This app is designed for Windows only.\n\n"
            f"Detected OS: {sys.platform}"
        )
        sys.exit(1)
    
    word_available = _check_word_available()
    if not word_available:
        _show_word_missing_dialog()


# ── Processing helpers ───────────────────────────────────────────────────────

def _wait_file_stable(filepath, timeout=6.0):
    """Block until *filepath* exists and its size stops changing."""
    deadline = time.monotonic() + timeout
    last_size = -1
    ok = 0
    while time.monotonic() < deadline:
        try:
            if os.path.isfile(filepath):
                sz = os.path.getsize(filepath)
                if sz == last_size and sz > 0:
                    ok += 1
                    if ok >= FILE_STABLE_CHECKS:
                        return
                else:
                    ok = 0
                    last_size = sz
        except Exception:
            ok = 0
        time.sleep(FILE_STABLE_POLL)


def _wait_doc_ready(doc):
    """Poll until Word finishes loading/laying out the document."""
    deadline = time.monotonic() + DOC_READY_TIMEOUT
    while time.monotonic() < deadline:
        try:
            _ = doc.Content.Start
            return
        except Exception:
            time.sleep(DOC_READY_POLL)


def _extract_html_body(html_source):
    if not html_source:
        return ""
    match = re.search(r'<body[^>]*>(.*)</body>', html_source,
                      re.DOTALL | re.IGNORECASE)
    return match.group(1) if match else html_source


# Hardcoded English month names — strftime("%B") is locale-dependent and
# produces non-English names on French, German, etc. Windows installs.
_MONTH_NAMES = (
    None, 'January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December',
)


def _safe_strftime(dt):
    """Format a datetime as 'DD Month YYYY HH:MM' using hardcoded English months."""
    return (f"{dt.day:02d} {_MONTH_NAMES[dt.month]} "
            f"{dt.year:04d} {dt.hour:02d}:{dt.minute:02d}")


def _format_email_datetime(dt):
    """Format Outlook COM date as Day Month Year HH:MM (e.g. 06 May 2025 16:17).

    Uses hardcoded English month names so the output is identical regardless
    of the Windows display-language / locale on the machine.
    """
    if dt is None:
        return "Unknown"
    # Python datetime or pywintypes.datetime
    try:
        return _safe_strftime(dt)
    except Exception:
        pass
    # OLE Automation date (float: days since 30 Dec 1899)
    try:
        if isinstance(dt, (int, float)):
            base = datetime.datetime(1899, 12, 30)
            d = base + datetime.timedelta(days=float(dt))
            return _safe_strftime(d)
    except Exception:
        pass
    # COM-style attributes (Year, Month, Day, Hour, Minute)
    try:
        y = getattr(dt, 'Year', None) or getattr(dt, 'year', None)
        m = getattr(dt, 'Month', None) or getattr(dt, 'month', None)
        d = getattr(dt, 'Day', None) or getattr(dt, 'day', None)
        hr = getattr(dt, 'Hour', None) or getattr(dt, 'hour', 0)
        mn = getattr(dt, 'Minute', None) or getattr(dt, 'minute', 0)
        if y is not None and m is not None and d is not None and y < 4000:
            month_name = _MONTH_NAMES[int(m)] if 1 <= int(m) <= 12 else str(m)
            return f"{int(d):02d} {month_name} {int(y):04d} {int(hr):02d}:{int(mn):02d}"
    except Exception:
        pass
    return str(dt)


# Expected pattern: "DD Month YYYY HH:MM"  e.g. "06 May 2025 16:17"
_DATE_PATTERN = re.compile(
    r'^\d{2} (?:' + '|'.join(_MONTH_NAMES[1:]) + r') \d{4} \d{2}:\d{2}$'
)


def _validate_date(date_str, filepath, source_label):
    """Check the extracted date matches 'DD Month YYYY HH:MM'.

    Returns None if valid, or a warning string if the date looks wrong.
    The caller is responsible for logging the warning.
    """
    filename = os.path.basename(filepath)
    if date_str == "Unknown":
        return (f"⚠ DATE WARNING [{filename}]: No date found "
                f"(tried: {source_label})")
    if not _DATE_PATTERN.match(date_str):
        return (f"⚠ DATE WARNING [{filename}]: Unexpected format '{date_str}' "
                f"(source: {source_label}) — expected 'DD Month YYYY HH:MM'")
    return None


def _word_com_path(path):
    """
    Absolute, normalized path for Word COM on Windows.

    Forward-slash paths (e.g. C:/Users/...) are valid in Python but Word often
    mis-resolves them as C:\\ + //server/... and fails with "couldn't find your file".

    Note: Tkinter's filedialog.askdirectory() on Windows typically returns paths
    with forward slashes, so this applies even when the user picks a folder in the GUI.
    """
    if not path:
        return path
    path = os.path.abspath(os.path.expanduser(str(path).strip()))
    path = os.path.normpath(path)
    if sys.platform == 'win32':
        path = path.replace('/', '\\')
    return path


# Formats that str(pywintypes.datetime) can produce depending on the
# Windows locale / pywin32 version.  We try each in order.
_PYWINTYPES_FORMATS = (
    "%Y-%m-%d %H:%M:%S",       # ISO  (most common)
    "%m/%d/%Y %H:%M:%S",       # US locale  24h
    "%d/%m/%Y %H:%M:%S",       # UK / EU locale 24h
    "%Y-%m-%dT%H:%M:%S",       # ISO with T separator
    "%d-%m-%Y %H:%M:%S",       # dash-separated EU
    "%m-%d-%Y %H:%M:%S",       # dash-separated US
    "%m/%d/%y %H:%M:%S",       # US locale short year
    "%d/%m/%y %H:%M:%S",       # EU locale short year
)

# Some locales produce 12-hour times with AM/PM — handled separately below.
_PYWINTYPES_FORMATS_AMPM = (
    "%m/%d/%Y %I:%M:%S %p",    # US 12-hour
    "%d/%m/%Y %I:%M:%S %p",    # UK 12-hour
    "%Y-%m-%d %I:%M:%S %p",    # ISO 12-hour (rare)
)

# Years that represent "not set" / null OLE dates
_BOGUS_YEARS = {1601, 1899, 1900, 4501}


def _is_plausible_date(dt):
    """Return True if *dt* looks like a real email date, not a null/epoch placeholder."""
    if dt is None:
        return False
    return 1970 <= dt.year <= 2100 and dt.year not in _BOGUS_YEARS


def _parse_pywintypes_date(val):
    """
    Reliably convert a pywintypes.datetime COM object to a plain Python datetime.

    str(pywintypes.datetime) output varies by Windows locale:
      - English/ISO:  "2025-05-06 16:17:00+00:00"
      - US locale:    "05/06/2025 16:17:00"
      - UK/EU locale: "06/05/2025 16:17:00"
      - US 12-hour:   "5/6/2025 4:17:00 PM"

    We try every known format so the date is parsed correctly on any machine.
    """
    if val is None:
        return None

    # 1) If it's already a Python datetime (pywintypes.datetime is a subclass),
    #    convert timezone-aware to local time, then strip tzinfo.
    if isinstance(val, datetime.datetime):
        try:
            dt = val
            if dt.tzinfo is not None:
                dt = dt.astimezone().replace(tzinfo=None)
            if _is_plausible_date(dt):
                return dt
        except (OverflowError, OSError, Exception):
            pass

    # 2) Direct attribute access — works on genuine pywintypes.datetime
    try:
        dt = datetime.datetime(
            year=val.year, month=val.month, day=val.day,
            hour=val.hour, minute=val.minute, second=val.second)
        if _is_plausible_date(dt):
            return dt
    except Exception:
        pass

    # 3) timetuple() — more robust on some pywin32 builds where .year etc throw
    try:
        tt = val.timetuple()
        dt = datetime.datetime(*tt[:6])
        if _is_plausible_date(dt):
            return dt
    except Exception:
        pass

    # 4) OLE Automation float (days since 30 Dec 1899)
    try:
        if isinstance(val, (int, float)):
            base = datetime.datetime(1899, 12, 30)
            dt = base + datetime.timedelta(days=float(val))
            if _is_plausible_date(dt):
                return dt
    except Exception:
        pass

    # 5) Parse string representation
    raw_full = str(val).strip()
    raw = raw_full[:19]  # drop timezone suffix for 24-hour formats

    for fmt in _PYWINTYPES_FORMATS:
        try:
            dt = datetime.datetime.strptime(raw, fmt)
            if _is_plausible_date(dt):
                return dt
        except (ValueError, TypeError):
            continue

    # 6) AM/PM formats need the full string (not truncated to 19 chars)
    #    Strip timezone suffix like "+00:00" but keep AM/PM
    raw_notz = re.sub(r'[+-]\d{2}:\d{2}$', '', raw_full).strip()
    for fmt in _PYWINTYPES_FORMATS_AMPM:
        try:
            dt = datetime.datetime.strptime(raw_notz, fmt)
            if _is_plausible_date(dt):
                return dt
        except (ValueError, TypeError):
            continue

    return None


def _parse_date_from_transport_headers(msg_com):
    """Extract Date from the raw SMTP transport headers stored in the .msg.

    Outlook stores the original RFC 2822 headers as PR_TRANSPORT_MESSAGE_HEADERS.
    This is the most reliable date source for sent items / first-in-chain emails
    where SentOn / ReceivedTime can be empty.
    """
    from email.utils import parsedate_to_datetime

    try:
        headers = msg_com.PropertyAccessor.GetProperty(
            "http://schemas.microsoft.com/mapi/proptag/0x007D001F")
    except Exception:
        return None
    if not headers:
        return None
    match = re.search(r'^Date:\s*(.+)$', headers, re.MULTILINE | re.IGNORECASE)
    if not match:
        return None
    try:
        dt = parsedate_to_datetime(match.group(1).strip())
        if dt is not None and dt.tzinfo is not None:
            dt = dt.astimezone().replace(tzinfo=None)
        if _is_plausible_date(dt):
            return dt
    except Exception:
        pass
    return None


# MAPI property tags that store dates — used as fallback when the COM model
# properties (SentOn, ReceivedTime, …) return null/bogus values.
_MAPI_DATE_TAGS = (
    # (display label, MAPI proptag URL)
    ('PR_CLIENT_SUBMIT_TIME',
     'http://schemas.microsoft.com/mapi/proptag/0x00390040'),
    ('PR_MESSAGE_DELIVERY_TIME',
     'http://schemas.microsoft.com/mapi/proptag/0x0E060040'),
    ('PR_CREATION_TIME',
     'http://schemas.microsoft.com/mapi/proptag/0x30070040'),
    ('PR_LAST_MODIFICATION_TIME',
     'http://schemas.microsoft.com/mapi/proptag/0x30080040'),
)


def _parse_date_from_mapi_props(msg_com):
    """Try raw MAPI property tags via PropertyAccessor.

    These can succeed when the higher-level model properties (SentOn etc.) fail,
    especially for sent items and calendar-originated messages.
    Returns (datetime, source_label) or (None, None).
    """
    for label, tag in _MAPI_DATE_TAGS:
        try:
            val = msg_com.PropertyAccessor.GetProperty(tag)
            if val is None:
                continue
            dt = _parse_pywintypes_date(val)
            if dt is not None:
                return dt, label
        except Exception:
            continue
    return None, None


# ── Reusable COM wrapper (one Word instance for the entire run) ──────────────

class _WordApp:
    """Manages a single Word COM instance, reused across all conversions."""

    def __init__(self):
        self._app = None

    def _ensure(self):
        if self._app is not None:
            try:
                _ = self._app.Visible
                return
            except Exception:
                self._app = None
        import comtypes.client
        self._app = comtypes.client.CreateObject('Word.Application')
        self._app.Visible = False

    def open_and_save_pdf(self, src_path, pdf_path, log_fn):
        """Open *src_path* in Word, save as PDF, close the document."""
        self._ensure()
        src_path = _word_com_path(src_path)
        pdf_path = _word_com_path(pdf_path)
        doc = self._open(src_path, log_fn)
        try:
            self._save_pdf(doc, pdf_path, log_fn)
        finally:
            try:
                doc.Close(False)
            except Exception:
                pass

    def quit(self):
        if self._app is not None:
            try:
                self._app.Quit()
            except Exception:
                pass
            self._app = None

    def _open(self, path, log_fn):
        path = _word_com_path(path)
        for attempt in range(RETRY_ATTEMPTS):
            try:
                self._ensure()
                if not os.path.isfile(path):
                    raise FileNotFoundError(f"Word cannot open missing file: {path}")
                doc = self._app.Documents.Open(path)
                if doc is None:
                    raise RuntimeError("Word returned None")
                _wait_doc_ready(doc)
                return doc
            except Exception:
                self._app = None
                delay = RETRY_DELAY_BASE * (attempt + 1)
                if attempt < RETRY_ATTEMPTS - 1:
                    log_fn(f"    Retry {attempt + 1}/{RETRY_ATTEMPTS}: "
                           f"Word open failed, waiting {delay}s ...")
                    time.sleep(delay)
                else:
                    raise

    def _save_pdf(self, doc, pdf_path, log_fn):
        for attempt in range(RETRY_ATTEMPTS):
            try:
                doc.SaveAs(pdf_path, FileFormat=17)
                _wait_file_stable(pdf_path)
                if os.path.isfile(pdf_path) and os.path.getsize(pdf_path) > 0:
                    return
                raise RuntimeError("PDF missing or empty after SaveAs")
            except Exception:
                delay = RETRY_DELAY_BASE * (attempt + 1)
                if attempt < RETRY_ATTEMPTS - 1:
                    log_fn(f"    Retry {attempt + 1}/{RETRY_ATTEMPTS}: "
                           f"PDF save failed, waiting {delay}s ...")
                    time.sleep(delay)
                else:
                    raise


def _finalize_email_html(from_display, to_field_plain, date_str_plain, subject_plain,
                         body_content, out_html_path):
    """Write EMAIL_HTML_TEMPLATE; escape braces so str.format() never breaks on CSS in bodies."""
    def _safe(s):
        if s is None:
            s = ""
        return str(s).replace('{', '{{').replace('}', '}}')

    to_esc = html_mod.escape(to_field_plain) if to_field_plain else "Unknown"
    date_esc = html_mod.escape(date_str_plain) if date_str_plain else "Unknown"
    subj_esc = html_mod.escape(subject_plain) if subject_plain else "(No Subject)"

    full_html = EMAIL_HTML_TEMPLATE.format(
        from_field=_safe(from_display),
        to_field=_safe(to_esc),
        date_field=_safe(date_esc),
        subject_field=_safe(subj_esc),
        body_html=_safe(body_content),
    )
    with open(out_html_path, 'w', encoding='utf-8') as f:
        f.write(full_html)


def _msg_to_html(src_path, out_html_path):
    """Extract .msg metadata via Outlook COM and write a clean HTML file."""
    import win32com.client

    outlook = win32com.client.Dispatch('Outlook.Application')
    # OpenSharedItem preserves the original sent/received dates;
    # CreateItemFromTemplate creates an unsent draft and loses them.
    abs_path = _word_com_path(src_path)
    msg = outlook.Session.OpenSharedItem(abs_path)

    subject = msg.Subject or "(No Subject)"
    sender_name = msg.SenderName or ""
    sender_email = msg.SenderEmailAddress or ""
    to_field = msg.To or ""

    if sender_email.startswith('/'):
        sender_email = ""

    if sender_email and sender_name:
        from_display = (f"{html_mod.escape(sender_name)} "
                        f"&lt;{html_mod.escape(sender_email)}&gt;")
    elif sender_name:
        from_display = html_mod.escape(sender_name)
    elif sender_email:
        from_display = html_mod.escape(sender_email)
    else:
        from_display = "Unknown"

    date_str = "Unknown"
    date_source = None
    for attr in ('SentOn', 'ReceivedTime', 'CreationTime', 'LastModificationTime'):
        try:
            val = getattr(msg, attr, None)
            if val is None:
                continue
            dt = _parse_pywintypes_date(val)
            if dt is None:
                continue
            date_str = _safe_strftime(dt)
            date_source = attr
            break
        except Exception:
            continue

    # Fallback: parse date from raw transport headers (most reliable for
    # sent items / first email in a chain where COM properties can be empty)
    if date_str == "Unknown":
        dt = _parse_date_from_transport_headers(msg)
        if dt is not None:
            date_str = _safe_strftime(dt)
            date_source = 'TransportHeaders'

    # Fallback: try raw MAPI property tags via PropertyAccessor
    if date_str == "Unknown":
        dt, mapi_label = _parse_date_from_mapi_props(msg)
        if dt is not None:
            date_str = _safe_strftime(dt)
            date_source = mapi_label

    # Last resort: use the raw str() of whichever COM date property has a value.
    # This may look ugly (e.g. "2025-05-06 16:17:00+00:00") but guarantees
    # the date field in the PDF is never blank.
    if date_str == "Unknown":
        for attr in ('SentOn', 'ReceivedTime', 'CreationTime', 'LastModificationTime'):
            try:
                val = getattr(msg, attr, None)
                if val is not None:
                    raw = str(val).strip()
                    if raw and raw != 'None':
                        date_str = raw
                        date_source = f'{attr} (raw)'
                        break
            except Exception:
                continue

    body_html = None
    try:
        body_html = msg.HTMLBody
    except Exception:
        pass

    if body_html:
        body_content = _extract_html_body(body_html)
    else:
        plain_body = msg.Body or ""
        body_content = html_mod.escape(plain_body).replace('\n', '<br>\n')

    # Close the msg item so Outlook doesn't accumulate open items
    try:
        msg.Close(1)  # olDiscard
    except Exception:
        pass

    date_warning = _validate_date(date_str, src_path,
                                    date_source or 'SentOn/ReceivedTime/CreationTime')
    _finalize_email_html(
        from_display, to_field, date_str, subject, body_content, out_html_path)
    return date_warning


def _msg_to_html_fallback(src_path, out_html_path):
    """Parse .msg using extract-msg (pure Python) when Outlook is not installed."""
    import extract_msg

    msg = extract_msg.Message(src_path)
    try:
        subject = msg.subject or "(No Subject)"
        sender_name = msg.senderName or ""
        sender_email = msg.sender or ""
        # extract-msg stores recipients as a string
        to_field = msg.to or ""

        if sender_email.startswith('/'):
            sender_email = ""

        if sender_email and sender_name:
            from_display = (f"{html_mod.escape(sender_name)} "
                            f"&lt;{html_mod.escape(sender_email)}&gt;")
        elif sender_name:
            from_display = html_mod.escape(sender_name)
        elif sender_email:
            from_display = html_mod.escape(sender_email)
        else:
            from_display = "Unknown"

        # extract-msg exposes .date as a string like "Mon, 6 May 2025 16:17:00 +0100"
        date_str = "Unknown"
        date_source = None
        raw_date = msg.date
        if raw_date:
            try:
                from email.utils import parsedate_to_datetime
                dt = parsedate_to_datetime(str(raw_date))
                if dt is not None and dt.tzinfo is not None:
                    dt = dt.astimezone().replace(tzinfo=None)
                if _is_plausible_date(dt):
                    date_str = _safe_strftime(dt)
                    date_source = 'msg.date'
            except Exception:
                date_str = str(raw_date).strip() or "Unknown"
                date_source = 'msg.date (raw)'

        # Fallback: try transport headers embedded in the .msg
        if date_str == "Unknown":
            try:
                headers = msg.headerDict or {}
                hdr_date = headers.get('Date') or headers.get('date')
                if hdr_date:
                    from email.utils import parsedate_to_datetime
                    dt = parsedate_to_datetime(str(hdr_date))
                    if dt is not None and dt.tzinfo is not None:
                        dt = dt.astimezone().replace(tzinfo=None)
                    if _is_plausible_date(dt):
                        date_str = _safe_strftime(dt)
                        date_source = 'headerDict.Date'
            except Exception:
                pass

        # Last resort: use the raw .date string as-is
        if date_str == "Unknown" and raw_date:
            raw = str(raw_date).strip()
            if raw and raw != 'None':
                date_str = raw
                date_source = 'msg.date (raw fallback)'

        body_html = None
        try:
            body_html = msg.htmlBody
            if isinstance(body_html, bytes):
                body_html = body_html.decode('utf-8', errors='replace')
        except Exception:
            pass

        if body_html:
            body_content = _extract_html_body(body_html)
        else:
            plain_body = msg.body or ""
            body_content = html_mod.escape(plain_body).replace('\n', '<br>\n')

        date_warning = _validate_date(date_str, src_path,
                                      date_source or 'msg.date')
        _finalize_email_html(
            from_display, to_field, date_str, subject, body_content, out_html_path)
        return date_warning
    finally:
        msg.close()


def _get_page_limit(filename, rules):
    # Match against stem only (name without extension) so e.g. "msg" doesn't
    # match every .msg file and "email" only matches names containing "email".
    stem = os.path.splitext(filename)[0].lower()
    for keyword, pages in rules.items():
        if keyword in stem:
            return pages
    return None


# ── Background worker ────────────────────────────────────────────────────────

def _run_bundle(input_dir, output_dir, rules, msg_queue, cancel_event):
    """Runs in a background thread. Communicates via msg_queue."""
    import pythoncom
    pythoncom.CoInitialize()

    def log_fn(text):
        msg_queue.put(("log", text))

    word = _WordApp()

    try:
        from pypdf import PdfReader, PdfWriter

        entries = sorted(
            [f for f in os.listdir(input_dir)
             if os.path.isfile(os.path.join(input_dir, f))
             and os.path.splitext(f)[1].lower() in SUPPORTED_EXT],
            key=lambda x: x.lower(),
        )

        if not entries:
            log_fn("No supported files found in the input folder.")
            msg_queue.put(("done", False, None))
            return

        total = len(entries)
        needs_word = any(
            os.path.splitext(f)[1].lower() in ('.doc', '.docx', '.msg')
            for f in entries)

        if needs_word:
            log_fn("Starting Word (reused for all conversions) ...")
            try:
                word._ensure()
            except Exception as e:
                log_fn("❌ ERROR: Microsoft Word could not be started.")
                log_fn("   Word is required to convert .doc, .docx, .msg files to PDF.")
                log_fn("   Please ensure Microsoft Word is installed:")
                log_fn("   → https://www.microsoft.com/office")
                log_fn(f"   {e}")
                msg_queue.put(("done", False, None))
                return

        log_fn(f"Found {total} file(s) to process.\n")

        temp_dir = tempfile.mkdtemp(prefix='BundleTool_')
        writer = PdfWriter()
        processed = 0

        for idx, filename in enumerate(entries, 1):
            if cancel_event.is_set():
                log_fn("\nCancelled by user.")
                break

            src = os.path.join(input_dir, filename)
            ext = os.path.splitext(filename)[1].lower()
            page_limit = _get_page_limit(filename, rules)
            limit_label = f" [max {page_limit} page(s)]" if page_limit else ""

            log_fn(f"  [{idx}/{total}] {filename}{limit_label}")
            msg_queue.put(("progress", idx, total))

            try:
                if ext == '.pdf':
                    pdf_path = src
                elif ext in ('.doc', '.docx'):
                    pdf_path = os.path.join(temp_dir, f"conv_{idx}.pdf")
                    word.open_and_save_pdf(src, pdf_path, log_fn)
                elif ext == '.msg':
                    html_path = os.path.join(temp_dir, f"conv_{idx}.html")
                    pdf_path = os.path.join(temp_dir, f"conv_{idx}.pdf")
                    # Try Outlook COM first; fall back to extract-msg if unavailable
                    try:
                        date_warn = _msg_to_html(src, html_path)
                    except Exception:
                        log_fn("    Outlook unavailable, using extract-msg fallback ...")
                        date_warn = _msg_to_html_fallback(src, html_path)
                    if date_warn:
                        log_fn(f"    {date_warn}")
                    word.open_and_save_pdf(html_path, pdf_path, log_fn)
                    try:
                        os.remove(html_path)
                    except Exception:
                        pass
                else:
                    continue

                reader = PdfReader(pdf_path)
                num_pages = len(reader.pages)

                if page_limit and page_limit < num_pages:
                    for i in range(page_limit):
                        writer.add_page(reader.pages[i])
                    log_fn(f"    Added {page_limit} of {num_pages} page(s).")
                else:
                    for page in reader.pages:
                        writer.add_page(page)
                    log_fn(f"    Added {num_pages} page(s).")

                processed += 1
            except Exception as e:
                err_msg = str(e)
                # Provide user-friendly hints for common errors
                if "Couldn't find your file" in err_msg or "couldn't find" in err_msg.lower():
                    log_fn(f"    FAILED: File path issue (Word COM path encoding)")
                    log_fn(f"    Hint: Try moving file to a path without special characters")
                elif "Word" in err_msg or "ActiveX" in err_msg:
                    log_fn(f"    FAILED: Microsoft Word error")
                    log_fn(f"    Hint: Ensure Microsoft Word is properly installed")
                else:
                    log_fn(f"    FAILED: {e}")
                log_fn(traceback.format_exc())

        word.quit()

        try:
            shutil.rmtree(temp_dir, ignore_errors=True)
        except Exception:
            pass

        if cancel_event.is_set() or processed == 0:
            if processed == 0 and not cancel_event.is_set():
                log_fn("\nAll files failed. No output created.")
            msg_queue.put(("done", False, None))
            return

        os.makedirs(output_dir, exist_ok=True)
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        out_path = os.path.join(output_dir, f"Bundle_{timestamp}.pdf")

        with open(out_path, 'wb') as f:
            writer.write(f)

        log_fn(f"\nBundle saved to:\n  {out_path}")
        log_fn(f"  ({processed} of {total} files merged)")
        msg_queue.put(("done", True, out_path))

    except Exception as e:
        msg_queue.put(("log", f"\nFATAL ERROR: {e}"))
        msg_queue.put(("log", traceback.format_exc()))
        msg_queue.put(("done", False, None))
    finally:
        word.quit()
        pythoncom.CoUninitialize()


# ── Config file helpers ──────────────────────────────────────────────────────

def _load_config_file(config_path):
    rules = []
    if not os.path.isfile(config_path):
        return rules
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith('#'):
                    continue
                if '=' not in line:
                    continue
                keyword, _, pages_str = line.partition('=')
                keyword = keyword.strip().lower()
                try:
                    pages = int(pages_str.strip())
                    if pages >= 1:
                        rules.append((keyword, pages))
                except ValueError:
                    pass
    except Exception:
        pass
    return rules


def _save_config_file(config_path, rules_dict):
    lines = [
        "# Bundle Script - Page Rules (config.txt)",
        "# Format: keyword = number_of_pages",
        "# Files not matching any keyword include ALL pages.",
        "",
    ]
    for kw, pg in rules_dict.items():
        lines.append(f"{kw} = {pg}")
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines) + '\n')
    except Exception:
        pass


# ── GUI ──────────────────────────────────────────────────────────────────────

ctk.set_appearance_mode("system")
ctk.set_default_color_theme("blue")

SECTION_PAD = {"padx": 18, "pady": (10, 0)}
INNER_PAD_X = 14
GREEN = "#2fa572"
ORANGE = "#d4912a"
RED = "#d43a2a"


class RuleRow(ctk.CTkFrame):
    """One keyword / page-count rule inside the rules list."""

    def __init__(self, master, keyword="", pages="1", on_delete=None, **kw):
        super().__init__(master, fg_color="transparent", **kw)

        self.keyword_entry = ctk.CTkEntry(
            self, placeholder_text="e.g. email", width=200)
        self.keyword_entry.pack(side="left", padx=(0, 8), pady=2)
        if keyword:
            self.keyword_entry.insert(0, keyword)

        self.pages_entry = ctk.CTkEntry(
            self, placeholder_text="1", width=55, justify="center")
        self.pages_entry.pack(side="left", padx=(0, 8), pady=2)
        if pages:
            self.pages_entry.insert(0, str(pages))

        self.del_btn = ctk.CTkButton(
            self, text="X", width=30, height=28,
            fg_color="#c0392b", hover_color="#922b21",
            font=ctk.CTkFont(size=12, weight="bold"),
            command=lambda: on_delete(self) if on_delete else None)
        self.del_btn.pack(side="left", pady=2)

    def get_rule(self):
        kw = self.keyword_entry.get().strip().lower()
        pg = self.pages_entry.get().strip()
        if not kw or not pg:
            return None
        try:
            pg_int = int(pg)
            return (kw, pg_int) if pg_int >= 1 else None
        except ValueError:
            return None

    def set_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self.keyword_entry.configure(state=state)
        self.pages_entry.configure(state=state)
        self.del_btn.configure(state=state)


class BundleApp(ctk.CTk):

    def __init__(self):
        super().__init__()

        self.title("Create Bundle  -  PDF Bundler")
        self.geometry("780x860")
        self.minsize(660, 700)

        self._root_dir = _base_dir()
        self._rule_rows: list[RuleRow] = []
        self._worker: threading.Thread | None = None
        self._cancel = threading.Event()
        self._queue = queue_mod.Queue()
        self._log_lines: list[str] = []

        self._build_ui()
        self._load_defaults()
        self._center_window()
        _check_system_requirements()
        self._poll_queue()

    # ── build UI ──

    def _build_ui(self):
        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)

        self._build_header()
        self._build_folders_section()
        self._build_rules_section()
        self._build_log_section()
        self._build_actions()

    def _build_header(self):
        frm = ctk.CTkFrame(self, fg_color="transparent")
        frm.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 0))

        ctk.CTkLabel(
            frm, text="Create Bundle",
            font=ctk.CTkFont(size=24, weight="bold"),
        ).pack(anchor="w")
        ctk.CTkLabel(
            frm,
            text="Merge PDFs, Word docs, and Outlook .msg emails into one PDF.",
            font=ctk.CTkFont(size=13), text_color="gray",
        ).pack(anchor="w", pady=(2, 0))

        # System status
        word_status = "✓" if _check_word_available() else "⚠"
        word_color = GREEN if _check_word_available() else ORANGE
        status_frame = ctk.CTkFrame(frm, fg_color="transparent")
        status_frame.pack(anchor="w", pady=(8, 0))
        ctk.CTkLabel(
            status_frame,
            text=f"{word_status} Microsoft Word",
            font=ctk.CTkFont(size=11), text_color=word_color,
        ).pack(side="left", padx=(0, 4))
        if not _check_word_available():
            ctk.CTkLabel(
                status_frame,
                text="(Required for .doc, .docx, .msg files)",
                font=ctk.CTkFont(size=10), text_color="gray",
            ).pack(side="left")

    # ── folders ──

    def _build_folders_section(self):
        frm = ctk.CTkFrame(self)
        frm.grid(row=1, column=0, sticky="ew", **SECTION_PAD)
        frm.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            frm, text="Folders",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).grid(row=0, column=0, columnspan=3, sticky="w",
               padx=INNER_PAD_X, pady=(12, 6))

        # Input
        ctk.CTkLabel(frm, text="Input Folder").grid(
            row=1, column=0, sticky="w", padx=(INNER_PAD_X, 8), pady=4)
        self._input_var = ctk.StringVar()
        self._input_entry = ctk.CTkEntry(
            frm, textvariable=self._input_var)
        self._input_entry.grid(row=1, column=1, sticky="ew", padx=4, pady=4)
        self._input_entry.bind("<FocusOut>", lambda _: self._refresh_file_count())
        self._input_entry.bind("<Return>", lambda _: self._refresh_file_count())

        ctk.CTkButton(frm, text="Browse", width=80,
                       command=self._browse_input).grid(
            row=1, column=2, padx=(4, INNER_PAD_X), pady=4)

        self._file_count_label = ctk.CTkLabel(
            frm, text="", font=ctk.CTkFont(size=11))
        self._file_count_label.grid(
            row=2, column=1, sticky="w", padx=6, pady=(0, 2))

        # Output
        ctk.CTkLabel(frm, text="Output Folder").grid(
            row=3, column=0, sticky="w", padx=(INNER_PAD_X, 8), pady=4)
        self._output_var = ctk.StringVar()
        self._output_entry = ctk.CTkEntry(
            frm, textvariable=self._output_var)
        self._output_entry.grid(row=3, column=1, sticky="ew", padx=4, pady=4)

        btn_frame = ctk.CTkFrame(frm, fg_color="transparent")
        btn_frame.grid(row=3, column=2, padx=(4, INNER_PAD_X), pady=4)

        ctk.CTkButton(btn_frame, text="Browse", width=80,
                       command=self._browse_output).pack(side="left")

        self._open_output_btn = ctk.CTkButton(
            frm, text="Open", width=56, height=28,
            fg_color="transparent", border_width=1,
            text_color=("gray30", "gray70"),
            border_color=("gray50", "gray60"),
            hover_color=("gray85", "gray30"),
            command=self._open_output_folder)
        self._open_output_btn.grid(
            row=4, column=1, sticky="w", padx=6, pady=(0, 12))

    # ── rules ──

    def _build_rules_section(self):
        frm = ctk.CTkFrame(self)
        frm.grid(row=2, column=0, sticky="ew", **SECTION_PAD)
        frm.grid_columnconfigure(0, weight=1)

        header = ctk.CTkFrame(frm, fg_color="transparent")
        header.grid(row=0, column=0, sticky="ew",
                     padx=INNER_PAD_X, pady=(12, 2))

        ctk.CTkLabel(
            header, text="Page Rules",
            font=ctk.CTkFont(size=15, weight="bold"),
        ).pack(side="left")
        ctk.CTkLabel(
            header,
            text="   Limit pages when a filename contains a keyword",
            font=ctk.CTkFont(size=11), text_color="gray",
        ).pack(side="left")

        col_hdr = ctk.CTkFrame(frm, fg_color="transparent")
        col_hdr.grid(row=1, column=0, sticky="w", padx=INNER_PAD_X)
        ctk.CTkLabel(col_hdr, text="Keyword", width=200,
                      font=ctk.CTkFont(size=11, weight="bold"),
                      anchor="w").pack(side="left", padx=(0, 8))
        ctk.CTkLabel(col_hdr, text="Pages", width=55,
                      font=ctk.CTkFont(size=11, weight="bold"),
                      anchor="center").pack(side="left")

        self._rules_container = ctk.CTkFrame(frm, fg_color="transparent")
        self._rules_container.grid(row=2, column=0, sticky="ew",
                                    padx=INNER_PAD_X)

        self._add_rule_btn = ctk.CTkButton(
            frm, text="+ Add Rule", width=120,
            fg_color="transparent", border_width=1,
            text_color=("gray30", "gray70"),
            border_color=("gray50", "gray60"),
            hover_color=("gray85", "gray30"),
            command=lambda: self._add_rule_row())
        self._add_rule_btn.grid(row=3, column=0, sticky="w",
                                 padx=INNER_PAD_X, pady=(6, 14))

    # ── log ──

    def _build_log_section(self):
        frm = ctk.CTkFrame(self)
        frm.grid(row=3, column=0, sticky="nsew", **SECTION_PAD)
        frm.grid_rowconfigure(1, weight=1)
        frm.grid_columnconfigure(0, weight=1)

        hdr = ctk.CTkFrame(frm, fg_color="transparent")
        hdr.grid(row=0, column=0, sticky="ew",
                  padx=INNER_PAD_X, pady=(10, 4))
        ctk.CTkLabel(hdr, text="Log",
                      font=ctk.CTkFont(size=15, weight="bold")).pack(side="left")

        ctk.CTkButton(
            hdr, text="Clear", width=56, height=26,
            fg_color="transparent", border_width=1,
            text_color=("gray30", "gray70"),
            border_color=("gray50", "gray60"),
            hover_color=("gray85", "gray30"),
            command=self._clear_log).pack(side="right")

        self._log_box = ctk.CTkTextbox(
            frm, state="disabled",
            font=ctk.CTkFont(family="Consolas", size=12))
        self._log_box.grid(row=1, column=0, sticky="nsew",
                            padx=INNER_PAD_X, pady=(0, 6))

        bar_frame = ctk.CTkFrame(frm, fg_color="transparent")
        bar_frame.grid(row=2, column=0, sticky="ew",
                        padx=INNER_PAD_X, pady=(0, 12))
        bar_frame.grid_columnconfigure(0, weight=1)

        self._progress_bar = ctk.CTkProgressBar(bar_frame)
        self._progress_bar.grid(row=0, column=0, sticky="ew", padx=(0, 10))
        self._progress_bar.set(0)

        self._progress_label = ctk.CTkLabel(
            bar_frame, text="0 / 0", width=70,
            font=ctk.CTkFont(size=11))
        self._progress_label.grid(row=0, column=1)

    # ── action button ──

    def _build_actions(self):
        frm = ctk.CTkFrame(self, fg_color="transparent")
        frm.grid(row=4, column=0, sticky="ew", padx=18, pady=(10, 18))
        frm.grid_columnconfigure(0, weight=1)

        self._action_btn = ctk.CTkButton(
            frm, text="Create Bundle", height=48,
            font=ctk.CTkFont(size=16, weight="bold"),
            command=self._on_action)
        self._action_btn.grid(row=0, column=0, sticky="ew")

        self._default_btn_fg = self._action_btn.cget("fg_color")
        self._default_btn_hover = self._action_btn.cget("hover_color")

    # ── defaults / config ──

    def _load_defaults(self):
        default_input = os.path.join(self._root_dir, 'INPUT')
        default_output = os.path.join(self._root_dir, 'OUTPUT')

        self._input_var.set(_word_com_path(default_input))
        self._output_var.set(_word_com_path(default_output))
        self._refresh_file_count()

        config_path = os.path.join(default_input, 'config.txt')
        loaded = _load_config_file(config_path)
        rules = loaded if loaded else [("email", 1)]

        for kw, pg in rules:
            self._add_rule_row(kw, str(pg))

    def _center_window(self):
        self.update_idletasks()
        w = self.winfo_width()
        h = self.winfo_height()
        x = (self.winfo_screenwidth() - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    # ── folder browsing ──

    def _browse_input(self):
        current = self._input_var.get().strip()
        initial = current if os.path.isdir(current) else self._root_dir
        path = filedialog.askdirectory(
            title="Select Input Folder", initialdir=initial)
        if not path:
            return
        path = _word_com_path(path)
        self._input_var.set(path)
        self._refresh_file_count()

        cfg = _load_config_file(os.path.join(path, 'config.txt'))
        if cfg:
            for row in self._rule_rows[:]:
                self._remove_rule_row(row)
            for kw, pg in cfg:
                self._add_rule_row(kw, str(pg))

    def _browse_output(self):
        current = self._output_var.get().strip()
        initial = current if os.path.isdir(current) else self._root_dir
        path = filedialog.askdirectory(
            title="Select Output Folder", initialdir=initial)
        if path:
            self._output_var.set(_word_com_path(path))

    def _open_output_folder(self):
        path = self._output_var.get().strip()
        if path and os.path.isdir(path):
            os.startfile(path)

    def _refresh_file_count(self):
        path = self._input_var.get().strip()
        if not path or not os.path.isdir(path):
            self._file_count_label.configure(
                text="  Folder not found", text_color=RED)
            return
        count = sum(
            1 for f in os.listdir(path)
            if os.path.isfile(os.path.join(path, f))
            and os.path.splitext(f)[1].lower() in SUPPORTED_EXT)
        if count == 0:
            self._file_count_label.configure(
                text="  No supported files found (pdf, doc, docx, msg)",
                text_color=ORANGE)
        else:
            self._file_count_label.configure(
                text=f"  {count} supported file(s) found", text_color=GREEN)

    # ── rule rows ──

    def _add_rule_row(self, keyword="", pages="1"):
        row = RuleRow(self._rules_container, keyword=keyword, pages=pages,
                      on_delete=self._remove_rule_row)
        row.pack(fill="x", pady=1)
        self._rule_rows.append(row)

    def _remove_rule_row(self, row):
        if row in self._rule_rows:
            self._rule_rows.remove(row)
        row.destroy()

    def _collect_rules(self) -> dict:
        rules = {}
        for row in self._rule_rows:
            result = row.get_rule()
            if result:
                rules[result[0]] = result[1]
        return rules

    # ── log ──

    def _append_log(self, text):
        self._log_lines.append(text)
        self._log_box.configure(state="normal")
        self._log_box.insert("end", text + "\n")
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _clear_log(self):
        self._log_lines.clear()
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")
        self._progress_bar.set(0)
        self._progress_label.configure(text="0 / 0")

    # ── queue polling ──

    def _poll_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                kind = msg[0]
                if kind == "log":
                    self._append_log(msg[1])
                elif kind == "progress":
                    idx, total = msg[1], msg[2]
                    self._progress_bar.set(idx / total)
                    self._progress_label.configure(text=f"{idx} / {total}")
                elif kind == "done":
                    self._on_done(msg[1], msg[2])
        except queue_mod.Empty:
            pass
        self.after(100, self._poll_queue)

    # ── start / cancel / done ──

    def _on_action(self):
        if self._worker and self._worker.is_alive():
            self._cancel.set()
            self._action_btn.configure(text="Cancelling ...", state="disabled")
            return

        input_dir = _word_com_path(self._input_var.get().strip())
        output_dir = _word_com_path(self._output_var.get().strip())
        self._input_var.set(input_dir)
        self._output_var.set(output_dir)

        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showerror("Error", "Please select a valid input folder.")
            return
        if not output_dir:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        rules = self._collect_rules()

        _save_config_file(os.path.join(input_dir, 'config.txt'), rules)

        self._clear_log()
        self._append_log("=" * 50)
        self._append_log("  Create Bundle - PDF Bundler")
        self._append_log("=" * 50)
        self._append_log("")
        
        # Show system capabilities
        if _check_word_available():
            self._append_log("✓ Microsoft Word is available")
        else:
            self._append_log("⚠ Microsoft Word is NOT available")
            self._append_log("  (PDFs only)")
        self._append_log("")

        if rules:
            self._append_log("\nPage rules:")
            for kw, pg in rules.items():
                self._append_log(f'  "{kw}" -> {pg} page(s)')
            self._append_log("")

        self._cancel.clear()
        self._set_controls_enabled(False)
        self._action_btn.configure(
            text="Cancel", fg_color="#c0392b", hover_color="#922b21")

        self._worker = threading.Thread(
            target=_run_bundle,
            args=(input_dir, output_dir, rules, self._queue, self._cancel),
            daemon=True)
        self._worker.start()

    def _on_done(self, success, out_path):
        self._action_btn.configure(
            text="Create Bundle", state="normal",
            fg_color=self._default_btn_fg,
            hover_color=self._default_btn_hover)
        self._set_controls_enabled(True)

        if success:
            self._progress_bar.set(1)

        self._refresh_file_count()

        output_dir = self._output_var.get().strip()
        if output_dir and self._log_lines:
            try:
                os.makedirs(output_dir, exist_ok=True)
                ts = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                log_path = os.path.join(output_dir, f"log_{ts}.txt")
                with open(log_path, 'w', encoding='utf-8') as f:
                    f.write('\n'.join(self._log_lines))
                self._append_log(f"\nLog saved: {log_path}")
            except Exception:
                pass

        if success and out_path:
            messagebox.showinfo(
                "Success",
                f"Bundle created successfully!\n\n{out_path}")

    def _set_controls_enabled(self, enabled):
        state = "normal" if enabled else "disabled"
        self._input_entry.configure(state=state)
        self._output_entry.configure(state=state)
        self._add_rule_btn.configure(state=state)
        self._open_output_btn.configure(state=state)
        for row in self._rule_rows:
            row.set_enabled(enabled)


# ── Entry point ──────────────────────────────────────────────────────────────

if __name__ == '__main__':
    app = BundleApp()
    app.mainloop()