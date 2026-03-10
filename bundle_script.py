"""
Create_Bundle – PDF Bundler (GUI)
Merges PDFs, Word docs, and Outlook .msg files into a single timestamped PDF.
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
RETRY_DELAY_BASE = 3
INTER_FILE_DELAY = 2.5          # pause between processing each file
DOC_OPEN_SETTLE = 3.0           # after Word opens a document
DOC_READY_POLL_INTERVAL = 0.5   # interval when waiting for doc ready
DOC_READY_MAX_WAIT = 10.0       # max seconds to wait for document to be ready
DOC_SAVE_SETTLE = 2.0          # after Word saves as PDF
FILE_STABLE_CHECK_INTERVAL = 0.25
FILE_STABLE_MIN_CHECKS = 2      # file size unchanged for this many checks
DOC_CLOSE_SETTLE = 1.5          # after closing a document
APP_QUIT_SETTLE = 2.0           # after quitting Word/Outlook before next use
OUTLOOK_MSG_LOAD_SETTLE = 2.5   # after opening .msg before reading properties
OUTLOOK_AFTER_MSG_SETTLE = 1.0  # after finishing with one .msg before next
HTML_TO_DISK_SETTLE = 1.0      # after writing HTML file before opening in Word

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


# ── Processing helpers ───────────────────────────────────────────────────────

def _wait_for_file_stable(filepath, timeout=8.0, check_interval=FILE_STABLE_CHECK_INTERVAL):
    """Wait until file exists and size is unchanged for FILE_STABLE_MIN_CHECKS checks."""
    deadline = time.monotonic() + timeout
    last_size = -1
    stable_count = 0
    while time.monotonic() < deadline:
        try:
            if os.path.isfile(filepath):
                size = os.path.getsize(filepath)
                if size == last_size and size >= 0:
                    stable_count += 1
                    if stable_count >= FILE_STABLE_MIN_CHECKS:
                        return
                else:
                    stable_count = 0
                    last_size = size
            else:
                stable_count = 0
        except Exception:
            stable_count = 0
        time.sleep(check_interval)


def _wait_for_doc_ready(doc, log_fn):
    """Give Word time to fully build the document after Open (e.g. finish layout)."""
    time.sleep(DOC_OPEN_SETTLE)
    deadline = time.monotonic() + DOC_READY_MAX_WAIT
    while time.monotonic() < deadline:
        try:
            # Touch content to force Word to finish loading; can trigger background layout
            _ = doc.Content.Start
            time.sleep(0.5)
            return
        except Exception:
            time.sleep(DOC_READY_POLL_INTERVAL)
    log_fn("    Note: document ready wait reached timeout; continuing anyway.")


def _create_word_app():
    import comtypes.client
    word = comtypes.client.CreateObject('Word.Application')
    word.Visible = False
    return word


def _open_doc_with_retry(word, path, log_fn):
    for attempt in range(RETRY_ATTEMPTS):
        try:
            doc = word.Documents.Open(path)
            if doc is None:
                raise RuntimeError("Word returned None after Documents.Open")
            _wait_for_doc_ready(doc, log_fn)
            return doc
        except Exception:
            delay = RETRY_DELAY_BASE * (attempt + 1)
            if attempt < RETRY_ATTEMPTS - 1:
                log_fn(f"    Retry {attempt + 1}/{RETRY_ATTEMPTS}: "
                       f"Word failed to open, waiting {delay}s ...")
                time.sleep(delay)
            else:
                raise


def _save_pdf_with_retry(doc, pdf_path, log_fn):
    for attempt in range(RETRY_ATTEMPTS):
        try:
            doc.SaveAs(pdf_path, FileFormat=17)
            time.sleep(DOC_SAVE_SETTLE)
            _wait_for_file_stable(pdf_path)
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


def _word_to_pdf(src_path, out_pdf_path, log_fn):
    word = _create_word_app()
    try:
        doc = _open_doc_with_retry(word, src_path, log_fn)
        _save_pdf_with_retry(doc, out_pdf_path, log_fn)
        doc.Close(False)
        time.sleep(DOC_CLOSE_SETTLE)
    finally:
        word.Quit()
        time.sleep(APP_QUIT_SETTLE)


def _extract_html_body(html_source):
    if not html_source:
        return ""
    match = re.search(r'<body[^>]*>(.*)</body>', html_source,
                      re.DOTALL | re.IGNORECASE)
    return match.group(1) if match else html_source


def _format_email_datetime(dt):
    if dt is None:
        return "Unknown"
    try:
        return dt.strftime("%d/%m/%Y %H:%M")
    except Exception:
        return str(dt)


def _msg_to_pdf(src_path, out_pdf_path, log_fn):
    import win32com.client

    outlook = win32com.client.Dispatch('Outlook.Application')
    msg = outlook.CreateItemFromTemplate(src_path)
    time.sleep(OUTLOOK_MSG_LOAD_SETTLE)

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

    sent_date = None
    try:
        sent_date = msg.SentOn
    except Exception:
        try:
            sent_date = msg.ReceivedTime
        except Exception:
            pass
    date_str = _format_email_datetime(sent_date)

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

    full_html = EMAIL_HTML_TEMPLATE.format(
        from_field=from_display,
        to_field=html_mod.escape(to_field) if to_field else "Unknown",
        date_field=html_mod.escape(date_str),
        subject_field=html_mod.escape(subject),
        body_html=body_content,
    )

    time.sleep(OUTLOOK_AFTER_MSG_SETTLE)

    html_path = out_pdf_path.replace('.pdf', '.html')
    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(full_html)
    time.sleep(HTML_TO_DISK_SETTLE)

    word = _create_word_app()
    try:
        doc = _open_doc_with_retry(word, html_path, log_fn)
        _save_pdf_with_retry(doc, out_pdf_path, log_fn)
        doc.Close(False)
        time.sleep(DOC_CLOSE_SETTLE)
    finally:
        word.Quit()
        time.sleep(APP_QUIT_SETTLE)

    try:
        os.remove(html_path)
    except Exception:
        pass


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
                    _word_to_pdf(src, pdf_path, log_fn)
                elif ext == '.msg':
                    pdf_path = os.path.join(temp_dir, f"conv_{idx}.pdf")
                    _msg_to_pdf(src, pdf_path, log_fn)
                else:
                    continue

                time.sleep(1.0)
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
                log_fn(f"    FAILED: {e}")
                log_fn(traceback.format_exc())

            time.sleep(INTER_FILE_DELAY)

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
            text="Merge PDFs, Word documents and Outlook emails into one PDF.",
            font=ctk.CTkFont(size=13), text_color="gray",
        ).pack(anchor="w", pady=(2, 0))

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

        self._input_var.set(default_input)
        self._output_var.set(default_output)
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
            self._output_var.set(path)

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

        input_dir = self._input_var.get().strip()
        output_dir = self._output_var.get().strip()

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
