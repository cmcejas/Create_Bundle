# Create Bundle – PDF Bundler

A Windows desktop app that merges **PDFs**, **Word documents** (`.doc`/`.docx`), and **emails** (Outlook **`.msg`** and standard **`.eml`**) into a single timestamped PDF. Built with Python, CustomTkinter, and Word (Outlook COM is only needed for `.msg`).

## Download (for end users)

If you just want to use the app (no build required):

- **Installer (recommended)** – Run **`Create_Bundle_Setup.exe`** from the [Output](Output) folder. No admin rights needed; it will create shortcuts and INPUT/OUTPUT folders.  
  → Direct link: `https://github.com/YOUR_ORG/Create_Bundle/raw/main/Output/Create_Bundle_Setup.exe` (replace with your repo URL after push.)
- **Portable** – Download **`Create_Bundle.exe`** from the [dist](dist) folder. Copy it anywhere and run; it will create INPUT and OUTPUT folders next to itself.

Both files are checked into the repo so you can share the GitHub link with colleagues.

## Features

- **GUI** – Simple interface: choose input/output folders, set page rules, run the bundle.
- **Page rules** – Limit how many pages are included per file by filename keyword (e.g. files with “email” in the name → 1 page). Configure in the app or via `config.txt` in the input folder.
- **Email formatting** – Converts `.msg` (Outlook) and `.eml` (RFC 822) to PDF with the same clean header (From, To, Date, Subject) and consistent layout.
- **Single exe** – One self-contained `.exe` you can copy to any Windows PC (no Python or installer required).
- **Optional installer** – Inno Setup installer, no admin rights; installs per user.

## Requirements

- **Windows** – **Word** is required to convert Word docs and emails to PDF. **Outlook** is only required for **`.msg`** files; **`.eml`** files are parsed in Python (no Outlook).
- To run from source: **Python 3.10+**, and the packages in `requirements.txt`.

## Run from source

```bash
cd BundleScript
pip install -r requirements.txt
python bundle_script.py
```

Create an **INPUT** folder next to the script (or set it in the app), put your PDFs, Word files, `.msg` / `.eml` emails there, set the **OUTPUT** folder, and click **Create Bundle**.

## Build the exe and installer

See **[docs/BUILD.md](docs/BUILD.md)** for:

1. Building the single-file exe with PyInstaller (all dependencies inside the exe).
2. Building the optional Inno Setup installer (per-user, no admin).

Outputs:

- **`dist/Create_Bundle.exe`** – Copy this one file to any PC and run it; it will create INPUT/OUTPUT next to itself if needed.
- **`Output/Create_Bundle_Setup.exe`** – Optional installer for Start Menu shortcut and folders.

## Project layout

```
BundleScript/
├── bundle_script.py      # Main app (GUI + bundling logic)
├── requirements.txt      # Python dependencies
├── Create_Bundle.spec    # PyInstaller spec (onefile)
├── Create_Bundle.iss     # Inno Setup script
├── README.md
├── .gitignore
└── docs/
    └── BUILD.md          # Build instructions
```

## License

Use and modify as you like.
