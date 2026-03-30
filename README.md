# Create Bundle – PDF Bundler

A Windows desktop app that merges **PDFs**, **Word documents** (`.doc`/`.docx`), and **emails** (Outlook **`.msg`** and standard **`.eml`**) into a single timestamped PDF. Built with Python, CustomTkinter, and Word (Outlook COM is only needed for `.msg`).

## System Requirements

- **Windows** (the app is Windows-only)
- **Microsoft Word** – Required to convert Word docs and Outlook `.msg` files to PDF. **Optional for PDF and `.eml` files.**
  - Download: https://www.microsoft.com/office
  - If Word is missing, the app will warn you at startup but still work with supported formats
- **.NET Runtime** – Usually pre-installed on Windows. If the app won't start, install the [Visual C++ Redistributable](https://support.microsoft.com/en-us/help/2977003/the-latest-supported-visual-c-downloads)

## Download (for end users)

If you just want to use the app (no build required):

- **Installer (recommended)** – Run **`Create_Bundle_Setup.exe`** from the [Output](Output) folder. No admin rights needed; it will create shortcuts and INPUT/OUTPUT folders. Includes automatic VC++ Runtime check.  
  → Direct link: `https://github.com/YOUR_ORG/Create_Bundle/raw/main/Output/Create_Bundle_Setup.exe` (replace with your repo URL after push.)
- **Portable** – Download **`Create_Bundle.exe`** from the [dist](dist) folder. Copy it anywhere and run; it will create INPUT and OUTPUT folders next to itself.

Both files are checked into the repo so you can share the GitHub link with colleagues.

## Features

- **GUI** – Simple interface: choose input/output folders, set page rules, run the bundle.
- **Automatic system checks** – Detects if Microsoft Word is available on startup; warns if missing but continues (useful for PDF-only workflows).
- **Page rules** – Limit how many pages are included per file by filename keyword (e.g. files with "email" in the name → 1 page). Configure in the app or via `config.txt` in the input folder.
- **Email formatting** – Converts `.msg` (Outlook) and `.eml` (RFC 822) to PDF with the same clean header (From, To, Date, Subject) and consistent layout.
- **Helpful error messages** – If something fails, the app provides clear hints (e.g., "Word not installed", "file path issue").
- **Single exe** – One self-contained `.exe` you can copy to any Windows PC (no Python or installer required).
- **Optional installer** – Inno Setup installer, no admin rights; installs per user.

## Run from source

```bash
cd BundleScript
pip install -r requirements.txt
python bundle_script.py
```

Create an **INPUT** folder next to the script (or set it in the app), put your PDFs, Word files, `.msg` / `.eml` emails there, set the **OUTPUT** folder, and click **Create Bundle**.

## Build the exe and installer

See **[docs/BUILD.md](docs/BUILD.md)** for:
What's Packaged in the `.exe`

Everything needed to run on any Windows PC, including:

- **Python 3.10+** runtime
- **CustomTkinter** - GUI framework (with dark-mode support)
- **PyPDF** - PDF reading/merging
- **comtypes** - Windows COM interface (for Word and Outlook), VC++ Runtime check,
- **pywin32** - Windows-native features
- **Pillow (PIL)** - Image support  
- **darkdetect** - System theme detection

The installer (`Create_Bundle_Setup.exe`) also handles:
- **Visual C++ Runtime** – Automatically checks and installs if needed
- **Start Menu shortcuts** – Quick access from Windows Start Menu
- **INPUT/OUTPUT folders** – Pre-created on install
- **Per-user installation** – No admin rights required

## 
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

## Troubleshooting

| Issue | Fix |
|-------|-----|
| "Microsoft Word Required" warning at startup | Word is not installed. Install from https://www.microsoft.com/office (Word is only needed for `.doc`, `.docx`, `.msg` files; PDFs and `.eml` work without it) |
| App won't start (black window then closes) | Install the [Visual C++ Redistributable](https://support.microsoft.com/en-us/help/2977003/the-latest-supported-visual-c-downloads) |
| "Couldn't find your file" error | Try renaming your file or moving it to a path without special characters (ü, ñ, etc.) |
| `.eml` files work but `.msg` files fail | Outlook is missing or damaged. Try repairing Office via `Settings > Apps > Installed apps > Microsoft Office > Repair` |

## License

Use and modify as you like.
