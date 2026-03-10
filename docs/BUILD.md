# Build instructions

## Prerequisites

- **Python 3.10+** with pip
- **Inno Setup 6** (for the installer): https://jrsoftware.org/isinfo.php

Install Python dependencies:

```bash
pip install -r requirements.txt
pip install pyinstaller
```

## 1. Build the single-file exe

From the repo root (`BundleScript/`):

```bash
pyinstaller Create_Bundle.spec --noconfirm
```

This produces **`dist/Create_Bundle.exe`** – a single executable with Python and all dependencies (CustomTkinter, Word/Outlook COM, pypdf) bundled inside. You can copy this one file to any Windows PC and run it; no `_internal` folder or installer needed.

## 2. Build the installer (optional)

Open **Inno Setup Compiler**, load **`Create_Bundle.iss`**, and click **Build → Compile**.

Or from the command line:

```bash
"C:\Program Files (x86)\Inno Setup 6\ISCC.exe" Create_Bundle.iss
```

The installer is created at **`Output/Create_Bundle_Setup.exe`**.

- **No admin required** – installs per user under `%LocalAppData%\Create Bundle`
- Copies the single exe and creates **INPUT** and **OUTPUT** folders
- Adds Start Menu (and optional Desktop) shortcut

## Summary

| Output | Use |
|--------|-----|
| `dist/Create_Bundle.exe` | Standalone app. Copy to any PC and run. |
| `Output/Create_Bundle_Setup.exe` | Optional installer for shortcuts and folder setup. |
