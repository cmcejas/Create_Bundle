# -*- mode: python ; coding: utf-8 -*-
# Build: pyinstaller Create_Bundle.spec
# Output: dist/Create_Bundle.exe (single exe with all dependencies - no _internal folder needed)

from PyInstaller.utils.hooks import collect_data_files

block_cipher = None

# CustomTkinter themes/fonts (required for GUI) - bundled into the exe
ctk_datas = collect_data_files('customtkinter')

a = Analysis(
    ['bundle_script.py'],
    pathex=[],
    binaries=[],
    datas=ctk_datas,
    hiddenimports=[
        'customtkinter',
        'darkdetect',
        'PIL',
        'PIL._tkinter_finder',
        'comtypes',
        'comtypes.client',
        'win32com',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'pypdf',
        'pypdf._reader',
        'pypdf._writer',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# One-file mode: everything (DLLs, Python, data) is inside this single exe.
# Copy this one exe to any machine and it runs without _internal or installer.
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Create_Bundle',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
