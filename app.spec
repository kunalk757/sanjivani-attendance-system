# -*- mode: python ; coding: utf-8 -*-
import os

block_cipher = None

# Base directory
BASE_DIR = os.path.abspath(SPECPATH)

added_datas = [
    ('haarcascade_frontalface_default.xml', '.'),
    ('sanjivani.ico', '.'),
    ('sanjivani.png', '.'),
    ('icon.ico', '.'),
    ('icon.png', '.'),
]

# Hidden imports required for standalone execution across different Windows systems
hidden_imports = [
    'cv2',
    'cv2.face',
    'numpy',
    'openpyxl',
    'openpyxl.styles',
    'openpyxl.workbook',
    'openpyxl.cell',
    'openpyxl.reader.excel',
    'PIL',
    'PIL.Image',
    'PIL.ImageTk',
    'tkinter',
    'tkinter.ttk',
    'tkinter.messagebox',
    'smtplib',
    'email',
    'email.mime',
    'email.mime.multipart',
    'email.mime.text',
    'email.mime.base',
    'email.encoders',
    'send_email',
]

a = Analysis(
    ['app.py'],
    pathex=[BASE_DIR],
    binaries=[],
    datas=added_datas,
    hiddenimports=hidden_imports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=['IPython', 'jupyter', 'notebook', 'pytest', 'scipy', 'matplotlib'],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='Sanjivani Attendance',
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
    icon='sanjivani.ico',
)
