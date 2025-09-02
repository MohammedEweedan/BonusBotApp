# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['bonus_bot_app.py'],
    pathex=['.'],
    binaries=[],
    datas=[],
    hiddenimports=[
        'pandas',
        'docx',
        'docx.oxml.text.run',
        'docx.text.run',
        'docx.shared',
        'win32com.client',
        'pythoncom',
        'pywintypes',
        'openpyxl',
        'xlsxwriter'
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

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='BonusBot',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Set to True if you want to see console output for debugging
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None  # Add icon='icon.ico' if you have an icon file
)