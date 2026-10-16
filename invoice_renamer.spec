# -*- mode: python ; coding: utf-8 -*-

a = Analysis(
    ['invoice_renamer.py'],
    pathex=[],
    binaries=[],
    datas=[],
    hiddenimports=[
        'invoice_renamer',
        'invoice_renamer.config',
        'invoice_renamer.ui',
        'invoice_renamer.setup_wizard',
        'invoice_renamer.worker',
        'invoice_renamer.extractor',
        'invoice_renamer.llm',
        'invoice_renamer.normalizer',
        'invoice_renamer.table',
        'invoice_renamer.hints',
        'invoice_renamer.logging_setup',
        'PIL._tkinter_finder',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='invoice-renamer',
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
