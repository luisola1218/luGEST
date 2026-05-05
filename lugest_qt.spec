# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = collect_submodules('impulse_mobile_api')
hiddenimports += collect_submodules('reportlab.graphics.barcode')


a = Analysis(
    ['lugest_qt_app.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app.ico', '.'),
        ('lugest_branding.json', '.'),
        ('VERSION', '.'),
        ('logo.jpg', '.'),
        ('Logos/logo.png', 'Logos'),
        ('Logos/image (1).jpg', 'Logos'),
    ],
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='lugest_qt',
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
    icon=['app.ico'],
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='lugest_qt',
)
