# -*- mode: python ; coding: utf-8 -*-

from PyInstaller.utils.hooks import collect_submodules

hiddenimports = ['main']
hiddenimports += collect_submodules('lugest_qt')
hiddenimports += collect_submodules('impulse_mobile_api')
hiddenimports += collect_submodules('reportlab.graphics.barcode')


a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('app.ico', '.'),
        ('lugest_branding.json', '.'),
        ('logo.jpg', '.'),
        ('Logos/logo.png', 'Logos'),
        ('Logos/image (1).jpg', 'Logos'),
        ('Logos/image.jpg', 'Logos'),
        ('Logos/logo(1).jpg', 'Logos'),
        ('Logos/Vidoinicial.mp4', 'Logos'),
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
    a.binaries,
    a.datas,
    [],
    name='main',
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
