# -*- mode: python ; coding: utf-8 -*-
import os

# 添加主要模組目錄路徑
main_path = os.path.dirname(os.path.abspath('program/src/gui.py'))

a = Analysis(
    ['program\\src\\gui.py'],
    pathex=[main_path],
    binaries=[],
    datas=[
        ('program\\123files_spec_rule', 'files_spec_rule'),
        ('program\\icon_image', 'icon_image'),
        ('program\\xlsx_to_json.py', '.'),
        ('icon.ico', '.')
    ],
    hiddenimports=['pandas', 'openpyxl', 'numpy', 'tkinter', 'xlsx_to_json'],
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
    name='XlsxToJson',
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
    icon=['icon.ico'],
)
