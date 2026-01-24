# -*- mode: python ; coding: utf-8 -*-

datas_list = [
    ('gui/styles/macos_style.qss', 'gui/styles'),
]

a = Analysis(
    ['gui/main.py'],
    pathex=[],
    binaries=[],
    datas=datas_list,
    hiddenimports=[
        'playwright',
        'playwright.async_api',
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
    [],
    exclude_binaries=True,
    name='足球彩票分析工具',
    icon='resources/icons/app_icon.ico',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=False,
    disable_windowed_traceback=False,
)
coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='足球彩票分析工具',
)
