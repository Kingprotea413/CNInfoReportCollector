# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

from cninfo_pipeline.template_registry import template_data_files

project_root = Path(globals().get("__file__", "CNInfoReportCollector.spec")).resolve().parent
icon_path = project_root / "assets" / "app_icon.ico"
template_datas = template_data_files(project_root)

a = Analysis(
    [str(project_root / 'app.py')],
    pathex=[],
    binaries=[],
    datas=template_datas,
    hiddenimports=[],
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
    name='CNInfoReportCollector',
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
    icon=str(icon_path),
)
