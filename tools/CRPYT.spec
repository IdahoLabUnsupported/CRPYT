# Copyright 2026, Battelle Energy Alliance, LLC. All Rights Reserved.
from pathlib import Path

from PyInstaller.building.build_main import Analysis, EXE, PYZ


project_root = Path.cwd()
script_path = project_root / "tools" / "CRPYT.py"
data_dir = project_root / "tools" / "data"
legacy_cache_dir = project_root / "tools" / "__pycache__"

datas = [
    (str(data_dir), "data"),
    (str(legacy_cache_dir), "__pycache__"),
]

hiddenimports = []

a = Analysis(
    [str(script_path)],
    pathex=[str(script_path.parent)],
    binaries=[],
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={"matplotlib": {"backends": ["TkAgg"]}},
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
    name="CRPYT",
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
