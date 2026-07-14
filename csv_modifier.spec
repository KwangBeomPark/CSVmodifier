# -*- mode: python ; coding: utf-8 -*-
from pathlib import Path
import re


project_root = Path(SPECPATH)
source = (project_root / 'csv_modifier.py').read_text(encoding='utf-8')
match = re.search(r'^__version__\s*=\s*["\']([^"\']+)["\']', source, re.MULTILINE)
if match is None:
    raise RuntimeError('Could not find __version__ in csv_modifier.py')
release_executable_name = f"App04_csv_modifier_v{match.group(1)}"

a = Analysis(
    ['csv_modifier.py'],
    pathex=[str(project_root)],
    binaries=[],
    datas=[
        (str(project_root / 'icon.ico'), '.'),
        (str(project_root / 'icon.png'), '.'),
    ],
    hiddenimports=[],
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
    a.zipfiles,
    a.datas,
    [],
    name=release_executable_name,
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
    icon=str(project_root / 'icon.ico'),
)
