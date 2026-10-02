# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_dir = Path(SPECPATH)

datas = []

# Include PowerShell scripts
powershell_dir = project_dir / "powershell"
if powershell_dir.exists():
    datas.append((str(powershell_dir), "powershell"))

# Include project assets if they exist
for folder in ["static", "assets", "icons", "images", "config"]:
    folder_path = project_dir / folder
    if folder_path.exists():
        datas.append((str(folder_path), folder))


a = Analysis(
    ["main.py"],
    pathex=[str(project_dir)],
    binaries=[],
    datas=datas,
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
    name="ManagementConsole",
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

    # Windows application icon
    icon=str(project_dir / "icons" / "ManagementConsole.ico"),
)