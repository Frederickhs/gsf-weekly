# -*- mode: python ; coding: utf-8 -*-

from pathlib import Path

project_root = Path(SPECPATH).parent.parent
scripts_dir = project_root / "Scripts"
version_file = project_root / "VERSION.txt"
version = version_file.read_text(encoding="utf-8").strip() if version_file.exists() else "0.0.0"
icon_file = project_root / "Packaging" / "macOS" / "SuccessFactorsWeeklyTrainingReport.icns"
icon_value = str(icon_file) if icon_file.exists() else None

analysis = Analysis(
    [str(scripts_dir / "gsf_app.py")],
    pathex=[str(scripts_dir)],
    binaries=[],
    datas=[],
    hiddenimports=[
        "openpyxl",
        "openpyxl.cell._writer",
        "openpyxl.styles",
        "openpyxl.worksheet._writer",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["pandas", "numpy"],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(analysis.pure)

exe = EXE(
    pyz,
    analysis.scripts,
    [],
    exclude_binaries=True,
    name="Success Factors Weekly Training Report",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

collection = COLLECT(
    exe,
    analysis.binaries,
    analysis.datas,
    strip=False,
    upx=False,
    name="Success Factors Weekly Training Report",
)

app = BUNDLE(
    collection,
    name="Success Factors Weekly Training Report.app",
    icon=icon_value,
    bundle_identifier="com.frederickhs.successfactorsweeklytrainingreport",
    version=version,
    info_plist={
        "CFBundleDisplayName": "Success Factors Weekly Training Report",
        "NSHighResolutionCapable": True,
    },
)
