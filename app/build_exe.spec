# -*- mode: python ; coding: utf-8 -*-
# Сборка: pyinstaller app/build_exe.spec  (из корня проекта)
from pathlib import Path

root = Path(SPECPATH).parent   # корень excel-compare

a = Analysis(
    [str(root / "app" / "main_gui.py")],
    pathex=[str(root / "app"), str(root / "src")],
    binaries=[],
    datas=[(str(root / "app" / "assets"), "assets")],
    hiddenimports=["pdfplumber", "pandas", "openpyxl", "pyxlsb"],
    hookspath=[],
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib"],
    noarchive=False,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz, a.scripts, [],
    exclude_binaries=True,
    name="excel-compare",
    debug=False, strip=False, upx=False,
    console=False,                       # GUI без консоли
    icon=str(root / "app" / "assets" / "icon.ico"),
)
coll = COLLECT(exe, a.binaries, a.datas,
               strip=False, upx=False, name="excel-compare")
