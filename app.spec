# PyInstaller spec — Mandats TXT→XLSX
# Usage: pyinstaller app.spec

import sys
from pathlib import Path
import streamlit, openpyxl

block_cipher = None

# Collect streamlit static assets
streamlit_dir = Path(streamlit.__file__).parent
openpyxl_dir  = Path(openpyxl.__file__).parent

added_datas = [
    # app source files (loaded at runtime, not imported)
    ("app.py",         "."),
    ("txt_to_xlsx.py", "."),
    # streamlit static + runtime assets
    (str(streamlit_dir / "static"),  "streamlit/static"),
    (str(streamlit_dir / "runtime"), "streamlit/runtime"),
    # openpyxl templates
    (str(openpyxl_dir / "templates"), "openpyxl/templates"),
]

a = Analysis(
    ["launcher.py"],
    pathex=[],
    binaries=[],
    datas=added_datas,
    hiddenimports=[
        "streamlit",
        "streamlit.web.cli",
        "streamlit.runtime.scriptrunner",
        "openpyxl",
        "openpyxl.styles",
        "openpyxl.utils",
        "altair",
        "pyarrow",
        "packaging",
        "click",
        "toml",
        "tornado",
        "PIL",
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name="Mandats-TXT-to-XLSX",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,          # no terminal window
    disable_windowed_traceback=False,
    icon=None,
)
