# -*- mode: python ; coding: utf-8 -*-
# Spec PyInstaller pour "Data Collector" (console BELVEDERE).
# Construit un exécutable unique (--onefile) avec l'icône Romana.
#
#   pip install pyinstaller
#   pyinstaller DataCollector.spec
#
# Le résultat est dans dist/ :  "Data Collector.exe" (Windows) ou l'app (Mac).

import sys
import os

block_cipher = None

# Ressources embarquées (lecture seule) : interface + logo + aide OCR.
datas = [
    ("templates/index.html", "templates"),
    ("static/logo.png", "static"),
    ("static/app.ico", "static"),
    ("static/ocr_vision.swift", "static"),
]

a = Analysis(
    ["DataCollector.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=["paths", "app", "extract", "letters", "store", "ocr",
                   "openpyxl", "docx"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=["tkinter", "matplotlib", "numpy", "PyQt5", "PySide2"],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

icon = os.path.join("static", "app.ico")

# Mode --onedir : démarrage rapide (~1-2 s). Produit un dossier "Data Collector"
# contenant l'exécutable + ses fichiers. L'utilisateur double-clique l'exe dedans.
exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="Data Collector",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,            # fenêtre console (affiche l'adresse). False pour la cacher.
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name="Data Collector",
)

# Sur macOS, produire aussi un .app
if sys.platform == "darwin":
    app = BUNDLE(
        coll,
        name="Data Collector.app",
        icon=os.path.join("static", "app.icns") if os.path.exists(os.path.join("static", "app.icns")) else icon,
        bundle_identifier="ma.romana.datacollector",
        info_plist={"LSUIElement": False},
    )
