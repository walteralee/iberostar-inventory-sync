# -*- mode: python ; coding: utf-8 -*-
#
# Receta de PyInstaller para generar la aplicación de escritorio.
# Uso (desde la raíz del repositorio):
#
#     pyinstaller packaging/iberostar.spec --noconfirm
#
# Resultado: dist/IberostarGestorPedidos/IberostarGestorPedidos.exe

from pathlib import Path

ROOT = Path(SPECPATH).parent
BACKEND = ROOT / "app" / "backend"

a = Analysis(
    [str(BACKEND / "desktop.py")],
    pathex=[str(BACKEND)],
    datas=[
        (str(ROOT / "app" / "frontend"), "app/frontend"),
        (str(ROOT / "storage" / "templates"), "storage/templates"),
    ],
    hiddenimports=["clr"],
    excludes=["tkinter", "pytest", "PIL"],
    noarchive=False,
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="IberostarGestorPedidos",
    icon=str(ROOT / "packaging" / "icon.ico"),
    version=str(ROOT / "build" / "version_info.txt"),
    console=False,
    upx=False,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    name="IberostarGestorPedidos",
    upx=False,
)
