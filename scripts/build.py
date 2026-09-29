"""
Compila la aplicación de escritorio y sus paquetes de distribución.

    python scripts/build.py

Genera en dist/:
    IberostarGestorPedidos/               aplicación (carpeta)
    IberostarGestorPedidos-Portable.zip   para llevar en un USB
    IberostarGestorPedidos-Setup.exe      instalador (si Inno Setup está
                                          instalado)

Los nombres no llevan versión para que el enlace de descarga directa del
README (releases/latest/download/...) sea siempre el mismo.

La versión se toma de app/backend/config/constants.py (PROJECT_VERSION),
que es la única fuente de verdad.
"""

from __future__ import annotations

from pathlib import Path
import os
import re
import shutil
import subprocess
import sys

ROOT = Path(__file__).resolve().parents[1]
BUILD_DIR = ROOT / "build"
DIST_DIR = ROOT / "dist"
APP_NAME = "IberostarGestorPedidos"

ISCC_CANDIDATES = (
    Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "Inno Setup 6" / "ISCC.exe",
    Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "Inno Setup 6" / "ISCC.exe",
    Path(os.environ.get("LOCALAPPDATA", "")) / "Programs" / "Inno Setup 6" / "ISCC.exe",
)


def read_version() -> str:
    constants = (ROOT / "app" / "backend" / "config" / "constants.py").read_text(
        encoding="utf-8"
    )
    match = re.search(r'PROJECT_VERSION\s*=\s*"(\d+)\.(\d+)\.(\d+)"', constants)

    if not match:
        sys.exit("No se encontró PROJECT_VERSION (formato X.Y.Z) en constants.py")

    return ".".join(match.groups())


def write_version_info(version: str) -> None:
    """Metadatos que Windows muestra en Propiedades > Detalles del .exe."""

    numbers = ", ".join(version.split(".") + ["0"])
    template = (ROOT / "packaging" / "version_info.template.txt").read_text(
        encoding="utf-8"
    )

    BUILD_DIR.mkdir(exist_ok=True)
    (BUILD_DIR / "version_info.txt").write_text(
        template.replace("{VERSION_TUPLE}", numbers).replace("{VERSION}", version),
        encoding="utf-8",
    )


def run(command: list[str]) -> None:
    print(f"\n> {' '.join(command)}", flush=True)
    subprocess.run(command, cwd=ROOT, check=True)


def find_iscc() -> Path | None:
    on_path = shutil.which("iscc")

    if on_path:
        return Path(on_path)

    return next((path for path in ISCC_CANDIDATES if path.is_file()), None)


def main() -> None:
    version = read_version()
    print(f"Compilando {APP_NAME} v{version}")

    write_version_info(version)

    run([
        sys.executable,
        "-m",
        "PyInstaller",
        str(ROOT / "packaging" / "iberostar.spec"),
        "--noconfirm",
        "--clean",
        "--distpath",
        str(DIST_DIR),
        "--workpath",
        str(BUILD_DIR / "pyinstaller"),
    ])

    portable = DIST_DIR / f"{APP_NAME}-Portable"
    print(f"\nCreando {portable.name}.zip", flush=True)
    shutil.make_archive(str(portable), "zip", DIST_DIR, APP_NAME)

    iscc = find_iscc()

    if iscc is None:
        print(
            "\nInno Setup no está instalado: se omite el instalador.\n"
            "Descárgalo de https://jrsoftware.org/isdl.php para generarlo."
        )
    else:
        run([
            str(iscc),
            f"/DAppVersion={version}",
            f"/O{DIST_DIR}",
            str(ROOT / "packaging" / "installer.iss"),
        ])

    print("\nListo. Resultados en dist/:")
    for item in sorted(DIST_DIR.iterdir()):
        print(f"  {item.name}")


if __name__ == "__main__":
    main()
