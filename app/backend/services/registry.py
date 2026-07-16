"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    registry.py

Descripción:
    Servicio encargado de gestionar el registro de albaranes
    importados.
"""

from datetime import datetime
from pathlib import Path
import json

from config.constants import MONTHS
from config.settings import REGISTRY_FILE
from models.pdf_metadata import PDFMetadata


class Registry:
    """
    Gestiona el registro de albaranes importados.
    """

    def __init__(self) -> None:
        """
        Inicializa el registro cargando el fichero JSON.
        """

        self._data = self._load()

    # ======================================================
    # PUBLIC
    # ======================================================

    @property
    def data(self) -> dict:
        """
        Devuelve el contenido completo del registro.
        """

        return self._data

    def exists(
        self,
        ibs_code: str,
    ) -> bool:
        """
        Comprueba si un albarán ya está registrado.

        Args:
            ibs_code: Código IBS del albarán.

        Returns:
            True si el albarán ya existe.
        """

        return ibs_code in self._data

    def is_synchronized(
        self,
        ibs_code: str,
    ) -> bool:
        """
        Comprueba si un albarán ya ha sido sincronizado.

        Args:
            ibs_code: Código IBS.

        Returns:
            True si ya fue sincronizado.
        """

        if not self.exists(
            ibs_code,
        ):
            return False

        return self._data[ibs_code]["synchronized"]

    def mark_as_synchronized(
        self,
        ibs_code: str,
    ) -> None:
        """
        Marca un albarán como sincronizado.

        Args:
            ibs_code: Código IBS.
        """

        print()
        print("-" * 100)
        print("ACTUALIZACIÓN DEL ESTADO DE SINCRONIZACIÓN")
        print("-" * 100)
        print(f"Código IBS    : {ibs_code}")
        print("Proceso       : Marcando albarán como sincronizado...")

        if self.exists(
            ibs_code,
        ):

            self._data[ibs_code]["synchronized"] = True

            print("IBS registrado: SÍ")
            print("Estado anterior: PENDIENTE")
            print("Estado actual  : SINCRONIZADO")

        else:

            print("IBS registrado: NO")
            print("Estado        : NO MODIFICADO")
            print(
                "Resultado     : No se puede marcar como sincronizado "
                "porque el IBS no existe."
            )

        print("-" * 100)

    def register(
        self,
        metadata: PDFMetadata,
        pdf_path: Path,
    ) -> None:
        """
        Registra un nuevo albarán importado.

        Args:
            metadata: Metadatos del PDF.
            pdf_path: Ruta donde se ha almacenado el PDF.
        """

        parsed_date = datetime.strptime(
            metadata.delivery_date,
            "%d/%m/%Y",
        )

        self._data[metadata.ibs_code] = {
            "delivery_date": metadata.delivery_date,
            "year": parsed_date.year,
            "month": MONTHS[parsed_date.month - 1],
            "sales_point": metadata.sales_point,
            "pdf": str(pdf_path),
            "synchronized": False,
        }

        print()
        print("-" * 100)
        print("REGISTRO DEL ALBARÁN")
        print("-" * 100)
        print(f"Código IBS     : {metadata.ibs_code}")
        print(f"Fecha          : {metadata.delivery_date}")
        print(f"Año            : {parsed_date.year}")
        print(f"Mes            : {MONTHS[parsed_date.month - 1]}")
        print(f"Punto de venta : {metadata.sales_point}")
        print(f"Archivo        : {pdf_path.name}")
        print("Sincronización : PENDIENTE")
        print("Estado         : AÑADIDO AL REGISTRY")
        print("-" * 100)

    def save(self) -> None:
        """
        Guarda el registro en disco.
        """

        print()
        print("-" * 100)
        print("ESCRITURA DEL REGISTRY")
        print("-" * 100)
        print(f"Archivo        : {REGISTRY_FILE.name}")
        print(f"Ruta           : {REGISTRY_FILE}")
        print(f"Registros      : {len(self._data)}")
        print("Proceso        : Guardando contenido...")

        with REGISTRY_FILE.open(
            "w",
            encoding="utf-8",
        ) as file:

            json.dump(
                self._data,
                file,
                indent=4,
                ensure_ascii=False,
            )

        print("Estado         : GUARDADO CORRECTAMENTE")
        print("-" * 100)

    # ======================================================
    # PRIVATE
    # ======================================================

    def _load(self) -> dict:
        """
        Carga el registro desde disco.

        Si el fichero no existe, lo crea automáticamente.

        Returns:
            Diccionario con el contenido del registro.
        """

        print()
        print("=" * 100)
        print("CARGA DEL REGISTRY")
        print("=" * 100)
        print(f"Archivo        : {REGISTRY_FILE.name}")
        print(f"Ruta           : {REGISTRY_FILE}")
        print("Proceso        : Comprobando archivo del Registry...")

        if not REGISTRY_FILE.exists():

            print("Archivo        : NO EXISTE")
            print("Proceso        : Creando nuevo Registry...")

            REGISTRY_FILE.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            REGISTRY_FILE.write_text(
                "{}",
                encoding="utf-8",
            )

            print("Estado         : REGISTRY CREADO")
            print("Registros      : 0")
            print("=" * 100)

            return {}

        print("Archivo        : ENCONTRADO")
        print("Proceso        : Leyendo registros...")

        with REGISTRY_FILE.open(
            "r",
            encoding="utf-8",
        ) as file:

            try:

                data = json.load(
                    file,
                )

                print("Lectura        : COMPLETADA")
                print(f"Registros      : {len(data)}")
                print("Estado         : REGISTRY CARGADO CORRECTAMENTE")
                print("=" * 100)

                return data

            except json.JSONDecodeError:

                print()
                print("!" * 100)
                print("ERROR DE LECTURA DEL REGISTRY")
                print("!" * 100)
                print(f"Archivo        : {REGISTRY_FILE.name}")
                print("Motivo         : El contenido JSON no es válido.")
                print("Resultado      : Se utilizará un Registry vacío.")
                print("Registros      : 0")
                print("!" * 100)

                return {}
