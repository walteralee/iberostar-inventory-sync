"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    registry.py

Descripción:
    Servicio encargado de gestionar el registro de albaranes
    importados.
"""

from pathlib import Path
import json

from config.settings import REGISTRY_FILE
from models.pdf_metadata import PDFMetadata

from datetime import datetime

from config.constants import MONTHS


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

        if not self.exists(ibs_code):
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

        if self.exists(ibs_code):

            self._data[ibs_code]["synchronized"] = True

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

    def save(self) -> None:
        """
        Guarda el registro en disco.
        """

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

        if not REGISTRY_FILE.exists():

            REGISTRY_FILE.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            REGISTRY_FILE.write_text(
                "{}",
                encoding="utf-8",
            )

            return {}

        with REGISTRY_FILE.open(
            "r",
            encoding="utf-8",
        ) as file:

            try:

                return json.load(file)

            except json.JSONDecodeError:

                return {}
