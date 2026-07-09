"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    importer.py

Descripción:
    Servicio encargado de importar nuevos albaranes PDF al sistema.
"""

from pathlib import Path
from tkinter import Tk, filedialog
from datetime import datetime

from config.constants import MONTHS
from config.settings import INPUT_PDFS_DIR
from config.constants import (
    PDF_IBS_CODE_LABEL,
    PDF_DATE_LABEL,
    PDF_DESTINATION_LABEL,
)

import pdfplumber
import shutil


class Importer:
    """
    Servicio encargado de importar nuevos PDF.
    """

    def run(self) -> list[Path]:
        """
        Ejecuta el proceso de importación de nuevos PDF.

        Returns:
            Lista de archivos PDF seleccionados.
        """

        pdf_files = self._select_pdf_files()

        if not pdf_files:

            print("\nNo se seleccionó ningún archivo PDF.\n")

            return []

        print("\n" + "=" * 100)
        print("PDF SELECCIONADOS")
        print("=" * 100)

        for index, pdf in enumerate(pdf_files, start=1):

            print(f"{index:03d} | {pdf}")

        print("=" * 100)

        # ==================================================
        # Extraer y mostrar la información básica de cada
        # PDF seleccionado.
        # ==================================================

        for index, pdf_file in enumerate(pdf_files, start=1):

            text = self._extract_text(pdf_file)

            ibs_code = self._extract_ibs_code(text)
            delivery_date = self._extract_date(text)
            sales_point = self._extract_sales_point(text)

            destination = self._build_destination_directory(
                delivery_date,
            )

            self._create_directory(destination)

            if self._delivery_exists(
                ibs_code,
                destination,
            ):

                status = "YA EXISTE"

            else:

                self._copy_pdf(
                    pdf_file,
                    destination,
                )

                status = "COPIADO"

            print("\n" + "=" * 100)
            print(f"PDF {index}")
            print("=" * 100)
            print(f"Archivo       : {pdf_file.name}")
            print(f"Código IBS    : {ibs_code}")
            print(f"Fecha         : {delivery_date}")
            print(f"Punto de venta: {sales_point}")
            print(f"Destino       : {destination}")
            print(f"Estado        : {status}")

        return pdf_files

    # ======================================================
    # PRIVATE
    # ======================================================

    def _select_pdf_files(self) -> list[Path]:
        """
        Abre el explorador de Windows para seleccionar PDF.

        Returns:
            Lista de rutas seleccionadas.
        """

        root = Tk()
        root.withdraw()
        root.attributes("-topmost", True)

        files = filedialog.askopenfilenames(
            title="Seleccionar albaranes PDF",
            filetypes=[
                ("Archivos PDF", "*.pdf"),
            ],
        )

        root.destroy()

        return [Path(file) for file in files]

    def _extract_text(
        self,
        pdf_file: Path,
    ) -> str:
        """
        Extrae el texto completo de un PDF.

        Args:
            pdf_file: Ruta del archivo PDF.

        Returns:
            Texto completo del PDF.
        """

        pages = []

        with pdfplumber.open(pdf_file) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:
                    pages.append(text)

        return "\n".join(pages)

    def _extract_ibs_code(
        self,
        text: str,
    ) -> str:
        """
        Extrae el código IBS del PDF.

        Args:
            text: Texto completo del PDF.

        Returns:
            Código IBS.
        """

        for line in text.splitlines():

            if line.startswith(PDF_IBS_CODE_LABEL):

                return line.replace(PDF_IBS_CODE_LABEL, "").strip()

        raise ValueError("Código IBS no encontrado.")

    def _extract_date(
        self,
        text: str,
    ) -> str:
        """
        Extrae la fecha del PDF.

        Args:
            text: Texto completo del PDF.

        Returns:
            Fecha del albarán.
        """

        for line in text.splitlines():

            if line.startswith(PDF_DATE_LABEL):

                return line.replace(PDF_DATE_LABEL, "").strip()

        raise ValueError("Fecha no encontrada.")

    def _extract_sales_point(
        self,
        text: str,
    ) -> str:
        """
        Extrae el punto de venta del PDF.

        Args:
            text: Texto completo del PDF.

        Returns:
            Nombre del punto de venta.
        """

        for line in text.splitlines():

            if line.startswith(PDF_DESTINATION_LABEL):

                return line.replace(PDF_DESTINATION_LABEL, "").strip()

        raise ValueError("Punto de venta no encontrado.")

    def _build_destination_directory(
        self,
        delivery_date: str,
    ) -> Path:
        """
        Construye la ruta donde debe almacenarse un PDF.

        Args:
            delivery_date: Fecha del albarán.

        Returns:
            Ruta de destino.
        """

        parsed_date = datetime.strptime(
            delivery_date,
            "%d/%m/%Y",
        )

        year = str(parsed_date.year)

        month = MONTHS[parsed_date.month - 1]

        day = str(parsed_date.day)

        return INPUT_PDFS_DIR / year / month / day

    def _create_directory(
        self,
        directory: Path,
    ) -> None:
        """
        Crea una carpeta si no existe.

        Args:
            directory: Carpeta a crear.
        """

        directory.mkdir(
            parents=True,
            exist_ok=True,
        )

    def _copy_pdf(
        self,
        pdf_file: Path,
        destination: Path,
    ) -> None:
        """
        Copia un PDF a la carpeta de destino.

        Args:
            ibs_code: Código IBS del albarán.
            destination: Carpeta destino.
        """

        shutil.copy2(
            pdf_file,
            destination / pdf_file.name,
        )

    def _delivery_exists(
        self,
        ibs_code: str,
        destination: Path,
    ) -> bool:
        """
        Comprueba si el albarán ya existe en la carpeta destino.

        Args:
            pdf_file: PDF seleccionado.
            destination: Carpeta destino.

        Returns:
            True si existe.
        """

        return (destination / pdf_file.name).exists()
