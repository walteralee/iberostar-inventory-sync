"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    importer.py

Descripción:
    Servicio encargado de importar nuevos albaranes PDF al sistema.
"""

from datetime import datetime
from pathlib import Path
from tkinter import Tk, filedialog
from services.registry import Registry

import shutil

import pdfplumber


from config.constants import (
    MONTHS,
    PDF_DATE_LABEL,
    PDF_DESTINATION_LABEL,
    PDF_IBS_CODE_LABEL,
)
from config.settings import INPUT_PDFS_DIR
from models.pdf_metadata import PDFMetadata


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
        registry = Registry()

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

            try:

                text = self._extract_text(pdf_file)

                metadata = self._extract_metadata(text)

                destination = self._build_destination_directory(
                    metadata.delivery_date,
                )

                self._create_directory(destination)

                if registry.exists(metadata.ibs_code):

                    status = "YA EXISTE"

                else:

                    self._copy_pdf(
                        pdf_file,
                        destination,
                    )

                    registry.register(
                        metadata,
                        destination / pdf_file.name,
                    )

                    status = "COPIADO"

                print("\n" + "=" * 100)
                print(f"PDF {index}")
                print("=" * 100)
                print(f"Archivo       : {pdf_file.name}")
                print(f"Código IBS    : {metadata.ibs_code}")
                print(f"Fecha         : {metadata.delivery_date}")
                print(f"Punto de venta: {metadata.sales_point}")
                print(f"Destino       : {destination}")
                print(f"Estado        : {status}")

            except Exception as error:

                print("\n" + "=" * 100)
                print(f"PDF {index}")
                print("=" * 100)
                print(f"Archivo       : {pdf_file.name}")
                print(f"Estado        : ERROR")
                print(f"Motivo        : {error}")

        registry.save()

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

    def _extract_metadata(
        self,
        text: str,
    ) -> PDFMetadata:
        """
        Extrae toda la información necesaria del PDF
        recorriendo el texto una única vez.
        """

        ibs_code = None
        delivery_date = None
        sales_point = None

        for line in text.splitlines():

            line = line.strip()

            if line.startswith(PDF_IBS_CODE_LABEL):

                ibs_code = line.replace(
                    PDF_IBS_CODE_LABEL,
                    "",
                ).strip()

            elif line.startswith(PDF_DATE_LABEL):

                delivery_date = line.replace(
                    PDF_DATE_LABEL,
                    "",
                ).strip()

            elif line.startswith(PDF_DESTINATION_LABEL):

                sales_point = line.replace(
                    PDF_DESTINATION_LABEL,
                    "",
                ).strip()

            if (
                ibs_code is not None
                and delivery_date is not None
                and sales_point is not None
            ):
                break

        if ibs_code is None:
            raise ValueError("Código IBS no encontrado.")

        if delivery_date is None:
            raise ValueError("Fecha no encontrada.")

        if sales_point is None:
            raise ValueError("Punto de venta no encontrado.")

        return PDFMetadata(
            ibs_code=ibs_code,
            delivery_date=delivery_date,
            sales_point=sales_point,
        )

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
            pdf_file: PDF seleccionado.
            destination: Carpeta destino.
        """

        shutil.copy2(
            pdf_file,
            destination / pdf_file.name,
        )
