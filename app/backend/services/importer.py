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
import shutil

import pdfplumber

from config.constants import (
    MONTHS,
    PDF_DATE_LABEL,
    PDF_DESTINATION_LABEL,
    PDF_IBS_CODE_LABEL,
    PDF_FILENAME_TEMPLATE,
)
from config.settings import INPUT_PDFS_DIR
from models.pdf_metadata import PDFMetadata
from services.registry import Registry


class Importer:
    """
    Servicio encargado de importar nuevos PDF.
    """

    def run(self) -> list[Path]:
        """
        Ejecuta el proceso de importación de nuevos PDF.

        Returns:
            Lista de archivos PDF importados correctamente.
        """

        registry = Registry()

        # ==================================================
        # SELECCIÓN DE ARCHIVOS
        # ==================================================

        print()
        print("=" * 100)
        print("1. SELECCIÓN DE ALBARANES PDF")
        print("=" * 100)
        print("Abriendo el explorador de archivos...")
        print("=" * 100)

        pdf_files = self._select_pdf_files()

        imported_pdf_files: list[Path] = []

        if not pdf_files:

            print()
            print("!" * 100)
            print("SELECCIÓN CANCELADA")
            print("!" * 100)
            print("No se seleccionó ningún archivo PDF.")
            print("No hay archivos para importar ni sincronizar.")
            print("!" * 100)

            return []

        # ==================================================
        # LISTA DE PDF ENCONTRADOS
        # ==================================================

        print()
        print("=" * 100)
        print("2. LISTA DE PDF ENCONTRADOS")
        print("=" * 100)

        for index, pdf_file in enumerate(
            pdf_files,
            start=1,
        ):
            print(f"{index:03d} | " f"{pdf_file.name}")

        print("-" * 100)
        print(f"Total de PDF seleccionados: {len(pdf_files)}")
        print("=" * 100)

        # Contadores exclusivamente informativos.
        imported_count = 0
        existing_count = 0
        error_count = 0

        # ==================================================
        # PROCESAMIENTO DE CADA PDF
        # ==================================================

        print()
        print("=" * 100)
        print("3. LECTURA DE METADATOS Y COMPROBACIÓN DEL REGISTRY")
        print("=" * 100)

        for index, pdf_file in enumerate(
            pdf_files,
            start=1,
        ):

            try:

                print()
                print("-" * 100)
                print(f"PDF {index:03d} DE {len(pdf_files):03d} " f"| {pdf_file.name}")
                print("-" * 100)

                # ==========================================
                # Lectura del PDF
                # ==========================================

                print("Proceso       : Leyendo contenido del PDF...")

                text = self._extract_text(
                    pdf_file,
                )

                print("Lectura       : COMPLETADA")

                # ==========================================
                # Extracción de metadatos
                # ==========================================

                print("Proceso       : Extrayendo metadatos...")

                metadata = self._extract_metadata(
                    text,
                )

                print("Metadatos     : ENCONTRADOS")
                print(f"Archivo       : {pdf_file.name}")
                print(f"Código IBS    : {metadata.ibs_code}")
                print(f"Fecha         : {metadata.delivery_date}")
                print(f"Punto de venta: {metadata.sales_point}")

                # ==========================================
                # Construcción del destino
                # ==========================================

                pdf_name = self._build_pdf_filename(
                    metadata.ibs_code,
                )

                pdf_path = self._build_destination_path(
                    metadata.delivery_date,
                    pdf_name,
                )

                self._create_directory(
                    pdf_path,
                )

                print(f"Nombre final  : {pdf_name}")
                print(f"Carpeta       : {pdf_path.parent}")

                # ==========================================
                # Comprobación del Registry
                # ==========================================

                print("Registry      : Comprobando código IBS...")

                if registry.exists(
                    metadata.ibs_code,
                ):

                    existing_count += 1

                    print(f"IBS registrado: SÍ")
                    print(f"Estado        : IGNORADO")
                    print("Resultado     : El PDF ya fue importado " "anteriormente.")

                else:

                    print(f"IBS registrado: NO")
                    print("Estado        : NUEVO")

                    # ======================================
                    # Copia del PDF
                    # ======================================

                    print("Proceso       : Copiando y renombrando PDF...")

                    self._copy_pdf(
                        pdf_file,
                        pdf_path,
                    )

                    print("Copia         : COMPLETADA")
                    print(f"Destino       : {pdf_path}")

                    # ======================================
                    # Registro del PDF
                    # ======================================

                    print("Proceso       : Registrando PDF...")

                    registry.register(
                        metadata,
                        pdf_path,
                    )

                    imported_pdf_files.append(
                        pdf_path,
                    )

                    imported_count += 1

                    print("Registry      : REGISTRADO")
                    print("Sincronización: PENDIENTE")
                    print("Resultado     : IMPORTADO CORRECTAMENTE")

                print("-" * 100)

            except Exception as error:

                error_count += 1

                print()
                print("!" * 100)
                print(f"ERROR DURANTE LA IMPORTACIÓN " f"DEL PDF {index:03d}")
                print("!" * 100)
                print(f"Archivo       : {pdf_file.name}")
                print("Estado        : ERROR")
                print(f"Motivo        : {error}")
                print("Resultado     : PDF NO IMPORTADO")
                print("!" * 100)

        # ==================================================
        # GUARDADO DEL REGISTRY
        # ==================================================

        print()
        print("=" * 100)
        print("4. GUARDADO DEL REGISTRY")
        print("=" * 100)
        print("Proceso       : Guardando cambios del Registry...")

        registry.save()

        print("Estado        : GUARDADO CORRECTAMENTE")
        print("=" * 100)

        # ==================================================
        # RESUMEN DE IMPORTACIÓN
        # ==================================================

        print()
        print("=" * 100)
        print("5. RESUMEN DE IMPORTACIÓN")
        print("=" * 100)
        print(f"PDF seleccionados : {len(pdf_files)}")
        print(f"PDF nuevos        : {imported_count}")
        print(f"PDF ya registrados: {existing_count}")
        print(f"PDF con errores   : {error_count}")
        print("-" * 100)

        if imported_pdf_files:

            print("PDF enviados al sincronizador:")

            for index, pdf_path in enumerate(
                imported_pdf_files,
                start=1,
            ):
                print(f"{index:03d} | " f"{pdf_path.name}")

        else:

            print("PDF enviados al sincronizador: NINGUNO")

        print("=" * 100)

        return imported_pdf_files

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
        root.attributes(
            "-topmost",
            True,
        )

        files = filedialog.askopenfilenames(
            title="Seleccionar albaranes PDF",
            filetypes=[
                (
                    "Archivos PDF",
                    "*.pdf",
                ),
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

        with pdfplumber.open(
            pdf_file,
        ) as pdf:

            for page in pdf.pages:

                text = page.extract_text()

                if text:
                    pages.append(
                        text,
                    )

        return "\n".join(
            pages,
        )

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

            if line.startswith(
                PDF_IBS_CODE_LABEL,
            ):

                ibs_code = line.replace(
                    PDF_IBS_CODE_LABEL,
                    "",
                ).strip()

            elif line.startswith(
                PDF_DATE_LABEL,
            ):

                delivery_date = line.replace(
                    PDF_DATE_LABEL,
                    "",
                ).strip()

            elif line.startswith(
                PDF_DESTINATION_LABEL,
            ):

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
            raise ValueError(
                "Código IBS no encontrado.",
            )

        if delivery_date is None:
            raise ValueError(
                "Fecha no encontrada.",
            )

        if sales_point is None:
            raise ValueError(
                "Punto de venta no encontrado.",
            )

        return PDFMetadata(
            ibs_code=ibs_code,
            delivery_date=delivery_date,
            sales_point=sales_point,
        )

    def _build_pdf_filename(
        self,
        ibs_code: str,
    ) -> str:
        """
        Construye el nombre estándar de un PDF.

        Args:
            ibs_code: Código IBS.

        Returns:
            Nombre del archivo PDF.
        """

        return PDF_FILENAME_TEMPLATE.format(
            ibs_code=ibs_code,
        )

    def _build_destination_path(
        self,
        delivery_date: str,
        pdf_name: str,
    ) -> Path:
        """
        Construye la ruta donde debe almacenarse un PDF.

        Args:
            delivery_date: Fecha del albarán.
            pdf_name: Nombre final del PDF.

        Returns:
            Ruta completa donde se almacenará el PDF.
        """

        parsed_date = datetime.strptime(
            delivery_date,
            "%d/%m/%Y",
        )

        year = str(
            parsed_date.year,
        )

        month = MONTHS[parsed_date.month - 1]

        day = str(
            parsed_date.day,
        )

        return INPUT_PDFS_DIR / year / month / day / pdf_name

    def _create_directory(
        self,
        pdf_path: Path,
    ) -> None:
        """
        Crea una carpeta si no existe.

        Args:
            pdf_path: Ruta del PDF.
        """

        pdf_path.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

    def _copy_pdf(
        self,
        source: Path,
        destination: Path,
    ) -> None:
        """
        Copia un PDF a la carpeta de destino.

        Args:
            source: PDF original.
            destination: Ruta donde se copiará el PDF.
        """

        shutil.copy2(
            source,
            destination,
        )
