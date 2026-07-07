"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    reader.py

Descripción:
    Lector de archivos PDF.

    Este módulo únicamente abre un archivo PDF y devuelve el documento para
    que posteriormente sea procesado por el parser.
"""

from pathlib import Path

import pdfplumber


class PDFReader:
    """
    Lector de documentos PDF.
    """

    def read(self, pdf_path: Path) -> pdfplumber.PDF:
        """
        Abre un documento PDF.

        Args:
            pdf_path: Ruta del archivo PDF.

        Returns:
            Documento PDF abierto.
        """

        return pdfplumber.open(pdf_path)