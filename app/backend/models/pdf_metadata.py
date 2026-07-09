from dataclasses import dataclass


@dataclass(slots=True)
class PDFMetadata:
    """
    Información básica necesaria para importar un PDF.
    """

    ibs_code: str
    delivery_date: str
    sales_point: str
