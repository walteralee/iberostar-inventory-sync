"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    constants.py

Descripción:
    Constantes globales del proyecto.

    Este archivo contiene únicamente valores fijos del dominio del proyecto.
    No debe contener rutas ni configuraciones modificables.
"""

# ==========================================================
# PROJECT
# ==========================================================

PROJECT_NAME = "Iberostar Inventory Synchronizer"
PROJECT_VERSION = "1.0.0"

# ==========================================================
# SALES POINTS
# ==========================================================

BAR_PISCINA = "Bar_Piscina"
BAR_SALON = "Bar_Salon"
COMEDOR = "Comedor"
MANOLETE = "Manolete"
STARCAFE = "Starcafe"

SALES_POINTS = (
    BAR_PISCINA,
    BAR_SALON,
    COMEDOR,
    MANOLETE,
    STARCAFE,
)

# ==========================================================
# PDF -> EXCEL MAPPING
# ==========================================================

SALES_POINT_MAPPING = {
    "XAN - Bar Piscina": BAR_PISCINA,
    "XAN - Bar Salón": BAR_SALON,
    "XAN - Comedor": COMEDOR,
    "XAN - Manolete": MANOLETE,
    "XAN - Starcafé": STARCAFE,
}

VALID_PDF_DEPARTMENTS = frozenset(
    SALES_POINT_MAPPING.keys()
)

# ==========================================================
# EXCEL
# ==========================================================

EXCEL_SHEET_NAME = "extraccion"

DAY_HEADER_ROW = 6
FIRST_PRODUCT_ROW = 7

PRODUCT_CODE_COLUMN = 1        # A
PRODUCT_NAME_COLUMN = 2        # B
CURRENT_STOCK_COLUMN = 3       # C

FIRST_DAY_COLUMN = 7           # G (día 1)

# ==========================================================
# PDF
# ==========================================================

PDF_DATE_LABEL = "Fecha:"
PDF_DESTINATION_LABEL = "Dpto. Destino:"

# ==========================================================
# FILE EXTENSIONS
# ==========================================================

PDF_EXTENSION = ".pdf"
EXCEL_EXTENSION = ".xlsx"

# ==========================================================
# MONTHS
# ==========================================================

MONTHS = (
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE",
)