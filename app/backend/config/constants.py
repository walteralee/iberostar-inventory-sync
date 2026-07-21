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
PROJECT_VERSION = "2.0.0"

# ==========================================================
# SALES POINTS
# ==========================================================

BAR_PISCINA = "Bar_Piscina"
BAR_SALON = "Bar_Salon"
COMEDOR = "Comedor"
MANOLETE = "Manolete"
STARCAFE = "Starcafe"
GENERALES = "Generales"

SALES_POINTS = (
    BAR_PISCINA,
    BAR_SALON,
    COMEDOR,
    MANOLETE,
    STARCAFE,
    GENERALES,
)

# ==========================================================
# SOURCE EXCEL -> SALES POINT MAPPING
# ==========================================================

SALES_POINT_MAPPING = {
    "BAR PISCINA": BAR_PISCINA,
    "BAR SALÓN": BAR_SALON,
    "COMEDOR": COMEDOR,
    "MANOLETE": MANOLETE,
    "STARCAFÉ": STARCAFE,
    "GENERALES BAR Y COMEDOR": GENERALES,
}

VALID_SALES_POINTS = frozenset(SALES_POINT_MAPPING.keys())

# ==========================================================
# VALID PRODUCT GROUPS
# ==========================================================

VALID_PRODUCT_GROUPS = frozenset(
    {
        "BEBIDAS",
        "ENVASES",
        "DROGUERÍA",
        "ALIMENTOS/COMIDA",
    }
)

# ==========================================================
# SOURCE EXCEL
# ==========================================================

SOURCE_EXCEL_EXTENSION = ".xlsx"

SOURCE_HEADER_ROW = 2

SOURCE_DATE_COLUMN = 1  # A
SOURCE_SALES_POINT_COLUMN = 6  # F
SOURCE_GROUP_COLUMN = 7  # G
SOURCE_PRODUCT_CODE_COLUMN = 10  # J
SOURCE_PRODUCT_NAME_COLUMN = 12  # L
SOURCE_FORMAT_COLUMN = 13  # M
SOURCE_QUANTITY_COLUMN = 16  # P
SOURCE_PRICE_COLUMN = 17  # Q

SOURCE_SALES_POINT_PREFIX = "XAN - "

# ==========================================================
# MONTHLY EXCEL
# ==========================================================

EXCEL_SHEET_NAME = "extraccion"

DAY_HEADER_ROW = 6
FIRST_PRODUCT_ROW = 7

PRODUCT_CODE_COLUMN = 1  # A
PRODUCT_NAME_COLUMN = 2  # B
CURRENT_STOCK_COLUMN = 3  # C
PRODUCT_FORMAT_COLUMN = 4  # D
PRODUCT_PRICE_COLUMN = 5  # E

FIRST_DAY_COLUMN = 7  # G

MONTH_EXTRACTION_COLUMN = 38  # AL
TOTAL_STOCK_COLUMN = 39  # AM
VALUE_COLUMN = 40  # AN

# ==========================================================
# REGISTRY
# ==========================================================

REGISTRY_FILENAME = "imported_deliveries.json"

# ==========================================================
# FILE EXTENSIONS
# ==========================================================

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

# ==========================================================
# EXCEL TEMPLATES
# ==========================================================

EXCEL_TEMPLATES = (
    BAR_PISCINA,
    BAR_SALON,
    COMEDOR,
    GENERALES,
    MANOLETE,
    STARCAFE,
)
