"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    helpers.py

Descripción:
    Funciones auxiliares reutilizables por distintos módulos del proyecto.
"""

from datetime import date, datetime

from config.constants import MONTHS


def normalize_text(text: str) -> str:
    """
    Normaliza un texto eliminando espacios sobrantes.

    Args:
        text: Texto a normalizar.

    Returns:
        Texto normalizado.
    """

    return " ".join(text.strip().split())


def parse_date(date_string: str) -> date:
    """
    Convierte una fecha en formato DD/MM/YYYY a un objeto date.

    Args:
        date_string: Fecha en formato DD/MM/YYYY.

    Returns:
        Objeto date.
    """

    return datetime.strptime(
        date_string,
        "%d/%m/%Y"
    ).date()


def get_day(value: date) -> int:
    """
    Devuelve el día del mes.

    Args:
        value: Fecha.

    Returns:
        Día del mes.
    """

    return value.day


def get_month_name(value: date) -> str:
    """
    Devuelve el nombre del mes en mayúsculas.

    Args:
        value: Fecha.

    Returns:
        Nombre del mes.
    """

    return MONTHS[value.month - 1]


def get_year(value: date) -> str:
    """
    Devuelve el año.

    Args:
        value: Fecha.

    Returns:
        Año en formato texto.
    """

    return str(value.year)