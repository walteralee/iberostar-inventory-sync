"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    product_codes.py

Descripción:
    Normalización única de códigos de producto, compartida por
    Importer, ExcelFinder y ProductManager.

    Antes de existir este módulo, cada clase tenía su propia versión
    de esta normalización con reglas ligeramente distintas (por
    ejemplo, solo una reconocía sufijos como ".00" o ".000"), lo que
    podía hacer que un mismo código no se reconociera como el mismo
    producto en dos sitios distintos y se creara una fila duplicada.
"""

from __future__ import annotations

import re
from math import isfinite

_DECIMAL_ZERO_SUFFIX = re.compile(r"\.0+$")


def normalize_product_code(value: object) -> str | None:
    """
    Normaliza una celda a un código de producto de solo dígitos.

    Acepta enteros, floats enteros y texto (incluyendo sufijos
    decimales compuestos solo por ceros, como ".0", ".00", ".000").

    Returns:
        El código normalizado, o ``None`` si el valor no puede
        interpretarse como un código de producto.
    """

    if value is None or isinstance(value, bool):
        return None

    if isinstance(value, int):
        return str(value)

    if isinstance(value, float):
        if not isfinite(value) or not value.is_integer():
            return None

        return str(int(value))

    text = str(value).strip()

    if not text:
        return None

    text = _DECIMAL_ZERO_SUFFIX.sub("", text)

    if not text.isdigit():
        return None

    return text
