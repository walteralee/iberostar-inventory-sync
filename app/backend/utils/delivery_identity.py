"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    delivery_identity.py

Descripción:
    Construcción única de la clave estable y la huella SHA-256 de
    una entrega, compartida por Registry y Synchronizer.

    Antes de existir este módulo, ambas clases mantenían copias
    independientes de esta lógica. Como el mecanismo de idempotencia
    del proyecto depende de que las dos huellas coincidan siempre
    byte a byte, cualquier divergencia futura entre las copias
    habría podido romper la recuperación de entregas ya aplicadas.
"""

from __future__ import annotations

import hashlib
import json
import unicodedata
from datetime import datetime
from math import isfinite
from typing import TYPE_CHECKING

if TYPE_CHECKING:
    from models.delivery import Delivery


def normalize_key_text(value: object) -> str:
    """
    Normaliza un texto para utilizarlo dentro de una clave estable.
    """

    normalized_value = unicodedata.normalize(
        "NFKD",
        str(value).strip(),
    )
    normalized_value = "".join(
        character
        for character in normalized_value
        if not unicodedata.combining(character)
    )
    return " ".join(normalized_value.casefold().split())


def normalize_payload_text(value: object) -> str:
    """
    Limpia espacios sin alterar mayúsculas, acentos ni contenido.
    """

    return " ".join(str(value).strip().split())


def canonical_number(value: object) -> str:
    """
    Produce la representación numérica estable usada en el hash.
    """

    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ValueError("El valor numérico debe ser int o float.")

    numeric_value = float(value)

    if not isfinite(numeric_value):
        raise ValueError("El valor numérico debe ser finito.")

    if numeric_value == 0:
        return "0"

    return format(numeric_value, ".15g")


def build_delivery_key(delivery: "Delivery") -> str:
    """
    Genera la clave estable ``fecha|punto_de_venta`` de una entrega.
    """

    delivery_date = delivery.delivery_date

    if isinstance(delivery_date, datetime):
        delivery_date = delivery_date.date()

    sales_point_name = normalize_key_text(delivery.sales_point.name)

    if not sales_point_name:
        raise ValueError("El punto de venta de la entrega está vacío.")

    return f"{delivery_date.isoformat()}|{sales_point_name}"


def build_payload_hash(delivery: "Delivery") -> str:
    """
    Genera la huella SHA-256 estable del contenido completo de una entrega.
    """

    products = sorted(
        (
            {
                "code": product.code.strip(),
                "name": normalize_payload_text(product.name),
                "format": normalize_payload_text(product.format),
                "price": canonical_number(product.price),
                "quantity": canonical_number(product.quantity),
            }
            for product in delivery.products
        ),
        key=lambda item: item["code"],
    )

    payload = {
        "delivery_key": build_delivery_key(delivery),
        "products": products,
    }

    serialized = json.dumps(
        payload,
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )

    return hashlib.sha256(serialized.encode("utf-8")).hexdigest()
