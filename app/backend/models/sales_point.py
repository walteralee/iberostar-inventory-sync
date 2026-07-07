"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    sales_point.py

Descripción:
    Modelo que representa un punto de venta.
"""

from dataclasses import dataclass


@dataclass(slots=True, frozen=True)
class SalesPoint:
    """
    Representa un punto de venta.
    """

    name: str

    def __str__(self) -> str:
        return self.name