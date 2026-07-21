"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    product.py

Descripción:
    Modelo que representa un producto procedente del Excel de Economato.

    El código del artículo es el identificador principal utilizado para
    localizar el producto dentro de los Excel mensuales y las plantillas.
"""

from dataclasses import dataclass


@dataclass(slots=True)
class Product:
    """
    Representa un producto suministrado a un punto de venta.
    """

    code: str
    name: str
    format: str
    price: float
    quantity: float
