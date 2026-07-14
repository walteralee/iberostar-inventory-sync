"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    product.py

Descripción:
    Modelo que representa un producto extraído de un albarán.

    El código del artículo es el identificador principal utilizado para
    localizar el producto dentro del Excel.
"""

from dataclasses import dataclass


@dataclass(slots=True)
class Product:
    """
    Representa un producto de un albarán.
    """

    code: int
    name: str
    format: str
    price: float
    quantity: float
