"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    delivery.py

Descripción:
    Modelo que representa un albarán de entrega.

    Un albarán pertenece a un único punto de venta y contiene todos los
    productos suministrados en dicho movimiento.
"""

from dataclasses import dataclass, field
from datetime import date

from models.product import Product
from models.sales_point import SalesPoint


@dataclass(slots=True)
class Delivery:
    """
    Representa un albarán de entrega.
    """

    sales_point: SalesPoint
    delivery_date: date
    products: list[Product] = field(default_factory=list)