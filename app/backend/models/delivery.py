"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    delivery.py

Descripción:
    Modelo que representa una entrega de Economato.

    Cada entrega pertenece a un único punto de venta, corresponde
    a una única fecha y agrupa todos los productos suministrados
    en dicho movimiento.
"""

from dataclasses import dataclass, field
from datetime import date

from models.product import Product
from models.sales_point import SalesPoint


@dataclass(slots=True)
class Delivery:
    """
    Representa una entrega de Economato
    correspondiente a un punto de venta
    en una fecha determinada.
    """

    sales_point: SalesPoint
    delivery_date: date
    products: list[Product] = field(default_factory=list)
