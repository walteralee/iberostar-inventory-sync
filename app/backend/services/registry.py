"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    registry.py

Descripción:
    Servicio encargado de mantener el registro persistente de las
    entregas importadas y sincronizadas.

    Cada entrega se identifica mediante la combinación de su fecha y su
    punto de venta. El registro conserva también el contenido completo de
    la entrega y una huella SHA-256, lo que permite:
        - Recuperar entregas pendientes sin volver a leer el Excel origen.
        - Detectar cambios de contenido bajo un mismo identificador.
        - Evitar que una entrega distinta se omita o sincronice por error.
        - Migrar de forma segura registros creados por versiones anteriores.
        - Guardar el JSON mediante sustitución atómica y copia de seguridad.
"""

from __future__ import annotations

from datetime import date, datetime, timezone
import hashlib
import json
import math
import os
from pathlib import Path
import shutil
from typing import Any
from uuid import uuid4

from config.constants import MONTHS
from config.settings import BACKUP_DIR, REGISTRY_FILE
from models.delivery import Delivery
from models.product import Product
from models.sales_point import SalesPoint
from utils.delivery_identity import (
    build_delivery_key,
    build_payload_hash,
    canonical_number,
    normalize_key_text,
    normalize_payload_text,
)


class RegistryConflictError(ValueError):
    """
    Se lanza cuando una entrega coincide en fecha y punto de venta con una
    ya registrada, pero su contenido (productos o cantidades) difiere.

    Es una subclase de ``ValueError`` para no romper a quienes ya capturan
    ese tipo genérico; permite a los llamantes que sí quieran distinguir
    este caso (por ejemplo, el Importer) omitir la entrega en conflicto sin
    detener el resto del proceso.
    """


class Registry:
    """
    Gestiona el estado persistente de las entregas.

    La API pública original se mantiene intacta:
        - ``exists(delivery)``
        - ``is_synchronized(delivery)``
        - ``register(delivery)``
        - ``mark_as_synchronized(delivery)``
        - ``save()``

    Además, el registro puede reconstruir las entregas pendientes mediante
    ``get_pending_deliveries()``.
    """

    _SCHEMA_VERSION = 2
    _HASH_ALGORITHM = "sha256"

    def __init__(
        self,
        *,
        registry_file: Path | str | None = None,
        backup_directory: Path | str | None = None,
        acquire_lock: bool = True,
    ) -> None:
        """
        Carga el Registry y, por defecto, bloquea el archivo para impedir
        que dos procesos modifiquen simultáneamente el mismo registro.

        Los parámetros opcionales están pensados para pruebas. La creación
        habitual continúa siendo ``Registry()``.
        """

        self._registry_file = Path(registry_file or REGISTRY_FILE)
        self._backup_directory = Path(backup_directory or (BACKUP_DIR / "registry"))
        self._lock_file = self._registry_file.with_suffix(
            f"{self._registry_file.suffix}.lock"
        )
        self._lock_handle = None
        self._closed = False

        if acquire_lock:
            self._acquire_process_lock()

        try:
            self._data: dict[str, dict[str, Any]] = self._load()
        except Exception:
            self.close()
            raise

    # ======================================================
    # CONTEXT MANAGER / LIFECYCLE
    # ======================================================

    def __enter__(self) -> Registry:
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()

    def __del__(self) -> None:
        try:
            self.close()
        except Exception:
            # Nunca se deben propagar excepciones desde __del__.
            pass

    def close(self) -> None:
        """Libera el bloqueo del Registry de forma idempotente."""

        if self._closed:
            return

        self._closed = True

        if self._lock_handle is None:
            return

        try:
            if os.name == "nt":
                import msvcrt

                self._lock_handle.seek(0)
                msvcrt.locking(
                    self._lock_handle.fileno(),
                    msvcrt.LK_UNLCK,
                    1,
                )
            else:
                import fcntl

                fcntl.flock(
                    self._lock_handle.fileno(),
                    fcntl.LOCK_UN,
                )
        finally:
            self._lock_handle.close()
            self._lock_handle = None

    # ======================================================
    # PUBLIC
    # ======================================================

    @property
    def data(self) -> dict[str, dict[str, Any]]:
        """
        Devuelve el diccionario interno del registro.

        Se conserva como objeto mutable por compatibilidad con el mecanismo
        de snapshot y restauración utilizado por ``Synchronizer``.
        """

        return self._data

    def exists(
        self,
        delivery: Delivery,
    ) -> bool:
        """
        Comprueba si la entrega está registrada.

        Cuando la clave ya existe, también verifica que su contenido coincida
        con el almacenado. Si una versión antigua no guardó los productos,
        el registro se enriquece en memoria con la entrega recibida.

        Lanza ``RegistryConflictError`` (subclase de ``ValueError``) si la
        clave existe pero con productos o cantidades diferentes. Los
        llamantes que quieran omitir la entrega en conflicto en vez de
        detener todo el proceso deben capturar ese tipo específico.
        """

        normalized_delivery = self._validate_delivery(delivery)
        delivery_key = build_delivery_key(normalized_delivery)
        entry = self._data.get(delivery_key)

        if entry is None:
            return False

        self._validate_or_attach_payload(
            delivery_key=delivery_key,
            entry=entry,
            delivery=normalized_delivery,
        )
        return True

    def is_synchronized(
        self,
        delivery: Delivery,
    ) -> bool:
        """Devuelve ``True`` si la entrega existe y ya fue sincronizada."""

        normalized_delivery = self._validate_delivery(delivery)
        delivery_key = build_delivery_key(normalized_delivery)
        entry = self._data.get(delivery_key)

        if entry is None:
            return False

        self._validate_or_attach_payload(
            delivery_key=delivery_key,
            entry=entry,
            delivery=normalized_delivery,
        )

        return bool(entry.get("synchronized", False))

    def register(
        self,
        delivery: Delivery,
    ) -> None:
        """
        Registra una entrega completa como pendiente.

        Si la clave ya existe, solo se acepta cuando la huella de contenido
        coincide. Una entrega diferente con la misma fecha y punto de venta
        provoca un error para evitar sincronizaciones ambiguas.

        Nota: esto significa que un segundo informe de Economato para el
        mismo día y punto de venta, importado en una ejecución posterior
        a la que ya sincronizó el primero, se rechaza en vez de fusionarse
        (es una decisión de diseño para no arriesgar la integridad de
        datos, no un error). Solo se fusionan correctamente los productos
        cuando todos los movimientos de un mismo día/punto de venta se
        importan juntos en una sola ejecución.
        """

        normalized_delivery = self._validate_delivery(delivery)
        delivery_key = build_delivery_key(normalized_delivery)
        delivery_date = self._as_date(normalized_delivery.delivery_date)
        sales_point_name = normalized_delivery.sales_point.name.strip()

        existing_entry = self._data.get(delivery_key)

        if existing_entry is not None:
            self._validate_or_attach_payload(
                delivery_key=delivery_key,
                entry=existing_entry,
                delivery=normalized_delivery,
            )

            self._print_register_existing(
                delivery_key=delivery_key,
                delivery_date=delivery_date,
                sales_point_name=sales_point_name,
                synchronized=bool(existing_entry.get("synchronized", False)),
            )
            return

        now = self._utc_now()
        products_payload = self._serialize_products(normalized_delivery.products)
        payload_hash = build_payload_hash(normalized_delivery)

        self._data[delivery_key] = {
            "schema_version": self._SCHEMA_VERSION,
            "delivery_date": delivery_date.strftime("%d/%m/%Y"),
            "delivery_date_iso": delivery_date.isoformat(),
            "year": delivery_date.year,
            "month": MONTHS[delivery_date.month - 1],
            "sales_point": sales_point_name,
            "product_count": len(products_payload),
            "products": products_payload,
            "payload_available": True,
            "payload_hash": payload_hash,
            "hash_algorithm": self._HASH_ALGORITHM,
            "synchronized": False,
            "registered_at_utc": now,
            "updated_at_utc": now,
            "synchronized_at_utc": None,
        }

        self._print_registered(
            delivery_key=delivery_key,
            delivery_date=delivery_date,
            sales_point_name=sales_point_name,
            product_count=len(products_payload),
        )

    def mark_as_synchronized(
        self,
        delivery: Delivery,
    ) -> None:
        """
        Marca una entrega existente como sincronizada.

        La operación es estricta: si la entrega no existe o su contenido no
        coincide con el registrado, se genera un error y no se modifica nada.
        """

        normalized_delivery = self._validate_delivery(delivery)
        delivery_key = build_delivery_key(normalized_delivery)
        delivery_date = self._as_date(normalized_delivery.delivery_date)
        sales_point_name = normalized_delivery.sales_point.name.strip()
        entry = self._data.get(delivery_key)

        if entry is None:
            raise ValueError(
                "No se puede marcar como sincronizada una entrega que no "
                f"existe en el Registry: {delivery_key}."
            )

        self._validate_or_attach_payload(
            delivery_key=delivery_key,
            entry=entry,
            delivery=normalized_delivery,
        )

        previous_status = bool(entry.get("synchronized", False))
        now = self._utc_now()

        entry["synchronized"] = True
        entry["updated_at_utc"] = now

        if not previous_status or not entry.get("synchronized_at_utc"):
            entry["synchronized_at_utc"] = now

        self._print_synchronization_update(
            delivery_key=delivery_key,
            delivery_date=delivery_date,
            sales_point_name=sales_point_name,
            previous_status=previous_status,
        )

    def save(self) -> None:
        """
        Valida y guarda el registro mediante sustitución atómica.

        Antes de reemplazar el JSON vigente se conserva una copia de seguridad
        validada en ``storage/backup/registry``.
        """

        self._ensure_open()
        normalized_data = self._validate_registry_data(self._data)

        print()
        print("-" * 100)
        print("ESCRITURA DEL REGISTRY")
        print("-" * 100)
        print(f"Archivo        : {self._registry_file.name}")
        print(f"Ruta           : {self._registry_file}")
        print(f"Registros      : {len(normalized_data)}")
        print("Proceso        : Validando y guardando contenido...")

        self._atomic_write(normalized_data)

        # Conservamos exactamente el mismo objeto dict para no romper los
        # snapshots externos que trabajan con ``registry.data``.
        self._data.clear()
        self._data.update(normalized_data)

        print("Estado         : GUARDADO CORRECTAMENTE")
        print("-" * 100)

    def get_delivery(
        self,
        delivery_or_key: Delivery | str,
    ) -> Delivery | None:
        """
        Reconstruye una entrega desde el Registry.

        Devuelve ``None`` si la clave no existe. Los registros heredados sin
        detalle de productos no pueden reconstruirse y generan un error claro.
        """

        if isinstance(delivery_or_key, Delivery):
            delivery_key = build_delivery_key(
                self._validate_delivery(delivery_or_key)
            )
        else:
            delivery_key = self._require_text(
                delivery_or_key,
                "identificador de entrega",
            )

        entry = self._data.get(delivery_key)

        if entry is None:
            return None

        return self._deserialize_delivery(
            delivery_key=delivery_key,
            entry=entry,
        )

    def get_pending_deliveries(self) -> list[Delivery]:
        """Reconstruye y devuelve todas las entregas pendientes."""

        deliveries: list[Delivery] = []

        for delivery_key, entry in sorted(self._data.items()):
            if bool(entry.get("synchronized", False)):
                continue

            deliveries.append(
                self._deserialize_delivery(
                    delivery_key=delivery_key,
                    entry=entry,
                )
            )

        return deliveries

    # ======================================================
    # DELIVERY SERIALIZATION
    # ======================================================

    def _validate_delivery(
        self,
        delivery: Delivery,
    ) -> Delivery:
        """Valida completamente una entrega sin modificarla."""

        if not isinstance(delivery, Delivery):
            raise ValueError(f"Objeto de entrega no válido: {type(delivery).__name__}.")

        delivery_date = self._as_date(delivery.delivery_date)

        if not isinstance(delivery.sales_point, SalesPoint):
            raise ValueError("La entrega no contiene un punto de venta válido.")

        sales_point_name = self._require_text(
            delivery.sales_point.name,
            "punto de venta",
        )

        if not isinstance(delivery.products, list) or not delivery.products:
            raise ValueError("La entrega no contiene ningún producto.")

        seen_codes: set[str] = set()

        for position, product in enumerate(delivery.products, start=1):
            if not isinstance(product, Product):
                raise ValueError(
                    f"El producto {position} no es una instancia válida de Product."
                )

            product_code = self._validate_product_code(
                product.code,
                position=position,
            )

            if product_code in seen_codes:
                raise ValueError(
                    f"El código {product_code} aparece repetido en la entrega."
                )

            seen_codes.add(product_code)

            self._require_text(
                product.name,
                f"nombre del producto {product_code}",
            )
            self._require_text(
                product.format,
                f"formato del producto {product_code}",
            )

            price = self._require_finite_number(
                product.price,
                f"precio del producto {product_code}",
            )
            quantity = self._require_finite_number(
                product.quantity,
                f"cantidad del producto {product_code}",
            )

            if price < 0:
                raise ValueError(
                    f"El precio del producto {product_code} no puede ser negativo."
                )

            if math.isclose(quantity, 0.0, abs_tol=1e-9):
                raise ValueError(
                    f"La cantidad del producto {product_code} no puede ser cero."
                )

        # Variables evaluadas expresamente para que los errores de fecha y
        # punto de venta se produzcan antes de acceder al almacenamiento.
        _ = delivery_date, sales_point_name
        return delivery

    def _serialize_products(
        self,
        products: list[Product],
    ) -> list[dict[str, Any]]:
        """Convierte productos a una representación JSON estable."""

        serialized_products = [
            {
                "code": product.code.strip(),
                "name": normalize_payload_text(product.name),
                "format": normalize_payload_text(product.format),
                "price": self._require_finite_number(product.price, "precio"),
                "quantity": self._require_finite_number(
                    product.quantity,
                    "cantidad",
                ),
            }
            for product in products
        ]

        return sorted(
            serialized_products,
            key=lambda item: item["code"],
        )

    def _deserialize_delivery(
        self,
        delivery_key: str,
        entry: dict[str, Any],
    ) -> Delivery:
        """Reconstruye un modelo Delivery a partir de una entrada validada."""

        if not bool(entry.get("payload_available", False)):
            raise ValueError(
                "La entrega no puede recuperarse porque fue registrada por una "
                "versión antigua que no guardaba sus productos. "
                f"Identificador: {delivery_key}."
            )

        products_data = entry.get("products")

        if not isinstance(products_data, list) or not products_data:
            raise ValueError(
                f"La entrega {delivery_key} no contiene productos recuperables."
            )

        delivery_date = date.fromisoformat(
            self._require_text(
                entry.get("delivery_date_iso"),
                "fecha ISO del Registry",
            )
        )
        sales_point_name = self._require_text(
            entry.get("sales_point"),
            "punto de venta del Registry",
        )

        products = [
            Product(
                code=self._validate_product_code(product_data.get("code")),
                name=self._require_text(
                    product_data.get("name"),
                    "nombre del producto almacenado",
                ),
                format=self._require_text(
                    product_data.get("format"),
                    "formato del producto almacenado",
                ),
                price=self._require_finite_number(
                    product_data.get("price"),
                    "precio del producto almacenado",
                ),
                quantity=self._require_finite_number(
                    product_data.get("quantity"),
                    "cantidad del producto almacenado",
                ),
            )
            for product_data in products_data
        ]

        delivery = Delivery(
            sales_point=SalesPoint(name=sales_point_name),
            delivery_date=delivery_date,
            products=products,
        )

        self._validate_delivery(delivery)

        expected_key = build_delivery_key(delivery)
        if expected_key != delivery_key:
            raise ValueError(
                "La clave del Registry no coincide con los datos almacenados: "
                f"{delivery_key}."
            )

        return delivery

    # ======================================================
    # PAYLOAD IDENTITY
    # ======================================================

    def _validate_or_attach_payload(
        self,
        delivery_key: str,
        entry: dict[str, Any],
        delivery: Delivery,
    ) -> None:
        """Compara la huella o completa una entrada heredada sin productos."""

        incoming_hash = build_payload_hash(delivery)
        stored_hash = entry.get("payload_hash")
        payload_available = bool(entry.get("payload_available", False))

        if payload_available and isinstance(stored_hash, str) and stored_hash:
            if stored_hash != incoming_hash:
                raise RegistryConflictError(
                    "El Registry ya contiene una entrega con la misma fecha y "
                    "punto de venta, pero con productos o cantidades diferentes. "
                    f"Identificador: {delivery_key}."
                )
            return

        # Migración progresiva: una entrada antigua conserva su estado, pero
        # recibe desde este momento el contenido completo verificable.
        products_payload = self._serialize_products(delivery.products)
        now = self._utc_now()

        entry["schema_version"] = self._SCHEMA_VERSION
        entry["product_count"] = len(products_payload)
        entry["products"] = products_payload
        entry["payload_available"] = True
        entry["payload_hash"] = incoming_hash
        entry["hash_algorithm"] = self._HASH_ALGORITHM
        entry["updated_at_utc"] = now
        entry.setdefault("registered_at_utc", now)
        entry.setdefault("synchronized_at_utc", None)

    def _build_payload_hash_from_entry(
        self,
        delivery_key: str,
        products: list[dict[str, Any]],
    ) -> str:
        """Calcula una huella directamente desde productos almacenados."""

        canonical_products = sorted(
            (
                {
                    "code": self._validate_product_code(product.get("code")),
                    "name": normalize_payload_text(product.get("name")),
                    "format": normalize_payload_text(product.get("format")),
                    "price": canonical_number(product.get("price")),
                    "quantity": canonical_number(product.get("quantity")),
                }
                for product in products
            ),
            key=lambda item: item["code"],
        )

        payload = {
            "delivery_key": delivery_key,
            "products": canonical_products,
        }

        serialized = json.dumps(
            payload,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        )

        return hashlib.sha256(serialized.encode("utf-8")).hexdigest()

    # ======================================================
    # REGISTRY VALIDATION / MIGRATION
    # ======================================================

    def _validate_registry_data(
        self,
        data: object,
    ) -> dict[str, dict[str, Any]]:
        """Valida y normaliza todo el documento antes de usarlo o guardarlo."""

        if not isinstance(data, dict):
            raise ValueError("El contenido del Registry debe ser un objeto JSON.")

        normalized_data: dict[str, dict[str, Any]] = {}

        for delivery_key, raw_entry in data.items():
            if not isinstance(delivery_key, str) or not delivery_key.strip():
                raise ValueError("El Registry contiene un identificador no válido.")

            if not isinstance(raw_entry, dict):
                raise ValueError(
                    f"La entrada {delivery_key!r} debe ser un objeto JSON."
                )

            normalized_data[delivery_key] = self._normalize_registry_entry(
                delivery_key=delivery_key,
                raw_entry=raw_entry,
            )

        return normalized_data

    def _normalize_registry_entry(
        self,
        delivery_key: str,
        raw_entry: dict[str, Any],
    ) -> dict[str, Any]:
        """Valida una entrada y migra estructuras heredadas a la versión 2."""

        key_date, key_sales_point = self._parse_delivery_key(delivery_key)

        stored_date = raw_entry.get("delivery_date_iso")
        if stored_date is None:
            legacy_date = raw_entry.get("delivery_date")
            parsed_date = self._parse_delivery_date(legacy_date).date()
        else:
            parsed_date = self._parse_delivery_date(stored_date).date()

        if parsed_date != key_date:
            raise ValueError(
                f"La fecha almacenada no coincide con la clave {delivery_key}."
            )

        sales_point_name = self._require_text(
            raw_entry.get("sales_point"),
            f"punto de venta de {delivery_key}",
        )

        if normalize_key_text(sales_point_name) != key_sales_point:
            raise ValueError(
                "El punto de venta almacenado no coincide con la clave "
                f"{delivery_key}."
            )

        synchronized = raw_entry.get("synchronized", False)
        if not isinstance(synchronized, bool):
            raise ValueError(
                f"El estado synchronized de {delivery_key} debe ser booleano."
            )

        raw_products = raw_entry.get("products")
        payload_available = bool(raw_entry.get("payload_available", False))

        # Formato heredado: ``products`` era solamente el contador.
        if isinstance(raw_products, int) and not isinstance(raw_products, bool):
            product_count = raw_products
            products: list[dict[str, Any]] = []
            payload_available = False
        elif isinstance(raw_products, list):
            products = self._normalize_stored_products(
                delivery_key=delivery_key,
                raw_products=raw_products,
            )
            product_count = len(products)
            payload_available = payload_available or bool(products)
        elif raw_products is None:
            product_count = self._require_non_negative_integer(
                raw_entry.get("product_count", 0),
                f"contador de productos de {delivery_key}",
            )
            products = []
            payload_available = False
        else:
            raise ValueError(
                f"El campo products de {delivery_key} no tiene un formato válido."
            )

        stored_product_count = raw_entry.get("product_count")
        if stored_product_count is not None:
            stored_product_count = self._require_non_negative_integer(
                stored_product_count,
                f"contador de productos de {delivery_key}",
            )

            if payload_available and stored_product_count != len(products):
                raise ValueError(
                    f"El contador de productos no coincide en {delivery_key}."
                )

            if not payload_available:
                product_count = stored_product_count

        payload_hash = raw_entry.get("payload_hash")

        if payload_available:
            if not products:
                raise ValueError(f"La entrega {delivery_key} declara un payload vacío.")

            calculated_hash = self._build_payload_hash_from_entry(
                delivery_key=delivery_key,
                products=products,
            )

            if payload_hash is not None and payload_hash != calculated_hash:
                raise ValueError(
                    f"La huella de contenido no coincide en {delivery_key}."
                )

            payload_hash = calculated_hash
        else:
            payload_hash = None

        registered_at = self._normalize_optional_timestamp(
            raw_entry.get("registered_at_utc")
        )
        updated_at = self._normalize_optional_timestamp(raw_entry.get("updated_at_utc"))
        synchronized_at = self._normalize_optional_timestamp(
            raw_entry.get("synchronized_at_utc")
        )

        return {
            "schema_version": self._SCHEMA_VERSION,
            "delivery_date": parsed_date.strftime("%d/%m/%Y"),
            "delivery_date_iso": parsed_date.isoformat(),
            "year": parsed_date.year,
            "month": MONTHS[parsed_date.month - 1],
            "sales_point": sales_point_name,
            "product_count": product_count,
            "products": products,
            "payload_available": payload_available,
            "payload_hash": payload_hash,
            "hash_algorithm": self._HASH_ALGORITHM,
            "synchronized": synchronized,
            "registered_at_utc": registered_at,
            "updated_at_utc": updated_at,
            "synchronized_at_utc": synchronized_at,
        }

    def _normalize_stored_products(
        self,
        delivery_key: str,
        raw_products: list[Any],
    ) -> list[dict[str, Any]]:
        """Valida los productos almacenados y devuelve una copia normalizada."""

        normalized_products: list[dict[str, Any]] = []
        seen_codes: set[str] = set()

        for position, raw_product in enumerate(raw_products, start=1):
            if not isinstance(raw_product, dict):
                raise ValueError(
                    f"El producto {position} de {delivery_key} no es un objeto JSON."
                )

            code = self._validate_product_code(
                raw_product.get("code"),
                position=position,
            )

            if code in seen_codes:
                raise ValueError(f"El código {code} está repetido en {delivery_key}.")

            seen_codes.add(code)

            name = self._require_text(
                raw_product.get("name"),
                f"nombre del producto {code}",
            )
            product_format = self._require_text(
                raw_product.get("format"),
                f"formato del producto {code}",
            )
            price = self._require_finite_number(
                raw_product.get("price"),
                f"precio del producto {code}",
            )
            quantity = self._require_finite_number(
                raw_product.get("quantity"),
                f"cantidad del producto {code}",
            )

            if price < 0:
                raise ValueError(
                    f"El precio del producto {code} no puede ser negativo."
                )

            if math.isclose(quantity, 0.0, abs_tol=1e-9):
                raise ValueError(f"La cantidad del producto {code} no puede ser cero.")

            normalized_products.append(
                {
                    "code": code,
                    "name": normalize_payload_text(name),
                    "format": normalize_payload_text(product_format),
                    "price": price,
                    "quantity": quantity,
                }
            )

        return sorted(
            normalized_products,
            key=lambda item: item["code"],
        )

    # ======================================================
    # FILE I/O
    # ======================================================

    def _load(self) -> dict[str, dict[str, Any]]:
        """Carga y valida el JSON, deteniéndose ante cualquier corrupción."""

        print()
        print("=" * 100)
        print("CARGA DEL REGISTRY")
        print("=" * 100)
        print(f"Archivo        : {self._registry_file.name}")
        print(f"Ruta           : {self._registry_file}")
        print("Proceso        : Comprobando archivo del Registry...")

        if not self._registry_file.exists():
            print("Archivo        : NO EXISTE")
            print("Proceso        : Creando nuevo Registry...")

            self._registry_file.parent.mkdir(parents=True, exist_ok=True)
            self._atomic_write({})

            print("Estado         : REGISTRY CREADO")
            print("Registros      : 0")
            print("=" * 100)
            return {}

        if not self._registry_file.is_file():
            raise RuntimeError(
                f"La ruta del Registry no es un archivo: {self._registry_file}"
            )

        print("Archivo        : ENCONTRADO")
        print("Proceso        : Leyendo y validando registros...")

        try:
            with self._registry_file.open("r", encoding="utf-8") as file:
                raw_data = json.load(file)

            data = self._validate_registry_data(raw_data)

        except (json.JSONDecodeError, OSError, TypeError, ValueError) as error:
            raise RuntimeError(
                "El Registry está dañado o contiene datos incompatibles. "
                "La ejecución se ha detenido para evitar pérdida de información "
                "o duplicación de cantidades."
            ) from error

        print("Lectura        : COMPLETADA")
        print(f"Registros      : {len(data)}")
        print("Estado         : REGISTRY CARGADO CORRECTAMENTE")
        print("=" * 100)

        return data

    def _atomic_write(
        self,
        data: dict[str, dict[str, Any]],
    ) -> None:
        """Escribe el JSON en un temporal del mismo directorio y lo reemplaza."""

        self._registry_file.parent.mkdir(parents=True, exist_ok=True)
        self._backup_directory.mkdir(parents=True, exist_ok=True)

        serialized = json.dumps(
            data,
            indent=4,
            ensure_ascii=False,
            sort_keys=True,
            allow_nan=False,
        )
        serialized += "\n"

        temporary_file = self._registry_file.with_name(
            f".{self._registry_file.name}.{uuid4().hex}.tmp"
        )

        try:
            with temporary_file.open("x", encoding="utf-8", newline="\n") as file:
                file.write(serialized)
                file.flush()
                os.fsync(file.fileno())

            if self._registry_file.is_file():
                backup_file = self._backup_directory / (
                    f"{self._registry_file.stem}.previous"
                    f"{self._registry_file.suffix}"
                )
                backup_temporary = backup_file.with_name(
                    f".{backup_file.name}.{uuid4().hex}.tmp"
                )

                try:
                    shutil.copy2(self._registry_file, backup_temporary)
                    backup_temporary.replace(backup_file)
                finally:
                    if backup_temporary.exists():
                        backup_temporary.unlink()

            temporary_file.replace(self._registry_file)
            self._fsync_directory(self._registry_file.parent)

        finally:
            if temporary_file.exists():
                temporary_file.unlink()

    def _acquire_process_lock(self) -> None:
        """Obtiene un bloqueo no bloqueante válido en Windows y Unix."""

        self._registry_file.parent.mkdir(parents=True, exist_ok=True)
        self._lock_handle = self._lock_file.open("a+b")

        try:
            if os.name == "nt":
                import msvcrt

                self._lock_handle.seek(0, os.SEEK_END)
                if self._lock_handle.tell() == 0:
                    self._lock_handle.write(b"\0")
                    self._lock_handle.flush()

                self._lock_handle.seek(0)
                msvcrt.locking(
                    self._lock_handle.fileno(),
                    msvcrt.LK_NBLCK,
                    1,
                )
            else:
                import fcntl

                fcntl.flock(
                    self._lock_handle.fileno(),
                    fcntl.LOCK_EX | fcntl.LOCK_NB,
                )

        except (OSError, BlockingIOError) as error:
            self._lock_handle.close()
            self._lock_handle = None
            raise RuntimeError(
                "El Registry ya está siendo utilizado por otra ejecución del "
                "programa. Cierra la otra instancia antes de continuar."
            ) from error

    def _fsync_directory(self, directory: Path) -> None:
        """Sincroniza los metadatos del directorio cuando el sistema lo permite."""

        if os.name == "nt":
            return

        try:
            descriptor = os.open(directory, os.O_RDONLY)
        except OSError:
            return

        try:
            os.fsync(descriptor)
        finally:
            os.close(descriptor)

    # ======================================================
    # KEYS / DATES / GENERIC VALIDATION
    # ======================================================

    def _parse_delivery_key(
        self,
        delivery_key: str,
    ) -> tuple[date, str]:
        """Separa y valida una clave persistida."""

        if "|" not in delivery_key:
            raise ValueError(f"Identificador de entrega no válido: {delivery_key!r}.")

        date_text, sales_point_key = delivery_key.split("|", maxsplit=1)

        try:
            parsed_date = date.fromisoformat(date_text)
        except ValueError as error:
            raise ValueError(
                f"Fecha no válida dentro de la clave {delivery_key!r}."
            ) from error

        sales_point_key = normalize_key_text(sales_point_key)
        if not sales_point_key:
            raise ValueError(
                f"Punto de venta vacío dentro de la clave {delivery_key!r}."
            )

        canonical_key = f"{parsed_date.isoformat()}|{sales_point_key}"
        if canonical_key != delivery_key:
            raise ValueError(
                f"La clave del Registry no está normalizada: {delivery_key!r}."
            )

        return parsed_date, sales_point_key

    def _as_date(
        self,
        value: object,
    ) -> date:
        """Convierte una fecha válida a ``date``."""

        if isinstance(value, datetime):
            return value.date()

        if isinstance(value, date):
            return value

        if isinstance(value, str):
            return self._parse_delivery_date(value).date()

        raise ValueError(f"Formato de fecha no válido: {value!r}.")

    def _parse_delivery_date(
        self,
        delivery_date: object,
    ) -> datetime:
        """Convierte los formatos históricos aceptados por el proyecto."""

        if isinstance(delivery_date, datetime):
            return delivery_date

        if isinstance(delivery_date, date):
            return datetime.combine(
                delivery_date,
                datetime.min.time(),
            )

        normalized_date = self._require_text(
            delivery_date,
            "fecha de entrega",
        )

        accepted_formats = (
            "%Y-%m-%d",
            "%d/%m/%Y",
            "%d-%m-%Y",
        )

        for date_format in accepted_formats:
            try:
                return datetime.strptime(normalized_date, date_format)
            except ValueError:
                continue

        raise ValueError(f"Formato de fecha no válido: {delivery_date!r}.")

    def _validate_product_code(
        self,
        value: object,
        *,
        position: int | None = None,
    ) -> str:
        """Valida un código compatible con Importer y ProductManager."""

        prefix = f" en la posición {position}" if position is not None else ""

        if not isinstance(value, str):
            raise ValueError(f"El código del producto{prefix} debe ser texto.")

        code = value.strip()

        if not code or not code.isdigit():
            raise ValueError(
                f"Código de producto no válido{prefix}: {value!r}. "
                "Debe contener únicamente dígitos."
            )

        return code

    def _require_text(
        self,
        value: object,
        field_name: str,
    ) -> str:
        """Devuelve un texto limpio y exige que no esté vacío."""

        if value is None:
            raise ValueError(f"El campo {field_name} está vacío.")

        normalized_value = " ".join(str(value).strip().split())

        if not normalized_value:
            raise ValueError(f"El campo {field_name} está vacío.")

        return normalized_value

    def _require_finite_number(
        self,
        value: object,
        field_name: str,
    ) -> float:
        """Exige un número real finito y rechaza booleanos."""

        if isinstance(value, bool) or not isinstance(value, (int, float)):
            raise ValueError(f"El campo {field_name} debe ser numérico.")

        numeric_value = float(value)

        if not math.isfinite(numeric_value):
            raise ValueError(f"El campo {field_name} debe ser un número finito.")

        return numeric_value

    def _require_non_negative_integer(
        self,
        value: object,
        field_name: str,
    ) -> int:
        """Valida contadores JSON."""

        if isinstance(value, bool) or not isinstance(value, int) or value < 0:
            raise ValueError(f"El campo {field_name} debe ser un entero no negativo.")

        return value

    def _normalize_optional_timestamp(
        self,
        value: object,
    ) -> str | None:
        """Valida timestamps ISO-8601 opcionales y los conserva normalizados."""

        if value is None or value == "":
            return None

        timestamp = self._require_text(value, "marca temporal")

        try:
            parsed = datetime.fromisoformat(timestamp.replace("Z", "+00:00"))
        except ValueError as error:
            raise ValueError(f"Marca temporal no válida: {timestamp!r}.") from error

        if parsed.tzinfo is None:
            parsed = parsed.replace(tzinfo=timezone.utc)
        else:
            parsed = parsed.astimezone(timezone.utc)

        return parsed.isoformat(timespec="seconds")

    def _utc_now(self) -> str:
        return datetime.now(timezone.utc).isoformat(timespec="seconds")

    def _ensure_open(self) -> None:
        if self._closed:
            raise RuntimeError("El Registry ya está cerrado.")

    # ======================================================
    # OUTPUT
    # ======================================================

    def _print_register_existing(
        self,
        delivery_key: str,
        delivery_date: date,
        sales_point_name: str,
        synchronized: bool,
    ) -> None:
        print()
        print("-" * 100)
        print("REGISTRO DE LA ENTREGA")
        print("-" * 100)
        print(f"Identificador   : {delivery_key}")
        print(f"Fecha           : {delivery_date.strftime('%d/%m/%Y')}")
        print(f"Punto de venta  : {sales_point_name}")
        print("Estado          : YA EXISTE EN EL REGISTRY")
        print("Sincronización  : " f"{'COMPLETADA' if synchronized else 'PENDIENTE'}")
        print("-" * 100)

    def _print_registered(
        self,
        delivery_key: str,
        delivery_date: date,
        sales_point_name: str,
        product_count: int,
    ) -> None:
        print()
        print("-" * 100)
        print("REGISTRO DE LA ENTREGA")
        print("-" * 100)
        print(f"Identificador   : {delivery_key}")
        print(f"Fecha           : {delivery_date.strftime('%d/%m/%Y')}")
        print(f"Año             : {delivery_date.year}")
        print(f"Mes             : {MONTHS[delivery_date.month - 1]}")
        print(f"Punto de venta  : {sales_point_name}")
        print(f"Productos       : {product_count}")
        print("Sincronización  : PENDIENTE")
        print("Estado          : AÑADIDA AL REGISTRY")
        print("-" * 100)

    def _print_synchronization_update(
        self,
        delivery_key: str,
        delivery_date: date,
        sales_point_name: str,
        previous_status: bool,
    ) -> None:
        print()
        print("-" * 100)
        print("ACTUALIZACIÓN DEL ESTADO DE SINCRONIZACIÓN")
        print("-" * 100)
        print(f"Identificador   : {delivery_key}")
        print(f"Fecha           : {delivery_date.strftime('%d/%m/%Y')}")
        print(f"Punto de venta  : {sales_point_name}")
        print("Proceso         : Marcando entrega como sincronizada...")
        print("Registrada      : SÍ")
        print(
            "Estado anterior : " f"{'SINCRONIZADO' if previous_status else 'PENDIENTE'}"
        )
        print("Estado actual   : SINCRONIZADO")
        print("-" * 100)
