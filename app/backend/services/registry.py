"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    registry.py

Descripción:
    Servicio encargado de gestionar el registro de las
    entregas importadas y sincronizadas.

    Cada entrega se identifica mediante la combinación de su
    fecha y su punto de venta.
"""

from datetime import date, datetime
import json
import unicodedata

from config.constants import MONTHS
from config.settings import REGISTRY_FILE
from models.delivery import Delivery


class Registry:
    """
    Gestiona el registro de entregas importadas y sincronizadas.
    """

    def __init__(self) -> None:
        """
        Inicializa el registro cargando el fichero JSON.
        """

        self._data = self._load()

    # ======================================================
    # PUBLIC
    # ======================================================

    @property
    def data(self) -> dict:
        """
        Devuelve el contenido completo del registro.
        """

        return self._data

    def exists(
        self,
        delivery: Delivery,
    ) -> bool:
        """
        Comprueba si una entrega ya está registrada.

        Args:
            delivery: Entrega que se desea comprobar.

        Returns:
            True si la entrega ya existe.
        """

        delivery_key = self._build_delivery_key(
            delivery,
        )

        return delivery_key in self._data

    def is_synchronized(
        self,
        delivery: Delivery,
    ) -> bool:
        """
        Comprueba si una entrega ya ha sido sincronizada.

        Args:
            delivery: Entrega que se desea comprobar.

        Returns:
            True si la entrega ya fue sincronizada.
        """

        delivery_key = self._build_delivery_key(
            delivery,
        )

        if delivery_key not in self._data:
            return False

        return bool(
            self._data[delivery_key].get(
                "synchronized",
                False,
            )
        )

    def mark_as_synchronized(
        self,
        delivery: Delivery,
    ) -> None:
        """
        Marca una entrega como sincronizada.

        Args:
            delivery: Entrega que se desea actualizar.
        """

        delivery_key = self._build_delivery_key(
            delivery,
        )

        parsed_date = self._parse_delivery_date(
            delivery.delivery_date,
        )

        delivery_date = parsed_date.strftime(
            "%d/%m/%Y",
        )

        sales_point_name = delivery.sales_point.name.strip()

        print()
        print("-" * 100)
        print("ACTUALIZACIÓN DEL ESTADO DE SINCRONIZACIÓN")
        print("-" * 100)
        print(f"Identificador   : {delivery_key}")
        print(f"Fecha           : {delivery_date}")
        print(f"Punto de venta  : {sales_point_name}")
        print("Proceso         : Marcando entrega como sincronizada...")

        if delivery_key in self._data:

            previous_status = bool(
                self._data[delivery_key].get(
                    "synchronized",
                    False,
                )
            )

            self._data[delivery_key]["synchronized"] = True

            print("Registrada      : SÍ")
            print(
                "Estado anterior : "
                f"{'SINCRONIZADO' if previous_status else 'PENDIENTE'}"
            )
            print("Estado actual   : SINCRONIZADO")

        else:

            print("Registrada      : NO")
            print("Estado          : NO MODIFICADO")
            print(
                "Resultado       : No se puede marcar como sincronizada "
                "porque la entrega no existe."
            )

        print("-" * 100)

    def register(
        self,
        delivery: Delivery,
    ) -> None:
        """
        Registra una nueva entrega importada.

        Args:
            delivery: Entrega que se añadirá al registro.
        """

        delivery_key = self._build_delivery_key(
            delivery,
        )

        parsed_date = self._parse_delivery_date(
            delivery.delivery_date,
        )

        delivery_date = parsed_date.strftime(
            "%d/%m/%Y",
        )

        month = MONTHS[parsed_date.month - 1]
        sales_point_name = delivery.sales_point.name.strip()

        if delivery_key in self._data:

            print()
            print("-" * 100)
            print("REGISTRO DE LA ENTREGA")
            print("-" * 100)
            print(f"Identificador   : {delivery_key}")
            print(f"Fecha           : {delivery_date}")
            print(f"Punto de venta  : {sales_point_name}")
            print("Estado          : YA EXISTE EN EL REGISTRY")
            print("-" * 100)

            return

        self._data[delivery_key] = {
            "delivery_date": delivery_date,
            "year": parsed_date.year,
            "month": month,
            "sales_point": sales_point_name,
            "products": len(delivery.products),
            "synchronized": False,
        }

        print()
        print("-" * 100)
        print("REGISTRO DE LA ENTREGA")
        print("-" * 100)
        print(f"Identificador   : {delivery_key}")
        print(f"Fecha           : {delivery_date}")
        print(f"Año             : {parsed_date.year}")
        print(f"Mes             : {month}")
        print(f"Punto de venta  : {sales_point_name}")
        print(f"Productos       : {len(delivery.products)}")
        print("Sincronización  : PENDIENTE")
        print("Estado          : AÑADIDA AL REGISTRY")
        print("-" * 100)

    def save(self) -> None:
        """
        Guarda el registro en disco.
        """

        print()
        print("-" * 100)
        print("ESCRITURA DEL REGISTRY")
        print("-" * 100)
        print(f"Archivo        : {REGISTRY_FILE.name}")
        print(f"Ruta           : {REGISTRY_FILE}")
        print(f"Registros      : {len(self._data)}")
        print("Proceso        : Guardando contenido...")

        REGISTRY_FILE.parent.mkdir(
            parents=True,
            exist_ok=True,
        )

        temporary_file = REGISTRY_FILE.with_suffix(
            f"{REGISTRY_FILE.suffix}.tmp",
        )

        with temporary_file.open(
            "w",
            encoding="utf-8",
        ) as file:

            json.dump(
                self._data,
                file,
                indent=4,
                ensure_ascii=False,
            )

        temporary_file.replace(
            REGISTRY_FILE,
        )

        print("Estado         : GUARDADO CORRECTAMENTE")
        print("-" * 100)

    # ======================================================
    # PRIVATE
    # ======================================================

    def _build_delivery_key(
        self,
        delivery: Delivery,
    ) -> str:
        """
        Genera el identificador único de una entrega.

        El identificador combina la fecha en formato ISO
        y el nombre normalizado del punto de venta.

        Args:
            delivery: Entrega de la que se generará la clave.

        Returns:
            Identificador único de la entrega.
        """

        parsed_date = self._parse_delivery_date(
            delivery.delivery_date,
        )

        sales_point_name = self._normalize_key_text(
            delivery.sales_point.name,
        )

        if not sales_point_name:
            raise ValueError(
                "El punto de venta de la entrega está vacío.",
            )

        return f"{parsed_date.strftime('%Y-%m-%d')}" f"|{sales_point_name}"

    def _normalize_key_text(
        self,
        value: object,
    ) -> str:
        """
        Normaliza un texto para utilizarlo dentro de una clave.
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

    def _load(self) -> dict:
        """
        Carga el registro desde disco.

        Si el fichero no existe, lo crea automáticamente.

        Returns:
            Diccionario con el contenido del registro.
        """

        print()
        print("=" * 100)
        print("CARGA DEL REGISTRY")
        print("=" * 100)
        print(f"Archivo        : {REGISTRY_FILE.name}")
        print(f"Ruta           : {REGISTRY_FILE}")
        print("Proceso        : Comprobando archivo del Registry...")

        if not REGISTRY_FILE.exists():

            print("Archivo        : NO EXISTE")
            print("Proceso        : Creando nuevo Registry...")

            REGISTRY_FILE.parent.mkdir(
                parents=True,
                exist_ok=True,
            )

            REGISTRY_FILE.write_text(
                "{}",
                encoding="utf-8",
            )

            print("Estado         : REGISTRY CREADO")
            print("Registros      : 0")
            print("=" * 100)

            return {}

        print("Archivo        : ENCONTRADO")
        print("Proceso        : Leyendo registros...")

        try:

            with REGISTRY_FILE.open(
                "r",
                encoding="utf-8",
            ) as file:

                data = json.load(
                    file,
                )

            if not isinstance(
                data,
                dict,
            ):
                raise ValueError(
                    "El contenido del Registry debe ser un objeto JSON.",
                )

            print("Lectura        : COMPLETADA")
            print(f"Registros      : {len(data)}")
            print("Estado         : REGISTRY CARGADO CORRECTAMENTE")
            print("=" * 100)

            return data

        except (
            json.JSONDecodeError,
            OSError,
            ValueError,
        ) as error:
            raise RuntimeError(
                "El Registry está dañado. "
                "La ejecución se ha detenido para evitar duplicar cantidades."
            ) from error

    def _parse_delivery_date(
        self,
        delivery_date: str | date | datetime,
    ) -> datetime:
        """
        Convierte la fecha de una entrega a datetime.

        Args:
            delivery_date: Fecha de la entrega.

        Returns:
            Fecha convertida a datetime.
        """

        if isinstance(
            delivery_date,
            datetime,
        ):
            return delivery_date

        if isinstance(
            delivery_date,
            date,
        ):
            return datetime.combine(
                delivery_date,
                datetime.min.time(),
            )

        normalized_date = str(
            delivery_date,
        ).strip()

        accepted_formats = (
            "%d/%m/%Y",
            "%Y-%m-%d",
            "%d-%m-%Y",
        )

        for date_format in accepted_formats:

            try:

                return datetime.strptime(
                    normalized_date,
                    date_format,
                )

            except ValueError:
                continue

        raise ValueError(
            f"Formato de fecha no válido: {delivery_date}",
        )
