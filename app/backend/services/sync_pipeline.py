"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    sync_pipeline.py

Descripción:
    Orquesta el proceso completo sobre un conjunto de Excel de
    Economato: importación, recuperación de entregas pendientes de
    ejecuciones anteriores y sincronización con los Excel mensuales.

    Devuelve un resumen serializable para que la capa de presentación
    (la API web) no tenga que conocer los detalles de cada servicio.
"""

from __future__ import annotations

from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any, Iterable

from services.importer import Importer, ImportSummary
from services.registry import Registry
from services.synchronizer import Synchronizer, _SynchronizationTotals
from utils.delivery_identity import build_delivery_key


@dataclass(slots=True)
class SyncResult:
    """Resultado consolidado de una ejecución completa."""

    import_summary: ImportSummary
    pending_warnings: list[str]
    synchronization_totals: _SynchronizationTotals

    @property
    def has_incidents(self) -> bool:
        return (
            self.import_summary.has_incidents()
            or bool(self.pending_warnings)
            or bool(self.synchronization_totals.error_deliveries)
        )

    def to_dict(self) -> dict[str, Any]:
        return {
            "has_incidents": self.has_incidents,
            "import_summary": asdict(self.import_summary),
            "pending_warnings": list(self.pending_warnings),
            "synchronization_totals": asdict(self.synchronization_totals),
        }


def run_sync(excel_files: Iterable[Path | str]) -> SyncResult:
    """
    Importa los Excel indicados y sincroniza todas las entregas
    pendientes, incluidas las que quedaron a medias en ejecuciones
    anteriores.

    Raises:
        RuntimeError: si el Registry está bloqueado por otra ejecución.
    """

    with Registry() as registry:
        importer = Importer(registry=registry)
        synchronizer = Synchronizer(registry=registry)

        deliveries = importer.run(excel_files)

        # Una entrega puede quedar registrada pero sin sincronizar si una
        # ejecución anterior falló o se interrumpió a mitad de camino. Como
        # el Importer solo compara contra los Excel recibidos ahora, esa
        # entrega no se retomaría nunca si su Excel de origen no se vuelve
        # a importar. Por eso se recupera del Registry todo lo pendiente.
        known_delivery_keys = {
            build_delivery_key(delivery) for delivery in deliveries
        }

        recovered_deliveries, pending_warnings = registry.get_pending_deliveries()

        deliveries = deliveries + [
            delivery
            for delivery in recovered_deliveries
            if build_delivery_key(delivery) not in known_delivery_keys
        ]

        synchronization_totals = synchronizer.run(deliveries)

    return SyncResult(
        import_summary=importer.last_summary,
        pending_warnings=pending_warnings,
        synchronization_totals=synchronization_totals,
    )
