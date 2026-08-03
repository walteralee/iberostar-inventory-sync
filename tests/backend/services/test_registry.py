"""
No existía ningún test para Registry pese a ser la pieza que garantiza
que una entrega no se sincronice dos veces ni se pierda si el proceso
falla a mitad de camino. Estos tests usan siempre registry_file /
backup_directory apuntando a tmp_path, nunca a storage/ real.
"""

import json
from datetime import date

import pytest

from models.delivery import Delivery
from models.product import Product
from models.sales_point import SalesPoint
from services.registry import Registry, RegistryConflictError


def _make_delivery(sales_point="Bar Piscina", day=1, code="1", quantity=48.0, name="A"):
    return Delivery(
        sales_point=SalesPoint(name=sales_point),
        delivery_date=date(2026, 7, day),
        products=[
            Product(code=code, name=name, format="UNIDAD", price=1.0, quantity=quantity)
        ],
    )


@pytest.fixture
def registry(tmp_path):
    reg = Registry(
        registry_file=tmp_path / "imported_deliveries.json",
        backup_directory=tmp_path / "backup",
        acquire_lock=False,
    )
    yield reg
    reg.close()


class TestRegisterAndExists:
    def test_a_delivery_does_not_exist_before_registering(self, registry):
        assert registry.exists(_make_delivery()) is False

    def test_registering_makes_it_exist_and_be_pending(self, registry):
        registry.register(_make_delivery())

        assert registry.exists(_make_delivery()) is True
        assert registry.is_synchronized(_make_delivery()) is False

    def test_registering_the_same_delivery_twice_is_a_no_op(self, registry):
        registry.register(_make_delivery())
        registry.register(_make_delivery())

        assert registry.exists(_make_delivery()) is True

    def test_conflicting_delivery_for_the_same_key_raises(self, registry):
        registry.register(_make_delivery(quantity=48.0))

        with pytest.raises(RegistryConflictError):
            registry.register(_make_delivery(quantity=999.0))

    def test_exists_also_raises_on_conflicting_content(self, registry):
        registry.register(_make_delivery(quantity=48.0))

        with pytest.raises(RegistryConflictError):
            registry.exists(_make_delivery(quantity=999.0))


class TestMarkAsSynchronized:
    def test_cannot_mark_an_unregistered_delivery(self, registry):
        with pytest.raises(ValueError):
            registry.mark_as_synchronized(_make_delivery())

    def test_marks_a_registered_delivery_as_synchronized(self, registry):
        delivery = _make_delivery()
        registry.register(delivery)

        registry.mark_as_synchronized(delivery)

        assert registry.is_synchronized(delivery) is True

    def test_marking_twice_keeps_the_first_synchronized_timestamp(self, registry):
        delivery = _make_delivery()
        registry.register(delivery)
        registry.mark_as_synchronized(delivery)
        registry.save()

        key = list(registry.data.keys())[0]
        first_timestamp = registry.data[key]["synchronized_at_utc"]

        registry.mark_as_synchronized(delivery)

        assert registry.data[key]["synchronized_at_utc"] == first_timestamp


class TestGetPendingDeliveries:
    def test_only_returns_unsynchronized_deliveries(self, registry):
        pending_delivery = _make_delivery(sales_point="Bar Piscina", day=1)
        synced_delivery = _make_delivery(sales_point="Comedor", day=2)

        registry.register(pending_delivery)
        registry.register(synced_delivery)
        registry.mark_as_synchronized(synced_delivery)

        pending, warnings = registry.get_pending_deliveries()

        assert warnings == []
        assert len(pending) == 1
        assert pending[0].sales_point.name == "Bar Piscina"

    def test_reconstructed_delivery_round_trips_through_save_and_reload(
        self, registry, tmp_path
    ):
        delivery = _make_delivery(quantity=3024.0, name="AZUCAR TERRON MORENO 4 GR")
        registry.register(delivery)
        registry.save()

        reloaded = Registry(
            registry_file=tmp_path / "imported_deliveries.json",
            backup_directory=tmp_path / "backup",
            acquire_lock=False,
        )
        pending, warnings = reloaded.get_pending_deliveries()

        assert warnings == []
        assert len(pending) == 1
        assert pending[0].products[0].quantity == 3024.0
        assert pending[0].products[0].name == "AZUCAR TERRON MORENO 4 GR"
        reloaded.close()


class TestGetDelivery:
    def test_returns_none_for_an_unknown_key(self, registry):
        assert registry.get_delivery("2026-07-01|inexistente") is None

    def test_returns_the_delivery_for_a_known_key(self, registry):
        delivery = _make_delivery()
        registry.register(delivery)

        found = registry.get_delivery("2026-07-01|bar piscina")

        assert found is not None
        assert found.sales_point.name == "bar piscina" or found.sales_point.name == "Bar Piscina"


class TestPersistedFileValidation:
    def test_save_produces_a_file_loadable_by_a_new_registry(self, tmp_path):
        first = Registry(
            registry_file=tmp_path / "imported_deliveries.json",
            backup_directory=tmp_path / "backup",
            acquire_lock=False,
        )
        first.register(_make_delivery())
        first.save()
        first.close()

        second = Registry(
            registry_file=tmp_path / "imported_deliveries.json",
            backup_directory=tmp_path / "backup",
            acquire_lock=False,
        )
        assert len(second.data) == 1
        second.close()

    def test_corrupted_json_stops_with_a_clear_error(self, tmp_path):
        registry_file = tmp_path / "imported_deliveries.json"
        registry_file.write_text("{ esto no es json valido")

        with pytest.raises(RuntimeError):
            Registry(
                registry_file=registry_file,
                backup_directory=tmp_path / "backup",
                acquire_lock=False,
            )

    def test_legacy_entry_without_products_cannot_be_reconstructed(self, tmp_path):
        registry_file = tmp_path / "imported_deliveries.json"
        registry_file.write_text(
            json.dumps(
                {
                    "2026-07-01|bar piscina": {
                        "delivery_date": "01/07/2026",
                        "delivery_date_iso": "2026-07-01",
                        "sales_point": "Bar Piscina",
                        "product_count": 3,
                        "products": 3,
                        "synchronized": False,
                    }
                }
            )
        )

        registry = Registry(
            registry_file=registry_file,
            backup_directory=tmp_path / "backup",
            acquire_lock=False,
        )

        pending, warnings = registry.get_pending_deliveries()

        assert pending == []
        assert len(warnings) == 1
        registry.close()

    def test_missing_registry_file_starts_empty(self, tmp_path):
        registry = Registry(
            registry_file=tmp_path / "no_existe" / "imported_deliveries.json",
            backup_directory=tmp_path / "backup",
            acquire_lock=False,
        )

        assert registry.data == {}
        registry.close()
