import math
from datetime import date

import pytest

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
from utils.product_codes import normalize_product_code


class TestNormalizeProductCode:
    def test_accepts_plain_integer(self):
        assert normalize_product_code(7585) == "7585"

    def test_accepts_integer_float(self):
        assert normalize_product_code(7585.0) == "7585"

    def test_rejects_non_integer_float(self):
        assert normalize_product_code(7585.5) is None

    def test_accepts_text_digits(self):
        assert normalize_product_code("7585") == "7585"

    def test_strips_decimal_zero_suffix_from_text(self):
        assert normalize_product_code("7585.00") == "7585"

    def test_rejects_text_with_a_real_decimal_part(self):
        assert normalize_product_code("7585.5") is None

    def test_rejects_none_and_bool(self):
        assert normalize_product_code(None) is None
        assert normalize_product_code(True) is None
        assert normalize_product_code(False) is None

    def test_rejects_non_numeric_text(self):
        assert normalize_product_code("ABC123") is None

    def test_rejects_empty_or_blank_text(self):
        assert normalize_product_code("") is None
        assert normalize_product_code("   ") is None

    def test_rejects_non_finite_float(self):
        assert normalize_product_code(math.nan) is None
        assert normalize_product_code(math.inf) is None

    def test_leading_zeros_in_text_are_preserved(self):
        # Documenta el comportamiento actual: un código de texto con cero a
        # la izquierda NO se normaliza igual que su equivalente numérico
        # (normalize_product_code(7) == "7"). No se ha observado ningún
        # código real con esta forma en las plantillas ni en el Registry.
        assert normalize_product_code("007") == "007"


class TestNormalizeKeyText:
    def test_casefolds_and_collapses_whitespace(self):
        assert normalize_key_text("  Bar   Piscina  ") == "bar piscina"

    def test_strips_accents(self):
        assert normalize_key_text("Almacén") == "almacen"


class TestNormalizePayloadText:
    def test_trims_and_collapses_whitespace_without_changing_case(self):
        assert normalize_payload_text("  Gin   Larios  ") == "Gin Larios"


class TestCanonicalNumber:
    def test_zero_is_always_the_string_zero(self):
        assert canonical_number(0) == "0"
        assert canonical_number(0.0) == "0"
        assert canonical_number(-0.0) == "0"

    def test_stable_representation_for_equal_values(self):
        assert canonical_number(48) == canonical_number(48.0)

    def test_rejects_bool(self):
        with pytest.raises(ValueError):
            canonical_number(True)

    def test_rejects_non_finite(self):
        with pytest.raises(ValueError):
            canonical_number(math.inf)


def _make_delivery(quantity=48.0, code="7585"):
    return Delivery(
        sales_point=SalesPoint(name="Bar Piscina"),
        delivery_date=date(2026, 7, 1),
        products=[
            Product(
                code=code,
                name="AZUCAR TERRON MORENO 4 GR",
                format="UNIDAD",
                price=0.0115,
                quantity=quantity,
            )
        ],
    )


class TestBuildDeliveryKey:
    def test_combines_date_and_normalized_sales_point(self):
        delivery = _make_delivery()

        assert build_delivery_key(delivery) == "2026-07-01|bar piscina"

    def test_is_insensitive_to_sales_point_casing_and_spacing(self):
        first = _make_delivery()
        second = Delivery(
            sales_point=SalesPoint(name="  BAR   PISCINA "),
            delivery_date=date(2026, 7, 1),
            products=first.products,
        )

        assert build_delivery_key(first) == build_delivery_key(second)


class TestBuildPayloadHash:
    def test_same_content_produces_the_same_hash(self):
        first = _make_delivery()
        second = _make_delivery()

        assert build_payload_hash(first) == build_payload_hash(second)

    def test_different_quantity_changes_the_hash(self):
        first = _make_delivery(quantity=48.0)
        second = _make_delivery(quantity=49.0)

        assert build_payload_hash(first) != build_payload_hash(second)

    def test_product_order_does_not_affect_the_hash(self):
        product_a = Product(code="1", name="A", format="UNIDAD", price=1.0, quantity=1.0)
        product_b = Product(code="2", name="B", format="UNIDAD", price=2.0, quantity=2.0)

        first = Delivery(
            sales_point=SalesPoint(name="Bar Piscina"),
            delivery_date=date(2026, 7, 1),
            products=[product_a, product_b],
        )
        second = Delivery(
            sales_point=SalesPoint(name="Bar Piscina"),
            delivery_date=date(2026, 7, 1),
            products=[product_b, product_a],
        )

        assert build_payload_hash(first) == build_payload_hash(second)
