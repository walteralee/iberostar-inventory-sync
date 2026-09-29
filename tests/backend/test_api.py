"""
Pruebas de la API web local.

La sincronización real se sustituye por un doble (``run_sync``) para no
tocar nunca storage/; aquí solo se comprueba el contrato HTTP: códigos
de estado, validación de entrada y mensajes de error en JSON.
"""

from io import BytesIO

import pytest

import api
from services.importer import ImportSummary
from services.sync_pipeline import SyncResult
from services.synchronizer import _SynchronizationTotals


@pytest.fixture
def client(tmp_path, monkeypatch):
    monkeypatch.setattr(api, "MONTHLY_EXCELS_DIR", tmp_path / "excels")
    return api.create_app().test_client()


def _upload(name="informe.xlsx"):
    return {"files": [(BytesIO(b"contenido"), name)]}


class TestInfo:
    def test_returns_name_and_version(self, client):
        payload = client.get("/api/info").get_json()

        assert payload["name"]
        assert payload["version"]

    def test_unknown_api_route_returns_json_404(self, client):
        response = client.get("/api/no-existe")

        assert response.status_code == 404
        assert "error" in response.get_json()

    def test_serves_the_frontend(self, client):
        response = client.get("/")

        assert response.status_code == 200
        assert b"<!doctype html>" in response.data.lower()


class TestSync:
    def test_rejects_a_request_without_files(self, client):
        response = client.post("/api/sync")

        assert response.status_code == 400
        assert "error" in response.get_json()

    def test_rejects_files_that_are_not_xlsx(self, client):
        response = client.post(
            "/api/sync",
            data=_upload("informe.csv"),
            content_type="multipart/form-data",
        )

        assert response.status_code == 400
        assert "informe.csv" in response.get_json()["error"]

    def test_returns_the_summary_of_the_pipeline(self, client, monkeypatch):
        received = {}

        def fake_run_sync(paths):
            received["paths"] = list(paths)
            assert all(path.is_file() for path in received["paths"])
            return SyncResult(
                import_summary=ImportSummary(),
                pending_warnings=[],
                synchronization_totals=_SynchronizationTotals(
                    synchronized_deliveries=2,
                    products_written=5,
                ),
            )

        monkeypatch.setattr(api, "run_sync", fake_run_sync)

        response = client.post(
            "/api/sync",
            data=_upload(),
            content_type="multipart/form-data",
        )
        payload = response.get_json()

        assert response.status_code == 200
        assert payload["has_incidents"] is False
        assert payload["synchronization_totals"]["synchronized_deliveries"] == 2

        # Los archivos subidos se eliminan al terminar.
        assert not any(path.exists() for path in received["paths"])

    def test_a_locked_registry_returns_409(self, client, monkeypatch):
        def locked(_paths):
            raise RuntimeError("bloqueado")

        monkeypatch.setattr(api, "run_sync", locked)

        response = client.post(
            "/api/sync",
            data=_upload(),
            content_type="multipart/form-data",
        )

        assert response.status_code == 409

    def test_unexpected_errors_return_a_readable_500(self, client, monkeypatch):
        def broken(_paths):
            raise KeyError("fallo interno")

        monkeypatch.setattr(api, "run_sync", broken)

        response = client.post(
            "/api/sync",
            data=_upload(),
            content_type="multipart/form-data",
        )

        assert response.status_code == 500
        assert "error inesperado" in response.get_json()["error"].lower()


class TestExport:
    def test_rejects_non_numeric_year_and_month(self, client):
        response = client.get("/api/export?sales_point=Comedor&year=x&month=y")

        assert response.status_code == 400

    def test_rejects_unknown_sales_points(self, client):
        response = client.get("/api/export?sales_point=Nope&year=2026&month=1")

        assert response.status_code == 400

    def test_missing_month_returns_404(self, client, monkeypatch, tmp_path):
        monkeypatch.setattr(
            api.ExcelTemplateManager,
            "get_excel_path",
            lambda self, sales_point, year, month: tmp_path / "no-existe.xlsx",
        )

        response = client.get("/api/export?sales_point=Comedor&year=2026&month=1")

        assert response.status_code == 404


class TestPeriods:
    def test_lists_only_valid_month_folders_in_calendar_order(self, client, tmp_path):
        for month in ("MARZO", "ENERO", "no-es-un-mes"):
            (tmp_path / "excels" / "2026" / month).mkdir(parents=True)

        assert client.get("/api/periods").get_json() == {"2026": ["ENERO", "MARZO"]}
