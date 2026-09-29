"""
Proyecto:
    Iberostar Inventory Synchronizer

Archivo:
    desktop.py

Descripción:
    Punto de entrada de la aplicación de escritorio.

    Arranca el servidor local (api.py) en un hilo en segundo plano y lo
    muestra dentro de una ventana nativa de Windows (WebView2, incluido
    en Windows 10 y 11). Si WebView2 no estuviera disponible, abre la
    aplicación en el navegador predeterminado y se cierra sola cuando se
    cierra la pestaña.

    También se encarga de:
        - Registrar la actividad y los errores en un archivo de log.
        - Impedir que se abran dos instancias a la vez.
        - Mostrar un mensaje claro si algo falla al arrancar.
"""

from __future__ import annotations

from logging.handlers import RotatingFileHandler
from pathlib import Path
import logging
import os
import shutil
import sys
import threading
import time
import webbrowser

WINDOW_TITLE = "Iberostar · Gestor de Pedidos"
MUTEX_NAME = "Local\\IberostarGestorDePedidos"

# En modo navegador, si la página deja de dar señales durante este
# tiempo se considera cerrada y la aplicación termina.
BROWSER_IDLE_TIMEOUT_SECONDS = 90

logger = logging.getLogger("desktop")


# ==========================================================
# UTILIDADES DE WINDOWS
# ==========================================================


def _show_message(message: str, *, error: bool = False) -> None:
    """Cuadro de diálogo nativo (la app no tiene consola)."""

    if os.name != "nt":
        print(message, file=sys.stderr)
        return

    import ctypes

    MB_ICONERROR = 0x10
    MB_ICONINFORMATION = 0x40

    ctypes.windll.user32.MessageBoxW(
        None,
        message,
        WINDOW_TITLE,
        MB_ICONERROR if error else MB_ICONINFORMATION,
    )


def _acquire_single_instance() -> object | None:
    """
    Devuelve un handle de mutex si esta es la única instancia, o None
    si la aplicación ya estaba abierta (en ese caso la trae al frente).
    """

    if os.name != "nt":
        return object()

    import ctypes

    ERROR_ALREADY_EXISTS = 183
    SW_RESTORE = 9

    kernel32 = ctypes.windll.kernel32
    handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)

    if kernel32.GetLastError() != ERROR_ALREADY_EXISTS:
        return handle

    user32 = ctypes.windll.user32
    existing_window = user32.FindWindowW(None, WINDOW_TITLE)

    if existing_window:
        user32.ShowWindow(existing_window, SW_RESTORE)
        user32.SetForegroundWindow(existing_window)

    return None


# ==========================================================
# LOGGING
# ==========================================================


def _configure_logging(logs_dir: Path) -> Path:
    log_file = logs_dir / "app.log"

    handler = RotatingFileHandler(
        log_file,
        maxBytes=1_000_000,
        backupCount=3,
        encoding="utf-8",
    )
    handler.setFormatter(
        logging.Formatter("%(asctime)s [%(levelname)s] %(name)s: %(message)s")
    )

    root = logging.getLogger()
    root.setLevel(logging.INFO)
    root.addHandler(handler)

    # Cada petición HTTP (incluidos los pings) no aporta nada al log.
    logging.getLogger("werkzeug").setLevel(logging.WARNING)

    return log_file


# ==========================================================
# SERVIDOR
# ==========================================================


class _LocalServer:
    """Servidor Flask en un hilo, en un puerto libre de 127.0.0.1."""

    def __init__(self, app) -> None:
        from werkzeug.serving import make_server

        self.last_activity = time.monotonic()

        @app.before_request
        def _track_activity() -> None:
            self.last_activity = time.monotonic()

        self._server = make_server("127.0.0.1", 0, app, threaded=True)
        self.url = f"http://127.0.0.1:{self._server.server_port}/"
        self._thread = threading.Thread(
            target=self._server.serve_forever,
            name="local-server",
            daemon=True,
        )

    def start(self) -> None:
        self._thread.start()
        logger.info("Servidor local escuchando en %s", self.url)

    def stop(self) -> None:
        self._server.shutdown()


# ==========================================================
# API PARA LA VENTANA NATIVA
# ==========================================================


class DesktopApi:
    """
    Funciones expuestas al frontend como ``window.pywebview.api``.

    Los atributos privados no se exponen a JavaScript.
    """

    def __init__(self) -> None:
        self._window = None

    def attach(self, window) -> None:
        self._window = window

    def save_export(self, sales_point: str, year: int, month: int) -> dict:
        """
        Pide al usuario dónde guardar el Excel mensual y lo copia allí.
        """

        import webview

        from api import ExportNotFoundError, resolve_export_path

        try:
            source = resolve_export_path(sales_point, int(year), int(month))
        except (ValueError, ExportNotFoundError) as error:
            return {"error": str(error)}

        selection = self._window.create_file_dialog(
            webview.SAVE_DIALOG,
            directory=str(Path.home() / "Desktop"),
            save_filename=source.name,
            file_types=("Excel (*.xlsx)",),
        )

        if not selection:
            return {"cancelled": True}

        destination = Path(selection if isinstance(selection, str) else selection[0])

        if destination.suffix.lower() != ".xlsx":
            destination = destination.with_suffix(".xlsx")

        try:
            shutil.copy2(source, destination)
        except PermissionError:
            return {
                "error": "No se pudo guardar el archivo. Si lo tienes abierto "
                "en Excel, ciérralo e inténtalo de nuevo."
            }
        except OSError as error:
            logger.exception("Fallo al exportar a %s", destination)
            return {"error": f"No se pudo guardar el archivo: {error}"}

        logger.info("Exportado %s -> %s", source.name, destination)
        return {"saved": str(destination)}


# ==========================================================
# ARRANQUE
# ==========================================================


def _run_native_window(server: _LocalServer) -> None:
    import webview

    webview.settings["ALLOW_DOWNLOADS"] = True
    webview.settings["SHOW_DEFAULT_MENUS"] = False

    desktop_api = DesktopApi()

    window = webview.create_window(
        WINDOW_TITLE,
        server.url,
        js_api=desktop_api,
        width=1100,
        height=800,
        min_size=(720, 560),
        background_color="#F4F6F9",
        text_select=True,
    )
    desktop_api.attach(window)

    webview.start(
        gui="edgechromium",
        private_mode=False,
        storage_path=str(_webview_storage_dir()),
    )


def _webview_storage_dir() -> Path:
    """
    Caché de WebView2 en una carpeta propia y reconocible
    (%LOCALAPPDATA%\\Iberostar Gestor de Pedidos\\WebView), que el
    desinstalador borra. Por defecto pywebview usaría %APPDATA%\\pywebview.
    """

    from config.settings import APP_DISPLAY_NAME

    local_appdata = os.environ.get("LOCALAPPDATA", "").strip()
    base = Path(local_appdata) if local_appdata else Path.home() / "AppData" / "Local"

    return base / APP_DISPLAY_NAME / "WebView"


def _run_in_browser(server: _LocalServer) -> None:
    webbrowser.open(server.url)

    # El frontend hace ping periódicamente; cuando deja de hacerlo, la
    # pestaña se ha cerrado y no tiene sentido seguir en segundo plano.
    while True:
        time.sleep(5)
        if time.monotonic() - server.last_activity > BROWSER_IDLE_TIMEOUT_SECONDS:
            logger.info("Sin actividad en el navegador: cerrando.")
            return


def main() -> int:
    instance = _acquire_single_instance()

    if instance is None:
        return 0

    try:
        # Importar settings crea la carpeta de datos y copia las plantillas.
        from config.settings import LOGS_DIR

        log_file = _configure_logging(LOGS_DIR)
    except Exception as error:  # Sin log todavía: solo queda avisar.
        _show_message(
            f"No se pudo preparar la carpeta de datos de la aplicación.\n\n{error}",
            error=True,
        )
        return 1

    try:
        from api import create_app
        from config.constants import PROJECT_VERSION

        logger.info("Iniciando versión %s", PROJECT_VERSION)

        server = _LocalServer(create_app())
        server.start()

        try:
            _run_native_window(server)
        except Exception:
            logger.exception("WebView2 no disponible; se usa el navegador.")
            _run_in_browser(server)

        server.stop()
        logger.info("Aplicación cerrada correctamente.")
        return 0

    except Exception as error:
        logger.exception("Error fatal al arrancar la aplicación.")
        _show_message(
            "La aplicación no ha podido iniciarse.\n\n"
            f"{type(error).__name__}: {error}\n\n"
            f"Hay más detalles en:\n{log_file}",
            error=True,
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
