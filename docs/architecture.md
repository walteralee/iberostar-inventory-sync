# Arquitectura

La aplicación es un programa de escritorio para Windows compuesto por un
backend en Python y una interfaz web que se muestra dentro de una ventana
nativa.

```
┌──────────────────────────── IberostarGestorPedidos.exe ────────────────────────────┐
│                                                                                     │
│   desktop.py ──► ventana nativa (pywebview + WebView2)                               │
│       │                     │  HTML / CSS / JS  (app/frontend)                       │
│       │                     ▼                                                       │
│       └───────► api.py  (Flask, solo 127.0.0.1, puerto libre aleatorio)              │
│                     │                                                               │
│                     ▼                                                               │
│             services/sync_pipeline.py                                               │
│                ├── Importer      lee y valida los informes de Economato             │
│                ├── Registry      historial persistente de entregas (JSON)           │
│                └── Synchronizer  escribe en los Excel mensuales y las plantillas    │
│                                                                                     │
└─────────────────────────────────────────────────────────────────────────────────────┘
                         │
                         ▼
     Documentos\Iberostar Gestor de Pedidos\
        ├── Excel mensuales\<año>\<MES>\<Punto>_<Mes>_<año>.xlsx
        ├── Plantillas\
        ├── Copias de seguridad\
        └── Sistema\ (registry, logs)
```

## Capas

| Carpeta | Responsabilidad |
|---|---|
| `app/backend/desktop.py` | Punto de entrada. Arranca el servidor en un hilo, abre la ventana, impide dos instancias y registra errores en `Sistema/logs/app.log`. |
| `app/backend/api.py` | API HTTP local y servidor del frontend. Todas las respuestas de error son JSON legibles. |
| `app/backend/services/` | Lógica de negocio: importación, registro y sincronización. |
| `app/backend/excel/` | Lectura y escritura de bajo nivel sobre `openpyxl` (índices de productos, inserción de filas conservando formato y fórmulas). |
| `app/backend/models/` | Modelos de datos: `Delivery`, `Product`, `SalesPoint`. |
| `app/backend/config/` | Constantes del dominio (`constants.py`) y rutas (`settings.py`). Ningún otro módulo contiene rutas. |
| `app/frontend/` | Interfaz (HTML, CSS y JavaScript sin dependencias). |
| `packaging/`, `scripts/` | Compilación del ejecutable y del instalador. |

## Flujo de una sincronización

1. **Importación.** Cada informe se procesa por separado: se localiza la hoja con
   datos, se valida y normaliza cada fila (fechas, números en formato español,
   puntos de venta y grupos de producto admitidos) y se agrupan los productos por
   **fecha + punto de venta**. Un archivo defectuoso se descarta sin afectar a
   los demás.
2. **Registry.** Cada entrega se compara con el historial:
   - misma clave y mismo contenido ya sincronizado → se omite;
   - misma clave con contenido distinto → conflicto, se informa y no se toca nada;
   - nueva → se registra como pendiente.

   También se recuperan las entregas que quedaron pendientes en ejecuciones
   anteriores.
3. **Sincronización.** Las entregas se agrupan por Excel mensual. Cada libro se
   abre **una sola vez**, se aplican todas sus entregas, se crean los productos
   nuevos (también en la plantilla) y se guarda con una única copia de seguridad.

## Garantías de integridad

- **Idempotencia.** Cada Excel mensual guarda en una hoja oculta
  (`__SYNC_STATE__`) la clave y la huella SHA-256 de cada entrega aplicada. Ese
  marcador se guarda en el mismo archivo que las cantidades, así que es
  imposible sumar una entrega dos veces aunque el Registry se pierda.
- **Guardado atómico.** Los Excel y el Registry se escriben en un archivo
  temporal y se sustituyen al final: nunca quedan a medio guardar.
- **Copias de seguridad** con retención automática antes de cada modificación.
- **Aislamiento de errores.** Si una entrega falla a mitad de escritura, el libro
  en memoria se descarta y se reaplican las demás desde disco, de modo que la
  entrega fallida no deja cantidades parciales.
- **Bloqueo de proceso.** El Registry usa un bloqueo de sistema operativo: dos
  sincronizaciones simultáneas son imposibles.

## Rutas: desarrollo frente a instalación

`config/settings.py` distingue entre:

- **Recursos** (solo lectura): frontend y plantillas originales. Dentro del
  ejecutable en la versión instalada.
- **Datos** (lectura y escritura): en `storage/` al trabajar con el código
  fuente, y en `Documentos\Iberostar Gestor de Pedidos\` en la versión
  instalada. En el primer arranque se copian ahí las plantillas originales.

La variable de entorno `IBEROSTAR_DATA_DIR` permite usar cualquier otra carpeta.
