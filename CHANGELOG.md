# Changelog

Todos los cambios relevantes del proyecto. El formato sigue
[Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/) y el proyecto
usa [versionado semántico](https://semver.org/lang/es/).

## [3.0.0] - 2026-09-29

### Añadido
- Aplicación de escritorio para Windows con ventana nativa (WebView2), icono
  propio e instalador que no requiere permisos de administrador.
- Versión portable en ZIP para ejecutar sin instalar.
- Enlace de descarga directa permanente en el README
  (`IberostarGestorPedidos-Setup.exe` / `IberostarGestorPedidos-Portable.zip`).
- Botón *Abrir carpeta de Excel* y diálogo *Guardar como* al exportar.
- Resumen de sincronización con cifras clave (entregas, productos escritos,
  productos nuevos).
- Registro de actividad y errores en `Sistema/logs/app.log`.
- Protección contra abrir la aplicación dos veces.
- Tests de la API web.
- CI con GitHub Actions: tests en cada push y publicación automática del
  instalador al crear un tag.

### Cambiado
- La sincronización agrupa las entregas por Excel mensual y abre, respalda y
  guarda cada libro una sola vez. Un informe de ocho meses pasa de más de una
  hora a unos 25 segundos.
- Los datos de la versión instalada se guardan en
  `Documentos\Iberostar Gestor de Pedidos`.
- La orquestación del proceso se centraliza en `services/sync_pipeline.py`.
- Todas las respuestas de error de la API son JSON con mensajes legibles.

### Eliminado
- Modo consola (`main.py`, `RUN.bat` anterior) y selector de archivos tkinter.
  `RUN.bat` arranca ahora la aplicación de escritorio desde el código fuente.

### Corregido
- El contador de entregas sincronizadas contaba dos veces las entregas
  recuperadas desde el Excel.

## [2.0.0]

- Nueva arquitectura basada en un único Excel consolidado de Economato en
  lugar de albaranes PDF.
- Registry de entregas con recuperación ante interrupciones.
- Primera interfaz web local.
