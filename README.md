# Iberostar

Herramienta de sincronización y automatización desarrollada en Python para gestionar la actualización automática de inventarios a partir de albaranes en formato PDF.

El objetivo del proyecto es leer los albaranes generados por el economato, identificar el punto de venta correspondiente, extraer los productos y cantidades suministradas y actualizar automáticamente el archivo Excel asociado.

---

# Características

- Lectura automática de archivos PDF.
- Identificación automática del punto de venta.
- Extracción de productos y cantidades.
- Actualización automática de archivos Excel.
- Arquitectura modular preparada para futuras ampliaciones.
- Sistema de copias de seguridad.
- Preparado para incorporar una interfaz gráfica en el futuro.

---

# Arquitectura

```
Iberostar/

app/
├── backend/
└── frontend/

storage/
docs/
tests/
```

---

# Flujo de funcionamiento

```
PDF

↓

Lectura

↓

Interpretación

↓

Punto de venta

↓

Excel correspondiente

↓

Actualización

↓

Guardar
```

---

# Tecnologías

- Python
- OpenPyXL
- PDFPlumber
- PyMuPDF
- Python Dotenv

---

# Estado del proyecto

Actualmente el proyecto se encuentra en fase de desarrollo.

Versión actual:

```
v0.1
```

---

# Licencia

Proyecto privado.