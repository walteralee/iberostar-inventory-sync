# Iberostar Inventory Synchronizer

Herramienta desarrollada en Python para sincronizar automáticamente los movimientos de mercancía del Excel de Economato (almacén central) con los Excel mensuales de control de stock de cada punto de venta del hotel.

El sistema identifica el punto de venta, agrupa los productos suministrados por fecha, y actualiza automáticamente el Excel mensual correspondiente, creando productos y plantillas nuevas cuando hace falta.

---

# 🚀 Instalación

```bash
git clone https://github.com/walteralee/iberostar-inventory-sync.git
cd iberostar-inventory-sync
.\RUN.bat
```

## 2. Crear un entorno virtual

Windows

```bash
python -m venv .venv
```

Activar el entorno

```bash
.venv\Scripts\activate
```

Linux / macOS

```bash
python3 -m venv .venv
source .venv/bin/activate
```

---

## 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

---

## 4. Preparar los datos

Colocar los Excel de Economato dentro de la estructura correspondiente:

```
storage/
└── input/
    └── excels/
```

Las plantillas maestras de cada punto de venta ya están en `storage/templates/`.

---

## 5. Ejecutar

```bash
python app/backend/main.py
```

o simplemente

```bash
RUN.bat
```

Al ejecutarse se abre un selector de archivos para elegir el o los Excel de Economato a importar.

---

# Características

- Selección y lectura automática de los Excel de Economato (detecta la hoja correcta sin depender del nombre de las cabeceras).
- Identificación automática del punto de venta y agrupación de movimientos por fecha.
- Creación automática del Excel mensual de cada punto de venta a partir de su plantilla, cuando no existe todavía.
- Creación automática de productos nuevos dentro del Excel, conservando formato y fórmulas.
- Registro persistente de entregas importadas y sincronizadas, con recuperación segura ante interrupciones.
- Copias de seguridad automáticas con retención antes de modificar cualquier Excel.
- Guardado atómico de todos los archivos modificados.
- Preparado para incorporar una interfaz gráfica (scaffold en `app/frontend/`, aún sin implementar).

---

# Arquitectura

```
Iberostar
│
├── app
│   ├── backend
│   └── frontend
│
├── storage
├── docs
├── tests
│
├── README.md
├── requirements.txt
└── RUN.bat
```

---

# Flujo de funcionamiento

```
Selección de los Excel de Economato

↓

Lectura, validación y normalización de movimientos

↓

Agrupación por fecha y punto de venta

↓

Comprobación contra el Registry (nuevas / pendientes / ya sincronizadas)

↓

Por cada entrega pendiente:
    localizar o crear el Excel mensual del punto de venta
    localizar o crear cada producto
    escribir la cantidad en la columna del día
    crear backup y guardar de forma atómica

↓

Actualizar el Registry
```

---

# Tecnologías

- Python
- OpenPyXL
- tkinter (selector de archivos, incluido en la librería estándar)

---

# Estado del proyecto

Actualmente el proyecto está en uso activo para la gestión diaria de inventario.

Versión actual

```
2.0.0
```

---

# Licencia

Proyecto privado.
