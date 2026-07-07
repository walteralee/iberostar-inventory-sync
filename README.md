# Iberostar Inventory Synchronizer

Herramienta desarrollada en Python para sincronizar automáticamente los albaranes PDF generados por Iberostar con sus correspondientes archivos Excel de inventario.

El sistema identifica el punto de venta, extrae los productos y cantidades suministradas y actualiza automáticamente el Excel correspondiente.

---

# 🚀 Instalación

## 1. Clonar el repositorio

```bash
git clone https://github.com/walteralee/iberostar-inventory-sync.git
```

Entrar en la carpeta del proyecto:

```bash
cd iberostar-inventory-sync
```

---

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

## 4. Configurar el proyecto

Crear el archivo `.env` a partir del ejemplo:

Windows

```bash
copy .env.example .env
```

Linux / macOS

```bash
cp .env.example .env
```

---

## 5. Preparar los datos

Colocar los archivos PDF y Excel dentro de la estructura correspondiente:

```
storage/
└── input/
    ├── excels/
    └── pdfs/
```

---

## 6. Ejecutar

```bash
python app/backend/main.py
```

o simplemente

```bash
RUN.bat
```

---

# Características

- Lectura automática de albaranes PDF.
- Identificación automática del punto de venta.
- Extracción automática de códigos y cantidades.
- Actualización automática de archivos Excel.
- Arquitectura modular y escalable.
- Preparado para incorporar interfaz gráfica.
- Sistema de copias de seguridad.
- Código organizado por capas.

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
PDF

↓

Extracción del texto

↓

Interpretación

↓

Identificación del punto de venta

↓

Extracción de productos

↓

Localización del Excel

↓

Actualización automática

↓

Guardar
```

---

# Tecnologías

- Python
- OpenPyXL
- PDFPlumber
- PyMuPDF
- python-dotenv

---

# Estado del proyecto

Actualmente el proyecto se encuentra en desarrollo.

Versión actual

```
v0.1.0
```

---

# Licencia

Proyecto privado.
