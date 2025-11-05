# 🧾 Herramienta Empresarial – Conversor de Archivos

Aplicación de escritorio desarrollada en **Python + Flet**, diseñada para automatizar la conversión de archivos **Excel y PDF a formato TXT**, facilitando la integración con sistemas de gestión internos.  
Esta versión es una **demostración funcional**, en desarrollo, orientada a mostrar la arquitectura y flujo completo del sistema.

---

## 🚀 Funcionalidades principales
- Conversión automática de archivos Excel a TXT con **layout configurable por cliente**.
- **Interfaz gráfica** intuitiva con login y roles (`admin` / `operador`).
- **Base de datos local SQLite** para usuarios, clientes y productos.
- Plantillas **HEAD / LINE** editables para compatibilidad con sistemas externos.
- Generación automática de archivos `.txt` en la carpeta `output/`.

---

## ⚙️ Tecnologías utilizadas
- **Lenguaje:** Python  
- **Framework:** [Flet](https://flet.dev)  
- **Librerías:** `pandas`, `openpyxl`, `sqlite3`, `PyPDF2` (planificada), `pytesseract` (opcional OCR)
- **Base de datos:** SQLite  
- **Sistema operativo objetivo:** Windows (compatible con Mac/Linux)

---

## 🧩 Estructura del proyecto

herramienta-empresarial/
│
├── app_flet_conversion.py # Código principal (UI, lógica, DB, conversión)
├── output/ # Archivos TXT generados (no se incluye en repo)
├── app_data.db # Base de datos SQLite (se genera automáticamente)
└── .gitignore


---

## 🧠 Objetivo y contexto
Desarrollado como **solución interna empresarial**, esta herramienta permite estandarizar archivos de pedidos provenientes de distintos clientes con distintos formatos (Excel, PDF) y adaptarlos a la estructura requerida por un sistema de gestión.  
El diseño modular permite agregar clientes, personalizar layouts y extender funcionalidades fácilmente.

---

## ⚙️ Ejecución
```bash
python -m venv .venv
.venv\Scripts\activate
pip install flet pandas openpyxl
python app_flet_conversion.py

🧱 Estado actual

🧪 Proyecto en desarrollo – versión demostrativa.
Incluye las principales funciones del conversor y la interfaz de usuario.

📫 Contacto

Autor: Joaquín Bach
📧 joaquinbach99@gmail.com

🔗 linkedin.com/in/joaquin-bach-89218b289
