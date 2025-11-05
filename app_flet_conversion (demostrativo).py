r"""
App de Conversión Excel/PDF -> TXT con Layout Patagonia Sunrise
==============================================================

Descripción rápida
------------------
Aplicación de escritorio (Windows focus, pero corre en Mac/Linux) hecha en **Python + Flet** con:

- **Login con roles**: `admin` (configura clientes/productos) y `operador` (solo convierte archivos).
- **Conversor**: Toma un Excel (y en el futuro PDF), aplica reglas por cliente y genera un archivo `.txt`
  con layout ancho fijo.
- **Base de datos local SQLite** para usuarios, clientes y productos/códigos.
- **Plantilla de Layout configurable por cliente**: así aseguramos compatibilidad con el sistema destino.
- Código modular para que puedas ampliarlo (activar licencias, logs, etc.).

Instalación
-----------
```bash
python -m venv .venv
.venv\Scripts\activate  # Windows
pip install flet pandas openpyxl pymupdf pytesseract pdf2image pillow
```
> *OCR (pytesseract) sólo si vas a procesar PDFs escaneados. Omitirlo por ahora.*

Ejecución
---------
```bash
python app_flet_conversion.py
```

Estructura de datos mínima esperada en Excel
-------------------------------------------
Hoja con encabezados reconocibles (no importa el orden, se detectan por nombre aproximado):

- **EAN** o **Codigo**
- **Descripcion**
- **CodigoInterno** (opcional; si no, se busca en la tabla productos del cliente)
- **Bultos** o **Cajas**
- **TotalUnidades** o calculado: Bultos * UnidadesXBulto
- **UnidadesXBulto**
- **Precio** (unitario)

Si faltan columnas, la app intenta completar desde la DB del cliente. Si aún así faltan datos obligatorios, marca error por fila.

Layout TXT (resumen)
--------------------
Cada archivo generado contiene:

- 1 línea **HEAD** (datos del pedido/cliente).
- N líneas **LINE#** (productos).

Para máxima compatibilidad guardamos un **layout por cliente** en la DB. Cada layout es una cadena plantilla con "tokens" reemplazables (ej: `{GLN_CLIENTE}`) y longitudes definidas. Si no hay plantilla personalizada, usamos un **layout genérico configurable**. 

Notas importantes
-----------------
- El sistema destino que consume el TXT **usa posiciones fijas**. Si no coincide, da error.
- Por eso, ESTA APP TE PERMITE **EDITAR EL LAYOUT** y testear antes de uso productivo.
- Guardamos siempre copia del TXT generado en carpeta `output/` junto con un log.

------------------------------------------------------------------
"""

import re
import os
import sys
import sqlite3
import hashlib
import json
import datetime as dt
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

import flet as ft
import pandas as pd

# ------------------------------------------------------------------
# Configuración básica de paths
# ------------------------------------------------------------------
APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "app_data.db"
OUTPUT_DIR = APP_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

# ------------------------------------------------------------------
# Utilidades generales
# ------------------------------------------------------------------

def hash_password(password: str) -> str:
    """Hash simple SHA256 (suficiente para uso interno; cambiar si querés más seguridad)."""
    return hashlib.sha256(password.encode("utf-8")).hexdigest()


def verify_password(password: str, hashed: str) -> bool:
    return hash_password(password) == hashed


def now_yyyymmdd() -> str:
    return dt.date.today().strftime("%Y%m%d")


def safe_int(v, default=0):
    try:
        return int(str(v).strip())
    except Exception:
        return default


def safe_float(v, default=0.0):
    try:
        return float(str(v).replace(',', '.').strip())
    except Exception:
        return default


# ------------------------------------------------------------------
# Inicialización de base de datos
# ------------------------------------------------------------------
INIT_SQL = [
    # Tabla usuarios
    """
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        role TEXT NOT NULL CHECK (role IN ('admin','operador')),
        active INTEGER NOT NULL DEFAULT 1
    );
    """,
    # Tabla clientes
    """
    CREATE TABLE IF NOT EXISTS clients (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        name_display TEXT NOT NULL,
        address TEXT NOT NULL,
        gln_cliente TEXT NOT NULL,
        gln_destino TEXT NOT NULL,
        gln_alternativo TEXT NOT NULL,
        codigo_cliente TEXT NOT NULL,
        cod_adic TEXT NOT NULL,
        default_emision_days_offset INTEGER DEFAULT 0,
        default_entrega_days_offset INTEGER DEFAULT 0,
        default_venc_days_offset INTEGER DEFAULT 10,
        layout_head TEXT,   -- plantilla HEAD json / formato
        layout_line TEXT    -- plantilla LINE json / formato
    );
    """,
    # Tabla productos por cliente
    """
    CREATE TABLE IF NOT EXISTS products (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        client_id INTEGER NOT NULL,
        ean TEXT NOT NULL,
        desc TEXT NOT NULL,
        cod_int TEXT NOT NULL,
        uxb INTEGER DEFAULT 1,
        precio REAL DEFAULT 0,
        active INTEGER DEFAULT 1,
        FOREIGN KEY (client_id) REFERENCES clients(id)
    );
    """,
]


def init_db():
    new_db = not DB_PATH.exists()
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    for sql in INIT_SQL:
        cur.executescript(sql)
    conn.commit()

    if new_db:
        seed_db(conn)
    conn.close()


def seed_db(conn: sqlite3.Connection):
    """Carga usuarios y cliente de ejemplo."""
    cur = conn.cursor()

    # Usuarios
    users = [
        ("admin", hash_password("admin123"), "admin"),
        ("operador", hash_password("usuario123"), "operador"),
    ]
    cur.executemany(
        "INSERT INTO users (username, password_hash, role) VALUES (?,?,?)",
        users,
    )

    # Cliente ejemplo Patagonia Sunrise - AMBA
    layout_head = json.dumps(default_head_layout_template(), ensure_ascii=False)
    layout_line = json.dumps(default_line_layout_template(), ensure_ascii=False)

    cur.execute(
        """
        INSERT INTO clients (
            name, name_display, address, gln_cliente, gln_destino, gln_alternativo,
            codigo_cliente, cod_adic, layout_head, layout_line
        ) VALUES (?,?,?,?,?,?,?,?,?,?)
        """,
        (
            "patagonia_amba",
            "Patagonia Sunrise - AMBA",
            "AU RICHIERI Y BOULOGNE SUR MER-MCBA",
            "7798355160007",
            "9930709088447",
            "7798355160311",
            "973995",  # ejemplo
            "000000",  # ejemplo
            layout_head,
            layout_line,
        ),
    )
    client_id = cur.lastrowid

    # Productos ejemplo
    productos = [
        (client_id, "7798162980843", "Palta Hass Grande 135 g", "18395929", 70, 1100.0),
        (client_id, "7798162980751", "Zapallo Anco 1 Un", "16405484", 8, 930.0),
        (client_id, "2979900003580", "Limon X Un", "40231318", 21, 260.0),
    ]
    cur.executemany(
        "INSERT INTO products (client_id, ean, desc, cod_int, uxb, precio) VALUES (?,?,?,?,?,?)",
        productos,
    )

    conn.commit()


# ------------------------------------------------------------------
# Plantillas por defecto (pueden editarse en Admin UI)
# ------------------------------------------------------------------

def default_head_layout_template() -> Dict[str, Any]:
    """Layout HEAD default.

    Representado como lista de campos: orden, length, filler, value_key.
    value_key: nombre de campo (ej: GLN_CLIENTE) que será reemplazado.
    filler: espacio (default) si valor menor.

    Usamos JSON para poder editar desde Admin.
    """
    return {
        "total_len": 512,
        "fields": [
            {"key": "LITERAL_HEAD", "value": "HEAD", "len": 4},
            {"key": "GLN_CLIENTE", "len": 13},
            {"key": "GLN_DESTINO", "len": 13},
            {"key": "GLN_ALT", "len": 13},
            {"key": "PAD1", "len": 3},
            {"key": "ORDER_NUM", "len": 7},
            {"key": "PAD2", "len": 2},
            {"key": "NAME_DISPLAY", "len": 40},
            {"key": "ADDRESS", "len": 100},
            {"key": "FECHA_EMISION", "len": 8},
            {"key": "SPC1", "value": "  ", "len": 2},
            {"key": "FECHA_ENTREGA", "len": 8},
            {"key": "SPC2", "value": "  ", "len": 2},
            {"key": "FECHA_VENC", "len": 8},
            {"key": "PAD3", "len": 71},
            {"key": "CODIGO_CLIENTE", "len": 7},
            {"key": "PAD4", "len": 76},
            {"key": "COD_ADIC", "len": 6},
            {"key": "PAD5", "len": 1},
            {"key": "FECHA_ENTREGA2", "len": 8},
            {"key": "SPC3", "value": "  ", "len": 2},
            {"key": "ORDER_NUM2", "len": 7},
            {"key": "TAIL_PAD", "len": 61},
        ],
    }


def default_line_layout_template() -> Dict[str, Any]:
    """Layout LINE default."""
    return {
        "total_len": 384,
        "fields": [
            {"key": "LINE_PREFIX", "len": 8},  # ej: LINE1
            {"key": "EAN", "len": 20},
            {"key": "DESC1", "len": 30},
            {"key": "DESC2", "len": 30},
            {"key": "DESC3", "len": 30},
            {"key": "COD_INT", "len": 15},
            {"key": "BULTOS", "len": 6},
            {"key": "TOTAL_U", "len": 6},
            {"key": "UXB", "len": 4},
            {"key": "PRECIO", "len": 8},
            {"key": "SUBTOTAL", "len": 8},
            {"key": "IVA", "len": 6},
            {"key": "PAD_ZEROS", "len": 108},  # relleno
        ],
    }


# ------------------------------------------------------------------
# Acceso a datos
# ------------------------------------------------------------------

def get_conn():
    return sqlite3.connect(DB_PATH)


def db_get_user(username: str) -> Optional[Tuple]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, username, password_hash, role, active FROM users WHERE username=?", (username,))
    row = cur.fetchone()
    conn.close()
    return row


def db_get_clients() -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, name, name_display, address, gln_cliente, gln_destino, gln_alternativo, codigo_cliente, cod_adic, layout_head, layout_line FROM clients WHERE 1")
    rows = cur.fetchall()
    conn.close()
    cols = ["id", "name", "name_display", "address", "gln_cliente", "gln_destino", "gln_alternativo", "codigo_cliente", "cod_adic", "layout_head", "layout_line"]
    res = []
    for r in rows:
        rec = dict(zip(cols, r))
        res.append(rec)
    return res

def db_update_client(client_id: int, name_display: str, gln_cliente: str, gln_destino: str, address: str = "", codigo_cliente: str = "", cod_adic: str = "") -> None:
    """Actualiza campos editables del cliente en la DB."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        UPDATE clients
        SET name_display = ?, gln_cliente = ?, gln_destino = ?, address = ?, codigo_cliente = ?, cod_adic = ?
        WHERE id = ?
    """, (name_display, gln_cliente, gln_destino, address, codigo_cliente, cod_adic, client_id))
    conn.commit()
    conn.close()


def db_create_client(name: str, name_display: str, address: str, gln_cliente: str, gln_destino: str, gln_alternativo: str = "", codigo_cliente: str = "", cod_adic: str = "", layout_head: Optional[str] = None, layout_line: Optional[str] = None) -> int:
    """Crea un nuevo cliente y devuelve su id."""
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO clients (
            name, name_display, address, gln_cliente, gln_destino, gln_alternativo, codigo_cliente, cod_adic, layout_head, layout_line
        ) VALUES (?,?,?,?,?,?,?,?,?,?)
    """, (name, name_display, address, gln_cliente, gln_destino, gln_alternativo, codigo_cliente, cod_adic, layout_head or "", layout_line or ""))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def db_get_client(client_id: int) -> Optional[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT id, name, name_display, address, gln_cliente, gln_destino, gln_alternativo, codigo_cliente, cod_adic, layout_head, layout_line FROM clients WHERE id=?", (client_id,))
    row = cur.fetchone()
    conn.close()
    if not row:
        return None
    cols = ["id", "name", "name_display", "address", "gln_cliente", "gln_destino", "gln_alternativo", "codigo_cliente", "cod_adic", "layout_head", "layout_line"]
    return dict(zip(cols, row))


def db_get_products_for_client(client_id: int) -> List[Dict[str, Any]]:
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("SELECT ean, desc, cod_int, uxb, precio FROM products WHERE client_id=? AND active=1", (client_id,))
    rows = cur.fetchall()
    conn.close()
    cols = ["ean", "desc", "cod_int", "uxb", "precio"]
    res = [dict(zip(cols, r)) for r in rows]
    return res


# ------------------------------------------------------------------
# Conversión: construcción HEAD / LINE con layout configurable
# ------------------------------------------------------------------

def pad(v: Any, length: int) -> str:
    s = "" if v is None else str(v)
    return s[:length].ljust(length)


def build_head(client: Dict[str, Any], order_num: str, fecha_emision: str, fecha_entrega: str, fecha_venc: str) -> str:
    layout = client.get("layout_head")
    if layout:
        layout = json.loads(layout)
    else:
        layout = default_head_layout_template()

    values = {
        "LITERAL_HEAD": "HEAD",
        "GLN_CLIENTE": client["gln_cliente"],
        "GLN_DESTINO": client["gln_destino"],
        "GLN_ALT": client["gln_alternativo"],
        "ORDER_NUM": order_num,
        "NAME_DISPLAY": client["name_display"],
        "ADDRESS": client["address"],
        "FECHA_EMISION": fecha_emision,
        "FECHA_ENTREGA": fecha_entrega,
        "FECHA_VENC": fecha_venc,
        "CODIGO_CLIENTE": client["codigo_cliente"],
        "COD_ADIC": client["cod_adic"],
        "FECHA_ENTREGA2": fecha_entrega,
        "ORDER_NUM2": order_num,
    }

    out = []
    for fld in layout["fields"]:
        key = fld["key"]
        length = fld["len"]
        if "value" in fld:
            out.append(pad(fld["value"], length))
        else:
            out.append(pad(values.get(key, ""), length))
    head = "".join(out)
    # asegurar largo total
    total = layout.get("total_len", len(head))
    head = head[:total].ljust(total)
    return head


def build_line(layout_line: Dict[str, Any], idx: int, rec: Dict[str, Any]) -> str:
    desc = rec.get("desc", "")
    values = {
        "LINE_PREFIX": f"LINE{idx}",
        "EAN": rec.get("ean", ""),
        "DESC1": desc,
        "DESC2": desc,
        "DESC3": desc,
        "COD_INT": rec.get("cod_int", ""),
        "BULTOS": rec.get("bultos", ""),
        "TOTAL_U": rec.get("total_u", ""),
        "UXB": rec.get("uxb", ""),
        "PRECIO": rec.get("precio", ""),
        "SUBTOTAL": rec.get("precio", ""),
        "IVA": "0.00",
        "PAD_ZEROS": "0",
    }

    out = []
    for fld in layout_line["fields"]:
        key = fld["key"]
        length = fld["len"]
        if key == "PAD_ZEROS":
            out.append("0" * length)
        else:
            out.append(pad(values.get(key, ""), length))

    line = "".join(out)
    total = layout_line.get("total_len", len(line))
    return line[:total].ljust(total)


# ------------------------------------------------------------------
# Lector Excel -> registros de productos
# ------------------------------------------------------------------
COLUMN_ALIASES = {
    "pedido": [
        "pedido", "nro_pedido", "orden", "nro_orden", "order", "order_number", "po_number", "po", "poid"
    ],
    "ean": [
        "ean", "gtin", "codigo", "cod", "codigo_barra", "codigo_ean", "barcode", "barcode_array", "barcode_number"
    ],
    "desc": [
        "descripcion", "desc", "detalle", "producto", "item_description", "product_name", "description", "name"
    ],
    "cod_int": [
        "cod_interno", "codigo_interno", "cod_int", "articulo", "sku", "item_code", "supplier_sku", "sku_id"
    ],
    "bultos": [
        "bultos", "cajas", "packs", "cjs", "shipper", "total_ordered_case", "cases_ordered", "ordered_cases"
    ],
    "total_u": [
        "total_unidades", "unidades", "cantidad", "cant_total", "total_units", "ordered_qty", "ordered_quantity", "ordered_qty"
    ],
    "uxb": [
        "unidadesxbulto", "uxb", "un_x_bulto", "pack_size", "units_per_case", "units_per_case"
    ],
    "precio": [
        "precio", "precio_unitario", "p_unit", "unit_price", "unit_cost", "net_cost", "discounted_unit_cost", "unitcost"
    ],
}

def normalize_col(colname: str) -> Optional[str]:
    """
    Normaliza el nombre de columna y trata de mapearlo a una de las llaves esperadas.
    Estrategia:
      1) limpiar (lower + replace espacios por _)
      2) comparar igualdad
      3) comparar si alias está contenido (in) en el nombre normalizado
      4) comparar si el nombre normalizado está contenido en alias (por si alias tiene prefijo)
    Esto hace el mapeo mucho más robusto frente a nombres reales como
    'barcode_array', 'product_name', 'supplier_sku', 'ordered_qty', etc.
    """
    if not colname:
        return None
    c = colname.strip().lower().replace(" ", "_")
    # primer pase: igualdad exacta
    for key, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            if c == a:
                return key
    # segundo pase: substring (alias dentro del nombre o nombre dentro del alias)
    for key, aliases in COLUMN_ALIASES.items():
        for a in aliases:
            if a in c or c in a:
                return key
    # tercer pase: palabras claves sueltas
    if "po" in c and ("num" in c or "number" in c or "order" in c):
        return "pedido"
    if "barcode" in c or "gtin" in c:
        return "ean"
    if "product" in c or "name" in c or "descripcion" in c:
        return "desc"
    if "sku" in c or "supplier" in c or "cod_int" in c:
        return "cod_int"
    if "case" in c or "bult" in c or "caja" in c:
        return "bultos"
    if "unit" in c and ("ordered" in c or "qty" in c or "cantidad" in c):
        return "total_u"
    if "units_per" in c or "pack_size" in c or "uxb" in c:
        return "uxb"
    if "price" in c or "cost" in c or "precio" in c:
        return "precio"
    return None


def read_excel_products(path: str, client_id: int) -> List[Dict[str, Any]]:
    df = pd.read_excel(path)

    
    # Detectar columnas (mejorado, con prioridades)
    mapping = {}
    for col in df.columns:
        nk = normalize_col(str(col))
        # guardar solo la primera columna encontrada para cada key (evita sobrescribir)
        if nk and nk not in mapping:
            mapping[nk] = col

    # Priorizar columnas explícitas si existen
    for col in df.columns:
        cnorm = str(col).strip().lower().replace(" ", "_")
        # Preferir una columna que contenga 'po_number' como 'pedido'
        if "po_number" == cnorm or "po_number" in cnorm:
            mapping["pedido"] = col
            break

    # Priorizar product_name / product / product title como 'desc'
    for col in df.columns:
        cnorm = str(col).strip().lower().replace(" ", "_")
        if ("product_name" == cnorm or "product_name" in cnorm) or ("product" == cnorm and "name" in cnorm) or ("producto" in cnorm):
            mapping["desc"] = col
            break

    # si no detectó desc, intentar detectar por alias (mantiene la lógica anterior si aplica)
    if "desc" not in mapping:
        for col in df.columns:
            nk = normalize_col(str(col))
            if nk == "desc":
                mapping["desc"] = col
                break

    # --- DETECTAR columnas de fecha PO (si existen) ---
    po_creation_col = None
    po_expected_col = None
    for col in df.columns:
        cn = str(col).strip().lower().replace(" ", "_")
        # ejemplos: 'po_creation_date', 'creation_date', 'po_created_at'
        if ("po_creation" in cn) or ("creation" in cn and "po" in cn) or ("created" in cn and "po" in cn):
            po_creation_col = col
        # ejemplos: 'po_expected_delivery_at', 'expected_delivery', 'po_expected'
        if ("po_expected" in cn) or ("expected_delivery" in cn) or ("expected" in cn and "delivery" in cn):
            po_expected_col = col
    # --- FIN DETECCIÓN ---


    # Productos del cliente (para completar datos si faltan en el Excel)
    prod_db = {p["ean"]: p for p in db_get_products_for_client(client_id)}

    records = []
    for _, row in df.iterrows():
        # Limpieza robusta del EAN (quita listas/representaciones con corchetes, comillas)
        raw_ean = row[mapping.get("ean", list(df.columns)[0])]
        if isinstance(raw_ean, (list, tuple)):
            ean_val = str(raw_ean[0]) if raw_ean else ""
        else:
            ean_val = str(raw_ean)
        # quitar corchetes y comillas si venían como texto
        ean = ean_val.strip().strip("[]'\" ")

        # Desc (nombre del producto) — usar la columna preferida si existe
        desc_col = mapping.get("desc", None)
        desc_val = row[desc_col] if desc_col else ""
        # algunos valores vienen como floats/nan; convertir a string limpia
        desc = "" if pd.isna(desc_val) else str(desc_val).strip()

        # --- Lectura / normalización de fechas PO ---
        def _fmt_date_cell(v):
            """Devuelve YYYYMMDD o '' si no es parseable."""
            if pd.isna(v) or v in (None, ""):
                return ""
            # si es Timestamp / datetime / date
            try:
                if hasattr(v, "strftime"):
                    return v.strftime("%Y%m%d")
            except Exception:
                pass
            # intentar parsear con pandas (maneja textos, excel floats, etc.)
            try:
                parsed = pd.to_datetime(v, errors="coerce")
                if pd.isna(parsed):
                    return ""
                return parsed.strftime("%Y%m%d")
            except Exception:
                return ""

        raw_creation = row[po_creation_col] if po_creation_col else None
        raw_expected = row[po_expected_col] if po_expected_col else None
        fecha_creation = _fmt_date_cell(raw_creation)
        fecha_expected = _fmt_date_cell(raw_expected)
        # --- FIN fechas ---


        rec = {
            "pedido": str(row[mapping.get("pedido", "")]).strip() if mapping.get("pedido") else "",
            "ean": ean,
            "desc": desc,
            "cod_int": str(row[mapping.get("cod_int", "")]).strip()
                        if mapping.get("cod_int")
                        else prod_db.get(ean, {}).get("cod_int", ""),
            "uxb": safe_int(row[mapping.get("uxb", "")])
                        if mapping.get("uxb") else prod_db.get(ean, {}).get("uxb", 1),
            "bultos": safe_int(row[mapping.get("bultos", "")])
                        if mapping.get("bultos") else 0,
            "total_u": safe_int(row[mapping.get("total_u", "")])
                        if mapping.get("total_u") else None,
            "precio": safe_float(row[mapping.get("precio", "")])
                        if mapping.get("precio") else prod_db.get(ean, {}).get("precio", 0),
            "po_creation_date": fecha_creation,
            "po_expected_delivery_at": fecha_expected,

        }


        # Calcular total unidades si falta
        if rec["total_u"] is None or rec["total_u"] == 0:
            rec["total_u"] = rec["bultos"] * rec["uxb"] if rec["bultos"] and rec["uxb"] else rec["bultos"]

        records.append(rec)
    return records




# ------------------------------------------------------------------
# Proceso completo de conversión
# ------------------------------------------------------------------

def convert_file_to_txt(input_path: str, client: Dict[str, Any],
                        output_dir: Path = OUTPUT_DIR) -> List[Path]:
    """
    Lee Excel -> genera uno o más TXT con layout exacto.
    Devuelve lista de rutas de archivos generados.
    """
    client_id = client["id"]
    records = read_excel_products(input_path, client_id)

    # --- Agrupar productos por número de pedido (versión definitiva) ---
    pedidos: Dict[str, List[Dict[str, Any]]] = {}
    current_po = None
    current_group = []

    # Importante: mantener el orden original del Excel
    for rec in records:
        nro_pedido = str(rec.get("pedido", "")).strip()

        # Si el número está vacío, forzar NO_PO (o podrías usar timestamp)
        if not nro_pedido:
            nro_pedido = "NO_PO"

        # Si cambia el número de pedido, guardar el grupo anterior
        if current_po is not None and nro_pedido != current_po:
            pedidos[current_po] = current_group
            current_group = []

        current_po = nro_pedido
        current_group.append(rec)

    # Guardar el último grupo
    if current_po is not None and current_group:
        pedidos[current_po] = current_group
    # --- FIN DEL BLOQUE ---

    
    layout_line = json.loads(client["layout_line"]) if client.get("layout_line") else default_line_layout_template()
    layout_head = json.loads(client["layout_head"]) if client.get("layout_head") else default_head_layout_template()

    generated_files = []

    for nro_pedido, items in pedidos.items():
        # --- Fechas para el HEAD: usar campos del primer item del pedido ---
        first_rec = items[0] if items else {}
        # fecha de creación (violeta) --> po_creation_date
        fecha_emision = first_rec.get("po_creation_date") or now_yyyymmdd()
        # fecha esperada de entrega (amarillo) --> po_expected_delivery_at
        fecha_entrega = first_rec.get("po_expected_delivery_at") or now_yyyymmdd()

        # fecha vencimiento (azul) = fecha_entrega + 5 días (si fecha_entrega válida),
        # si no se puede parsear, fallback a fecha_emision + 10 días
        fecha_venc = ""
        if fecha_entrega:
            try:
                dt_entrega = dt.datetime.strptime(fecha_entrega, "%Y%m%d").date()
                fecha_venc = (dt_entrega + dt.timedelta(days=5)).strftime("%Y%m%d")
            except Exception:
                try:
                    dt_emision = dt.datetime.strptime(fecha_emision, "%Y%m%d").date()
                    fecha_venc = (dt_emision + dt.timedelta(days=10)).strftime("%Y%m%d")
                except Exception:
                    fecha_venc = (dt.date.today() + dt.timedelta(days=10)).strftime("%Y%m%d")
        else:
            # sin fecha_entrega, fallback a emision+10d
            try:
                dt_emision = dt.datetime.strptime(fecha_emision, "%Y%m%d").date()
                fecha_venc = (dt_emision + dt.timedelta(days=10)).strftime("%Y%m%d")
            except Exception:
                fecha_venc = (dt.date.today() + dt.timedelta(days=10)).strftime("%Y%m%d")
        # --- fin fechas ---


        # HEAD
        head = build_head(client, nro_pedido, fecha_emision, fecha_entrega, fecha_venc)

        # LINEs (una por producto del pedido)
        lines = []
        for idx, rec in enumerate(items, start=1):
            line_data = {
                "ean": rec.get("ean", ""),
                "desc": rec.get("desc", ""),
                "cod_int": rec.get("cod_int", ""),
                "uxb": rec.get("uxb", 1),
                "bultos": rec.get("bultos", 0),
                "total_u": rec.get("total_u", 0),
                "precio": f"{safe_float(rec.get('precio', 0)):.2f}",
                "subtotal": f"{safe_float(rec.get('precio', 0)) * safe_int(rec.get('total_u', 0)):.2f}",
            }

            line = build_line(layout_line, idx, line_data)
            lines.append(line)

        # Unir contenido
        content = head + "\n" + "\n".join(lines)

        # -------------------------
        # Nombre de archivo seguro
        # -------------------------
        # Sanitizar nro_pedido para evitar caracteres inválidos en Windows (":", "/", "\" etc.)
        # Reemplazamos cualquier secuencia de caracteres no alfanuméricos por '_'
        safe_nro = re.sub(r'[^A-Za-z0-9\-_.]', '_', str(nro_pedido)).strip('_')
        if not safe_nro:
            # fallback: timestamp
            safe_nro = dt.datetime.now().strftime("%Y%m%d%H%M%S")

        # opcional: limitar longitud
        safe_nro = safe_nro[:64]

        fname = f"ORDERS_{safe_nro}_{client['gln_cliente']}_{client['gln_destino']}_{client['cod_adic']}.txt"
        out_path = output_dir / fname

        # debug: mostrar ruta que se va a escribir (temporal, podés comentar luego)
        print(f"Escribiendo TXT: {out_path}")

        with open(out_path, "w", encoding="utf-8", newline="\n") as f:
            f.write(content)

        generated_files.append(out_path)

    return generated_files



# ------------------------------------------------------------------
# Estado global simple
# ------------------------------------------------------------------
class AppState:
    def __init__(self):
        self.user = None  # (id, username, role)
        self.clients = db_get_clients()
        self.selected_client_id: Optional[int] = self.clients[0]["id"] if self.clients else None

    @property
    def is_admin(self) -> bool:
        return bool(self.user) and self.user["role"] == "admin"

    def set_user(self, row):
        if row:
            self.user = {
                "id": row[0],
                "username": row[1],
                "password_hash": row[2],
                "role": row[3],
                "active": row[4],
            }
        else:
            self.user = None


# ------------------------------------------------------------------
# UI: Login
# ------------------------------------------------------------------

def login_view(page: ft.Page, state: AppState):
    page.views.clear()

    username_tf = ft.TextField(label="Usuario", autofocus=True, width=250)
    password_tf = ft.TextField(label="Contraseña", password=True, can_reveal_password=True, width=250)
    error_txt = ft.Text("", color=ft.Colors.RED, size=12)

    def do_login(e=None):
        uname = username_tf.value.strip()
        pwd = password_tf.value
        row = db_get_user(uname)
        if not row:
            error_txt.value = "Usuario no encontrado"
            page.update()
            return
        if row[4] != 1:
            error_txt.value = "Usuario inactivo"
            page.update()
            return
        if not verify_password(pwd, row[2]):
            error_txt.value = "Contraseña incorrecta"
            page.update()
            return
        state.set_user(row)
        main_view(page, state)

    login_btn = ft.ElevatedButton("Ingresar", on_click=do_login)
    page.views.append(
        ft.View(
            "/login",
            controls=[
                ft.Column([
                    ft.Text("Ingreso al Sistema", size=20, weight=ft.FontWeight.BOLD),
                    username_tf,
                    password_tf,
                    login_btn,
                    error_txt,
                ], alignment=ft.MainAxisAlignment.CENTER, horizontal_alignment=ft.CrossAxisAlignment.CENTER),
            ],
            vertical_alignment=ft.MainAxisAlignment.CENTER,
            horizontal_alignment=ft.CrossAxisAlignment.CENTER,
        )
    )
    page.update()


# ------------------------------------------------------------------
# UI: Principal
# ------------------------------------------------------------------

def main_view(page: ft.Page, state: AppState):
    page.views.clear()

    # Tabs: Conversión (siempre) + Admin (solo admin)
    tabs = []
    conv_tab = ft.Tab(text="Conversor", content=conversion_tab(page, state))
    tabs.append(conv_tab)
    if state.is_admin:
        admin_tab = ft.Tab(text="Admin", content=admin_tab_content(page, state))
        tabs.append(admin_tab)

    logout_btn = ft.ElevatedButton("Salir", icon=ft.Icons.LOGOUT, on_click=lambda e: logout(page, state))

    page.views.append(
        ft.View(
            "/main",
            controls=[
                ft.Row([ft.Text(f"Usuario: {state.user['username']} ({state.user['role']})", size=14), logout_btn]),
                ft.Tabs(tabs=tabs, expand=1),
            ],
        )
    )
    page.update()


def logout(page: ft.Page, state: AppState):
    state.set_user(None)
    login_view(page, state)


# ------------------------------------------------------------------
# Conversor TAB
# ------------------------------------------------------------------

def conversion_tab(page: ft.Page, state: AppState) -> ft.Column:
    clients = state.clients
    dd_clients = ft.Dropdown(
        label="Cliente",
        width=400,
        options=[ft.dropdown.Option(str(c["id"]), c["name_display"]) for c in clients],
        value=str(state.selected_client_id) if state.selected_client_id else None,
    )

    picked_file_path = {"path": None}
    picked_lbl = ft.Text("Ningún archivo seleccionado", italic=True, size=12)

    def on_pick_result(e: ft.FilePickerResultEvent):
        if e.files:
            picked_file_path["path"] = e.files[0].path
            picked_lbl.value = e.files[0].name
        else:
            picked_file_path["path"] = None
            picked_lbl.value = "Ningún archivo seleccionado"
        page.update()

    file_picker = ft.FilePicker(on_result=on_pick_result)
    page.overlay.append(file_picker)

    pick_btn = ft.ElevatedButton("Seleccionar Excel", icon=ft.Icons.UPLOAD_FILE, on_click=lambda e: file_picker.pick_files(allow_multiple=False))

    status_txt = ft.Text("", size=12)

    def do_convert(e):
        # Mensaje inicial y refresco inmediato para mostrar "Procesando..."
        status_txt.value = "Procesando..."
        page.update()

        if not picked_file_path["path"]:
            status_txt.value = "Seleccioná un archivo primero."
            page.update()
            return

        client_id = int(dd_clients.value)
        client = db_get_client(client_id)
        if not client:
            status_txt.value = "Cliente inválido."
            page.update()
            return

        try:
            generated_files = convert_file_to_txt(
                input_path=picked_file_path["path"],
                client=client,
                output_dir=OUTPUT_DIR,
            )
        except Exception as exc:
            status_txt.value = f"Error durante la conversión: {exc}"
            page.update()
            return

        # refrescar estado global si es necesario
        state.clients = db_get_clients()

        n = len(generated_files)
        if n == 0:
            msg = "No se generaron archivos."
        elif n <= 10:
            msg = f"Se generaron {n} archivos: " + ", ".join([f.name for f in generated_files])
        else:
            names_preview = ", ".join([f.name for f in generated_files[:10]])
            msg = f"Se generaron {n} archivos. Ejemplos: {names_preview}..."

        status_txt.value = msg
        page.update()

        try:
            page.snack_bar = ft.SnackBar(ft.Text(f"Conversión finalizada: {n} archivos"), open=True, bgcolor=ft.colors.SURFACE_VARIANT)
            page.update()
        except Exception:
            pass

    convert_btn = ft.ElevatedButton("Convertir", icon=ft.Icons.PLAY_ARROW, on_click=do_convert)

    # Envolvemos el bloque de inputs en un Container con padding superior para separar de la barra superior
    form_column = ft.Column([
        dd_clients,
        ft.Row([pick_btn, picked_lbl]),
        convert_btn,
        status_txt,
    ], scroll=ft.ScrollMode.AUTO, expand=False)

    return ft.Container(
        padding=ft.padding.only(top=16, left=12, right=12),
        content=form_column
    )


# ------------------------------------------------------------------
# Admin TAB
# ------------------------------------------------------------------

def admin_tab_content(page: ft.Page, state: AppState) -> ft.Column:
    # Construye filas (cada fila incluye un botón "Editar" para cargar el cliente en el formulario)
    def build_table_rows():
        rows = []
        for c in state.clients:
            cid = c["id"]
            rows.append(
                ft.DataRow(
                    cells=[
                        ft.DataCell(ft.Text(str(c["id"]))),
                        ft.DataCell(ft.Text(c["name_display"])),
                        ft.DataCell(ft.Text(c["gln_cliente"])),
                        ft.DataCell(ft.Text(c["gln_destino"])),
                        # Columna acción: botón Editar (compatible con versiones antiguas y nuevas)
                        ft.DataCell(
                            ft.IconButton(
                                icon=ft.Icons.EDIT,
                                tooltip="Editar",
                                on_click=lambda e, _cid=cid: load_client_into_form(_cid)
                            )
                        ),
                    ]
                )
            )
        return rows

    # Holder para id seleccionado
    sel_client_id = {"id": None}

    # Campos del formulario de edición
    id_txt = ft.Text(value="", disabled=True)
    name_tf = ft.TextField(label="Nombre a mostrar", width=360)
    gln_cliente_tf = ft.TextField(label="GLN Cliente", width=360)
    gln_destino_tf = ft.TextField(label="GLN Destino", width=360)
    address_tf = ft.TextField(label="Address (opcional)", width=360)
    codigo_cliente_tf = ft.TextField(label="Código Cliente (opcional)", width=360)
    cod_adic_tf = ft.TextField(label="Cod Adic (opcional)", width=360)

    edit_status = ft.Text("", size=12)

    # Tabla inicial (NO usar on_row_selected ni key)
    tbl = ft.DataTable(
        columns=[
            ft.DataColumn(ft.Text("ID")),
            ft.DataColumn(ft.Text("Nombre")),
            ft.DataColumn(ft.Text("GLN Cliente")),
            ft.DataColumn(ft.Text("GLN Destino")),
            ft.DataColumn(ft.Text("Acción")),
        ],
        rows=build_table_rows(),
        width=980,
    )

    # Función que carga un cliente en el formulario (llamada por el botón Editar)
    def load_client_into_form(cid: int):
        sel_client_id["id"] = cid
        client = db_get_client(cid)
        if client:
            id_txt.value = str(client["id"])
            name_tf.value = client.get("name_display", "")
            gln_cliente_tf.value = client.get("gln_cliente", "")
            gln_destino_tf.value = client.get("gln_destino", "")
            address_tf.value = client.get("address", "")
            codigo_cliente_tf.value = client.get("codigo_cliente", "")
            cod_adic_tf.value = client.get("cod_adic", "")
            edit_status.value = f"Cliente {cid} cargado para editar."
        else:
            edit_status.value = "Cliente no encontrado."
        page.update()

    def refresh_clients_and_table(e=None):
        state.clients = db_get_clients()
        tbl.rows = build_table_rows()
        edit_status.value = "Clientes actualizados."
        page.update()

    def save_changes(e):
        if not sel_client_id["id"]:
            edit_status.value = "Seleccioná un cliente para editar."
            page.update()
            return
        try:
            db_update_client(
                client_id=sel_client_id["id"],
                name_display=name_tf.value.strip(),
                gln_cliente=gln_cliente_tf.value.strip(),
                gln_destino=gln_destino_tf.value.strip(),
                address=address_tf.value.strip(),
                codigo_cliente=codigo_cliente_tf.value.strip(),
                cod_adic=cod_adic_tf.value.strip(),
            )
            edit_status.value = "Cambios guardados."
            refresh_clients_and_table()
        except Exception as ex:
            edit_status.value = f"Error al guardar: {ex}"
            page.update()

    save_btn = ft.ElevatedButton("Guardar cambios", icon=ft.Icons.SAVE, on_click=save_changes)

    # Diálogo para crear nuevo cliente
    def open_new_client_dialog(e=None):
        new_name_tf = ft.TextField(label="Nombre (internal, sin espacios)", width=360, value="")
        new_name_display_tf = ft.TextField(label="Nombre a mostrar", width=360, value="")
        new_address_tf = ft.TextField(label="Address", width=360, value="")
        new_gln_cliente_tf = ft.TextField(label="GLN Cliente", width=360, value="")
        new_gln_destino_tf = ft.TextField(label="GLN Destino", width=360, value="")
        new_cod_cliente_tf = ft.TextField(label="Código Cliente", width=360, value="")
        new_cod_adic_tf = ft.TextField(label="Cod Adic (opcional)", width=360, value="")

        new_status = ft.Text("", size=12)

        def create_new(ev):
            internal_name = new_name_tf.value.strip()
            display = new_name_display_tf.value.strip()
            gln_c = new_gln_cliente_tf.value.strip()
            gln_d = new_gln_destino_tf.value.strip()
            if not internal_name or not display or not gln_c or not gln_d:
                new_status.value = "Completá los campos obligatorios (nombre, nombre a mostrar, GLNs)."
                page.update()
                return
            try:
                new_id = db_create_client(
                    name=internal_name,
                    name_display=display,
                    address=new_address_tf.value.strip(),
                    gln_cliente=gln_c,
                    gln_destino=gln_d,
                    gln_alternativo="",
                    codigo_cliente=new_cod_cliente_tf.value.strip(),
                    cod_adic=new_cod_adic_tf.value.strip(),
                    layout_head=json.dumps(default_head_layout_template(), ensure_ascii=False),
                    layout_line=json.dumps(default_line_layout_template(), ensure_ascii=False),
                )
                new_status.value = f"Cliente creado con id {new_id}."
                page.dialog.open = False
                refresh_clients_and_table()
                page.update()
            except Exception as ex:
                new_status.value = f"Error creando cliente: {ex}"
                page.update()

        create_btn = ft.ElevatedButton("Crear cliente", on_click=create_new)
        cancel_btn = ft.TextButton("Cancelar", on_click=lambda ev: (setattr(page.dialog, "open", False), page.update()))

        dialog_content = ft.Column([
            ft.Text("Nuevo Cliente", weight=ft.FontWeight.BOLD),
            new_name_tf,
            new_name_display_tf,
            new_address_tf,
            new_gln_cliente_tf,
            new_gln_destino_tf,
            new_cod_cliente_tf,
            new_cod_adic_tf,
            ft.Row([create_btn, cancel_btn]),
            new_status,
        ])

        page.dialog = ft.AlertDialog(content=dialog_content, actions=[], on_dismiss=lambda ev: None)
        page.dialog.open = True
        page.update()

    new_client_btn = ft.ElevatedButton("Agregar nuevo cliente", icon=ft.Icons.ADD, on_click=open_new_client_dialog)
    refresh_btn = ft.ElevatedButton("Refrescar", icon=ft.Icons.REFRESH, on_click=refresh_clients_and_table)

    left_col = ft.Column([tbl, ft.Row([refresh_btn, new_client_btn])], expand=True)
    right_col = ft.Column([
        ft.Text("Editar cliente seleccionado", size=14, weight=ft.FontWeight.BOLD),
        id_txt,
        name_tf,
        gln_cliente_tf,
        gln_destino_tf,
        address_tf,
        codigo_cliente_tf,
        cod_adic_tf,
        ft.Row([save_btn]),
        edit_status,
    ], width=420)

    return ft.Row([left_col, right_col], alignment=ft.MainAxisAlignment.START)


# ------------------------------------------------------------------
# App Flet main()
# ------------------------------------------------------------------

def main(page: ft.Page):
    page.title = "Conversor Pedidos"
    page.window_width = 900
    page.window_height = 700
    page.horizontal_alignment = ft.CrossAxisAlignment.CENTER

    # init DB (solo una vez al inicio del proceso Python)
    if not hasattr(main, "_db_init_done"):
        init_db()
        main._db_init_done = True

    state = AppState()
    login_view(page, state)


# ------------------------------------------------------------------
# Entry point
# ------------------------------------------------------------------
if __name__ == "__main__":
    ft.app(target=main)
