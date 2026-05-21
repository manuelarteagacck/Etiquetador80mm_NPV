# app.py - Aplicación Etiquetador 80mm NPV
# (Versión rediseñada por Gemini)
# Ejecutar: python app.py
print("Aplicación Etiquetador80mm NPV lista para ejecutar.")


# -*- coding: utf-8 -*-
"""
Etiquetador 80 mm — NPV
Python 3.11+ (Windows)

Requisitos clave:
- Tkinter UI para búsqueda/escaneo/edición con diseño mejorado.
- Conexión a SQL Server (pyodbc, ODBC Driver 17+)
- Impresión ESC/POS (python-escpos) y fallback por Spooler Windows (win32print)
- Persistencia en config.json
- Listo para empaquetar con PyInstaller:  pyinstaller -F -n Etiquetador80mm app.py
"""

import json
import os
import re
import sys
import textwrap
import datetime
import traceback
from dataclasses import dataclass, asdict
from typing import Optional, Dict, Any, List

APP_TITLE = "Etiquetador 80mm - TIENDA NPV"
CONFIG_FILENAME = "config.json"
FERNET_KEY_FILENAME = "config.key"
PROMO_BACKGROUND_FILENAME = "back5.png"
PROMO_TEMPLATE_SIZE = (1585, 992)


def _missing_dependency_exit(module_name: str, package_name: str) -> None:
    message = (
        f"Falta el modulo '{module_name}'.\n\n"
        "Instala las dependencias en el mismo Python con el que ejecutas la app:\n"
        f"  {sys.executable} -m pip install -r requirements.txt\n\n"
        "O instala solo este paquete:\n"
        f"  {sys.executable} -m pip install {package_name}"
    )
    print(message, file=sys.stderr)
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, message, APP_TITLE, 0x10)
        except Exception:
            pass
    raise SystemExit(1)

# --- Impresión / Windows ---
try:
    import win32print
    import win32con
    import win32ui
except ModuleNotFoundError as exc:
    if exc.name and exc.name.startswith("win32"):
        _missing_dependency_exit(exc.name, "pywin32")
    raise

# --- ESC/POS ---
try:
    from escpos.printer import Serial, Usb, Network, Dummy
except ModuleNotFoundError as exc:
    if exc.name and exc.name.startswith("escpos"):
        _missing_dependency_exit(exc.name, "python-escpos==3.1")
    raise

# --- DB ---
try:
    import pyodbc
except ModuleNotFoundError as exc:
    if exc.name == "pyodbc":
        _missing_dependency_exit(exc.name, "pyodbc")
    raise

# --- Cifrado de secretos ---
try:
    from cryptography.fernet import Fernet, InvalidToken
except ModuleNotFoundError as exc:
    if exc.name == "cryptography":
        _missing_dependency_exit(exc.name, "cryptography")
    raise

# --- UI ---
import tkinter as tk
from tkinter import ttk, messagebox, font
try:
    from PIL import Image, ImageTk, ImageWin
except ModuleNotFoundError:
    Image = None
    ImageTk = None
    ImageWin = None
try:
    from ttkthemes import ThemedTk
except ModuleNotFoundError as exc:
    if exc.name == "ttkthemes":
        _missing_dependency_exit(exc.name, "ttkthemes==3.2.2")
    raise

APP_TITLE = "Etiquetador 80mm — TIENDA NPV"
CONFIG_FILENAME = "config.json"
FERNET_KEY_FILENAME = "config.key"


# =============================
# Utilidades de configuración
# =============================
DEFAULT_CONFIG = {
    "conexion_odbc": {
        # Ejemplo con autenticación integrada:
        # "dsn": "",  # si usas DSN
        "server": ".\\POS",
        "database": "NPV",
        "trusted_connection": False,
        "username": "sa",
        "password_encrypted": "",
        "driver": "{ODBC Driver 17 for SQL Server}"
    },
    "impresora": {
        "modo": "default",  # "default" | "named"
        "printer_name": ""  # usado cuando modo="named"
    },
    "escpos": {
        # Método físico de conexión a la térmica cuando se use modo escpos
        "mode": "spooler_windows",  # "serial" | "usb" | "network" | "spooler_windows"
        "serial": {"com": "COM3", "baudrate": 9600, "bytesize": 8, "parity": "N", "stopbits": 1, "timeout": 1},
        "usb": {"vendor_id": 0x0000, "product_id": 0x0000, "interface": 0, "in_ep": 0x82, "out_ep": 0x01},
        "network": {"ip": "192.168.1.50", "port": 9100}
    },
    "tamaños": {
        "size_producto": [2, 1],   # [width, height]
        "size_precio": [2, 2],
        "size_codigos": [1, 1],
        "size_vigencia": [1, 1]
    },
    "espaciados": {
        "espacio_arriba_precio": 0,
        "espacio_abajo_precio": 0
    },
    "corte": "partial",  # "partial" | "full"
    "vigencia_default": "Vigente {MES_ABR} {YYYY}",  # plantilla con macros
    "preferencias": {
        "auto_print_on_scan": False,
        "individual_print": False
    }
}


def month_abbr_es(dt: datetime.date) -> str:
    meses = ["ENE","FEB","MAR","ABR","MAY","JUN","JUL","AGO","SEP","OCT","NOV","DIC"]
    return meses[dt.month - 1]


def resolve_vigencia_template(tpl: str) -> str:
    today = datetime.date.today()
    rep = {
        "{YYYY}": f"{today.year:04d}",
        "{YY}": f"{today.year%100:02d}",
        "{MM}": f"{today.month:02d}",
        "{MES_ABR}": month_abbr_es(today),
    }
    out = tpl
    for k, v in rep.items():
        out = out.replace(k, v)
    return out


def load_config() -> dict:
    path = config_path()
    if not os.path.exists(path):
        cfg = json.loads(json.dumps(DEFAULT_CONFIG))
        _hydrate_db_password(cfg)
        save_config(cfg)
        return cfg
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    # merge simple (rellena faltantes con defaults)
    def deep_merge(a, b):
        for k, v in b.items():
            if k not in a:
                a[k] = v
            else:
                if isinstance(v, dict) and isinstance(a[k], dict):
                    deep_merge(a[k], v)
        return a
    cfg = deep_merge(data, json.loads(json.dumps(DEFAULT_CONFIG)))
    password_migrated = _hydrate_db_password(cfg)
    if password_migrated:
        save_config(cfg)
    return cfg


def save_config(cfg: dict):
    disk_cfg = _config_for_disk(cfg)
    with open(config_path(), "w", encoding="utf-8") as f:
        json.dump(disk_cfg, f, ensure_ascii=False, indent=2)


def resource_path(filename: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, filename)


def app_base_path() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def config_path() -> str:
    return os.path.join(app_base_path(), CONFIG_FILENAME)


def fernet_key_path() -> str:
    return os.path.join(app_base_path(), FERNET_KEY_FILENAME)


def log_path() -> str:
    return os.path.join(app_base_path(), "log.txt")


def _load_or_create_fernet_key() -> bytes:
    path = fernet_key_path()
    if os.path.exists(path):
        with open(path, "rb") as f:
            key = f.read().strip()
        try:
            Fernet(key)
        except Exception as exc:
            raise ValueError(f"La llave Fernet no es valida: {path}") from exc
        return key

    key = Fernet.generate_key()
    with open(path, "wb") as f:
        f.write(key)
    try:
        os.chmod(path, 0o600)
    except Exception:
        pass
    return key


def _get_fernet() -> Fernet:
    return Fernet(_load_or_create_fernet_key())


def _encrypt_secret(value: str) -> str:
    if not value:
        return ""
    return _get_fernet().encrypt(str(value).encode("utf-8")).decode("ascii")


def _decrypt_secret(token: str) -> str:
    if not token:
        return ""
    try:
        return _get_fernet().decrypt(str(token).encode("ascii")).decode("utf-8")
    except InvalidToken as exc:
        raise ValueError(
            "No se pudo descifrar conexion_odbc.password_encrypted. "
            f"Verifica que {FERNET_KEY_FILENAME} corresponda a este config.json."
        ) from exc


def _hydrate_db_password(cfg: dict) -> bool:
    c = cfg.setdefault("conexion_odbc", {})
    had_plain_password = "password" in c
    plain_password = c.get("password")
    encrypted_password = c.get("password_encrypted")

    if had_plain_password and plain_password:
        c["password"] = str(plain_password)
        return True
    if encrypted_password:
        c["password"] = _decrypt_secret(encrypted_password)
    else:
        c["password"] = ""
    return had_plain_password


def _config_for_disk(cfg: dict) -> dict:
    disk_cfg = json.loads(json.dumps(cfg))
    c = disk_cfg.setdefault("conexion_odbc", {})
    password = c.pop("password", None)
    if password:
        existing_token = c.get("password_encrypted")
        if existing_token:
            try:
                if _decrypt_secret(existing_token) == str(password):
                    return disk_cfg
            except ValueError:
                pass
        c["password_encrypted"] = _encrypt_secret(str(password))
    else:
        c["password_encrypted"] = c.get("password_encrypted", "")
    return disk_cfg


def _sanitize_connection_string(conn_str: str) -> str:
    return re.sub(r"(PWD=)[^;]*", r"\1***", conn_str or "", flags=re.IGNORECASE)


def write_log(context: str, exc: Exception = None, extra: str = "") -> None:
    try:
        lines = [
            "=" * 80,
            datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            context,
        ]
        if extra:
            lines.append(extra)
        if exc is not None:
            lines.append("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)).rstrip())
        with open(log_path(), "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n")
    except Exception:
        pass


def get_promo_background_path() -> str:
    return resource_path(PROMO_BACKGROUND_FILENAME)


# =============================
# Conexión a SQL Server
# =============================
def _build_connection_string(cfg: dict) -> str:
    c = cfg["conexion_odbc"]
    driver = c.get("driver") or "{ODBC Driver 17 for SQL Server}"
    pieces = [f"DRIVER={driver}"]
    if c.get("server"):
        pieces.append(f"SERVER={c['server']}")
    if c.get("database"):
        pieces.append(f"DATABASE={c['database']}")
    if c.get("trusted_connection"):
        pieces.append("Trusted_Connection=yes")
    else:
        pieces.append(f"UID={c.get('username','')}")
        pieces.append(f"PWD={c.get('password','')}")
    # Si usa DSN, el driver puede omitirse y se puede retornar "DSN=..."
    if c.get("dsn"):
        return f"DSN={c['dsn']}"
    return ";".join(pieces) + ";"


def _get_connection(cfg: dict):
    conn_str = _build_connection_string(cfg)
    return pyodbc.connect(conn_str, timeout=5)


def _normalize_upc_digits(s: Optional[str]) -> str:
    if not s:
        return ""
    return re.sub(r"[^0-9]", "", s)


def _format_price(value) -> str:
    try:
        v = float(value)
        return f"{v:.2f}"
    except Exception:
        return ""

def _format_date(value) -> str:
    if not value:
        return ""
    if isinstance(value, datetime.datetime):
        return value.strftime("%d/%m/%Y")
    if isinstance(value, datetime.date):
        return value.strftime("%d/%m/%Y")
    return str(value)

def _price_to_float(value) -> Optional[float]:
    try:
        if value is None:
            return None
        digits = re.sub(r"[^0-9.]", "", str(value))
        if not digits:
            return None
        return float(digits)
    except Exception:
        return None


def _format_currency(value) -> str:
    text = str(value or "").strip()
    if not text:
        return "$0.00"
    if text.startswith("$"):
        return text
    number = _price_to_float(text)
    if number is None:
        return f"${text}"
    return f"${number:.2f}"


def _split_price(value) -> tuple:
    """Devuelve (entero, decimales) de un valor de precio."""
    v = _price_to_float(value)
    if v is None:
        return ("", "")
    parts = f"{v:.2f}".split(".")
    return (parts[0], parts[1])

def search_items(term: str, cfg: dict = None) -> List[Dict[str, Any]]:
    """
    Busca artículos por UPC, ARTICULO, o DESCRIPCION (LIKE).
    Devuelve una lista de dicts, cada uno con ARTICULO, DESCRIPCION, PRECIO, UPC, VIGENCIA.
    Maneja la obtención del precio más reciente para cada artículo.
    """
    cfg = cfg or load_config()
    conn = None
    results = []
    term_digits = _normalize_upc_digits(term)
    like_term = f"%{term}%"

    sql = """
    WITH ArticulosEncontrados AS (
        -- Find all unique ARTICULO keys from the three search methods
        SELECT ARTICULO FROM NPV.dbo.NPVFDArticulos WHERE ARTICULO = ?
        UNION
        SELECT ARTICULO FROM NPV.dbo.NPVFDArticulos WHERE DESCRIPCION LIKE ?
        UNION
        SELECT ARTICULO FROM NPV.dbo.NPVFDArticulosEquivalentes WHERE EQUIVALENTE = ?
    ),
    PreciosRankeados AS (
        -- For those articles, find their latest price
        SELECT
            pv.ARTICULO,
            pv.PRECIO,
            pv.PRECIOSVENTA,
            pv.TIPOIMPUESTO1,
            pv.TIPOIMPUESTO2,
            pv.TIPOIMPUESTO3,
            pv.TIPOIMPUESTO4,
            pv.FECHAINICIO,
            ROW_NUMBER() OVER(PARTITION BY pv.ARTICULO ORDER BY pv.FECHAINICIO DESC) as rn
        FROM NPV.dbo.NPVFDPreciosVenta pv
        INNER JOIN ArticulosEncontrados ae ON pv.ARTICULO = ae.ARTICULO
        WHERE (pv.FECHAINICIO IS NULL OR pv.FECHAINICIO <= GETDATE())
    )
    -- Join everything together to get the final details
    SELECT
        a.ARTICULO,
        a.DESCRIPCION,
        pr.PRECIOSVENTA AS PRECIO,
        promo.PRECIOUNITARIO *
            (1 + COALESCE(i1.PORCENTAJE, 0) / 100.0) *
            (1 + COALESCE(i2.PORCENTAJE, 0) / 100.0) *
            (1 + COALESCE(i3.PORCENTAJE, 0) / 100.0) *
            (1 + COALESCE(i4.PORCENTAJE, 0) / 100.0) AS PRECIO_ESPECIAL,
        CASE WHEN promo.PRECIOUNITARIO IS NULL THEN 0 ELSE 1 END AS TIENE_PRECIO_ESPECIAL,
        promo.PROMO_FECHAINICIO,
        promo.PROMO_FECHAFINAL,
        (SELECT TOP 1 e.EQUIVALENTE FROM NPV.dbo.NPVFDArticulosEquivalentes e WHERE e.ARTICULO = a.ARTICULO) as UPC,
        pr.FECHAINICIO
    FROM NPV.dbo.NPVFDArticulos a
    JOIN PreciosRankeados pr ON a.ARTICULO = pr.ARTICULO
    OUTER APPLY (
        SELECT TOP 1
            pl.PRECIOUNITARIO,
            pe.VIGENCIAINICIO AS PROMO_FECHAINICIO,
            pe.VIGENCIAFINAL AS PROMO_FECHAFINAL
        FROM NPV.dbo.NPVFDPromocionLineas pl
        INNER JOIN NPV.dbo.NPVFDPromocionEncabezado pe ON pe.CLAVE = pl.CLAVE
        WHERE pl.CLAVEARTICULO = a.ARTICULO
          AND pl.CANTIDAD = 1
          AND pe.VIGENCIAFINAL >= GETDATE()
          AND pe.CLASE = 'Z002'
        ORDER BY pl.CLAVE DESC
    ) promo
    LEFT JOIN NPV.dbo.NPVFDImpuestos i1 ON i1.IMPUESTO = pr.TIPOIMPUESTO1
    LEFT JOIN NPV.dbo.NPVFDImpuestos i2 ON i2.IMPUESTO = pr.TIPOIMPUESTO2
    LEFT JOIN NPV.dbo.NPVFDImpuestos i3 ON i3.IMPUESTO = pr.TIPOIMPUESTO3
    LEFT JOIN NPV.dbo.NPVFDImpuestos i4 ON i4.IMPUESTO = pr.TIPOIMPUESTO4
    WHERE pr.rn = 1
    ORDER BY a.DESCRIPCION;
    """

    try:
        conn = _get_connection(cfg)
        cursor = conn.cursor()
        
        rows = cursor.execute(sql, term, like_term, term_digits).fetchall()

        for row in rows:
            fecha_vigencia = row.FECHAINICIO.strftime("%d/%m/%Y") if row.FECHAINICIO else datetime.date.today().strftime("%d/%m/%Y")
            precio_venta = _format_price(row.PRECIO)
            precio_especial = _format_price(row.PRECIO_ESPECIAL) or precio_venta
            
            item = {
                "ARTICULO": str(row.ARTICULO).strip(),
                "DESCRIPCION": str(row.DESCRIPCION or "").strip(),
                "PRECIO": precio_venta,
                "PRECIO_ESPECIAL": precio_especial,
                "TIENE_PRECIO_ESPECIAL": bool(row.TIENE_PRECIO_ESPECIAL),
                "UPC": _normalize_upc_digits(row.UPC or ""),
                "VIGENCIA": f"Valido a partir de: {fecha_vigencia} Aplican TyC",
                "PROMO_FECHAINICIO": row.PROMO_FECHAINICIO,
                "PROMO_FECHAFINAL": row.PROMO_FECHAFINAL,
            }
            results.append(item)

    except Exception as e:
        try:
            conn_info = _sanitize_connection_string(_build_connection_string(cfg))
        except Exception:
            conn_info = "No se pudo construir connection string."
        write_log(
            "ERROR search_items",
            e,
            extra=f"term={term}\nconnection={conn_info}",
        )
        print(f"[ERROR search_items] {e}")
        return []
    finally:
        if conn:
            conn.close()
    
    return results


# ==============================
# Impresoras (detección)
# ==============================
def detect_default_printer() -> str:
    try:
        return win32print.GetDefaultPrinter()
    except Exception:
        return ""


def list_printers() -> List[str]:
    flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
    try:
        printers = win32print.EnumPrinters(flags, None, 2)
        names = [p["pPrinterName"] for p in printers]
        # De-duplicar preservando orden
        seen = set()
        out = []
        for n in names:
            if n not in seen:
                out.append(n)
                seen.add(n)
        return out
    except Exception:
        return []


# ==============================
# Impresión con GDI de Windows
# ==============================

def _get_selected_printer_name(cfg: dict) -> str:
    modo = (cfg.get("impresora", {}).get("modo") or "default").lower()
    if modo == "default":
        return detect_default_printer()
    return cfg.get("impresora", {}).get("printer_name") or ""


def _item_for_individual_label(item_dict: Dict[str, Any]) -> Dict[str, Any]:
    item = dict(item_dict or {})
    item["PRECIO_ESPECIAL"] = ""
    item["TIENE_PRECIO_ESPECIAL"] = False
    item["PROMO_TERMINOS"] = ""
    item["PROMO_TERMINOS2"] = ""
    return item


def draw_promo_label_gdi(hDC, item_dict: Dict[str, Any], printable_width: int) -> bool:
    if not item_dict.get("PRECIO_ESPECIAL") or Image is None or ImageWin is None:
        return False

    bg_path = get_promo_background_path()
    if not os.path.exists(bg_path):
        print(f"[DEBUG] No se encontro la imagen de fondo en: {bg_path}")
        return False

    try:
        with Image.open(bg_path) as bg:
            bg = bg.convert("RGB")
            img_w, img_h = bg.size
            printable_height = int(printable_width * img_h / img_w)
            bg_resized = bg.resize((printable_width, printable_height), Image.LANCZOS)
            dib = ImageWin.Dib(bg_resized)
            dib.draw(hDC.GetHandleOutput(), (0, 0, printable_width, printable_height))
    except Exception as e:
        print(f"[ERROR] Fallo al cargar imagen de fondo: {e}")
        return False

    # IMPORTANTE: Usar PROMO_TEMPLATE_SIZE para el escalado de coordenadas de texto
    design_w, design_h = PROMO_TEMPLATE_SIZE
    sx = printable_width / design_w
    sy = printable_height / design_h

    def rect(left, top, right, bottom):
        return (int(left * sx), int(top * sy), int(right * sx), int(bottom * sy))

    def font_height(px):
        return max(12, int(px * sy))

    precio = item_dict.get("PRECIO", "$0.00")
    precio_especial = item_dict.get("PRECIO_ESPECIAL", "")
    
    p_antes_int, p_antes_dec = _split_price(precio)
    p_ahora_int, p_ahora_dec = _split_price(precio_especial)
    
    ahorro_val = max(0, (_price_to_float(precio) or 0) - (_price_to_float(precio_especial) or 0))
    p_ahorra_int, p_ahorra_dec = _split_price(ahorro_val)
    promo_terminos = item_dict.get("PROMO_TERMINOS", "")
    promo_terminos2 = item_dict.get("PROMO_TERMINOS2", "")

    hDC.SetTextColor(0x111111) # Color oscuro para el texto
    hDC.SetBkColor(0xFFFFFF)
    hDC.SetBkMode(win32con.OPAQUE)

    font_savings = win32ui.CreateFont({"name": "Arial", "height": font_height(100), "weight": win32con.FW_BOLD})
    font_savings_dec = win32ui.CreateFont({"name": "Arial", "height": font_height(100), "weight": win32con.FW_BOLD})
    hDC.SelectObject(font_savings)
    hDC.DrawText(p_ahorra_int, rect(940, 70, 1245, 200), win32con.DT_RIGHT | win32con.DT_SINGLELINE | win32con.DT_VCENTER)
    hDC.SelectObject(font_savings_dec)
    hDC.DrawText(p_ahorra_dec, rect(1310, 70, 1460, 200), win32con.DT_LEFT | win32con.DT_SINGLELINE | win32con.DT_VCENTER)

    font_now = win32ui.CreateFont({"name": "Arial", "height": font_height(166), "weight": win32con.FW_BOLD})
    font_now_dec = win32ui.CreateFont({"name": "Arial", "height": font_height(166), "weight": win32con.FW_BOLD})
    hDC.SelectObject(font_now)
    hDC.DrawText(p_ahora_int, rect(300, 285, 1065, 480), win32con.DT_RIGHT | win32con.DT_SINGLELINE | win32con.DT_VCENTER)
    hDC.SelectObject(font_now_dec)
    hDC.DrawText(p_ahora_dec, rect(1210, 285, 1405, 480), win32con.DT_LEFT | win32con.DT_SINGLELINE | win32con.DT_VCENTER)

    font_before = win32ui.CreateFont({"name": "Arial", "height": font_height(66), "weight": win32con.FW_BOLD})
    font_before_dec = win32ui.CreateFont({"name": "Arial", "height": font_height(66), "weight": win32con.FW_BOLD})
    hDC.SelectObject(font_before)
    hDC.DrawText(p_antes_int, rect(1020, 720, 1230, 810), win32con.DT_RIGHT | win32con.DT_SINGLELINE | win32con.DT_VCENTER)
    hDC.SelectObject(font_before_dec)
    hDC.DrawText(p_antes_dec, rect(1300, 720, 1390, 810), win32con.DT_LEFT | win32con.DT_SINGLELINE | win32con.DT_VCENTER)

    hDC.SetBkMode(win32con.TRANSPARENT)

    desc = (item_dict.get("DESCRIPCION") or "").upper()
    if desc:
        font_desc = win32ui.CreateFont({"name": "Arial", "height": font_height(40), "weight": win32con.FW_BOLD})
        hDC.SelectObject(font_desc)
        hDC.DrawText(desc, rect(80, 535, 1505, 590), win32con.DT_CENTER | win32con.DT_SINGLELINE | win32con.DT_VCENTER)

    if promo_terminos:
        font_terms = win32ui.CreateFont({"name": "Arial", "height": font_height(28), "weight": win32con.FW_BOLD})
        hDC.SelectObject(font_terms)
        hDC.DrawText(promo_terminos, rect(95, 865, 1490, 900), win32con.DT_LEFT | win32con.DT_SINGLELINE | win32con.DT_VCENTER)
        if promo_terminos2:
            hDC.DrawText(promo_terminos2, rect(95, 900, 1490, 960), win32con.DT_LEFT | win32con.DT_WORDBREAK)

    return True


def print_label_gdi(item_dict: Dict[str, Any], config: dict, copies: int = 1, show_errors: bool = True):
    """
    Imprime la etiqueta final usando GDI, basado en el área de impresión real.
    """
    printer_name = _get_selected_printer_name(config)
    if not printer_name:
        raise RuntimeError("No se pudo determinar la impresora a usar.")

    try:
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        
        # Obtener el área de impresión real en píxeles
        printable_width = hDC.GetDeviceCaps(win32con.HORZRES)
        hDC.SetMapMode(win32con.MM_TEXT)
        hDC.SetBkMode(win32con.TRANSPARENT)

        for i in range(copies):
            hDC.StartDoc(f"Etiqueta NPV ({i+1}/{copies})")
            hDC.StartPage()

            if draw_promo_label_gdi(hDC, item_dict, printable_width):
                hDC.EndPage()
                hDC.EndDoc()
                continue

            y_pos = 20 # Margen superior

            # --- 2. Dibujar Descripción ---
            try:
                font_desc = win32ui.CreateFont({"name": "Arial", "height": 46, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_desc)
                desc = (item_dict.get("DESCRIPCION") or "").upper()
                print(f"Dibujando descripción: {desc}") # <-- DEBUG
                rect = (10, y_pos, printable_width - 10, y_pos + 200)
                hDC.DrawText(desc, rect, win32con.DT_CENTER | win32con.DT_WORDBREAK)
                y_pos += 70 # Espacio ajustado para la descripción
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar la descripción: {e}")

            # --- 3. Dibujar Precio ---
            try:
                font_price_height = 77
                font_price = win32ui.CreateFont({"name": "Arial", "height": font_price_height, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_price)
                precio = _format_currency(item_dict.get("PRECIO", "$0.00"))
                precio_especial = _format_currency(item_dict.get("PRECIO_ESPECIAL", "")) if item_dict.get("PRECIO_ESPECIAL") else ""
                if precio_especial:
                    precio_base_num = _price_to_float(precio) or 0
                    precio_especial_num = _price_to_float(precio_especial) or 0
                    ahorro = max(0, precio_base_num - precio_especial_num)
                    font_label = win32ui.CreateFont({"name": "Arial", "height": 24, "weight": win32con.FW_BOLD})
                    font_price = win32ui.CreateFont({"name": "Arial", "height": 54, "weight": win32con.FW_BOLD})
                    font_save = win32ui.CreateFont({"name": "Arial", "height": 22, "weight": win32con.FW_BOLD})
                    hDC.SelectObject(font_label)
                    hDC.DrawText(f"ANTES: {precio}", (10, y_pos, printable_width - 10, y_pos + 32), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 35
                    hDC.SelectObject(font_price)
                    hDC.DrawText(f"AHORA: {precio_especial}", (10, y_pos, printable_width - 10, y_pos + 66), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 60
                    hDC.SelectObject(font_save)
                    hDC.DrawText(f"AHORRAS: {_format_price(ahorro)}", (10, y_pos, printable_width - 10, y_pos + 30), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 30 - font_price_height
                else:
                    rect = (10, y_pos, printable_width - 10, y_pos + 100)
                    hDC.DrawText(precio, rect, win32con.DT_CENTER | win32con.DT_SINGLELINE)
                
                # Se usa la configuración de espaciado. El avance es altura de fuente + espacio.
                espaciados = config.get("espaciados", {})
                espacio_abajo = espaciados.get("espacio_abajo_precio", 0)
                y_pos += font_price_height + espacio_abajo
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar el precio: {e}")

            # --- 4. Dibujar Pie de página ---
            try:
                font_footer = win32ui.CreateFont({"name": "Arial", "height": 24, "weight": win32con.FW_NORMAL})
                hDC.SelectObject(font_footer)
                
                footer_items = [
                    item_dict.get("VIGENCIA", ""),
                    f"{item_dict.get('ARTICULO', '')}   {item_dict.get('UPC', '')}"
                ]

                for text in footer_items:
                    if not text.strip(): continue
                    rect = (10, y_pos, printable_width - 10, y_pos + 30)
                    hDC.DrawText(text, rect, win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 25
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar el pie de página: {e}")

            hDC.EndPage()
            hDC.EndDoc()
        hDC.DeleteDC()
        print("[INFO] Se envió el trabajo de impresión final.")
    except Exception as e:
        print(f"[ERROR] Falló la impresión GDI: {e}")
        if show_errors:
            messagebox.showerror(APP_TITLE, f"Error de GDI: {e}")
        else:
            write_log("ERROR print_label_gdi", e)


def print_label_gdi_small(item_dict: Dict[str, Any], config: dict, copies: int = 1, show_errors: bool = True):
    """
    Imprime la etiqueta en formato angosto.
    """
    item_dict = _item_for_individual_label(item_dict)
    auto_individual_print = bool(item_dict.get("_AUTO_INDIVIDUAL_PRINT"))
    printer_name = _get_selected_printer_name(config)
    if not printer_name:
        raise RuntimeError("No se pudo determinar la impresora a usar.")

    try:
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printer_name)
        
        printable_width = hDC.GetDeviceCaps(win32con.HORZRES)
        hDC.SetMapMode(win32con.MM_TEXT)
        hDC.SetBkMode(win32con.TRANSPARENT)

        for i in range(copies):
            hDC.StartDoc(f"Etiqueta NPV Individual ({i+1}/{copies})")
            hDC.StartPage()

            y_pos = 20 # Margen superior (como el original)
            horizontal_margin = 40 # Margen horizontal más grande para hacerlo más angosto

            # --- 2. Dibujar Descripción ---
            try:
                desc_font_height = 34 if auto_individual_print else 40
                desc_advance = 112 if auto_individual_print else 70
                font_desc = win32ui.CreateFont({"name": "Arial", "height": desc_font_height, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_desc)
                desc = (item_dict.get("DESCRIPCION") or "").upper()
                rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 200)
                hDC.DrawText(desc, rect, win32con.DT_CENTER | win32con.DT_WORDBREAK)
                y_pos += desc_advance
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar la descripción (individual): {e}")

            # --- 3. Dibujar Precio ---
            try:
                font_price_height = 77
                font_price = win32ui.CreateFont({"name": "Arial", "height": font_price_height, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_price)
                precio = _format_currency(item_dict.get("PRECIO", "$0.00"))
                precio_especial = item_dict.get("PRECIO_ESPECIAL", "")
                if precio_especial:
                    precio_base_num = _price_to_float(precio) or 0
                    precio_especial_num = _price_to_float(precio_especial) or 0
                    ahorro = max(0, precio_base_num - precio_especial_num)
                    font_label = win32ui.CreateFont({"name": "Arial", "height": 22, "weight": win32con.FW_BOLD})
                    font_price = win32ui.CreateFont({"name": "Arial", "height": 48, "weight": win32con.FW_BOLD})
                    font_save = win32ui.CreateFont({"name": "Arial", "height": 20, "weight": win32con.FW_BOLD})
                    hDC.SelectObject(font_label)
                    hDC.DrawText(f"ANTES: {precio}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 30), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 28
                    hDC.SelectObject(font_price)
                    hDC.DrawText(f"AHORA: {precio_especial}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 60), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 56
                    hDC.SelectObject(font_save)
                    hDC.DrawText(f"AHORRAS: {_format_price(ahorro)}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 28), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 28 - font_price_height
                else:
                    rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 100)
                    hDC.DrawText(precio, rect, win32con.DT_CENTER | win32con.DT_SINGLELINE)
                
                espaciados = config.get("espaciados", {})
                espacio_abajo = espaciados.get("espacio_abajo_precio", 0)
                y_pos += font_price_height + espacio_abajo
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar el precio (individual): {e}")

            # --- 4. Dibujar Pie de página ---
            try:
                font_footer = win32ui.CreateFont({"name": "Arial", "height": 21, "weight": win32con.FW_NORMAL})
                hDC.SelectObject(font_footer)
                
                footer_items = [
                    item_dict.get("VIGENCIA", ""),
                    f"{item_dict.get('ARTICULO', '')}   {item_dict.get('UPC', '')}"
                ]

                for text in footer_items:
                    if not text.strip(): continue
                    rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 30)
                    hDC.DrawText(text, rect, win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 25
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar el pie de página (individual): {e}")

            hDC.EndPage()
            hDC.EndDoc()
        hDC.DeleteDC()
        print("[INFO] Se envió el trabajo de impresión (individual) final.")
    except Exception as e:
        print(f"[ERROR] Falló la impresión GDI (individual): {e}")
        if show_errors:
            messagebox.showerror(APP_TITLE, f"Error de GDI (individual): {e}")
        else:
            write_log("ERROR print_label_gdi_small", e)


# ==============================
# Construcción de etiqueta (Helper)
# ==============================
def build_label_text(item_dict: Dict[str, Any], config: dict) -> Dict[str, Any]:
    """
    Normaliza y devuelve campos para vista previa / impresión.
    """
    if not item_dict:
        return {}

    descripcion = (item_dict.get("DESCRIPCION") or "").strip()
    precio      = _format_price(item_dict.get("PRECIO", "").replace("$", "") if isinstance(item_dict.get("PRECIO"), str) else item_dict.get("PRECIO"))
    precio_especial = _format_price(item_dict.get("PRECIO_ESPECIAL", "").replace("$", "") if isinstance(item_dict.get("PRECIO_ESPECIAL"), str) else item_dict.get("PRECIO_ESPECIAL"))
    articulo    = (item_dict.get("ARTICULO") or "").strip()
    upc         = _normalize_upc_digits(item_dict.get("UPC"))

    if not precio:
        # Precio inválido
        precio = ""
    if not precio_especial:
        precio_especial = precio

    # Vigencia: usa la que viene del item si existe, si no, la de por defecto.
    vigencia = item_dict.get("VIGENCIA")
    if not vigencia:
        vigencia = resolve_vigencia_template(config.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"]))

    promo_inicio = item_dict.get("PROMO_FECHAINICIO")
    promo_final = item_dict.get("PROMO_FECHAFINAL")
    promo_inicio_txt = _format_date(promo_inicio)
    promo_final_txt = _format_date(promo_final)
    promo_terminos = ""
    if promo_inicio_txt and promo_final_txt:
        promo_terminos = f"Terminos y Condiciones: valido del {promo_inicio_txt} al {promo_final_txt}"
    promo_terminos2 = "No acumulable. Sujeto a cambios sin aviso."

    return {
        "DESCRIPCION": descripcion,
        "PRECIO": precio,
        "PRECIO_ESPECIAL": precio_especial,
        "TIENE_PRECIO_ESPECIAL": bool(item_dict.get("TIENE_PRECIO_ESPECIAL")),
        "PROMO_FECHAINICIO": promo_inicio,
        "PROMO_FECHAFINAL": promo_final,
        "PROMO_TERMINOS": promo_terminos,
        "PROMO_TERMINOS2": promo_terminos2,
        "ARTICULO": articulo,
        "UPC": upc,
        "VIGENCIA": vigencia
    }


# ==============================
# UI (Tkinter)
# ==============================

class ResultSelectionWindow(tk.Toplevel):
    def __init__(self, parent, results, callback):
        super().__init__(parent)
        self.title("Seleccionar Artículo")
        self.geometry("800x450")
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#F0F0F0")

        self.callback = callback
        self.results = results
        
        # --- Style ---
        style = ttk.Style(self)
        # Inherit colors from parent if possible, otherwise define them
        try:
            primary_color = parent.PRIMARY_COLOR
            primary_active_color = parent.PRIMARY_ACTIVE_COLOR
            light_text_color = parent.LIGHT_TEXT_COLOR
        except AttributeError:
            primary_color = "#0078D7"
            primary_active_color = "#005A9E"
            light_text_color = "#FFFFFF"

        style.configure("Treeview", rowheight=25, font=("Segoe UI", 10))
        style.configure("Treeview.Heading", font=("Segoe UI", 11, "bold"))
        style.map("Treeview", background=[("selected", primary_color)])

        # Standard Button
        style.configure("TButton", font=("Segoe UI", 10, "bold"), padding=(10, 5), relief="flat")
        style.map("TButton", background=[('active', '#E0E0E0')])

        # Accent Button
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"), padding=(12, 6), relief="flat")
        style.map("Accent.TButton",
                  foreground=[('!disabled', light_text_color)],
                  background=[('!disabled', primary_color),
                              ('active', primary_active_color)])

        # --- Widgets ---
        self.tree = ttk.Treeview(self, columns=("articulo", "descripcion", "upc"), show="headings")
        self.tree.heading("articulo", text="Artículo")
        self.tree.heading("descripcion", text="Descripción")
        self.tree.heading("upc", text="UPC")
        self.tree.column("articulo", width=120, anchor="w")
        self.tree.column("descripcion", width=480, anchor="w")
        self.tree.column("upc", width=150, anchor="w")
        self.tree.pack(expand=True, fill="both", padx=15, pady=15)

        for item in self.results:
            self.tree.insert("", "end", values=(item["ARTICULO"], item["DESCRIPCION"], item["UPC"]))

        self.tree.bind("<Double-1>", self.on_select)

        # --- Buttons ---
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=15, pady=(0, 15))
        
        select_button = ttk.Button(btn_frame, text="Seleccionar", command=self.on_select)
        select_button.pack(side="right")
        
        cancel_button = ttk.Button(btn_frame, text="Cancelar", command=self.on_close)
        cancel_button.pack(side="right", padx=(0, 10))


        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.wait_window()

    def on_select(self, event=None):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        selected_tree_item = selected_items[0]
        selected_values = self.tree.item(selected_tree_item, "values")
        
        selected_articulo = selected_values[0]
        for result_item in self.results:
            if result_item["ARTICULO"] == selected_articulo:
                self.callback(result_item)
                self.destroy()
                return

    def on_close(self):
        self.callback(None) # Indicate no selection was made
        self.destroy()


class LabelPrintPreviewWindow(tk.Toplevel):
    def __init__(self, parent, item: Dict[str, Any], individual: bool = False):
        super().__init__(parent)
        self.title("Vista previa de impresion")
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#F0F0F0")
        self.resizable(False, False)

        self.item = _item_for_individual_label(item) if individual else item
        self.individual = individual

        label_width = 500 if individual else 660
        if self.item.get("PRECIO_ESPECIAL"):
            label_height = int(label_width * PROMO_TEMPLATE_SIZE[1] / PROMO_TEMPLATE_SIZE[0])
        else:
            label_height = 285
        window_width = label_width + 60
        window_height = label_height + 95
        self.geometry(f"{window_width}x{window_height}")

        title = "Etiqueta individual (angosta) - ESC para salir" if individual else "Etiqueta 80mm - ESC para salir"
        ttk.Label(self, text=title, font=("Segoe UI", 12, "bold")).pack(anchor="w", padx=20, pady=(16, 8))

        self.canvas = tk.Canvas(
            self,
            width=label_width,
            height=label_height,
            bg="#D9D9D9",
            highlightthickness=0,
        )
        self.canvas.pack(padx=20, pady=(0, 14))
        self._draw_label(label_width, label_height)

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=20, pady=(0, 16))

        ttk.Button(btn_frame, text="Cerrar", command=self.destroy).pack(side="right")

        self.bind("<Escape>", lambda event: self.destroy())
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        self.focus_force()
        self.lift()
        self.attributes("-topmost", True)
        self.after(100, lambda: self.attributes("-topmost", False))
        self.wait_window()

    def _fit_text_size(self, text: str, max_chars: int, base_size: int, min_size: int) -> int:
        text_len = len(text or "")
        if text_len <= max_chars:
            return base_size
        return max(min_size, base_size - ((text_len - max_chars) // 6 + 1) * 2)

    def _draw_label(self, label_width: int, label_height: int):
        self.canvas.delete("all")
        self._promo_photo = None

        if self.item.get("PRECIO_ESPECIAL") and Image is not None and ImageTk is not None and os.path.exists(get_promo_background_path()):
            bg = Image.open(get_promo_background_path()).convert("RGB")
            bg = bg.resize((label_width, label_height), Image.LANCZOS)
            self._promo_photo = ImageTk.PhotoImage(bg)
            self.canvas.create_image(0, 0, image=self._promo_photo, anchor="nw")

            template_width, template_height = PROMO_TEMPLATE_SIZE
            sx = label_width / template_width
            sy = label_height / template_height

            def cx(value):
                return value * sx

            def cy(value):
                return value * sy

            price = self.item.get("PRECIO") or "$0.00"
            special_price = self.item.get("PRECIO_ESPECIAL") or ""
            ahorro_val = max(0, (_price_to_float(price) or 0) - (_price_to_float(special_price) or 0))
            
            p_ahorra_int, p_ahorra_dec = _split_price(ahorro_val)
            p_ahora_int, p_ahora_dec = _split_price(special_price)
            p_antes_int, p_antes_dec = _split_price(price)

            desc = (self.item.get("DESCRIPCION") or "").upper()
            promo_terminos = self.item.get("PROMO_TERMINOS") or ""
            promo_terminos2 = self.item.get("PROMO_TERMINOS2") or ""

            # AHORRA
            self.canvas.create_rectangle(cx(940), cy(70), cx(1245), cy(200), fill="#FFFFFF", outline="")
            self.canvas.create_rectangle(cx(1310), cy(70), cx(1460), cy(200), fill="#FFFFFF", outline="")
            self.canvas.create_text(
                cx(1245), cy(135), text=p_ahorra_int, anchor="e",
                font=("Arial", int(100 * sy), "bold"), fill="#111111"
            )
            self.canvas.create_text(
                cx(1310), cy(135), text=p_ahorra_dec, anchor="w",
                font=("Arial", int(100 * sy), "bold"), fill="#111111"
            )

            # PRECIO NUEVO
            self.canvas.create_rectangle(cx(300), cy(285), cx(1065), cy(480), fill="#FFFFFF", outline="")
            self.canvas.create_rectangle(cx(1210), cy(285), cx(1405), cy(480), fill="#FFFFFF", outline="")
            self.canvas.create_text(
                cx(1065), cy(382), text=p_ahora_int, anchor="e",
                font=("Arial", int(166 * sy), "bold"), fill="#111111"
            )
            self.canvas.create_text(
                cx(1210), cy(382), text=p_ahora_dec, anchor="w",
                font=("Arial", int(166 * sy), "bold"), fill="#111111"
            )

            # PRECIO ANTERIOR
            self.canvas.create_rectangle(cx(1020), cy(720), cx(1230), cy(810), fill="#FFFFFF", outline="")
            self.canvas.create_rectangle(cx(1300), cy(720), cx(1390), cy(810), fill="#FFFFFF", outline="")
            self.canvas.create_text(
                cx(1230), cy(765), text=p_antes_int, anchor="e",
                font=("Arial", int(66 * sy), "bold"), fill="#111111"
            )
            self.canvas.create_text(
                cx(1300), cy(765), text=p_antes_dec, anchor="w",
                font=("Arial", int(66 * sy), "bold"), fill="#111111"
            )

            if desc:
                self.canvas.create_text(
                    cx(792),
                    cy(563),
                    text=desc,
                    width=cx(1425),
                    anchor="center",
                    justify="center",
                    font=("Arial", max(10, int(40 * sy)), "bold"),
                    fill="#111111",
                )

            if promo_terminos:
                self.canvas.create_text(
                    cx(95),
                    cy(865),
                    text=promo_terminos,
                    width=cx(1395),
                    anchor="nw",
                    justify="left",
                    font=("Arial", max(9, int(28 * sy)), "bold"),
                    fill="#111111",
                )
                if promo_terminos2:
                    self.canvas.create_text(
                        cx(95),
                        cy(900),
                        text=promo_terminos2,
                        width=cx(1395),
                        anchor="nw",
                        justify="left",
                        font=("Arial", max(9, int(28 * sy)), "bold"),
                        fill="#111111",
                    )

            return

        page_pad = 12
        x1 = page_pad
        y1 = page_pad
        x2 = label_width - page_pad
        y2 = label_height - page_pad

        self.canvas.create_rectangle(x1, y1, x2, y2, fill="#FFFFFF", outline="#222222", width=2)

        margin = 52 if self.individual else 18
        text_left = x1 + margin
        text_right = x2 - margin
        center_x = label_width / 2

        desc = (self.item.get("DESCRIPCION") or "").upper()
        price = _format_currency(self.item.get("PRECIO") or "$0.00")
        special_price = self.item.get("PRECIO_ESPECIAL") or ""
        vigencia = self.item.get("VIGENCIA") or ""
        codes = f"{self.item.get('ARTICULO', '')}   {self.item.get('UPC', '')}".strip()

        desc_size = self._fit_text_size(desc, 49, 26 if self.individual else 28, 15)
        self.canvas.create_text(
            center_x,
            y1 + 24,
            text=desc,
            width=text_right - text_left,
            anchor="n",
            justify="center",
            font=("Arial", desc_size, "bold"),
            fill="#111111",
        )

        if special_price:
            price_num = _price_to_float(price) or 0
            special_num = _price_to_float(special_price) or 0
            savings = _format_price(max(0, price_num - special_num))
            self.canvas.create_text(
                center_x,
                y1 + 96,
                text=f"ANTES: {price}",
                width=text_right - text_left,
                anchor="n",
                justify="center",
                font=("Arial", 18, "bold"),
                fill="#111111",
            )
            self.canvas.create_text(
                center_x,
                y1 + 128,
                text=f"AHORA: {special_price}",
                width=text_right - text_left,
                anchor="n",
                justify="center",
                font=("Arial", 42, "bold"),
                fill="#111111",
            )
            self.canvas.create_text(
                center_x,
                y1 + 190,
                text=f"AHORRAS: {savings}",
                width=text_right - text_left,
                anchor="n",
                justify="center",
                font=("Arial", 16, "bold"),
                fill="#111111",
            )
            footer_y = y1 + 224
        else:
            price_size = self._fit_text_size(price, 8, 58, 42)
            self.canvas.create_text(
                center_x,
                y1 + 102,
                text=price,
                width=text_right - text_left,
                anchor="n",
                justify="center",
                font=("Arial", price_size, "bold"),
                fill="#111111",
            )
            footer_y = y1 + 205

        footer_size = 13 if self.individual else 14
        if vigencia:
            self.canvas.create_text(
                center_x,
                footer_y,
                text=vigencia,
                width=text_right - text_left,
                anchor="n",
                justify="center",
                font=("Arial", footer_size),
                fill="#111111",
            )
            footer_y += 26

        if codes:
            self.canvas.create_text(
                center_x,
                footer_y,
                text=codes,
                width=text_right - text_left,
                anchor="n",
                justify="center",
                font=("Arial", footer_size),
                fill="#111111",
            )


class App(ThemedTk):
    def __init__(self):
        super().__init__(theme="arc")
        self.title(APP_TITLE)
        screen_height = self.winfo_screenheight()
        window_height = min(720, max(520, screen_height - 120))
        self.geometry(f"900x{window_height}")
        self.minsize(720, 480)
        self.resizable(True, True)

        self.config_data = load_config()
        self.current_built_item = {}
        self.configure_styles()
        self.create_widgets()
        self.populate_printers()
        self.refresh_default_printer_note()

    def configure_styles(self):
        style = ttk.Style(self)
        self.font_normal = font.Font(family="Segoe UI", size=10)
        self.font_bold = font.Font(family="Segoe UI", size=10, weight="bold")
        self.font_h1 = font.Font(family="Segoe UI", size=12, weight="bold")

        # Colors
        self.BG_COLOR = "#F0F0F0"
        self.PRIMARY_COLOR = "#0078D7"
        self.PRIMARY_ACTIVE_COLOR = "#005A9E" # Darker blue for hover
        self.TEXT_COLOR = "#333333"
        self.LIGHT_TEXT_COLOR = "#FFFFFF"

        self.configure(bg=self.BG_COLOR)

        style.configure("TLabel", font=self.font_normal, background=self.BG_COLOR, foreground=self.TEXT_COLOR)
        
        # Standard Button
        style.configure("TButton", font=self.font_bold, padding=(10, 5), relief="flat")
        style.map("TButton",
            background=[('active', '#E0E0E0')],
            foreground=[('!disabled', self.TEXT_COLOR)])

        # Accent Button (for primary actions)
        style.configure("Accent.TButton", font=self.font_bold, padding=(12, 6), relief="flat")
        style.map("Accent.TButton",
                  foreground=[('!disabled', self.LIGHT_TEXT_COLOR)],
                  background=[('!disabled', self.PRIMARY_COLOR),
                              ('active', self.PRIMARY_ACTIVE_COLOR)])
        
        style.configure("TLabelframe", font=self.font_h1, background=self.BG_COLOR, borderwidth=1, relief="groove")
        style.configure("TLabelframe.Label", font=self.font_h1, background=self.BG_COLOR, foreground=self.PRIMARY_COLOR)
        
        style.configure("TEntry", font=self.font_normal, padding=5)
        style.map("TEntry",
            bordercolor=[('focus', self.PRIMARY_COLOR)],
            borderwidth=[('focus', 1)])

        style.configure("TCombobox", font=self.font_normal, padding=5)
        style.configure("TRadiobutton", font=self.font_normal, background=self.BG_COLOR)
        style.configure("TCheckbutton", font=self.font_normal, background=self.BG_COLOR)
        style.configure("TSpinbox", font=self.font_normal, padding=5)

    def create_widgets(self):
        scroll_container = ttk.Frame(self)
        scroll_container.pack(expand=True, fill="both")
        scroll_container.columnconfigure(0, weight=1)
        scroll_container.rowconfigure(0, weight=1)

        self.main_canvas = tk.Canvas(scroll_container, bg=self.BG_COLOR, highlightthickness=0)
        self.main_scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=self.main_canvas.yview)
        self.main_canvas.configure(yscrollcommand=self.main_scrollbar.set)
        self.main_canvas.grid(row=0, column=0, sticky="nsew")
        self.main_scrollbar.grid(row=0, column=1, sticky="ns")

        main_frame = ttk.Frame(self.main_canvas, padding=20)
        self.main_canvas_window = self.main_canvas.create_window((0, 0), window=main_frame, anchor="nw")
        main_frame.bind("<Configure>", self._update_main_scrollregion)
        self.main_canvas.bind("<Configure>", self._resize_main_canvas_window)
        self.main_canvas.bind_all("<MouseWheel>", self._on_main_mousewheel)

        # --- Búsqueda ---
        search_frame = ttk.LabelFrame(main_frame, text="Búsqueda (Artículo, UPC o Descripción)")
        search_frame.pack(fill="x", pady=(0, 20))
        
        self.entry_search = ttk.Entry(search_frame)
        self.entry_search.pack(fill="x", expand=True, padx=10, pady=10)
        self.entry_search.bind("<Return>", self.on_unified_search)
        self.entry_search.focus()

        # --- Vista Previa ---
        preview_frame = ttk.LabelFrame(main_frame, text="Vista Previa (Editable)")
        preview_frame.pack(fill="both", expand=True, pady=(0, 20))
        preview_frame.columnconfigure(1, weight=1)
        preview_frame.columnconfigure(3, weight=1)

        pad_x = 10
        pad_y = 8

        r = 0
        ttk.Label(preview_frame, text="Producto/Descripción:").grid(row=r, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.txt_producto = tk.Text(preview_frame, height=3, width=60, font=self.font_normal, relief="solid", borderwidth=1)
        self.txt_producto.grid(row=r, column=1, columnspan=3, sticky="ew", padx=pad_x, pady=pad_y)
        self.txt_producto.configure(state="disabled")

        r += 1
        ttk.Label(preview_frame, text="Precio ($xx.xx):").grid(row=r, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_precio = ttk.Entry(preview_frame, width=25)
        self.entry_precio.grid(row=r, column=1, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_precio.configure(state="readonly")

        ttk.Label(preview_frame, text="Código interno (ARTICULO):").grid(row=r, column=2, sticky="e", padx=pad_x, pady=pad_y)
        self.entry_codigo = ttk.Entry(preview_frame, width=25)
        self.entry_codigo.grid(row=r, column=3, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_codigo.configure(state="readonly")

        r += 1
        ttk.Label(preview_frame, text="Precio especial ($xx.xx):").grid(row=r, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_precio_especial = ttk.Entry(preview_frame, width=25)
        self.entry_precio_especial.grid(row=r, column=1, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_precio_especial.configure(state="readonly")

        r += 1
        ttk.Label(preview_frame, text="UPC (numérico):").grid(row=r, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_upc = ttk.Entry(preview_frame, width=25)
        self.entry_upc.grid(row=r, column=1, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_upc.configure(state="readonly")

        # Campo interno: se mantiene la vigencia para impresion, pero ya no se captura en pantalla.
        self.entry_vigencia = ttk.Entry(preview_frame, width=25)

        r += 1
        ttk.Label(preview_frame, text="Copias (1-20):").grid(row=r, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.spin_copias = ttk.Spinbox(preview_frame, from_=1, to=20, width=8)
        self.spin_copias.set(1)
        self.spin_copias.grid(row=r, column=1, sticky="w", padx=pad_x, pady=pad_y)

        # --- Impresora ---
        printer_frame = ttk.LabelFrame(main_frame, text="Configuración de Impresión")
        printer_frame.pack(fill="x", pady=(0, 20))
        printer_frame.columnconfigure(1, weight=1)

        self.config_data.setdefault("impresora", {})["modo"] = "default"
        self.printer_mode = tk.StringVar(value="default")
        self.rb_default = ttk.Radiobutton(printer_frame, text="Usar predeterminada", value="default", variable=self.printer_mode, command=self.on_printer_mode_change)
        self.rb_default.grid(row=0, column=0, sticky="w", padx=pad_x, pady=pad_y)

        self.rb_named = ttk.Radiobutton(printer_frame, text="Elegir de la lista", value="named", variable=self.printer_mode, command=self.on_printer_mode_change)
        self.rb_named.grid(row=1, column=0, sticky="w", padx=pad_x, pady=pad_y)

        self.cmb_printers = ttk.Combobox(printer_frame, state="disabled", width=40)
        self.cmb_printers.grid(row=1, column=1, sticky="ew", padx=pad_x, pady=pad_y)

        self.lbl_default = ttk.Label(printer_frame, text="", font=self.font_normal)
        self.lbl_default.grid(row=0, column=1, sticky="w", padx=pad_x, pady=pad_y)

        preferencias = self.config_data.setdefault("preferencias", {})

        self.auto_print_var = tk.BooleanVar(value=bool(preferencias.get("auto_print_on_scan", False)))
        self.chk_auto_print = ttk.Checkbutton(printer_frame, text="Impresión automática al escanear", variable=self.auto_print_var)
        self.chk_auto_print.grid(row=2, column=0, columnspan=2, sticky="w", padx=pad_x, pady=pad_y)

        self.individual_print_var = tk.BooleanVar(value=bool(preferencias.get("individual_print", False)))
        self.chk_individual_print = ttk.Checkbutton(printer_frame, text="Imprimir etiqueta individual (angosta)", variable=self.individual_print_var)
        self.chk_individual_print.grid(row=3, column=0, columnspan=2, sticky="w", padx=pad_x, pady=pad_y)

        # --- Botones de Acción ---
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill="x")

        self.btn_cerrar = ttk.Button(buttons_frame, text="Cerrar", command=self.on_close)
        self.btn_cerrar.pack(side="right", padx=(10, 0))

        self.btn_imprimir = ttk.Button(buttons_frame, text="Imprimir", command=self.on_print_click)
        self.btn_imprimir.pack(side="right", padx=(10, 0))

        self.btn_preview_print = ttk.Button(buttons_frame, text="Vista previa", command=self.on_preview_print_click)
        self.btn_preview_print.pack(side="right", padx=(10, 0))
        
        self.btn_test = ttk.Button(buttons_frame, text="Probar Impresión", command=self.on_test_print)
        self.btn_test.pack(side="right")

        # Cargar vigencia default y bloquear/activar combobox impresoras correctamente
        self.entry_vigencia.delete(0, tk.END)
        self.entry_vigencia.insert(0, resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"])))
        self.on_printer_mode_change()

    # --------- Helpers UI ----------
    def _update_main_scrollregion(self, event=None):
        self.main_canvas.configure(scrollregion=self.main_canvas.bbox("all"))

    def _resize_main_canvas_window(self, event):
        self.main_canvas.itemconfigure(self.main_canvas_window, width=event.width)

    def _on_main_mousewheel(self, event):
        try:
            if event.widget.winfo_toplevel() is not self:
                return
        except Exception:
            return
        self.main_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _clear_placeholder(self, entry: tk.Entry, text: str):
        if entry.get() == text:
            entry.delete(0, tk.END)

    def _restore_placeholder(self, entry: tk.Entry, text: str):
        if not entry.get():
            entry.insert(0, text)

    def _size_to_str(self, size):
        return f"{int(size[0])}x{int(size[1])}"

    def _str_to_size(self, s):
        m = re.match(r"^\s*(\d+)x(\d+)\s*$", s or "")
        if not m:
            return [1, 1]
        return [int(m.group(1)), int(m.group(2))]

    def populate_printers(self):
        printers = list_printers()
        self.cmb_printers["values"] = printers
        sel = self.config_data["impresora"].get("printer_name") or ""
        if sel and sel in printers:
            self.cmb_printers.set(sel)
        elif printers:
            self.cmb_printers.set(printers[0])

    def refresh_default_printer_note(self):
        d = detect_default_printer()
        self.lbl_default.config(text=f"Predeterminada: {d or '—'}")

    def on_printer_mode_change(self):
        if self.printer_mode.get() == "named":
            self.cmb_printers.configure(state="readonly")
        else:
            self.cmb_printers.configure(state="disabled")

    def _is_individual_print_enabled(self) -> bool:
        return bool(self.individual_print_var.get())

    def _send_label_to_printer(
        self,
        item: Dict[str, Any],
        copies: int,
        individual_print: Optional[bool] = None,
        show_errors: bool = True,
    ):
        if individual_print is None:
            individual_print = self._is_individual_print_enabled()

        if individual_print:
            print_label_gdi_small(item, self.config_data, copies=copies, show_errors=show_errors)
        else:
            print_label_gdi(item, self.config_data, copies=copies, show_errors=show_errors)

    def _item_for_silent_print(self, item: Dict[str, Any], individual_print: bool) -> Dict[str, Any]:
        item_to_print = build_label_text(item, self.config_data)
        if individual_print:
            item_to_print = _item_for_individual_label(item_to_print)
            item_to_print["_AUTO_INDIVIDUAL_PRINT"] = True
            return item_to_print

        if not item_to_print.get("TIENE_PRECIO_ESPECIAL"):
            item_to_print["PRECIO_ESPECIAL"] = ""
            item_to_print["PROMO_TERMINOS"] = ""
            item_to_print["PROMO_TERMINOS2"] = ""
        return item_to_print

    # --------- Eventos ----------
    def on_unified_search(self, event=None):
        term = (self.entry_search.get() or "").strip()
        if not term:
            return

        auto_print = self.auto_print_var.get()
        individual_print = self._is_individual_print_enabled()
        try:
            results = search_items(term, self.config_data)
            
            if not results:
                if not auto_print:
                    messagebox.showinfo(APP_TITLE, "No se encontraron artículos.")
                self._clear_preview()
            elif len(results) == 1:
                self._show_item_in_preview(results[0], suppress_no_special_message=auto_print)
                if auto_print:
                    self.after(
                        80,
                        lambda item=results[0], individual=individual_print: self._print_item_silent(
                            item,
                            individual_print=individual,
                        ),
                    )
            else:
                if not auto_print:
                    # Multiple results, open selection window
                    ResultSelectionWindow(self, results, self._show_item_in_preview)

        except Exception as e:
            write_log("ERROR on_unified_search", e, extra=f"term={term}")
            if not auto_print:
                messagebox.showerror(APP_TITLE, f"Error en la búsqueda: {e}")
        
        # Clear search box after search
        self.entry_search.delete(0, tk.END)


    def on_test_print(self):
        # Construir datos de ejemplo
        mock = {
            "DESCRIPCION": "PRODUCTO DE PRUEBA 80MM",
            "PRECIO": "$99.90",
            "PRECIO_ESPECIAL": "",
            "ARTICULO": "TEST001",
            "UPC": "123456789012",
            "VIGENCIA": resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"]))
        }
        try:
            copies = int(self.spin_copias.get())
            self._update_config_from_ui()
            self._send_label_to_printer(mock, copies=copies)
            messagebox.showinfo(APP_TITLE, "Impresión de prueba enviada.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Error en impresión de prueba: {e}")

    def _validate_special_price(self, item: Dict[str, Any]) -> bool:
        special_price = item.get("PRECIO_ESPECIAL") or ""
        if not special_price:
            messagebox.showwarning(APP_TITLE, "El precio especial es un dato obligatorio.")
            self.entry_precio_especial.focus_set()
            return False

        base = _price_to_float(item.get("PRECIO"))
        special = _price_to_float(special_price)
        if base is None or special is None:
            messagebox.showwarning(APP_TITLE, "Valida el precio especial; no se pudo leer como importe.")
            return False
        if special > base:
            messagebox.showwarning(
                APP_TITLE,
                "Valida el precio especial: es mayor que el precio normal de venta."
            )
            return False
        return True

    def on_print_click(self):
        item = self._collect_item_from_preview()
        if not item.get("DESCRIPCION") or not item.get("PRECIO"):
            messagebox.showwarning(APP_TITLE, "No se puede imprimir sin Producto y Precio.")
            return
        if not self._is_individual_print_enabled() and not self._validate_special_price(item):
            return
        try:
            # Validaciones
            price_digits = re.sub(r"[^0-9.]", "", item["PRECIO"])
            if not price_digits:
                raise ValueError("Precio inválido.")
            copies = int(self.spin_copias.get())
            if copies < 1 or copies > 20:
                raise ValueError("Copias fuera de rango (1..20).")

            self._update_config_from_ui()
            self._send_label_to_printer(item, copies=copies)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Error al imprimir: {e}")

    def _print_current_item_silent(self, individual_print: Optional[bool] = None):
        item = self._collect_item_from_preview()
        if individual_print is None:
            individual_print = self._is_individual_print_enabled()
        self._print_item_silent(item, individual_print=individual_print)

    def _print_item_silent(self, item: Dict[str, Any], individual_print: bool):
        item = self._item_for_silent_print(item, individual_print)
        if not item.get("DESCRIPCION") or not item.get("PRECIO"):
            return

        try:
            price_digits = re.sub(r"[^0-9.]", "", item["PRECIO"])
            if not price_digits:
                return
            copies = int(self.spin_copias.get())
            if copies < 1 or copies > 20:
                return

            self._update_config_from_ui()
            self._send_label_to_printer(
                item,
                copies=copies,
                individual_print=individual_print,
                show_errors=False,
            )
        except Exception as e:
            write_log("ERROR auto_print", e, extra=f"item={item}")

    def on_preview_print_click(self):
        item = self._collect_item_from_preview()
        if not item.get("DESCRIPCION") or not item.get("PRECIO"):
            messagebox.showwarning(APP_TITLE, "No se puede mostrar la vista previa sin Producto y Precio.")
            return
        if not self._is_individual_print_enabled() and not self._validate_special_price(item):
            return
        try:
            price_digits = re.sub(r"[^0-9.]", "", item["PRECIO"])
            if not price_digits:
                raise ValueError("Precio invalido.")
            self._update_config_from_ui()
            LabelPrintPreviewWindow(self, item, individual=self._is_individual_print_enabled())
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Error al mostrar vista previa: {e}")

    def on_close(self):
        try:
            self._update_config_from_ui()
            save_config(self.config_data)
        except Exception as e:
            print(f"[WARN] No se pudo guardar config: {e}")
        self.destroy()

    # --------- Lógica búsqueda/preview ----------
    def _set_locked_preview_fields_state(self, state: str):
        self.txt_producto.configure(state="normal" if state == "normal" else "disabled")
        for entry in (
            self.entry_precio,
            self.entry_precio_especial,
            self.entry_codigo,
            self.entry_upc,
        ):
            entry.configure(state=state)

    def _lock_preview_fields(self):
        self._set_locked_preview_fields_state("readonly")
        self.txt_producto.configure(state="disabled")

    def _unlock_preview_fields(self):
        self._set_locked_preview_fields_state("normal")

    def _clear_preview(self):
        self.current_built_item = {}
        self._unlock_preview_fields()
        self.txt_producto.delete("1.0", tk.END)
        self.entry_precio.delete(0, tk.END)
        self.entry_precio_especial.delete(0, tk.END)
        self.entry_codigo.delete(0, tk.END)
        self.entry_upc.delete(0, tk.END)
        self._lock_preview_fields()
        self.entry_vigencia.delete(0, tk.END)
        self.entry_vigencia.insert(0, resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"])))


    def _show_item_in_preview(self, item: Optional[Dict[str, Any]], suppress_no_special_message: bool = False):
        if not item: # This can happen if the selection window is closed
            return
            
        # Build text for preview
        built = build_label_text(item, self.config_data)
        self.current_built_item = built
        if not suppress_no_special_message and not built.get("TIENE_PRECIO_ESPECIAL"):
            messagebox.showinfo(APP_TITLE, "No existen precios especiales actualizados.")

        self._unlock_preview_fields()

        # Producto/Descripción
        self.txt_producto.delete("1.0", tk.END)
        self.txt_producto.insert("1.0", built.get("DESCRIPCION", ""))

        # Precio
        self.entry_precio.delete(0, tk.END)
        self.entry_precio.insert(0, built.get("PRECIO", ""))

        self.entry_precio_especial.delete(0, tk.END)
        self.entry_precio_especial.insert(0, built.get("PRECIO_ESPECIAL", ""))

        # Código interno
        self.entry_codigo.delete(0, tk.END)
        self.entry_codigo.insert(0, built.get("ARTICULO", ""))

        # UPC
        self.entry_upc.delete(0, tk.END)
        self.entry_upc.insert(0, built.get("UPC", ""))

        self._lock_preview_fields()

        # Vigencia
        self.entry_vigencia.delete(0, tk.END)
        self.entry_vigencia.insert(0, built.get("VIGENCIA", ""))

    def _collect_item_from_preview(self) -> Dict[str, Any]:
        descripcion = self.txt_producto.get("1.0", tk.END).strip()
        precio_raw = (self.entry_precio.get() or "").strip()
        precio_especial_raw = (self.entry_precio_especial.get() or "").strip()
        # Normalizar precio a $xx.xx
        precio_digits = re.sub(r"[^0-9.]", "", precio_raw)
        precio_fmt = _format_price(precio_digits) if precio_digits else ""
        precio_especial_digits = re.sub(r"[^0-9.]", "", precio_especial_raw)
        precio_especial_fmt = _format_price(precio_especial_digits) if precio_especial_digits else ""
        if not precio_especial_fmt:
            precio_especial_fmt = precio_fmt

        articulo = (self.entry_codigo.get() or "").strip()
        upc = _normalize_upc_digits(self.entry_upc.get() or "")
        vig = (self.entry_vigencia.get() or "").strip()
        if not vig:
            vig = resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"]))
        collected = {
            "DESCRIPCION": descripcion,
            "PRECIO": precio_fmt,
            "PRECIO_ESPECIAL": precio_especial_fmt,
            "ARTICULO": articulo,
            "UPC": upc,
            "VIGENCIA": vig
        }
        for key in ("TIENE_PRECIO_ESPECIAL", "PROMO_FECHAINICIO", "PROMO_FECHAFINAL", "PROMO_TERMINOS", "PROMO_TERMINOS2"):
            if key in self.current_built_item:
                collected[key] = self.current_built_item[key]
        return collected

    def _update_config_from_ui(self):
        # Impresora
        self.config_data["impresora"]["modo"] = self.printer_mode.get()
        if self.printer_mode.get() == "named":
            self.config_data["impresora"]["printer_name"] = self.cmb_printers.get()
        else:
            self.config_data["impresora"]["printer_name"] = ""

        preferencias = self.config_data.setdefault("preferencias", {})
        preferencias["auto_print_on_scan"] = bool(self.auto_print_var.get())
        preferencias["individual_print"] = self._is_individual_print_enabled()

        # The "tamaños" and "espaciados" sections are removed from the new UI,
        # so we don't need to update them anymore. If you want to keep them,
        # you would need to add them back to the create_widgets method.
        
        # Example for one of them if you add it back:
        # self.config_data["espaciados"]["espacio_abajo_precio"] = int(self.spin_down.get())


# ==============================
# Entry point
# ==============================
def main():
    try:
        app = App()
        app.protocol("WM_DELETE_WINDOW", app.on_close)
        app.mainloop()
    except Exception as e:
        write_log("ERROR main", e)
        raise


if __name__ == "__main__":
    main()
