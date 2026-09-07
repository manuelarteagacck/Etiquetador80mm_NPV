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
import math
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
PARAMCONF_FILENAME = "paramconf.json"
CONFIG_LOAD_WARNINGS = []
PROMO_BACKGROUND_FILENAME = "back5.png"
PROMO_TEMPLATE_SIZE = (1585, 992)
LOGUITO_FILENAME = "loguito.png"
APP_ICON_FILENAME = "ticket_printer.ico"
PROMO_DESCRIPTION_FONT_SIZE = 60


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
    import win32gui
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
PRICE_SOURCE_TABLE = "NPV.dbo.NPVFDPreciosVenta"
PRICE_BACKUP_TABLE = "NPV.dbo.NPVFDPreciosVentaBkp"
PRICE_NEW_DAILY_TABLE = "NPV.dbo.NPVFDPreciosNuevosDiarios"


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
        "auto_print_on_scan": True,
        "individual_print": False,
        "special_price_print": False
    }
}


DEFAULT_PRINT_CONFIG = {
    "impresora": {
        "modo": "default",
        "printer_name": ""
    },
    "escpos": {
        "mode": "spooler_windows",
        "serial": {"com": "COM3", "baudrate": 9600, "bytesize": 8, "parity": "N", "stopbits": 1, "timeout": 1},
        "usb": {"vendor_id": 0x0000, "product_id": 0x0000, "interface": 0, "in_ep": 0x82, "out_ep": 0x01},
        "network": {"ip": "192.168.1.50", "port": 9100}
    },
    "tamaños": {
        "size_producto": [2, 1],
        "size_precio": [2, 2],
        "size_codigos": [1, 1],
        "size_vigencia": [1, 1]
    },
    "espaciados": {
        "espacio_arriba_precio": 0,
        "espacio_abajo_precio": 0
    },
    "corte": "partial",
    "vigencia_default": "Vigente {MES_ABR} {YYYY}",
    "preferencias": {
        "auto_print_on_scan": True,
        "individual_print": False,
        "special_price_print": False
    },
    "layout_impresion": {
        "version": 2,
        "papel": {
            "auto_height": True,
            "bottom_margin": 12,
            "minimum_height": 180
        },
        "lineas_division": {
            "enabled": True,
            "top_y": 6,
            "margin": 10,
            "thickness": 3
        },
        "normal": {
            "top_y": 0,
            "draw_top_divider": False,
            "horizontal_margin": 10,
            "descripcion_font_height": 46,
            "descripcion_rect_height": 200,
            "descripcion_advance": 70,
            "precio_font_height": 92,
            "precio_rect_height": 118,
            "precio_especial_label_font_height": 24,
            "precio_especial_price_font_height": 54,
            "precio_especial_save_font_height": 22,
            "footer_font_height": 24,
            "footer_rect_height": 30,
            "footer_advance": 25,
            "barcode_margin": 35,
            "barcode_height": 58,
            "barcode_text_height": 20
        },
        "angosta": {
            "top_y": 0,
            "draw_top_divider": False,
            "horizontal_margin": 48,
            "descripcion_font_height": 34,
            "descripcion_font_height_auto": 30,
            "descripcion_rect_height": 200,
            "descripcion_advance": 58,
            "descripcion_advance_auto": 68,
            "precio_font_height": 70,
            "precio_rect_height": 92,
            "precio_especial_label_font_height": 20,
            "precio_especial_price_font_height": 43,
            "precio_especial_save_font_height": 18,
            "footer_font_height": 19,
            "footer_rect_height": 30,
            "footer_advance": 22,
            "barcode_height": 46,
            "barcode_text_height": 16
        }
    }
}

PRINT_CONFIG_KEYS = tuple(DEFAULT_PRINT_CONFIG.keys())


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


def _clone_json(data):
    return json.loads(json.dumps(data))


def _deep_merge(target: dict, defaults: dict) -> bool:
    changed = False
    for key, value in defaults.items():
        if key not in target:
            target[key] = _clone_json(value)
            changed = True
        elif isinstance(value, dict) and isinstance(target.get(key), dict):
            changed = _deep_merge(target[key], value) or changed
    return changed


def _migrate_print_layout(layout: dict) -> bool:
    """Actualiza configuraciones conservadas por instalaciones anteriores."""
    if not isinstance(layout, dict):
        return False

    try:
        version = int(layout.get("version", 1))
    except (TypeError, ValueError):
        version = 1

    changed = False
    if version < 2:
        # Version 2: la descripcion inicia en el borde imprimible y las
        # etiquetas normal/angosta ya no incluyen la linea de corte superior.
        for section_name in ("normal", "angosta"):
            section = layout.setdefault(section_name, {})
            if not isinstance(section, dict):
                section = {}
                layout[section_name] = section
            section["top_y"] = 0
            section["draw_top_divider"] = False
        layout["version"] = 2
        changed = True

    return changed


def _extract_print_config(cfg: dict) -> dict:
    print_cfg = _clone_json(DEFAULT_PRINT_CONFIG)
    for key in PRINT_CONFIG_KEYS:
        if key in cfg:
            print_cfg[key] = _clone_json(cfg[key])
            if key == "layout_impresion":
                _migrate_print_layout(print_cfg[key])
            if isinstance(print_cfg[key], dict) and isinstance(DEFAULT_PRINT_CONFIG.get(key), dict):
                _deep_merge(print_cfg[key], DEFAULT_PRINT_CONFIG[key])
    return print_cfg


def _apply_print_config(cfg: dict, print_cfg: dict) -> dict:
    for key in PRINT_CONFIG_KEYS:
        if key in print_cfg:
            cfg[key] = _clone_json(print_cfg[key])
    return cfg


def load_print_config(seed_cfg: dict = None) -> dict:
    path = paramconf_path()
    print_cfg = _extract_print_config(seed_cfg or {}) if seed_cfg else _clone_json(DEFAULT_PRINT_CONFIG)
    must_save = False

    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8-sig") as f:
                data = json.load(f)
            if isinstance(data, dict):
                print_cfg = data
                layout = print_cfg.setdefault("layout_impresion", {})
                must_save = _migrate_print_layout(layout)
                must_save = _deep_merge(print_cfg, DEFAULT_PRINT_CONFIG) or must_save
            else:
                must_save = True
        except Exception as exc:
            _remember_config_warning(
                f"No se pudo leer {PARAMCONF_FILENAME}. Se usaran parametros de impresion por defecto.",
                exc,
            )
            must_save = True
    else:
        must_save = True

    if must_save:
        save_print_config(print_cfg)
    return print_cfg


def save_print_config(cfg: dict):
    print_cfg = _extract_print_config(cfg or {})
    with open(paramconf_path(), "w", encoding="utf-8") as f:
        json.dump(print_cfg, f, ensure_ascii=False, indent=2)


def load_config() -> dict:
    CONFIG_LOAD_WARNINGS.clear()
    path = config_path()
    if not os.path.exists(path):
        cfg = _clone_json(DEFAULT_CONFIG)
        _apply_print_config(cfg, load_print_config(cfg))
        _hydrate_db_password(cfg)
        save_config(cfg)
        return cfg
    with open(path, "r", encoding="utf-8-sig") as f:
        data = json.load(f)
    _deep_merge(data, DEFAULT_CONFIG)
    cfg = _apply_print_config(data, load_print_config(data))
    password_migrated = _hydrate_db_password(cfg)
    if password_migrated:
        save_config(cfg)
    return cfg


def save_config(cfg: dict):
    save_print_config(cfg)
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


def paramconf_path() -> str:
    return os.path.join(app_base_path(), PARAMCONF_FILENAME)


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


def _remember_config_warning(message: str, exc: Exception = None) -> None:
    CONFIG_LOAD_WARNINGS.append(message)
    try:
        write_log("WARN config", exc, extra=message)
    except Exception:
        pass


def pop_config_load_warnings() -> list:
    warnings = list(CONFIG_LOAD_WARNINGS)
    CONFIG_LOAD_WARNINGS.clear()
    return warnings


def _hydrate_db_password(cfg: dict) -> bool:
    c = cfg.setdefault("conexion_odbc", {})
    had_plain_password = "password" in c
    plain_password = c.get("password")
    encrypted_password = c.get("password_encrypted")

    if had_plain_password and plain_password:
        c["password"] = str(plain_password)
        return True
    if encrypted_password:
        try:
            c["password"] = _decrypt_secret(encrypted_password)
        except ValueError as exc:
            c["password"] = ""
            c["password_encrypted"] = ""
            c.setdefault("password_encrypted_invalid", encrypted_password)
            _remember_config_warning(
                "No se pudo descifrar la contrasena guardada de SQL Server. "
                "El programa se abrira con la contrasena vacia. "
                f"Si necesitas recuperar la contrasena anterior, busca el {FERNET_KEY_FILENAME} "
                "que corresponde a este config.json; el token original quedo guardado en "
                "conexion_odbc.password_encrypted_invalid.",
                exc,
            )
            return True
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
    return app_or_resource_path(PROMO_BACKGROUND_FILENAME)


def apply_window_icon(window) -> None:
    icon_path = app_or_resource_path(APP_ICON_FILENAME)
    if not os.path.exists(icon_path):
        return
    try:
        window.iconbitmap(icon_path)
    except Exception as exc:
        write_log("ERROR apply_window_icon", exc, extra=f"icon={icon_path}")


def app_or_resource_path(filename: str) -> str:
    candidate = os.path.join(app_base_path(), filename)
    if os.path.exists(candidate):
        return candidate
    return resource_path(filename)


def load_ui_image(filename: str, max_size: tuple):
    if Image is None or ImageTk is None:
        return None
    path = app_or_resource_path(filename)
    if not os.path.exists(path):
        return None
    try:
        with Image.open(path) as img:
            img = img.convert("RGBA")
            img.thumbnail(max_size, Image.LANCZOS)
            return ImageTk.PhotoImage(img)
    except Exception as exc:
        write_log("ERROR load_ui_image", exc, extra=f"image={path}")
        return None


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


def _promo_price_text(value) -> str:
    """Devuelve un precio sin signo para ocupar el espacio de la plantilla."""
    integer, decimals = _split_price(value)
    if not integer:
        return "0.00"
    return f"{integer}.{decimals}"


def _quantity_to_float(value) -> Optional[float]:
    try:
        if value is None:
            return None
        return float(value)
    except Exception:
        return None


def _format_quantity(value) -> str:
    qty = _quantity_to_float(value)
    if qty is None:
        return ""
    if qty.is_integer():
        return str(int(qty))
    return f"{qty:g}"


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
        CASE WHEN COALESCE(existencia.DISPONIBLE, 0) > 1 THEN
            promo.PRECIOUNITARIO *
                (1 + COALESCE(i1.PORCENTAJE, 0) / 100.0) *
                (1 + COALESCE(i2.PORCENTAJE, 0) / 100.0) *
                (1 + COALESCE(i3.PORCENTAJE, 0) / 100.0) *
                (1 + COALESCE(i4.PORCENTAJE, 0) / 100.0)
        END AS PRECIO_ESPECIAL,
        CASE
            WHEN promo.PRECIOUNITARIO IS NULL THEN 0
            WHEN COALESCE(existencia.DISPONIBLE, 0) <= 1 THEN 0
            ELSE 1
        END AS TIENE_PRECIO_ESPECIAL,
        promo.PROMO_CANTIDAD,
        promo.PROMO_FECHAINICIO,
        promo.PROMO_FECHAFINAL,
        existencia.DISPONIBLE,
        (SELECT TOP 1 e.EQUIVALENTE FROM NPV.dbo.NPVFDArticulosEquivalentes e WHERE e.ARTICULO = a.ARTICULO) as UPC,
        pr.FECHAINICIO
    FROM NPV.dbo.NPVFDArticulos a
    JOIN PreciosRankeados pr ON a.ARTICULO = pr.ARTICULO
    OUTER APPLY (
        SELECT TOP 1
            pl.CANTIDAD AS PROMO_CANTIDAD,
            pl.PRECIOUNITARIO,
            pe.VIGENCIAINICIO AS PROMO_FECHAINICIO,
            pe.VIGENCIAFINAL AS PROMO_FECHAFINAL
        FROM NPV.dbo.NPVFDPromocionLineas pl
        INNER JOIN NPV.dbo.NPVFDPromocionEncabezado pe ON pe.CLAVE = pl.CLAVE
        WHERE pl.CLAVEARTICULO = a.ARTICULO
          AND COALESCE(pl.CANTIDAD, 0) > 0
          AND pl.CANTIDAD = 1
          AND pl.PRECIOUNITARIO IS NOT NULL
          AND pl.STATUS = 'A'
          AND pe.STATUS = 'A'
          AND pe.VIGENCIAINICIO <= GETDATE()
          AND pe.VIGENCIAFINAL >= GETDATE()
          AND pe.CLASE = 'Z002'
        ORDER BY
            CASE WHEN pl.CANTIDAD = 1 THEN 0 ELSE 1 END,
            pl.PRECIOUNITARIO ASC,
            pl.CLAVE DESC
    ) promo
    OUTER APPLY (
        SELECT MAX(COALESCE(ex.DISPONIBLE, 0)) AS DISPONIBLE
        FROM NPV.dbo.NPVFDExistencias ex
        WHERE ex.ARTICULO LIKE a.ARTICULO
    ) existencia
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
            tiene_precio_especial = bool(row.TIENE_PRECIO_ESPECIAL)
            precio_especial = _format_price(row.PRECIO_ESPECIAL) if tiene_precio_especial else ""
            precio_venta_num = _price_to_float(precio_venta)
            precio_especial_num = _price_to_float(precio_especial)
            if (
                tiene_precio_especial
                and precio_venta_num is not None
                and precio_especial_num is not None
                and precio_especial_num >= precio_venta_num
            ):
                tiene_precio_especial = False
                precio_especial = ""
            
            item = {
                "ARTICULO": str(row.ARTICULO).strip(),
                "DESCRIPCION": str(row.DESCRIPCION or "").strip(),
                "PRECIO": precio_venta,
                "PRECIO_ESPECIAL": precio_especial,
                "TIENE_PRECIO_ESPECIAL": tiene_precio_especial,
                "UPC": _normalize_upc_digits(row.UPC or ""),
                "VIGENCIA": f"Valido a partir de: {fecha_vigencia} Aplican TyC",
                "PROMO_CANTIDAD": row.PROMO_CANTIDAD if tiene_precio_especial else None,
                "PROMO_FECHAINICIO": row.PROMO_FECHAINICIO if tiene_precio_especial else None,
                "PROMO_FECHAFINAL": row.PROMO_FECHAFINAL if tiene_precio_especial else None,
                "DISPONIBLE": row.DISPONIBLE,
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


def _quote_sql_identifier(name: str) -> str:
    if not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", str(name or "")):
        raise ValueError(f"Identificador SQL no valido: {name!r}")
    return f"[{name}]"


def _get_source_price_column(cursor) -> str:
    rows = cursor.execute(
        """
        SELECT COLUMN_NAME
        FROM NPV.INFORMATION_SCHEMA.COLUMNS
        WHERE TABLE_SCHEMA = 'dbo'
          AND TABLE_NAME = 'NPVFDPreciosVenta'
          AND COLUMN_NAME IN ('PRECIOVENTA', 'PRECIOSVENTA', 'PRECIO')
        """
    ).fetchall()
    available = {str(row.COLUMN_NAME).upper(): str(row.COLUMN_NAME) for row in rows}
    for candidate in ("PRECIOVENTA", "PRECIOSVENTA", "PRECIO"):
        if candidate in available:
            return available[candidate]
    raise RuntimeError("No se encontro una columna de precio en NPV.dbo.NPVFDPreciosVenta.")


def _ensure_price_backup_comparison_schema(cursor, price_column: str) -> bool:
    """
    Prepara el respaldo para comparar la clave ARTICULO + FECHAINICIO.

    Devuelve True cuando fue necesario crear o migrar la estructura. Si la
    tabla no existe se guarda, por articulo, el ultimo precio cuya FECHAINICIO
    sea estrictamente anterior al dia actual.
    """
    quoted_price = _quote_sql_identifier(price_column)
    table_exists = cursor.execute(
        f"SELECT OBJECT_ID(N'{PRICE_BACKUP_TABLE}', N'U')"
    ).fetchval()

    if not table_exists:
        cursor.execute(
            f"""
            SELECT TOP (0)
                pv.ARTICULO,
                pv.{quoted_price} AS PRECIOVENTA,
                pv.FECHAINICIO,
                pv.FECHAALTA,
                CAST(GETDATE() AS datetime) AS FECHACOPIA
            INTO {PRICE_BACKUP_TABLE}
            FROM {PRICE_SOURCE_TABLE} AS pv;

            ;WITH Fuente AS (
                SELECT
                    pv.ARTICULO,
                    pv.{quoted_price} AS PRECIOVENTA,
                    pv.FECHAINICIO,
                    pv.FECHAALTA,
                    ROW_NUMBER() OVER (
                        PARTITION BY pv.ARTICULO
                        ORDER BY
                            pv.FECHAINICIO DESC,
                            pv.FECHAALTA DESC,
                            pv.{quoted_price} DESC
                    ) AS rn
                FROM {PRICE_SOURCE_TABLE} AS pv
                WHERE pv.FECHAINICIO < CONVERT(date, GETDATE())
            )
            INSERT INTO {PRICE_BACKUP_TABLE} (
                ARTICULO,
                PRECIOVENTA,
                FECHAINICIO,
                FECHAALTA,
                FECHACOPIA
            )
            SELECT
                ARTICULO,
                PRECIOVENTA,
                FECHAINICIO,
                FECHAALTA,
                GETDATE()
            FROM Fuente
            WHERE rn = 1;
            """
        )
        return True

    start_date_exists = cursor.execute(
        """
        SELECT COL_LENGTH(
            N'NPV.dbo.NPVFDPreciosVentaBkp',
            N'FECHAINICIO'
        )
        """
    ).fetchval()
    if start_date_exists is not None:
        return False

    # La version anterior del respaldo guardaba el precio y FECHAALTA pero no
    # FECHAINICIO. Se conserva esa instantanea y se recupera la fecha del
    # registro fuente que mejor coincide, de modo que los cambios pendientes
    # sigan apareciendo despues de actualizar la aplicacion.
    cursor.execute(
        f"""
        ALTER TABLE {PRICE_BACKUP_TABLE}
            ADD FECHAINICIO datetime NULL;
        """
    )
    cursor.execute(
        f"""
        UPDATE b
        SET FECHAINICIO = origen.FECHAINICIO
        FROM {PRICE_BACKUP_TABLE} AS b
        OUTER APPLY (
            SELECT TOP (1)
                pv.FECHAINICIO
            FROM {PRICE_SOURCE_TABLE} AS pv
            WHERE pv.ARTICULO = b.ARTICULO
            ORDER BY
                CASE
                    WHEN (
                        pv.FECHAALTA = b.FECHAALTA
                        OR (pv.FECHAALTA IS NULL AND b.FECHAALTA IS NULL)
                    )
                    AND (
                        pv.{quoted_price} = b.PRECIOVENTA
                        OR (pv.{quoted_price} IS NULL AND b.PRECIOVENTA IS NULL)
                    )
                    THEN 0
                    WHEN pv.FECHAALTA = b.FECHAALTA THEN 1
                    WHEN pv.{quoted_price} = b.PRECIOVENTA THEN 2
                    ELSE 3
                END,
                CASE
                    WHEN pv.FECHAINICIO IS NULL OR pv.FECHAINICIO <= GETDATE()
                    THEN 0 ELSE 1
                END,
                pv.FECHAINICIO DESC,
                pv.FECHAALTA DESC
        ) AS origen;

        ;WITH LineaBaseHistorica AS (
            SELECT
                pv.ARTICULO,
                pv.{quoted_price} AS PRECIOVENTA,
                pv.FECHAINICIO,
                pv.FECHAALTA,
                ROW_NUMBER() OVER (
                    PARTITION BY pv.ARTICULO
                    ORDER BY
                        pv.FECHAINICIO DESC,
                        pv.FECHAALTA DESC,
                        pv.{quoted_price} DESC
                ) AS rn
            FROM {PRICE_SOURCE_TABLE} AS pv
            WHERE pv.FECHAINICIO < CONVERT(date, GETDATE())
              AND NOT EXISTS (
                    SELECT 1
                    FROM {PRICE_BACKUP_TABLE} AS b
                    WHERE b.ARTICULO = pv.ARTICULO
              )
        )
        INSERT INTO {PRICE_BACKUP_TABLE} (
            ARTICULO,
            PRECIOVENTA,
            FECHAINICIO,
            FECHAALTA,
            FECHACOPIA
        )
        SELECT
            ARTICULO,
            PRECIOVENTA,
            FECHAINICIO,
            FECHAALTA,
            GETDATE()
        FROM LineaBaseHistorica
        WHERE rn = 1;
        """
    )
    return True


def _ensure_new_daily_prices_table(cursor, price_column: str) -> bool:
    quoted_price = _quote_sql_identifier(price_column)
    table_exists = cursor.execute(
        f"SELECT OBJECT_ID(N'{PRICE_NEW_DAILY_TABLE}', N'U')"
    ).fetchval()
    if table_exists:
        return False

    cursor.execute(
        f"""
        SELECT TOP (0)
            CONVERT(date, GETDATE()) AS FECHADETECCION,
            pv.ARTICULO,
            CAST(a.DESCRIPCION AS varchar(255)) AS DESCRIPCION,
            pv.{quoted_price} AS PRECIOVENTA_ANTERIOR,
            pv.{quoted_price} AS PRECIOVENTA,
            pv.FECHAINICIO,
            pv.FECHAALTA,
            CAST(GETDATE() AS datetime) AS FECHACOPIA_BKP,
            CAST(GETDATE() AS datetime) AS FECHAREGISTRO
        INTO {PRICE_NEW_DAILY_TABLE}
        FROM {PRICE_SOURCE_TABLE} AS pv
        LEFT JOIN NPV.dbo.NPVFDArticulos AS a
            ON a.ARTICULO = pv.ARTICULO;
        """
    )
    return True


def _refresh_new_daily_prices_if_changed(cursor, price_column: str) -> tuple:
    """
    Reemplaza totalmente la tabla diaria solo cuando hay claves nuevas.

    El respaldo contiene el ultimo precio anterior a hoy. La comparacion toma
    los precios de hoy en adelante que no tengan la misma clave
    ARTICULO + FECHAINICIO en el respaldo. Si no hay diferencias, el contenido
    previo de NPVFDPreciosNuevosDiarios permanece intacto.
    """
    quoted_price = _quote_sql_identifier(price_column)
    cursor.execute(
        f"""
        SET XACT_ABORT ON;
        BEGIN TRANSACTION;

        ;WITH Fuente AS (
            SELECT
                pv.ARTICULO,
                pv.{quoted_price} AS PRECIOVENTA,
                pv.FECHAINICIO,
                pv.FECHAALTA,
                ROW_NUMBER() OVER (
                    PARTITION BY pv.ARTICULO, pv.FECHAINICIO
                    ORDER BY pv.FECHAALTA DESC, pv.{quoted_price} DESC
                ) AS rn
            FROM {PRICE_SOURCE_TABLE} AS pv
            WHERE pv.FECHAINICIO >= CONVERT(date, GETDATE())
              AND NOT EXISTS (
                SELECT 1
                FROM {PRICE_BACKUP_TABLE} AS b
                WHERE b.ARTICULO = pv.ARTICULO
                  AND (
                        b.FECHAINICIO = pv.FECHAINICIO
                     OR (b.FECHAINICIO IS NULL AND pv.FECHAINICIO IS NULL)
                  )
            )
        )
        SELECT
            ARTICULO,
            PRECIOVENTA,
            FECHAINICIO,
            FECHAALTA,
            rn
        INTO #PreciosCambiados
        FROM Fuente;

        DECLARE @changed_rows int = (
            SELECT COUNT(1)
            FROM #PreciosCambiados
            WHERE rn = 1
        );
        DECLARE @inserted_rows int = 0;

        IF @changed_rows > 0
        BEGIN
            DELETE FROM {PRICE_NEW_DAILY_TABLE};

            INSERT INTO {PRICE_NEW_DAILY_TABLE} (
                FECHADETECCION,
                ARTICULO,
                DESCRIPCION,
                PRECIOVENTA_ANTERIOR,
                PRECIOVENTA,
                FECHAINICIO,
                FECHAALTA,
                FECHACOPIA_BKP,
                FECHAREGISTRO
            )
            SELECT
                CONVERT(date, GETDATE()),
                pc.ARTICULO,
                COALESCE(NULLIF(a.DESCRIPCION, ''), pc.ARTICULO),
                anterior.PRECIOVENTA,
                pc.PRECIOVENTA,
                pc.FECHAINICIO,
                pc.FECHAALTA,
                anterior.FECHACOPIA,
                GETDATE()
            FROM #PreciosCambiados AS pc
            LEFT JOIN NPV.dbo.NPVFDArticulos AS a
                ON a.ARTICULO = pc.ARTICULO
            OUTER APPLY (
                SELECT TOP (1)
                    b.PRECIOVENTA,
                    b.FECHACOPIA
                FROM {PRICE_BACKUP_TABLE} AS b
                WHERE b.ARTICULO = pc.ARTICULO
                ORDER BY
                    CASE
                        WHEN b.FECHAINICIO <= pc.FECHAINICIO THEN 0
                        ELSE 1
                    END,
                    b.FECHAINICIO DESC,
                    b.FECHAALTA DESC
            ) AS anterior
            WHERE pc.rn = 1;

            SET @inserted_rows = @@ROWCOUNT;
        END;

        COMMIT TRANSACTION;

        SELECT
            @changed_rows AS changed_rows,
            @inserted_rows AS inserted_rows;
        """
    )
    while cursor.description is None:
        if not cursor.nextset():
            return 0, 0
    row = cursor.fetchone()
    if not row:
        return 0, 0
    return int(row.changed_rows or 0), int(row.inserted_rows or 0)


def _load_latest_new_daily_price_items(cursor) -> List[Dict[str, Any]]:
    rows = cursor.execute(
        f"""
        ;WITH UltimoLote AS (
            SELECT MAX(FECHADETECCION) AS FECHADETECCION
            FROM {PRICE_NEW_DAILY_TABLE}
        ),
        PreciosGuardados AS (
            SELECT
                nd.ARTICULO,
                COALESCE(NULLIF(nd.DESCRIPCION, ''), a.DESCRIPCION, nd.ARTICULO) AS DESCRIPCION,
                nd.PRECIOVENTA,
                nd.FECHAINICIO,
                e.EQUIVALENTE AS UPC,
                ROW_NUMBER() OVER (
                    PARTITION BY nd.ARTICULO, nd.FECHAINICIO
                    ORDER BY nd.FECHAREGISTRO DESC, nd.PRECIOVENTA DESC
                ) AS rn
            FROM {PRICE_NEW_DAILY_TABLE} AS nd
            CROSS JOIN UltimoLote AS ul
            LEFT JOIN NPV.dbo.NPVFDArticulos AS a
                ON a.ARTICULO = nd.ARTICULO
            OUTER APPLY (
                SELECT TOP (1)
                    ae.EQUIVALENTE
                FROM NPV.dbo.NPVFDArticulosEquivalentes AS ae
                WHERE ae.ARTICULO = nd.ARTICULO
                ORDER BY ae.EQUIVALENTE
            ) AS e
            WHERE nd.FECHADETECCION = ul.FECHADETECCION
        )
        SELECT
            ARTICULO,
            DESCRIPCION,
            PRECIOVENTA,
            FECHAINICIO,
            UPC
        FROM PreciosGuardados
        WHERE rn = 1
        ORDER BY FECHAINICIO, DESCRIPCION, ARTICULO;
        """
    ).fetchall()

    results = []
    for row in rows:
        fecha_vigencia = (
            row.FECHAINICIO.strftime("%d/%m/%Y")
            if row.FECHAINICIO
            else datetime.date.today().strftime("%d/%m/%Y")
        )
        results.append({
            "ARTICULO": str(row.ARTICULO).strip(),
            "DESCRIPCION": str(row.DESCRIPCION or "").strip(),
            "PRECIO": _format_price(row.PRECIOVENTA),
            "PRECIO_ESPECIAL": "",
            "TIENE_PRECIO_ESPECIAL": False,
            "UPC": _normalize_upc_digits(row.UPC or ""),
            "VIGENCIA": f"Valido a partir de: {fecha_vigencia} Aplican TyC",
        })
    return results


def fetch_new_daily_price_items(cfg: dict = None) -> List[Dict[str, Any]]:
    """
    Actualiza condicionalmente NPVFDPreciosNuevosDiarios y carga su ultimo lote.
    """
    cfg = cfg or load_config()
    conn = None

    try:
        conn = _get_connection(cfg)
        cursor = conn.cursor()
        price_column = _get_source_price_column(cursor)
        backup_changed = _ensure_price_backup_comparison_schema(cursor, price_column)
        daily_changed = _ensure_new_daily_prices_table(cursor, price_column)
        if backup_changed or daily_changed:
            conn.commit()

        changed_rows, inserted_rows = _refresh_new_daily_prices_if_changed(
            cursor,
            price_column,
        )
        conn.commit()
        print(
            "[INFO] Comparacion de precios al abrir: "
            f"diferencias={changed_rows}; reemplazados={inserted_rows}."
        )
        return _load_latest_new_daily_price_items(cursor)
    except Exception as e:
        if conn:
            try:
                conn.rollback()
            except Exception:
                pass
        try:
            conn_info = _sanitize_connection_string(_build_connection_string(cfg))
        except Exception:
            conn_info = "No se pudo construir connection string."
        write_log(
            "ERROR fetch_new_daily_price_items",
            e,
            extra=f"connection={conn_info}",
        )
        raise
    finally:
        if conn:
            conn.close()


# ==============================
# Impresoras (detección)
# ==============================
def fetch_current_special_price_items(cfg: dict = None) -> List[Dict[str, Any]]:
    """
    Lee todos los articulos con precio especial vigente al momento de consulta.
    """
    cfg = cfg or load_config()
    conn = None
    results = []

    sql = """
    ;WITH PromosVigentes AS (
        SELECT
            pl.CLAVEARTICULO AS ARTICULO,
            pl.CLAVE,
            pl.CANTIDAD AS PROMO_CANTIDAD,
            pl.PRECIOUNITARIO,
            pe.CLASE AS CLASE_PROMOCION,
            pe.VIGENCIAINICIO AS PROMO_FECHAINICIO,
            pe.VIGENCIAFINAL AS PROMO_FECHAFINAL,
            ROW_NUMBER() OVER (
                PARTITION BY pl.CLAVEARTICULO
                ORDER BY pl.PRECIOUNITARIO ASC, pl.CLAVE DESC
            ) AS rn
        FROM NPV.dbo.NPVFDPromocionLineas pl
        INNER JOIN NPV.dbo.NPVFDPromocionEncabezado pe
            ON pe.CLAVE = pl.CLAVE
        WHERE COALESCE(pl.CANTIDAD, 0) > 0
          AND pl.CANTIDAD = 1
          AND pl.PRECIOUNITARIO IS NOT NULL
          AND pl.STATUS = 'A'
          AND pe.STATUS = 'A'
          AND pe.VIGENCIAINICIO <= GETDATE()
          AND pe.VIGENCIAFINAL >= GETDATE()
          AND pe.CLASE = 'Z002'
    ),
    PreciosRankeados AS (
        SELECT
            pv.ARTICULO,
            pv.PRECIO,
            pv.PRECIOSVENTA,
            pv.TIPOIMPUESTO1,
            pv.TIPOIMPUESTO2,
            pv.TIPOIMPUESTO3,
            pv.TIPOIMPUESTO4,
            pv.FECHAINICIO,
            ROW_NUMBER() OVER (
                PARTITION BY pv.ARTICULO
                ORDER BY pv.FECHAINICIO DESC
            ) AS rn
        FROM NPV.dbo.NPVFDPreciosVenta pv
        WHERE pv.FECHAINICIO IS NULL OR pv.FECHAINICIO <= GETDATE()
    )
    SELECT
        a.ARTICULO,
        a.DESCRIPCION,
        pr.PRECIOSVENTA AS PRECIO,
        promo.PRECIOUNITARIO *
            (1 + COALESCE(i1.PORCENTAJE, 0) / 100.0) *
            (1 + COALESCE(i2.PORCENTAJE, 0) / 100.0) *
            (1 + COALESCE(i3.PORCENTAJE, 0) / 100.0) *
            (1 + COALESCE(i4.PORCENTAJE, 0) / 100.0) AS PRECIO_ESPECIAL,
        promo.PROMO_CANTIDAD,
        promo.PROMO_FECHAINICIO,
        promo.PROMO_FECHAFINAL,
        promo.CLAVE AS CLAVE_PROMOCION,
        promo.CLASE_PROMOCION,
        existencia.DISPONIBLE,
        e.EQUIVALENTE AS UPC,
        pr.FECHAINICIO
    FROM PromosVigentes promo
    INNER JOIN NPV.dbo.NPVFDArticulos a
        ON a.ARTICULO = promo.ARTICULO
    INNER JOIN PreciosRankeados pr
        ON pr.ARTICULO = promo.ARTICULO
       AND pr.rn = 1
    OUTER APPLY (
        SELECT TOP 1 ae.EQUIVALENTE
        FROM NPV.dbo.NPVFDArticulosEquivalentes ae
        WHERE ae.ARTICULO = promo.ARTICULO
        ORDER BY ae.EQUIVALENTE
    ) e
    OUTER APPLY (
        SELECT MAX(COALESCE(ex.DISPONIBLE, 0)) AS DISPONIBLE
        FROM NPV.dbo.NPVFDExistencias ex
        WHERE ex.ARTICULO LIKE promo.ARTICULO
    ) existencia
    LEFT JOIN NPV.dbo.NPVFDImpuestos i1 ON i1.IMPUESTO = pr.TIPOIMPUESTO1
    LEFT JOIN NPV.dbo.NPVFDImpuestos i2 ON i2.IMPUESTO = pr.TIPOIMPUESTO2
    LEFT JOIN NPV.dbo.NPVFDImpuestos i3 ON i3.IMPUESTO = pr.TIPOIMPUESTO3
    LEFT JOIN NPV.dbo.NPVFDImpuestos i4 ON i4.IMPUESTO = pr.TIPOIMPUESTO4
    WHERE promo.rn = 1
      AND COALESCE(existencia.DISPONIBLE, 0) > 1
    ORDER BY a.DESCRIPCION, a.ARTICULO;
    """

    try:
        conn = _get_connection(cfg)
        cursor = conn.cursor()
        rows = cursor.execute(sql).fetchall()
        for row in rows:
            precio_venta = _format_price(row.PRECIO)
            precio_especial = _format_price(row.PRECIO_ESPECIAL)
            precio_venta_num = _price_to_float(precio_venta)
            precio_especial_num = _price_to_float(precio_especial)
            if (
                not precio_venta
                or not precio_especial
                or precio_venta_num is None
                or precio_especial_num is None
                or precio_especial_num >= precio_venta_num
            ):
                continue

            promo_inicio_txt = _format_date(row.PROMO_FECHAINICIO)
            promo_final_txt = _format_date(row.PROMO_FECHAFINAL)
            vigencia = ""
            if promo_inicio_txt and promo_final_txt:
                vigencia = f"Promocion vigente: {promo_inicio_txt} - {promo_final_txt}"

            results.append({
                "ARTICULO": str(row.ARTICULO).strip(),
                "DESCRIPCION": str(row.DESCRIPCION or "").strip(),
                "PRECIO": precio_venta,
                "PRECIO_ESPECIAL": precio_especial,
                "TIENE_PRECIO_ESPECIAL": True,
                "UPC": _normalize_upc_digits(row.UPC or ""),
                "VIGENCIA": vigencia,
                "DISPONIBLE": row.DISPONIBLE,
                "PROMO_CANTIDAD": row.PROMO_CANTIDAD,
                "PROMO_FECHAINICIO": row.PROMO_FECHAINICIO,
                "PROMO_FECHAFINAL": row.PROMO_FECHAFINAL,
                "CLAVE_PROMOCION": str(row.CLAVE_PROMOCION or "").strip(),
                "CLASE_PROMOCION": str(row.CLASE_PROMOCION or "").strip(),
            })
    except Exception as e:
        try:
            conn_info = _sanitize_connection_string(_build_connection_string(cfg))
        except Exception:
            conn_info = "No se pudo construir connection string."
        write_log(
            "ERROR fetch_current_special_price_items",
            e,
            extra=f"connection={conn_info}",
        )
        raise
    finally:
        if conn:
            conn.close()

    return results


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
    item["PROMO_CANTIDAD"] = None
    item["PROMO_TERMINOS"] = ""
    item["PROMO_TERMINOS2"] = ""
    return item


CODE128_PATTERNS = [
    "212222", "222122", "222221", "121223", "121322", "131222",
    "122213", "122312", "132212", "221213", "221312", "231212",
    "112232", "122132", "122231", "113222", "123122", "123221",
    "223211", "221132", "221231", "213212", "223112", "312131",
    "311222", "321122", "321221", "312212", "322112", "322211",
    "212123", "212321", "232121", "111323", "131123", "131321",
    "112313", "132113", "132311", "211313", "231113", "231311",
    "112133", "112331", "132131", "113123", "113321", "133121",
    "313121", "211331", "231131", "213113", "213311", "213131",
    "311123", "311321", "331121", "312113", "312311", "332111",
    "314111", "221411", "431111", "111224", "111422", "121124",
    "121421", "141122", "141221", "112214", "112412", "122114",
    "122411", "142112", "142211", "241211", "221114", "413111",
    "241112", "134111", "111242", "121142", "121241", "114212",
    "124112", "124211", "411212", "421112", "421211", "212141",
    "214121", "412121", "111143", "111341", "131141", "114113",
    "114311", "411113", "411311", "113141", "114131", "311141",
    "411131", "211412", "211214", "211232", "2331112",
]


def _barcode_value_for_item(item_dict: Dict[str, Any]) -> str:
    upc = _normalize_upc_digits(item_dict.get("UPC"))
    if upc:
        return upc
    return re.sub(r"[^\x20-\x7E]", "", str(item_dict.get("ARTICULO") or "").strip())


def _code128_b_widths(value: str) -> List[int]:
    text = re.sub(r"[^\x20-\x7E]", "", str(value or "").strip())
    if not text:
        return []

    codes = [104]  # Start Code B
    codes.extend(ord(ch) - 32 for ch in text)

    checksum = codes[0]
    for pos, code in enumerate(codes[1:], start=1):
        checksum += code * pos
    codes.append(checksum % 103)
    codes.append(106)  # Stop

    widths: List[int] = []
    for code in codes:
        widths.extend(int(width) for width in CODE128_PATTERNS[code])
    return widths


def draw_code128_barcode_gdi(
    hDC,
    value: str,
    left: int,
    top: int,
    right: int,
    bar_height: int = 58,
    text_height: int = 20,
) -> int:
    widths = _code128_b_widths(value)
    if not widths:
        return 0

    quiet_zone_modules = 10
    total_modules = sum(widths) + quiet_zone_modules * 2
    available_width = max(1, right - left)
    module_width = max(1, available_width // total_modules)
    barcode_width = total_modules * module_width
    start_x = left + max(0, (available_width - barcode_width) // 2)
    x = start_x + quiet_zone_modules * module_width

    hDC.FillSolidRect((left, top, right, top + bar_height + text_height + 6), 0xFFFFFF)

    draw_bar = True
    for width in widths:
        segment_width = width * module_width
        if draw_bar:
            hDC.FillSolidRect((x, top, x + segment_width, top + bar_height), 0x000000)
        x += segment_width
        draw_bar = not draw_bar

    font_text = win32ui.CreateFont({"name": "Arial", "height": text_height, "weight": win32con.FW_NORMAL})
    hDC.SelectObject(font_text)
    hDC.SetTextColor(0x000000)
    hDC.SetBkMode(win32con.TRANSPARENT)
    hDC.DrawText(
        str(value),
        (left, top + bar_height + 2, right, top + bar_height + text_height + 4),
        win32con.DT_CENTER | win32con.DT_SINGLELINE,
    )
    return bar_height + text_height + 6


def _print_layout_section(config: dict, section: str) -> dict:
    layout = (config or {}).get("layout_impresion") or {}
    section_cfg = layout.get(section) or {}
    default_cfg = DEFAULT_PRINT_CONFIG["layout_impresion"].get(section, {})
    merged = _clone_json(section_cfg) if isinstance(section_cfg, dict) else {}
    _deep_merge(merged, default_cfg)
    return merged


def _param_int(params: dict, key: str, default: int, minimum: int = None) -> int:
    try:
        value = int(params.get(key, default))
    except Exception:
        value = default
    if minimum is not None:
        value = max(minimum, value)
    return value


def _param_bool(params: dict, key: str, default: bool) -> bool:
    value = params.get(key, default)
    if isinstance(value, str):
        return value.strip().lower() not in ("0", "false", "no", "off")
    return bool(value)


def _estimate_label_height_dots(
    item_dict: Dict[str, Any],
    config: dict,
    printable_width: int,
    section: str,
    auto_individual_print: bool = False,
) -> int:
    """
    Calcula el alto real que ocupa la etiqueta.

    En etiquetas normales y angostas el contenido comienza en el borde
    imprimible; solo se conserva un margen pequeno tras el ultimo elemento.
    """
    paper_params = _print_layout_section(config, "papel")
    minimum_height = _param_int(paper_params, "minimum_height", 180, 1)
    bottom_margin = _param_int(paper_params, "bottom_margin", 12, 0)

    if (
        section == "normal"
        and item_dict.get("PRECIO_ESPECIAL")
        and Image is not None
        and ImageWin is not None
        and os.path.exists(get_promo_background_path())
    ):
        promo_height = math.ceil(printable_width * PROMO_TEMPLATE_SIZE[1] / PROMO_TEMPLATE_SIZE[0])
        return max(minimum_height, promo_height)

    params = _print_layout_section(config, section)
    y_pos = _param_int(params, "top_y", 0, 0)
    content_bottom = 0

    if _param_bool(params, "draw_top_divider", False):
        divider_params = _print_layout_section(config, "lineas_division")
        if _param_bool(divider_params, "enabled", True):
            content_bottom = (
                _param_int(divider_params, "top_y", 6, 0)
                + _param_int(divider_params, "thickness", 3, 1)
            )

    if section == "angosta" and auto_individual_print:
        desc_advance = _param_int(params, "descripcion_advance_auto", 68, 0)
    else:
        desc_advance = _param_int(
            params,
            "descripcion_advance",
            70 if section == "normal" else 58,
            0,
        )
    desc_rect_height = _param_int(params, "descripcion_rect_height", 200, 1)
    content_bottom = max(content_bottom, y_pos + desc_rect_height)
    y_pos += desc_advance

    font_price_height = _param_int(
        params,
        "precio_font_height",
        92 if section == "normal" else 70,
        1,
    )
    precio_especial = bool(item_dict.get("PRECIO_ESPECIAL"))
    if precio_especial:
        # Los tres renglones ANTES/AHORA/AHORRAS avanzan 125 puntos en el
        # formato normal y 100 en el angosto.
        special_height = 125 if section == "normal" else 100
        content_bottom = max(content_bottom, y_pos + special_height)
        y_pos += special_height
    else:
        price_rect_height = _param_int(
            params,
            "precio_rect_height",
            118 if section == "normal" else 92,
            1,
        )
        content_bottom = max(content_bottom, y_pos + price_rect_height)
        y_pos += font_price_height

    y_pos += _param_int((config or {}).get("espaciados") or {}, "espacio_abajo_precio", 0, 0)

    footer_rect_height = _param_int(params, "footer_rect_height", 30, 1)
    footer_advance = _param_int(
        params,
        "footer_advance",
        25 if section == "normal" else 22,
        0,
    )
    footer_items = [
        item_dict.get("VIGENCIA", ""),
        f"{item_dict.get('ARTICULO', '')}   {item_dict.get('UPC', '')}",
    ]
    for text in footer_items:
        if not str(text or "").strip():
            continue
        content_bottom = max(content_bottom, y_pos + footer_rect_height)
        y_pos += footer_advance

    if not precio_especial and _barcode_value_for_item(item_dict):
        y_pos += 4
        barcode_height = _param_int(
            params,
            "barcode_height",
            58 if section == "normal" else 46,
            1,
        )
        barcode_text_height = _param_int(
            params,
            "barcode_text_height",
            20 if section == "normal" else 16,
            1,
        )
        content_bottom = max(content_bottom, y_pos + barcode_height + barcode_text_height + 6)

    return max(minimum_height, content_bottom + bottom_margin)


def _create_label_printer_dc(
    printer_name: str,
    item_dict: Dict[str, Any],
    config: dict,
    section: str,
    auto_individual_print: bool = False,
):
    """
    Crea un DC temporal con el alto justo de la etiqueta.

    El DEVMODE se aplica solo al trabajo actual. No cambia el formulario ni
    las preferencias permanentes configuradas en Windows.
    """
    fallback_dc = win32ui.CreateDC()
    fallback_dc.CreatePrinterDC(printer_name)

    paper_params = _print_layout_section(config, "papel")
    if not _param_bool(paper_params, "auto_height", True):
        return fallback_dc

    printable_width = fallback_dc.GetDeviceCaps(win32con.HORZRES)
    dpi_x = max(1, fallback_dc.GetDeviceCaps(win32con.LOGPIXELSX))
    dpi_y = max(1, fallback_dc.GetDeviceCaps(win32con.LOGPIXELSY))
    desired_height = _estimate_label_height_dots(
        item_dict,
        config,
        printable_width,
        section,
        auto_individual_print=auto_individual_print,
    )

    printer_handle = None
    try:
        printer_handle = win32print.OpenPrinter(printer_name)
        printer_info = win32print.GetPrinter(printer_handle, 2)
        devmode = printer_info.get("pDevMode")
        if devmode is None:
            raise RuntimeError("El controlador no devolvio parametros DEVMODE.")

        # PaperWidth/PaperLength usan decimas de milimetro. math.ceil evita
        # perder un punto por redondeo en controladores termicos de 203 dpi.
        paper_width = max(1, math.ceil(printable_width * 254 / dpi_x))
        paper_length = max(1, math.ceil(desired_height * 254 / dpi_y))
        devmode.Fields |= (
            win32con.DM_PAPERSIZE
            | win32con.DM_PAPERWIDTH
            | win32con.DM_PAPERLENGTH
        )
        devmode.PaperSize = 0
        devmode.PaperWidth = paper_width
        devmode.PaperLength = paper_length

        raw_dc = win32gui.CreateDC("WINSPOOL", printer_name, devmode)
        sized_dc = win32ui.CreateDCFromHandle(raw_dc)
        actual_height = sized_dc.GetDeviceCaps(win32con.VERTRES)
        if actual_height < desired_height:
            sized_dc.DeleteDC()
            raise RuntimeError(
                f"El controlador acepto solo {actual_height} de {desired_height} puntos."
            )

        fallback_dc.DeleteDC()
        print(
            "[INFO] Alto automatico de etiqueta: "
            f"{actual_height} puntos ({paper_length / 10:.1f} mm)."
        )
        return sized_dc
    except Exception as exc:
        print(
            "[WARN] La impresora no acepto el alto automatico; "
            f"se usara su formulario actual: {exc}"
        )
        return fallback_dc
    finally:
        if printer_handle is not None:
            try:
                win32print.ClosePrinter(printer_handle)
            except Exception:
                pass


def draw_print_area_divider(hDC, printable_width: int, y: int, config: dict = None, margin: int = None, thickness: int = None) -> None:
    params = _print_layout_section(config, "lineas_division")
    if not _param_bool(params, "enabled", True):
        return
    if margin is None:
        margin = _param_int(params, "margin", 10, 0)
    if thickness is None:
        thickness = _param_int(params, "thickness", 3, 1)
    y = max(0, int(y))
    left = max(0, int(margin))
    right = max(left + 1, int(printable_width) - left)
    try:
        hDC.FillSolidRect((left, y, right, y + thickness), 0x000000)
    except Exception as e:
        print(f"[WARN] No se pudo dibujar linea divisoria: {e}")


def draw_print_area_dividers(hDC, printable_width: int, bottom_y: int = None, config: dict = None, top_y: int = None) -> None:
    params = _print_layout_section(config, "lineas_division")
    if top_y is None:
        top_y = _param_int(params, "top_y", 6, 0)
    draw_print_area_divider(hDC, printable_width, top_y, config)


def draw_promo_label_gdi(hDC, item_dict: Dict[str, Any], printable_width: int, config: dict = None) -> bool:
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
    
    ahorro_val = max(0, (_price_to_float(precio) or 0) - (_price_to_float(precio_especial) or 0))
    texto_ahorro = _promo_price_text(ahorro_val)
    texto_ahora = _promo_price_text(precio_especial)
    texto_antes = _promo_price_text(precio)
    promo_terminos = item_dict.get("PROMO_TERMINOS", "")
    promo_terminos2 = item_dict.get("PROMO_TERMINOS2", "")

    hDC.SetTextColor(0x111111) # Color oscuro para el texto
    hDC.SetBkColor(0xFFFFFF)
    hDC.SetBkMode(win32con.OPAQUE)

    font_savings = win32ui.CreateFont({"name": "Arial", "height": font_height(100), "weight": win32con.FW_BOLD})
    hDC.SelectObject(font_savings)
    hDC.DrawText(texto_ahorro, rect(1010, 65, 1480, 205), win32con.DT_CENTER | win32con.DT_SINGLELINE | win32con.DT_VCENTER)

    font_now = win32ui.CreateFont({"name": "Arial", "height": font_height(166), "weight": win32con.FW_BOLD})
    hDC.SelectObject(font_now)
    hDC.DrawText(texto_ahora, rect(340, 315, 1480, 500), win32con.DT_CENTER | win32con.DT_SINGLELINE | win32con.DT_VCENTER)

    font_before = win32ui.CreateFont({"name": "Arial", "height": font_height(66), "weight": win32con.FW_BOLD})
    hDC.SelectObject(font_before)
    hDC.DrawText(texto_antes, rect(1010, 715, 1410, 815), win32con.DT_CENTER | win32con.DT_SINGLELINE | win32con.DT_VCENTER)

    hDC.SetBkMode(win32con.TRANSPARENT)

    desc = (item_dict.get("DESCRIPCION") or "").upper()
    if desc:
        font_desc = win32ui.CreateFont({"name": "Arial", "height": font_height(PROMO_DESCRIPTION_FONT_SIZE), "weight": win32con.FW_BOLD})
        hDC.SelectObject(font_desc)
        hDC.DrawText(desc, rect(80, 528, 1505, 605), win32con.DT_CENTER | win32con.DT_SINGLELINE | win32con.DT_VCENTER)

    if promo_terminos:
        font_terms = win32ui.CreateFont({"name": "Arial", "height": font_height(28), "weight": win32con.FW_BOLD})
        hDC.SelectObject(font_terms)
        hDC.DrawText(promo_terminos, rect(95, 865, 1490, 900), win32con.DT_LEFT | win32con.DT_SINGLELINE | win32con.DT_VCENTER)
        if promo_terminos2:
            hDC.DrawText(promo_terminos2, rect(95, 900, 1490, 960), win32con.DT_LEFT | win32con.DT_WORDBREAK)

    draw_print_area_dividers(hDC, printable_width, config=config)
    return True


def print_label_gdi(item_dict: Dict[str, Any], config: dict, copies: int = 1, show_errors: bool = True):
    """
    Imprime la etiqueta final usando GDI, basado en el área de impresión real.
    """
    printer_name = _get_selected_printer_name(config)
    if not printer_name:
        raise RuntimeError("No se pudo determinar la impresora a usar.")

    try:
        hDC = _create_label_printer_dc(
            printer_name,
            item_dict,
            config,
            "normal",
        )
        
        # Obtener el área de impresión real en píxeles
        printable_width = hDC.GetDeviceCaps(win32con.HORZRES)
        hDC.SetMapMode(win32con.MM_TEXT)
        hDC.SetBkMode(win32con.TRANSPARENT)
        normal_params = _print_layout_section(config, "normal")
        horizontal_margin = _param_int(normal_params, "horizontal_margin", 10, 0)

        for i in range(copies):
            hDC.StartDoc(f"Etiqueta NPV ({i+1}/{copies})")
            hDC.StartPage()

            if draw_promo_label_gdi(hDC, item_dict, printable_width, config):
                hDC.EndPage()
                hDC.EndDoc()
                continue

            if _param_bool(normal_params, "draw_top_divider", False):
                divider_params = _print_layout_section(config, "lineas_division")
                draw_print_area_divider(hDC, printable_width, _param_int(divider_params, "top_y", 6, 0), config)
            y_pos = _param_int(normal_params, "top_y", 0, 0)
            normal_price_label = not bool(item_dict.get("PRECIO_ESPECIAL"))

            # --- 2. Dibujar Descripción ---
            try:
                desc_font_height = _param_int(normal_params, "descripcion_font_height", 46, 1)
                desc_rect_height = _param_int(normal_params, "descripcion_rect_height", 200, 1)
                desc_advance = _param_int(normal_params, "descripcion_advance", 70, 0)
                font_desc = win32ui.CreateFont({"name": "Arial", "height": desc_font_height, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_desc)
                desc = (item_dict.get("DESCRIPCION") or "").upper()
                print(f"Dibujando descripción: {desc}") # <-- DEBUG
                rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + desc_rect_height)
                hDC.DrawText(desc, rect, win32con.DT_CENTER | win32con.DT_WORDBREAK)
                y_pos += desc_advance
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar la descripción: {e}")

            # --- 3. Dibujar Precio ---
            try:
                font_price_height = _param_int(normal_params, "precio_font_height", 92, 1)
                font_price = win32ui.CreateFont({"name": "Arial", "height": font_price_height, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_price)
                precio = _format_currency(item_dict.get("PRECIO", "$0.00"))
                precio_especial = _format_currency(item_dict.get("PRECIO_ESPECIAL", "")) if item_dict.get("PRECIO_ESPECIAL") else ""
                if precio_especial:
                    precio_base_num = _price_to_float(precio) or 0
                    precio_especial_num = _price_to_float(precio_especial) or 0
                    ahorro = max(0, precio_base_num - precio_especial_num)
                    font_label = win32ui.CreateFont({"name": "Arial", "height": _param_int(normal_params, "precio_especial_label_font_height", 24, 1), "weight": win32con.FW_BOLD})
                    font_price = win32ui.CreateFont({"name": "Arial", "height": _param_int(normal_params, "precio_especial_price_font_height", 54, 1), "weight": win32con.FW_BOLD})
                    font_save = win32ui.CreateFont({"name": "Arial", "height": _param_int(normal_params, "precio_especial_save_font_height", 22, 1), "weight": win32con.FW_BOLD})
                    hDC.SelectObject(font_label)
                    hDC.DrawText(f"ANTES: {precio}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 32), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 35
                    hDC.SelectObject(font_price)
                    hDC.DrawText(f"AHORA: {precio_especial}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 66), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 60
                    hDC.SelectObject(font_save)
                    hDC.DrawText(f"AHORRAS: {_format_price(ahorro)}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 30), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 30 - font_price_height
                else:
                    rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + _param_int(normal_params, "precio_rect_height", 118, 1))
                    hDC.DrawText(precio, rect, win32con.DT_CENTER | win32con.DT_SINGLELINE)
                
                # Se usa la configuración de espaciado. El avance es altura de fuente + espacio.
                espaciados = config.get("espaciados", {})
                espacio_abajo = espaciados.get("espacio_abajo_precio", 0)
                y_pos += font_price_height + espacio_abajo
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar el precio: {e}")

            # --- 4. Dibujar Pie de página ---
            try:
                footer_rect_height = _param_int(normal_params, "footer_rect_height", 30, 1)
                footer_advance = _param_int(normal_params, "footer_advance", 25, 0)
                font_footer = win32ui.CreateFont({"name": "Arial", "height": _param_int(normal_params, "footer_font_height", 24, 1), "weight": win32con.FW_NORMAL})
                hDC.SelectObject(font_footer)
                
                footer_items = [
                    item_dict.get("VIGENCIA", ""),
                    f"{item_dict.get('ARTICULO', '')}   {item_dict.get('UPC', '')}"
                ]

                for text in footer_items:
                    if not text.strip(): continue
                    rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + footer_rect_height)
                    hDC.DrawText(text, rect, win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += footer_advance

                if normal_price_label:
                    barcode_value = _barcode_value_for_item(item_dict)
                    if barcode_value:
                        y_pos += 4
                        barcode_margin = _param_int(normal_params, "barcode_margin", 35, 0)
                        y_pos += draw_code128_barcode_gdi(
                            hDC,
                            barcode_value,
                            barcode_margin,
                            y_pos,
                            printable_width - barcode_margin,
                            bar_height=_param_int(normal_params, "barcode_height", 58, 1),
                            text_height=_param_int(normal_params, "barcode_text_height", 20, 1),
                        )
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
        hDC = _create_label_printer_dc(
            printer_name,
            item_dict,
            config,
            "angosta",
            auto_individual_print=auto_individual_print,
        )
        
        printable_width = hDC.GetDeviceCaps(win32con.HORZRES)
        hDC.SetMapMode(win32con.MM_TEXT)
        hDC.SetBkMode(win32con.TRANSPARENT)
        small_params = _print_layout_section(config, "angosta")

        for i in range(copies):
            hDC.StartDoc(f"Etiqueta NPV Individual ({i+1}/{copies})")
            hDC.StartPage()

            if _param_bool(small_params, "draw_top_divider", False):
                divider_params = _print_layout_section(config, "lineas_division")
                draw_print_area_divider(hDC, printable_width, _param_int(divider_params, "top_y", 6, 0), config)
            y_pos = _param_int(small_params, "top_y", 0, 0)
            normal_price_label = not bool(item_dict.get("PRECIO_ESPECIAL"))
            horizontal_margin = _param_int(small_params, "horizontal_margin", 48, 0)

            # --- 2. Dibujar Descripción ---
            try:
                desc_font_height = _param_int(small_params, "descripcion_font_height_auto", 30, 1) if auto_individual_print else _param_int(small_params, "descripcion_font_height", 34, 1)
                desc_advance = _param_int(small_params, "descripcion_advance_auto", 68, 0) if auto_individual_print else _param_int(small_params, "descripcion_advance", 58, 0)
                desc_rect_height = _param_int(small_params, "descripcion_rect_height", 200, 1)
                font_desc = win32ui.CreateFont({"name": "Arial", "height": desc_font_height, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_desc)
                desc = (item_dict.get("DESCRIPCION") or "").upper()
                rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + desc_rect_height)
                hDC.DrawText(desc, rect, win32con.DT_CENTER | win32con.DT_WORDBREAK)
                y_pos += desc_advance
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar la descripción (individual): {e}")

            # --- 3. Dibujar Precio ---
            try:
                font_price_height = _param_int(small_params, "precio_font_height", 70, 1)
                font_price = win32ui.CreateFont({"name": "Arial", "height": font_price_height, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_price)
                precio = _format_currency(item_dict.get("PRECIO", "$0.00"))
                precio_especial = item_dict.get("PRECIO_ESPECIAL", "")
                if precio_especial:
                    precio_base_num = _price_to_float(precio) or 0
                    precio_especial_num = _price_to_float(precio_especial) or 0
                    ahorro = max(0, precio_base_num - precio_especial_num)
                    font_label = win32ui.CreateFont({"name": "Arial", "height": _param_int(small_params, "precio_especial_label_font_height", 20, 1), "weight": win32con.FW_BOLD})
                    font_price = win32ui.CreateFont({"name": "Arial", "height": _param_int(small_params, "precio_especial_price_font_height", 43, 1), "weight": win32con.FW_BOLD})
                    font_save = win32ui.CreateFont({"name": "Arial", "height": _param_int(small_params, "precio_especial_save_font_height", 18, 1), "weight": win32con.FW_BOLD})
                    hDC.SelectObject(font_label)
                    hDC.DrawText(f"ANTES: {precio}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 27), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 25
                    hDC.SelectObject(font_price)
                    hDC.DrawText(f"AHORA: {precio_especial}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 54), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 50
                    hDC.SelectObject(font_save)
                    hDC.DrawText(f"AHORRAS: {_format_price(ahorro)}", (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 25), win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += 25 - font_price_height
                else:
                    rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + _param_int(small_params, "precio_rect_height", 92, 1))
                    hDC.DrawText(precio, rect, win32con.DT_CENTER | win32con.DT_SINGLELINE)
                
                espaciados = config.get("espaciados", {})
                espacio_abajo = espaciados.get("espacio_abajo_precio", 0)
                y_pos += font_price_height + espacio_abajo
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar el precio (individual): {e}")

            # --- 4. Dibujar Pie de página ---
            try:
                footer_rect_height = _param_int(small_params, "footer_rect_height", 30, 1)
                footer_advance = _param_int(small_params, "footer_advance", 22, 0)
                font_footer = win32ui.CreateFont({"name": "Arial", "height": _param_int(small_params, "footer_font_height", 19, 1), "weight": win32con.FW_NORMAL})
                hDC.SelectObject(font_footer)
                
                footer_items = [
                    item_dict.get("VIGENCIA", ""),
                    f"{item_dict.get('ARTICULO', '')}   {item_dict.get('UPC', '')}"
                ]

                for text in footer_items:
                    if not text.strip(): continue
                    rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + footer_rect_height)
                    hDC.DrawText(text, rect, win32con.DT_CENTER | win32con.DT_SINGLELINE)
                    y_pos += footer_advance

                if normal_price_label:
                    barcode_value = _barcode_value_for_item(item_dict)
                    if barcode_value:
                        y_pos += 4
                        y_pos += draw_code128_barcode_gdi(
                            hDC,
                            barcode_value,
                            horizontal_margin,
                            y_pos,
                            printable_width - horizontal_margin,
                            bar_height=_param_int(small_params, "barcode_height", 46, 1),
                            text_height=_param_int(small_params, "barcode_text_height", 16, 1),
                        )
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
    tiene_precio_especial = bool(item_dict.get("TIENE_PRECIO_ESPECIAL")) or bool(precio_especial)
    articulo    = (item_dict.get("ARTICULO") or "").strip()
    upc         = _normalize_upc_digits(item_dict.get("UPC"))

    if not precio:
        # Precio inválido
        precio = ""
    if not tiene_precio_especial:
        precio_especial = ""

    # Vigencia: usa la que viene del item si existe, si no, la de por defecto.
    vigencia = item_dict.get("VIGENCIA")
    if not vigencia:
        vigencia = resolve_vigencia_template(config.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"]))

    promo_inicio = item_dict.get("PROMO_FECHAINICIO")
    promo_final = item_dict.get("PROMO_FECHAFINAL")
    promo_cantidad = item_dict.get("PROMO_CANTIDAD")
    promo_inicio_txt = _format_date(promo_inicio)
    promo_final_txt = _format_date(promo_final)
    promo_terminos = ""
    if promo_inicio_txt and promo_final_txt:
        qty = _quantity_to_float(promo_cantidad)
        if qty and qty > 1:
            promo_terminos = (
                f"Terminos y Condiciones: compra minima {_format_quantity(qty)} pzas, "
                f"valido del {promo_inicio_txt} al {promo_final_txt}"
            )
        else:
            promo_terminos = f"Terminos y Condiciones: valido del {promo_inicio_txt} al {promo_final_txt}"
    promo_terminos2 = "No acumulable. Sujeto a cambios sin aviso."

    return {
        "DESCRIPCION": descripcion,
        "PRECIO": precio,
        "PRECIO_ESPECIAL": precio_especial,
        "TIENE_PRECIO_ESPECIAL": tiene_precio_especial,
        "PROMO_CANTIDAD": promo_cantidad,
        "PROMO_FECHAINICIO": promo_inicio,
        "PROMO_FECHAFINAL": promo_final,
        "PROMO_TERMINOS": promo_terminos,
        "PROMO_TERMINOS2": promo_terminos2,
        "DISPONIBLE": item_dict.get("DISPONIBLE"),
        "ARTICULO": articulo,
        "UPC": upc,
        "VIGENCIA": vigencia
    }


def _tree_sort_key(value):
    text = str(value or "").strip()
    date_match = re.search(r"\b(\d{1,2})/(\d{1,2})/(\d{4})\b", text)
    if date_match:
        day, month, year = (int(part) for part in date_match.groups())
        return (0, year, month, day, text.casefold())

    number_text = text.replace("$", "").replace(",", "").strip()
    if re.fullmatch(r"-?\d+(?:\.\d+)?", number_text):
        try:
            return (1, float(number_text), text.casefold())
        except Exception:
            pass

    return (2, text.casefold())


def enable_treeview_sorting(tree, columns, labels):
    sort_state = {"column": None, "descending": False}

    def refresh_headings():
        for column in columns:
            suffix = ""
            if sort_state["column"] == column:
                suffix = " v" if sort_state["descending"] else " ^"
            tree.heading(
                column,
                text=f"{labels.get(column, column)}{suffix}",
                command=lambda col=column: sort_by_column(col),
            )

    def sort_by_column(column):
        descending = sort_state["column"] == column and not sort_state["descending"]
        sort_state["column"] = column
        sort_state["descending"] = descending

        rows = [
            (tree.set(item_id, column), item_id)
            for item_id in tree.get_children("")
        ]
        rows.sort(key=lambda row: _tree_sort_key(row[0]), reverse=descending)

        for index, (_, item_id) in enumerate(rows):
            tree.move(item_id, "", index)

        refresh_headings()

    refresh_headings()


def get_window_work_area(window) -> tuple:
    if os.name == "nt":
        try:
            import ctypes

            class RECT(ctypes.Structure):
                _fields_ = [
                    ("left", ctypes.c_long),
                    ("top", ctypes.c_long),
                    ("right", ctypes.c_long),
                    ("bottom", ctypes.c_long),
                ]

            rect = RECT()
            if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(rect), 0):
                return rect.left, rect.top, rect.right - rect.left, rect.bottom - rect.top
        except Exception:
            pass

    return 0, 0, window.winfo_screenwidth(), window.winfo_screenheight()


def fit_window_to_screen(window, desired_width: int, desired_height: int, min_width: int, min_height: int, margin: int = 24):
    window.update_idletasks()
    work_left, work_top, work_width, work_height = get_window_work_area(window)
    max_width = max(360, work_width - (margin * 2))
    max_height = max(320, work_height - (margin * 2))
    safe_min_width = min(min_width, max_width)
    safe_min_height = min(min_height, max_height)
    width = min(max(desired_width, safe_min_width), max_width)
    height = min(max(desired_height, safe_min_height), max_height)
    x = work_left + max(margin, int((work_width - width) / 2))
    y = work_top + max(margin, int((work_height - height) / 2))
    window.minsize(int(safe_min_width), int(safe_min_height))
    window.maxsize(int(work_width), int(work_height))
    window.geometry(f"{int(width)}x{int(height)}+{int(x)}+{int(y)}")


CHECKBOX_UNCHECKED = "[ ]"
CHECKBOX_CHECKED = "[X]"
CHECKBOX_SELECTED_TAG = "checked"


class CheckboxTreeSelectionMixin:
    def _setup_checkbox_selection(self):
        self._checked_tree_items = set()
        self.tree.bind("<Button-1>", self._on_checkbox_tree_click)
        self.tree.tag_configure(CHECKBOX_SELECTED_TAG, background="#E8F2FF")

    def _set_checkbox_checked(self, item_id: str, checked: bool):
        if not item_id:
            return
        values = list(self.tree.item(item_id, "values"))
        if not values:
            return
        values[0] = CHECKBOX_CHECKED if checked else CHECKBOX_UNCHECKED
        self.tree.item(
            item_id,
            values=values,
            tags=(CHECKBOX_SELECTED_TAG,) if checked else (),
        )
        if checked:
            self._checked_tree_items.add(item_id)
        else:
            self._checked_tree_items.discard(item_id)

    def _on_checkbox_tree_click(self, event):
        if self.tree.identify_region(event.x, event.y) not in ("cell", "tree"):
            return None
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return None
        self._set_checkbox_checked(item_id, item_id not in self._checked_tree_items)
        return "break"

    def _set_all_checkbox_items(self, checked: bool):
        for item_id in self.tree.get_children(""):
            self._set_checkbox_checked(item_id, checked)

    def _checked_item_indices(self) -> List[int]:
        checked = set(getattr(self, "_checked_tree_items", set()))
        return [
            self._tree_item_to_index[item_id]
            for item_id in self.tree.get_children("")
            if item_id in checked and item_id in self._tree_item_to_index
        ]


# ==============================
# UI (Tkinter)
# ==============================

class ResultSelectionWindow(tk.Toplevel):
    def __init__(self, parent, results, callback):
        super().__init__(parent)
        self.title("Seleccionar Artículo")
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
        enable_treeview_sorting(
            self.tree,
            ("articulo", "descripcion", "upc"),
            {"articulo": "Articulo", "descripcion": "Descripcion", "upc": "UPC"},
        )
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
        fit_window_to_screen(self, desired_width=800, desired_height=450, min_width=640, min_height=360)
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


class NewDailyPricesSelectionWindow(CheckboxTreeSelectionMixin, tk.Toplevel):
    def __init__(self, parent, items: List[Dict[str, Any]]):
        super().__init__(parent)
        self.title("Precios Nuevos")
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#F0F0F0")

        self.items = items
        self.selected_items = None
        self._tree_item_to_index = {}

        ttk.Label(
            self,
            text="Precios nuevos disponibles",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", padx=15, pady=(14, 6))

        tree_frame = ttk.Frame(self)
        tree_frame.pack(expand=True, fill="both", padx=15, pady=(0, 10))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        columns = ("seleccion", "articulo", "descripcion", "precio", "vigencia", "upc")
        sortable_columns = ("articulo", "descripcion", "precio", "vigencia", "upc")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="none")
        self.tree.heading("seleccion", text="")
        self.tree.heading("articulo", text="Articulo")
        self.tree.heading("descripcion", text="Descripcion")
        self.tree.heading("precio", text="Precio")
        self.tree.heading("vigencia", text="Vigencia")
        self.tree.heading("upc", text="UPC")
        enable_treeview_sorting(
            self.tree,
            sortable_columns,
            {
                "articulo": "Articulo",
                "descripcion": "Descripcion",
                "precio": "Precio",
                "vigencia": "Vigencia",
                "upc": "UPC",
            },
        )
        self.tree.column("seleccion", width=48, anchor="center", stretch=False)
        self.tree.column("articulo", width=120, anchor="w", stretch=False)
        self.tree.column("descripcion", width=360, anchor="w")
        self.tree.column("precio", width=90, anchor="e", stretch=False)
        self.tree.column("vigencia", width=220, anchor="w", stretch=False)
        self.tree.column("upc", width=140, anchor="w", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=x_scrollbar.set)
        self._setup_checkbox_selection()

        for index, item in enumerate(self.items):
            iid = self.tree.insert(
                "",
                "end",
                values=(
                    CHECKBOX_UNCHECKED,
                    item.get("ARTICULO", ""),
                    item.get("DESCRIPCION", ""),
                    _format_currency(item.get("PRECIO", "")),
                    item.get("VIGENCIA", ""),
                    item.get("UPC", ""),
                ),
            )
            self._tree_item_to_index[iid] = index

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=15, pady=(0, 15))

        ttk.Button(btn_frame, text="Seleccionar todos", command=self.select_all).pack(side="left")
        ttk.Button(btn_frame, text="Limpiar seleccion", command=self.clear_selection).pack(side="left", padx=(10, 0))
        ttk.Button(btn_frame, text="Cancelar", command=self.on_close).pack(side="right", padx=(10, 0))
        ttk.Button(btn_frame, text="Imprimir seleccionados", command=self.on_print_selected).pack(side="right")

        self.bind("<Escape>", lambda event: self.on_close())
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        fit_window_to_screen(self, desired_width=980, desired_height=520, min_width=720, min_height=400)
        self.wait_window()

    def select_all(self):
        self._set_all_checkbox_items(True)

    def clear_selection(self):
        self._set_all_checkbox_items(False)

    def on_print_selected(self):
        selected_indices = self._checked_item_indices()
        if not selected_indices:
            messagebox.showwarning(APP_TITLE, "Selecciona al menos un articulo para imprimir.", parent=self)
            return
        self.selected_items = [self.items[index] for index in selected_indices]
        self.destroy()

    def on_close(self):
        self.selected_items = None
        self.destroy()


class CurrentSpecialPricesSelectionWindow(CheckboxTreeSelectionMixin, tk.Toplevel):
    def __init__(self, parent, items: List[Dict[str, Any]]):
        super().__init__(parent)
        self.title("Precios Especiales Vigentes")
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#F0F0F0")

        self.items = items
        self.selected_items = None
        self._tree_item_to_index = {}

        ttk.Label(
            self,
            text="Precios especiales vigentes",
            font=("Segoe UI", 12, "bold"),
        ).pack(anchor="w", padx=15, pady=(14, 6))

        tree_frame = ttk.Frame(self)
        tree_frame.pack(expand=True, fill="both", padx=15, pady=(0, 10))
        tree_frame.columnconfigure(0, weight=1)
        tree_frame.rowconfigure(0, weight=1)

        columns = ("seleccion", "articulo", "descripcion", "precio", "especial", "disponible", "vigencia", "upc")
        sortable_columns = ("articulo", "descripcion", "precio", "especial", "disponible", "vigencia", "upc")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="none")
        self.tree.heading("seleccion", text="")
        self.tree.heading("articulo", text="Articulo")
        self.tree.heading("descripcion", text="Descripcion")
        self.tree.heading("precio", text="Precio")
        self.tree.heading("especial", text="Precio especial")
        self.tree.heading("disponible", text="Disponible")
        self.tree.heading("vigencia", text="Vigencia")
        self.tree.heading("upc", text="UPC")
        enable_treeview_sorting(
            self.tree,
            sortable_columns,
            {
                "articulo": "Articulo",
                "descripcion": "Descripcion",
                "precio": "Precio",
                "especial": "Precio especial",
                "disponible": "Disponible",
                "vigencia": "Vigencia",
                "upc": "UPC",
            },
        )
        self.tree.column("seleccion", width=48, anchor="center", stretch=False)
        self.tree.column("articulo", width=120, anchor="w", stretch=False)
        self.tree.column("descripcion", width=320, anchor="w")
        self.tree.column("precio", width=90, anchor="e", stretch=False)
        self.tree.column("especial", width=110, anchor="e", stretch=False)
        self.tree.column("disponible", width=90, anchor="e", stretch=False)
        self.tree.column("vigencia", width=230, anchor="w", stretch=False)
        self.tree.column("upc", width=140, anchor="w", stretch=False)
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        x_scrollbar = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        x_scrollbar.grid(row=1, column=0, sticky="ew")
        self.tree.configure(yscrollcommand=scrollbar.set, xscrollcommand=x_scrollbar.set)
        self._setup_checkbox_selection()

        for index, item in enumerate(self.items):
            iid = self.tree.insert(
                "",
                "end",
                values=(
                    CHECKBOX_UNCHECKED,
                    item.get("ARTICULO", ""),
                    item.get("DESCRIPCION", ""),
                    _format_currency(item.get("PRECIO", "")),
                    _format_currency(item.get("PRECIO_ESPECIAL", "")),
                    _format_quantity(item.get("DISPONIBLE")),
                    item.get("VIGENCIA", ""),
                    item.get("UPC", ""),
                ),
            )
            self._tree_item_to_index[iid] = index

        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill="x", padx=15, pady=(0, 15))

        ttk.Button(btn_frame, text="Seleccionar todos", command=self.select_all).pack(side="left")
        ttk.Button(btn_frame, text="Limpiar seleccion", command=self.clear_selection).pack(side="left", padx=(10, 0))
        ttk.Button(btn_frame, text="Cancelar", command=self.on_close).pack(side="right", padx=(10, 0))
        ttk.Button(btn_frame, text="Imprimir seleccionados", command=self.on_print_selected).pack(side="right")

        self.bind("<Escape>", lambda event: self.on_close())
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        fit_window_to_screen(self, desired_width=1100, desired_height=560, min_width=760, min_height=420)
        self.wait_window()

    def select_all(self):
        self._set_all_checkbox_items(True)

    def clear_selection(self):
        self._set_all_checkbox_items(False)

    def on_print_selected(self):
        selected_indices = self._checked_item_indices()
        if not selected_indices:
            messagebox.showwarning(APP_TITLE, "Selecciona al menos un articulo para imprimir.", parent=self)
            return
        self.selected_items = [self.items[index] for index in selected_indices]
        self.destroy()

    def on_close(self):
        self.selected_items = None
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
            # La vista previa refleja el formato compacto sin cabecera ni
            # linea de corte superior.
            label_height = 300 if individual else 360
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

    def _draw_code128_barcode_canvas(
        self,
        value: str,
        left: float,
        top: float,
        right: float,
        bar_height: float,
        text_size: int,
    ) -> float:
        widths = _code128_b_widths(value)
        if not widths:
            return 0

        quiet_zone_modules = 10
        total_modules = sum(widths) + quiet_zone_modules * 2
        available_width = max(1, right - left)
        module_width = available_width / total_modules
        barcode_width = total_modules * module_width
        x = left + max(0, (available_width - barcode_width) / 2) + quiet_zone_modules * module_width

        self.canvas.create_rectangle(left, top, right, top + bar_height + text_size + 8, fill="#FFFFFF", outline="")

        draw_bar = True
        for width in widths:
            segment_width = width * module_width
            if draw_bar:
                self.canvas.create_rectangle(x, top, x + segment_width, top + bar_height, fill="#111111", outline="")
            x += segment_width
            draw_bar = not draw_bar

        self.canvas.create_text(
            (left + right) / 2,
            top + bar_height + 4,
            text=str(value),
            anchor="n",
            font=("Arial", text_size),
            fill="#111111",
        )
        return bar_height + text_size + 8

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
            
            texto_ahorro = _promo_price_text(ahorro_val)
            texto_ahora = _promo_price_text(special_price)
            texto_antes = _promo_price_text(price)

            desc = (self.item.get("DESCRIPCION") or "").upper()
            promo_terminos = self.item.get("PROMO_TERMINOS") or ""
            promo_terminos2 = self.item.get("PROMO_TERMINOS2") or ""

            # AHORRA
            self.canvas.create_rectangle(cx(1010), cy(65), cx(1480), cy(205), fill="#FFFFFF", outline="")
            self.canvas.create_text(
                cx(1245), cy(135), text=texto_ahorro, anchor="center",
                font=("Arial", -max(1, int(100 * sy)), "bold"), fill="#111111"
            )

            # PRECIO NUEVO
            self.canvas.create_rectangle(cx(340), cy(315), cx(1480), cy(500), fill="#FFFFFF", outline="")
            self.canvas.create_text(
                cx(910), cy(407), text=texto_ahora, anchor="center",
                font=("Arial", -max(1, int(166 * sy)), "bold"), fill="#111111"
            )

            # PRECIO ANTERIOR
            self.canvas.create_rectangle(cx(1010), cy(715), cx(1410), cy(815), fill="#FFFFFF", outline="")
            self.canvas.create_text(
                cx(1210), cy(765), text=texto_antes, anchor="center",
                font=("Arial", -max(1, int(66 * sy)), "bold"), fill="#111111"
            )

            if desc:
                self.canvas.create_text(
                    cx(792),
                    cy(563),
                    text=desc,
                    width=cx(1425),
                    anchor="center",
                    justify="center",
                    font=("Arial", -max(10, int(PROMO_DESCRIPTION_FONT_SIZE * sy)), "bold"),
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
                    font=("Arial", -max(9, int(28 * sy)), "bold"),
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
                        font=("Arial", -max(9, int(28 * sy)), "bold"),
                        fill="#111111",
                    )

            return

        page_pad = 12
        x1 = page_pad
        y1 = page_pad
        x2 = label_width - page_pad
        y2 = label_height - page_pad

        self.canvas.create_rectangle(x1, y1, x2, y2, fill="#FFFFFF", outline="#222222", width=2)

        margin = 60 if self.individual else 18
        text_left = x1 + margin
        text_right = x2 - margin
        center_x = label_width / 2

        desc = (self.item.get("DESCRIPCION") or "").upper()
        price = _format_currency(self.item.get("PRECIO") or "$0.00")
        special_price = self.item.get("PRECIO_ESPECIAL") or ""
        vigencia = self.item.get("VIGENCIA") or ""
        codes = f"{self.item.get('ARTICULO', '')}   {self.item.get('UPC', '')}".strip()

        desc_top = y1 + 8
        desc_size = self._fit_text_size(desc, 49, 23 if self.individual else 28, 14 if self.individual else 15)
        self.canvas.create_text(
            center_x,
            desc_top,
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
            price_size = self._fit_text_size(price, 8, 56 if self.individual else 68, 42 if self.individual else 50)
            self.canvas.create_text(
                center_x,
                y1 + (62 if self.individual else 86),
                text=price,
                width=text_right - text_left,
                anchor="n",
                justify="center",
                font=("Arial", price_size, "bold"),
                fill="#111111",
            )
            footer_y = y1 + (156 if self.individual else 204)

        footer_size = 12 if self.individual else 14
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
            footer_y += 23

        if not special_price:
            barcode_value = _barcode_value_for_item(self.item)
            if barcode_value:
                self._draw_code128_barcode_canvas(
                    barcode_value,
                    text_left,
                    footer_y + 4,
                    text_right,
                    42 if self.individual else 50,
                    10 if self.individual else 12,
                )


class App(ThemedTk):
    def __init__(self):
        super().__init__(theme="arc")
        self.title(APP_TITLE)
        apply_window_icon(self)
        fit_window_to_screen(self, desired_width=900, desired_height=720, min_width=720, min_height=480, margin=16)
        self.resizable(True, True)

        self.config_data = load_config()
        self.current_built_item = {}
        self.configure_styles()
        self.create_widgets()
        self.populate_printers()
        self.refresh_default_printer_note()
        self._show_config_load_warnings()
        self.after(500, self._check_new_prices_on_startup)

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

    def _show_config_load_warnings(self):
        warnings = pop_config_load_warnings()
        if warnings:
            messagebox.showwarning(APP_TITLE, "\n\n".join(warnings))

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

        self.loguito_photo = load_ui_image(LOGUITO_FILENAME, (190, 76))
        self.lbl_loguito = None
        if self.loguito_photo:
            self.lbl_loguito = ttk.Label(preview_frame, image=self.loguito_photo)
            self.lbl_loguito.grid(row=r, column=2, columnspan=2, sticky="w", padx=pad_x, pady=(0, pad_y))

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
        preferencias["auto_print_on_scan"] = True

        self.auto_print_var = tk.BooleanVar(value=True)
        self.chk_auto_print = ttk.Checkbutton(printer_frame, text="Impresión automática al escanear", variable=self.auto_print_var)
        self.chk_auto_print.grid(row=2, column=0, columnspan=2, sticky="w", padx=pad_x, pady=pad_y)

        self.individual_print_var = tk.BooleanVar(value=bool(preferencias.get("individual_print", False)))
        self.chk_individual_print = ttk.Checkbutton(
            printer_frame,
            text="Imprimir etiqueta individual (angosta)",
            variable=self.individual_print_var,
            command=self.on_individual_print_change,
        )
        self.chk_individual_print.grid(row=3, column=0, columnspan=2, sticky="w", padx=pad_x, pady=pad_y)

        self.special_price_print_var = tk.BooleanVar(value=bool(preferencias.get("special_price_print", False)))
        self.chk_special_price_print = ttk.Checkbutton(
            printer_frame,
            text="Imprimir etiqueta de precio especial",
            variable=self.special_price_print_var,
            command=self.on_special_price_print_change,
        )
        self.chk_special_price_print.grid(row=4, column=0, columnspan=2, sticky="w", padx=pad_x, pady=pad_y)
        self.special_price_print_var.set(False)
        self.chk_special_price_print.configure(state="disabled")

        # --- Botones de Acción ---
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill="x")

        self.btn_cerrar = ttk.Button(buttons_frame, text="Cerrar", command=self.on_close)
        self.btn_cerrar.pack(side="right", padx=(10, 0))

        self.btn_imprimir = ttk.Button(buttons_frame, text="Imprimir", command=self.on_print_click)
        self.btn_imprimir.pack(side="right", padx=(10, 0))

        self.btn_preview_print = ttk.Button(buttons_frame, text="Vista previa", command=self.on_preview_print_click)
        self.btn_preview_print.pack(side="right", padx=(10, 0))

        self.btn_precios_nuevos = ttk.Button(buttons_frame, text="Precios Nuevos", command=self.on_new_prices_print_click)
        self.btn_precios_nuevos.pack(side="right", padx=(10, 0))

        self.btn_precios_especiales_vigentes = ttk.Button(
            buttons_frame,
            text="Precios Especiales Vigentes",
            command=self.on_current_special_prices_click,
        )
        self.btn_precios_especiales_vigentes.pack(side="right", padx=(10, 0))
        
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

    def _is_special_price_print_enabled(self) -> bool:
        return bool(self.special_price_print_var.get())

    def _item_has_special_price(self, item: Dict[str, Any]) -> bool:
        return bool((item or {}).get("PRECIO_ESPECIAL"))

    def _sync_special_price_print_option(self, item: Dict[str, Any]):
        self.special_price_print_var.set(False)
        if self._item_has_special_price(item):
            self.chk_special_price_print.configure(state="normal")
            return
        self.chk_special_price_print.configure(state="disabled")

    def on_individual_print_change(self):
        pass

    def on_special_price_print_change(self):
        if self._is_special_price_print_enabled():
            if not self._item_has_special_price(self.current_built_item):
                self.special_price_print_var.set(False)
                messagebox.showwarning(
                    APP_TITLE,
                    "El articulo seleccionado no tiene precio especial vigente.",
                )
                return

    def _item_for_selected_print_format(
        self,
        item: Dict[str, Any],
        individual_print: Optional[bool] = None,
        special_price_print: Optional[bool] = None,
        auto_individual_print: bool = False,
    ) -> Dict[str, Any]:
        if individual_print is None:
            individual_print = self._is_individual_print_enabled()
        if special_price_print is None:
            special_price_print = self._is_special_price_print_enabled()

        item_to_print = dict(item or {})
        special_price_print = bool(special_price_print and self._item_has_special_price(item_to_print))
        if special_price_print:
            individual_print = False

        if individual_print:
            item_to_print = _item_for_individual_label(item_to_print)
            if auto_individual_print:
                item_to_print["_AUTO_INDIVIDUAL_PRINT"] = True
            return item_to_print

        if not special_price_print:
            item_to_print["PRECIO_ESPECIAL"] = ""
            item_to_print["TIENE_PRECIO_ESPECIAL"] = False
            item_to_print["PROMO_CANTIDAD"] = None
            item_to_print["PROMO_TERMINOS"] = ""
            item_to_print["PROMO_TERMINOS2"] = ""
        return item_to_print

    def _send_label_to_printer(
        self,
        item: Dict[str, Any],
        copies: int,
        individual_print: Optional[bool] = None,
        show_errors: bool = True,
    ):
        if individual_print is None:
            individual_print = self._is_individual_print_enabled()

        if self._item_has_special_price(item):
            individual_print = False

        if individual_print:
            print_label_gdi_small(item, self.config_data, copies=copies, show_errors=show_errors)
        else:
            print_label_gdi(item, self.config_data, copies=copies, show_errors=show_errors)

    def _item_for_silent_print(
        self,
        item: Dict[str, Any],
        individual_print: bool,
        special_price_print: bool,
    ) -> Dict[str, Any]:
        item_to_print = build_label_text(item, self.config_data)
        return self._item_for_selected_print_format(
            item_to_print,
            individual_print=individual_print,
            special_price_print=special_price_print,
            auto_individual_print=True,
        )

    # --------- Eventos ----------
    def on_unified_search(self, event=None):
        term = (self.entry_search.get() or "").strip()
        if not term:
            return

        auto_print = self.auto_print_var.get()
        individual_print = self._is_individual_print_enabled()
        special_price_print = self._is_special_price_print_enabled()
        try:
            results = search_items(term, self.config_data)
            
            if not results:
                if not auto_print:
                    messagebox.showinfo(APP_TITLE, "No se encontraron artículos.")
                self._clear_preview()
            elif len(results) == 1:
                self._show_item_in_preview(results[0])
                if auto_print:
                    individual_print = self._is_individual_print_enabled()
                    special_price_print = self._is_special_price_print_enabled()
                    self.after(
                        80,
                        lambda item=results[0], individual=individual_print, special=special_price_print: self._print_item_silent(
                            item,
                            individual_print=individual,
                            special_price_print=special,
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
            individual_print = self._is_individual_print_enabled()
            special_price_print = self._is_special_price_print_enabled()
            if special_price_print:
                mock["PRECIO_ESPECIAL"] = "$79.90"
                mock["TIENE_PRECIO_ESPECIAL"] = True
            mock = self._item_for_selected_print_format(
                mock,
                individual_print=individual_print,
                special_price_print=special_price_print,
            )
            self._send_label_to_printer(mock, copies=copies, individual_print=individual_print)
            messagebox.showinfo(APP_TITLE, "Impresión de prueba enviada.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Error en impresión de prueba: {e}")

    def _check_new_prices_on_startup(self):
        if self.winfo_exists():
            self.on_new_prices_print_click(show_empty=False)

    def on_new_prices_print_click(self, show_empty: bool = True):
        try:
            copies = int(self.spin_copias.get())
            if copies < 1 or copies > 20:
                raise ValueError("Copias fuera de rango (1..20).")

            self._update_config_from_ui()
            items = fetch_new_daily_price_items(self.config_data)
            if not items:
                if show_empty:
                    messagebox.showinfo(APP_TITLE, "No hay cambios de precios pendientes para imprimir.")
                return

            selection_window = NewDailyPricesSelectionWindow(self, items)
            selected_items = selection_window.selected_items
            if not selected_items:
                return

            individual_print = self._is_individual_print_enabled()
            self.btn_precios_nuevos.configure(state="disabled")
            self.update_idletasks()

            printed_items = 0
            for item in selected_items:
                item_to_print = self._item_for_silent_print(
                    item,
                    individual_print=individual_print,
                    special_price_print=False,
                )
                if not item_to_print.get("DESCRIPCION") or not item_to_print.get("PRECIO"):
                    continue
                self._send_label_to_printer(
                    item_to_print,
                    copies=copies,
                    individual_print=individual_print,
                    show_errors=False,
                )
                printed_items += 1

            self._show_item_in_preview(selected_items[0])
            messagebox.showinfo(
                APP_TITLE,
                f"Se enviaron {printed_items} articulo(s) de precios nuevos a impresion.",
            )
        except Exception as e:
            write_log("ERROR on_new_prices_print_click", e)
            messagebox.showerror(APP_TITLE, f"Error al imprimir precios nuevos: {e}")
        finally:
            if hasattr(self, "btn_precios_nuevos"):
                self.btn_precios_nuevos.configure(state="normal")

    def on_current_special_prices_click(self):
        try:
            copies = int(self.spin_copias.get())
            if copies < 1 or copies > 20:
                raise ValueError("Copias fuera de rango (1..20).")

            self._update_config_from_ui()
            items = fetch_current_special_price_items(self.config_data)
            if not items:
                messagebox.showinfo(APP_TITLE, "No hay precios especiales vigentes para hoy.")
                return

            selection_window = CurrentSpecialPricesSelectionWindow(self, items)
            selected_items = selection_window.selected_items
            if not selected_items:
                return

            self.btn_precios_especiales_vigentes.configure(state="disabled")
            self.update_idletasks()

            printed_items = 0
            skipped_items = 0
            for item in selected_items:
                item_to_print = self._item_for_silent_print(
                    item,
                    individual_print=False,
                    special_price_print=True,
                )
                disponible = _quantity_to_float(item_to_print.get("DISPONIBLE"))
                if (
                    not item_to_print.get("DESCRIPCION")
                    or not item_to_print.get("PRECIO")
                    or not item_to_print.get("PRECIO_ESPECIAL")
                    or (disponible is not None and disponible <= 1)
                ):
                    skipped_items += 1
                    continue
                self._send_label_to_printer(
                    item_to_print,
                    copies=copies,
                    individual_print=False,
                    show_errors=False,
                )
                printed_items += 1

            self._show_item_in_preview(selected_items[0])
            skipped_text = f" Se omitieron {skipped_items} sin disponible mayor a 1." if skipped_items else ""
            messagebox.showinfo(
                APP_TITLE,
                f"Se enviaron {printed_items} articulo(s) de precios especiales a impresion.{skipped_text}",
            )
        except Exception as e:
            write_log("ERROR on_current_special_prices_click", e)
            messagebox.showerror(APP_TITLE, f"Error al imprimir precios especiales vigentes: {e}")
        finally:
            if hasattr(self, "btn_precios_especiales_vigentes"):
                self.btn_precios_especiales_vigentes.configure(state="normal")

    def _validate_special_price(self, item: Dict[str, Any]) -> bool:
        special_price = item.get("PRECIO_ESPECIAL") or ""
        if not special_price:
            messagebox.showwarning(
                APP_TITLE,
                "El articulo seleccionado no tiene precio especial vigente.",
            )
            return False

        disponible = _quantity_to_float(item.get("DISPONIBLE"))
        if disponible is not None and disponible <= 1:
            messagebox.showwarning(
                APP_TITLE,
                "El articulo seleccionado no tiene existencia disponible mayor a 1.",
            )
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
        individual_print = self._is_individual_print_enabled()
        special_price_print = self._is_special_price_print_enabled()
        if special_price_print and not self._validate_special_price(item):
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
            item = self._item_for_selected_print_format(
                item,
                individual_print=individual_print,
                special_price_print=special_price_print,
            )
            self._send_label_to_printer(item, copies=copies, individual_print=individual_print)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Error al imprimir: {e}")

    def _print_current_item_silent(
        self,
        individual_print: Optional[bool] = None,
        special_price_print: Optional[bool] = None,
    ):
        item = self._collect_item_from_preview()
        if individual_print is None:
            individual_print = self._is_individual_print_enabled()
        if special_price_print is None:
            special_price_print = self._is_special_price_print_enabled()
        self._print_item_silent(
            item,
            individual_print=individual_print,
            special_price_print=special_price_print,
        )

    def _print_item_silent(self, item: Dict[str, Any], individual_print: bool, special_price_print: bool):
        item = self._item_for_silent_print(item, individual_print, special_price_print)
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
        individual_print = self._is_individual_print_enabled()
        special_price_print = self._is_special_price_print_enabled()
        if special_price_print and not self._validate_special_price(item):
            return
        try:
            price_digits = re.sub(r"[^0-9.]", "", item["PRECIO"])
            if not price_digits:
                raise ValueError("Precio invalido.")
            self._update_config_from_ui()
            item = self._item_for_selected_print_format(
                item,
                individual_print=individual_print,
                special_price_print=special_price_print,
            )
            LabelPrintPreviewWindow(self, item, individual=individual_print)
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
        self.special_price_print_var.set(False)
        self.chk_special_price_print.configure(state="disabled")
        self._unlock_preview_fields()
        self.txt_producto.delete("1.0", tk.END)
        self.entry_precio.delete(0, tk.END)
        self.entry_precio_especial.delete(0, tk.END)
        self.entry_codigo.delete(0, tk.END)
        self.entry_upc.delete(0, tk.END)
        self._lock_preview_fields()
        self.entry_vigencia.delete(0, tk.END)
        self.entry_vigencia.insert(0, resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"])))


    def _show_item_in_preview(self, item: Optional[Dict[str, Any]]):
        if not item: # This can happen if the selection window is closed
            return
            
        # Build text for preview
        built = build_label_text(item, self.config_data)
        self.current_built_item = built
        self._sync_special_price_print_option(built)

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
        for key in ("TIENE_PRECIO_ESPECIAL", "PROMO_CANTIDAD", "PROMO_FECHAINICIO", "PROMO_FECHAFINAL", "PROMO_TERMINOS", "PROMO_TERMINOS2", "DISPONIBLE"):
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
        preferencias["special_price_print"] = self._is_special_price_print_enabled()

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
