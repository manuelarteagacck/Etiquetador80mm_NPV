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
from dataclasses import dataclass, asdict
from typing import Optional, Dict, Any, List

# --- Impresión / Windows ---
import win32print
import win32con
import win32ui

# --- ESC/POS ---
from escpos.printer import Serial, Usb, Network, Dummy

# --- DB ---
import pyodbc

# --- UI ---
import tkinter as tk
from tkinter import ttk, messagebox, font
from ttkthemes import ThemedTk

APP_TITLE = "Etiquetador 80mm — TIENDA NPV"
CONFIG_FILENAME = "config.json"


# =============================
# Utilidades de configuración
# =============================
DEFAULT_CONFIG = {
    "conexion_odbc": {
        # Ejemplo con autenticación integrada:
        # "dsn": "",  # si usas DSN
        "server": "MI_SERVIDOR_SQL",
        "database": "NPV",
        "trusted_connection": True,
        "username": "",
        "password": "",
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
    "vigencia_default": "Vigente {MES_ABR} {YYYY}"  # plantilla con macros
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
    if not os.path.exists(CONFIG_FILENAME):
        with open(CONFIG_FILENAME, "w", encoding="utf-8") as f:
            json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
        return json.loads(json.dumps(DEFAULT_CONFIG))
    with open(CONFIG_FILENAME, "r", encoding="utf-8") as f:
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
    return deep_merge(data, json.loads(json.dumps(DEFAULT_CONFIG)))


def save_config(cfg: dict):
    with open(CONFIG_FILENAME, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


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
        return f"${v:0.2f}"
    except Exception:
        return ""

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
            pv.PRECIOSVENTA,
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
        (SELECT TOP 1 e.EQUIVALENTE FROM NPV.dbo.NPVFDArticulosEquivalentes e WHERE e.ARTICULO = a.ARTICULO) as UPC,
        pr.FECHAINICIO
    FROM NPV.dbo.NPVFDArticulos a
    JOIN PreciosRankeados pr ON a.ARTICULO = pr.ARTICULO
    WHERE pr.rn = 1
    ORDER BY a.DESCRIPCION;
    """

    try:
        conn = _get_connection(cfg)
        cursor = conn.cursor()
        
        rows = cursor.execute(sql, term, like_term, term_digits).fetchall()

        for row in rows:
            fecha_vigencia = row.FECHAINICIO.strftime("%d/%m/%Y") if row.FECHAINICIO else datetime.date.today().strftime("%d/%m/%Y")
            
            item = {
                "ARTICULO": str(row.ARTICULO).strip(),
                "DESCRIPCION": str(row.DESCRIPCION or "").strip(),
                "PRECIO": _format_price(row.PRECIO),
                "UPC": _normalize_upc_digits(row.UPC or ""),
                "VIGENCIA": f"Valido a partir de: {fecha_vigencia} Aplican TyC"
            }
            results.append(item)

    except Exception as e:
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

def print_label_gdi(item_dict: Dict[str, Any], config: dict, copies: int = 1):
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

            y_pos = 20 # Margen superior

            # --- 2. Dibujar Descripción ---
            try:
                font_desc = win32ui.CreateFont({"name": "Arial", "height": 42, "weight": win32con.FW_BOLD})
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
                precio = item_dict.get("PRECIO", "$0.00")
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
        messagebox.showerror(APP_TITLE, f"Error de GDI: {e}")


def print_label_gdi_small(item_dict: Dict[str, Any], config: dict, copies: int = 1):
    """
    Imprime la etiqueta en formato angosto.
    """
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
                font_desc = win32ui.CreateFont({"name": "Arial", "height": 36, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_desc)
                desc = (item_dict.get("DESCRIPCION") or "").upper()
                rect = (horizontal_margin, y_pos, printable_width - horizontal_margin, y_pos + 200)
                hDC.DrawText(desc, rect, win32con.DT_CENTER | win32con.DT_WORDBREAK)
                y_pos += 70 # Avance original
            except Exception as e:
                print(f"[ERROR] No se pudo dibujar la descripción (individual): {e}")

            # --- 3. Dibujar Precio ---
            try:
                font_price_height = 77
                font_price = win32ui.CreateFont({"name": "Arial", "height": font_price_height, "weight": win32con.FW_BOLD})
                hDC.SelectObject(font_price)
                precio = item_dict.get("PRECIO", "$0.00")
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
        messagebox.showerror(APP_TITLE, f"Error de GDI (individual): {e}")


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
    articulo    = (item_dict.get("ARTICULO") or "").strip()
    upc         = _normalize_upc_digits(item_dict.get("UPC"))

    if not precio:
        # Precio inválido
        precio = ""

    # Vigencia: usa la que viene del item si existe, si no, la de por defecto.
    vigencia = item_dict.get("VIGENCIA")
    if not vigencia:
        vigencia = resolve_vigencia_template(config.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"]))


    return {
        "DESCRIPCION": descripcion,
        "PRECIO": precio,
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


class App(ThemedTk):
    def __init__(self):
        super().__init__(theme="arc")
        self.title(APP_TITLE)
        self.geometry("900x680")
        self.resizable(False, False)

        self.config_data = load_config()
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
        main_frame = ttk.Frame(self, padding=20)
        main_frame.pack(expand=True, fill="both")

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

        r += 1
        ttk.Label(preview_frame, text="Precio ($xx.xx):").grid(row=r, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_precio = ttk.Entry(preview_frame, width=25)
        self.entry_precio.grid(row=r, column=1, sticky="w", padx=pad_x, pady=pad_y)

        ttk.Label(preview_frame, text="Código interno (ARTICULO):").grid(row=r, column=2, sticky="e", padx=pad_x, pady=pad_y)
        self.entry_codigo = ttk.Entry(preview_frame, width=25)
        self.entry_codigo.grid(row=r, column=3, sticky="w", padx=pad_x, pady=pad_y)

        r += 1
        ttk.Label(preview_frame, text="UPC (numérico):").grid(row=r, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.entry_upc = ttk.Entry(preview_frame, width=25)
        self.entry_upc.grid(row=r, column=1, sticky="w", padx=pad_x, pady=pad_y)

        ttk.Label(preview_frame, text="Vigencia:").grid(row=r, column=2, sticky="e", padx=pad_x, pady=pad_y)
        self.entry_vigencia = ttk.Entry(preview_frame, width=25)
        self.entry_vigencia.grid(row=r, column=3, sticky="w", padx=pad_x, pady=pad_y)

        r += 1
        ttk.Label(preview_frame, text="Copias (1-20):").grid(row=r, column=0, sticky="w", padx=pad_x, pady=pad_y)
        self.spin_copias = ttk.Spinbox(preview_frame, from_=1, to=20, width=8)
        self.spin_copias.set(1)
        self.spin_copias.grid(row=r, column=1, sticky="w", padx=pad_x, pady=pad_y)

        # --- Impresora ---
        printer_frame = ttk.LabelFrame(main_frame, text="Configuración de Impresión")
        printer_frame.pack(fill="x", pady=(0, 20))
        printer_frame.columnconfigure(1, weight=1)

        self.printer_mode = tk.StringVar(value=self.config_data["impresora"]["modo"])
        self.rb_default = ttk.Radiobutton(printer_frame, text="Usar predeterminada", value="default", variable=self.printer_mode, command=self.on_printer_mode_change)
        self.rb_default.grid(row=0, column=0, sticky="w", padx=pad_x, pady=pad_y)

        self.rb_named = ttk.Radiobutton(printer_frame, text="Elegir de la lista", value="named", variable=self.printer_mode, command=self.on_printer_mode_change)
        self.rb_named.grid(row=1, column=0, sticky="w", padx=pad_x, pady=pad_y)

        self.cmb_printers = ttk.Combobox(printer_frame, state="disabled", width=40)
        self.cmb_printers.grid(row=1, column=1, sticky="ew", padx=pad_x, pady=pad_y)

        self.lbl_default = ttk.Label(printer_frame, text="", font=self.font_normal)
        self.lbl_default.grid(row=0, column=1, sticky="w", padx=pad_x, pady=pad_y)

        self.auto_print_var = tk.BooleanVar(value=False)
        self.chk_auto_print = ttk.Checkbutton(printer_frame, text="Impresión automática al escanear", variable=self.auto_print_var)
        self.chk_auto_print.grid(row=2, column=0, columnspan=2, sticky="w", padx=pad_x, pady=pad_y)

        self.individual_print_var = tk.BooleanVar(value=False)
        self.chk_individual_print = ttk.Checkbutton(printer_frame, text="Imprimir etiqueta individual (angosta)", variable=self.individual_print_var)
        self.chk_individual_print.grid(row=3, column=0, columnspan=2, sticky="w", padx=pad_x, pady=pad_y)

        # --- Botones de Acción ---
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill="x")

        self.btn_cerrar = ttk.Button(buttons_frame, text="Cerrar", command=self.on_close)
        self.btn_cerrar.pack(side="right", padx=(10, 0))

        self.btn_imprimir = ttk.Button(buttons_frame, text="Imprimir", command=self.on_print_click)
        self.btn_imprimir.pack(side="right", padx=(10, 0))
        
        self.btn_test = ttk.Button(buttons_frame, text="Probar Impresión", command=self.on_test_print)
        self.btn_test.pack(side="right")

        # Cargar vigencia default y bloquear/activar combobox impresoras correctamente
        self.entry_vigencia.delete(0, tk.END)
        self.entry_vigencia.insert(0, resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"])))
        self.on_printer_mode_change()

    # --------- Helpers UI ----------
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

    # --------- Eventos ----------
    def on_unified_search(self, event=None):
        term = (self.entry_search.get() or "").strip()
        if not term:
            return

        try:
            results = search_items(term, self.config_data)
            
            if not results:
                messagebox.showinfo(APP_TITLE, "No se encontraron artículos.")
                self._clear_preview()
            elif len(results) == 1:
                self._show_item_in_preview(results[0])
                if self.auto_print_var.get():
                    self.on_print_click()
            else:
                # Multiple results, open selection window
                ResultSelectionWindow(self, results, self._show_item_in_preview)

        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Error en la búsqueda: {e}")
        
        # Clear search box after search
        self.entry_search.delete(0, tk.END)


    def on_test_print(self):
        # Construir datos de ejemplo
        mock = {
            "DESCRIPCION": "PRODUCTO DE PRUEBA 80MM",
            "PRECIO": "$99.90",
            "ARTICULO": "TEST001",
            "UPC": "123456789012",
            "VIGENCIA": resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"]))
        }
        try:
            copies = int(self.spin_copias.get())
            self._update_config_from_ui()
            if self.individual_print_var.get():
                print_label_gdi_small(mock, self.config_data, copies=copies)
            else:
                print_label_gdi(mock, self.config_data, copies=copies)
            messagebox.showinfo(APP_TITLE, "Impresión de prueba enviada.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Error en impresión de prueba: {e}")

    def on_print_click(self):
        item = self._collect_item_from_preview()
        if not item.get("DESCRIPCION") or not item.get("PRECIO"):
            messagebox.showwarning(APP_TITLE, "No se puede imprimir sin Producto y Precio.")
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
            if self.individual_print_var.get():
                print_label_gdi_small(item, self.config_data, copies=copies)
            else:
                print_label_gdi(item, self.config_data, copies=copies)
            messagebox.showinfo(APP_TITLE, "Impresión enviada.")
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Error al imprimir: {e}")

    def on_close(self):
        try:
            self._update_config_from_ui()
            save_config(self.config_data)
        except Exception as e:
            print(f"[WARN] No se pudo guardar config: {e}")
        self.destroy()

    # --------- Lógica búsqueda/preview ----------
    def _clear_preview(self):
        self.txt_producto.delete("1.0", tk.END)
        self.entry_precio.delete(0, tk.END)
        self.entry_codigo.delete(0, tk.END)
        self.entry_upc.delete(0, tk.END)
        self.entry_vigencia.delete(0, tk.END)
        self.entry_vigencia.insert(0, resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"])))


    def _show_item_in_preview(self, item: Optional[Dict[str, Any]]):
        if not item: # This can happen if the selection window is closed
            return
            
        # Build text for preview
        built = build_label_text(item, self.config_data)

        # Producto/Descripción
        self.txt_producto.delete("1.0", tk.END)
        self.txt_producto.insert("1.0", built.get("DESCRIPCION", ""))

        # Precio
        self.entry_precio.delete(0, tk.END)
        self.entry_precio.insert(0, built.get("PRECIO", ""))

        # Código interno
        self.entry_codigo.delete(0, tk.END)
        self.entry_codigo.insert(0, built.get("ARTICULO", ""))

        # UPC
        self.entry_upc.delete(0, tk.END)
        self.entry_upc.insert(0, built.get("UPC", ""))

        # Vigencia
        self.entry_vigencia.delete(0, tk.END)
        self.entry_vigencia.insert(0, built.get("VIGENCIA", ""))

    def _collect_item_from_preview(self) -> Dict[str, Any]:
        descripcion = self.txt_producto.get("1.0", tk.END).strip()
        precio_raw = (self.entry_precio.get() or "").strip()
        # Normalizar precio a $xx.xx
        precio_digits = re.sub(r"[^0-9.]", "", precio_raw)
        precio_fmt = _format_price(precio_digits) if precio_digits else ""

        articulo = (self.entry_codigo.get() or "").strip()
        upc = _normalize_upc_digits(self.entry_upc.get() or "")
        vig = (self.entry_vigencia.get() or "").strip()
        if not vig:
            vig = resolve_vigencia_template(self.config_data.get("vigencia_default", DEFAULT_CONFIG["vigencia_default"]))
        return {
            "DESCRIPCION": descripcion,
            "PRECIO": precio_fmt,
            "ARTICULO": articulo,
            "UPC": upc,
            "VIGENCIA": vig
        }

    def _update_config_from_ui(self):
        # Impresora
        self.config_data["impresora"]["modo"] = self.printer_mode.get()
        if self.printer_mode.get() == "named":
            self.config_data["impresora"]["printer_name"] = self.cmb_printers.get()
        else:
            self.config_data["impresora"]["printer_name"] = ""

        # The "tamaños" and "espaciados" sections are removed from the new UI,
        # so we don't need to update them anymore. If you want to keep them,
        # you would need to add them back to the create_widgets method.
        
        # Example for one of them if you add it back:
        # self.config_data["espaciados"]["espacio_abajo_precio"] = int(self.spin_down.get())


# ==============================
# Entry point
# ==============================
def main():
    app = App()
    app.protocol("WM_DELETE_WINDOW", app.on_close)
    app.mainloop()


if __name__ == "__main__":
    main()
