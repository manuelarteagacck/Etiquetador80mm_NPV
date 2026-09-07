import ctypes
from pathlib import Path
import sys
import tkinter as tk
import time

import win32con
import win32gui
import win32ui
from PIL import Image, ImageGrab


PROJECT_ROOT = Path(__file__).resolve().parent.parent
OUTPUT_DIR = Path(__file__).resolve().parent / "screens"
sys.path.insert(0, str(PROJECT_ROOT))

from app import App, LabelPrintPreviewWindow


def find_window(title_fragment: str):
    matches = []

    def collect(hwnd, _):
        title = win32gui.GetWindowText(hwnd)
        if win32gui.IsWindowVisible(hwnd) and title_fragment.lower() in title.lower():
            matches.append((hwnd, title))

    win32gui.EnumWindows(collect, None)
    if not matches:
        raise RuntimeError(f"No se encontro una ventana con: {title_fragment}")
    matches.sort(key=lambda item: len(item[1]))
    return matches[0]


def capture(title_fragment: str, filename: str):
    hwnd, title = find_window(title_fragment)
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.5)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top
    hwnd_dc = win32gui.GetWindowDC(hwnd)
    source_dc = win32ui.CreateDCFromHandle(hwnd_dc)
    memory_dc = source_dc.CreateCompatibleDC()
    bitmap = win32ui.CreateBitmap()
    bitmap.CreateCompatibleBitmap(source_dc, width, height)
    memory_dc.SelectObject(bitmap)
    try:
        rendered = ctypes.windll.user32.PrintWindow(hwnd, memory_dc.GetSafeHdc(), 2)
        if not rendered:
            raise RuntimeError("PrintWindow no devolvio contenido")
        bitmap_info = bitmap.GetInfo()
        bitmap_bits = bitmap.GetBitmapBits(True)
        image = Image.frombuffer(
            "RGB",
            (bitmap_info["bmWidth"], bitmap_info["bmHeight"]),
            bitmap_bits,
            "raw",
            "BGRX",
            0,
            1,
        )
    except Exception:
        image = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    finally:
        win32gui.DeleteObject(bitmap.GetHandle())
        memory_dc.DeleteDC()
        source_dc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwnd_dc)
    output = OUTPUT_DIR / filename
    image.save(output)
    print(f"CAPTURE={output} title={title!r} size={image.size}")


def show_preview(root, item, filename: str):
    original_wait_window = LabelPrintPreviewWindow.wait_window
    LabelPrintPreviewWindow.wait_window = lambda *_args, **_kwargs: None
    try:
        preview = LabelPrintPreviewWindow(root, item, individual=False)
    finally:
        LabelPrintPreviewWindow.wait_window = original_wait_window
    root.update_idletasks()
    root.update()
    capture("Vista previa de impresion", filename)
    preview.destroy()
    root.update()


def main():
    ctypes.windll.user32.SetProcessDPIAware()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    original_startup_check = App._check_new_prices_on_startup
    App._check_new_prices_on_startup = lambda _self: None
    try:
        root = App()
    finally:
        App._check_new_prices_on_startup = original_startup_check

    root.update_idletasks()
    root.update()

    normal_item = {
        "DESCRIPCION": "ALLIVIAX 550 MG.",
        "PRECIO": "$150.00",
        "PRECIO_ESPECIAL": "",
        "ARTICULO": "FA1020104",
        "UPC": "650240013805",
        "VIGENCIA": "Valido a partir de: 26/07/2026 Aplican TyC",
    }
    root.auto_print_var.set(False)
    root._show_item_in_preview(normal_item)
    root.update_idletasks()
    root.update()
    capture("Etiquetador 80mm", "12_articulo_cargado_vigente.png")
    show_preview(root, normal_item, "13_vista_previa_vigente.png")

    root.printer_mode.set("named")
    root.on_printer_mode_change()
    root.individual_print_var.set(True)
    root.on_individual_print_change()
    root.update_idletasks()
    root.update()
    capture("Etiquetador 80mm", "14_configuracion_impresion_vigente.png")

    promo_item = {
        "DESCRIPCION": "BE LIGHT JAMAICA 1L.",
        "PRECIO": "$27.00",
        "PRECIO_ESPECIAL": "$25.00",
        "TIENE_PRECIO_ESPECIAL": True,
        "ARTICULO": "AG1020056",
        "UPC": "7501031000000",
        "VIGENCIA": "Promocion vigente: 06/02/2026 - 31/12/2026 Aplican TyC",
    }
    root.printer_mode.set("default")
    root.on_printer_mode_change()
    root.individual_print_var.set(False)
    root.on_individual_print_change()
    root._show_item_in_preview(promo_item)
    root.special_price_print_var.set(True)
    root.on_special_price_print_change()
    root.update_idletasks()
    root.update()
    capture("Etiquetador 80mm", "17_articulo_promocion_vigente.png")
    show_preview(root, promo_item, "18_vista_previa_promocion_vigente.png")

    root.destroy()


if __name__ == "__main__":
    main()
