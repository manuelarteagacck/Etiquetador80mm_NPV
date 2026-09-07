import ctypes
import time
from pathlib import Path

import win32api
import win32clipboard
import win32con
import win32gui
from PIL import ImageGrab


OUTPUT_DIR = Path(__file__).resolve().parent / "screens"
APP_TITLE = "Etiquetador 80mm"


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


def activate(hwnd):
    win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception:
        pass
    time.sleep(0.4)


def capture(title_fragment: str, filename: str):
    hwnd, title = find_window(title_fragment)
    activate(hwnd)
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    image = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
    output = OUTPUT_DIR / filename
    image.save(output)
    print(f"CAPTURE={output} title={title!r} size={image.size}")
    return hwnd


def close_window(title_fragment: str):
    hwnd, _ = find_window(title_fragment)
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
    time.sleep(0.8)


def click_relative(hwnd, x: int, y: int):
    left, top, _, _ = win32gui.GetWindowRect(hwnd)
    activate(hwnd)
    win32api.SetCursorPos((left + x, top + y))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(0.35)


def paste_text(text: str):
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
    finally:
        win32clipboard.CloseClipboard()

    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(ord("V"), 0, 0, 0)
    win32api.keybd_event(ord("V"), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.2)


def press_enter():
    win32api.keybd_event(win32con.VK_RETURN, 0, 0, 0)
    win32api.keybd_event(win32con.VK_RETURN, 0, win32con.KEYEVENTF_KEYUP, 0)


def clear_focused_text():
    win32api.keybd_event(win32con.VK_CONTROL, 0, 0, 0)
    win32api.keybd_event(ord("A"), 0, 0, 0)
    win32api.keybd_event(ord("A"), 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_CONTROL, 0, win32con.KEYEVENTF_KEYUP, 0)
    win32api.keybd_event(win32con.VK_BACK, 0, 0, 0)
    win32api.keybd_event(win32con.VK_BACK, 0, win32con.KEYEVENTF_KEYUP, 0)
    time.sleep(0.2)


def main():
    ctypes.windll.user32.SetProcessDPIAware()
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    # La lista automatica ya se documento por separado; se cierra sin imprimir.
    try:
        close_window("Precios Nuevos")
    except RuntimeError:
        pass

    app_hwnd = capture(APP_TITLE, "11_principal_vigente.png")

    # Desactiva temporalmente la impresion automatica antes de consultar.
    click_relative(app_hwnd, 50, 585)
    click_relative(app_hwnd, 420, 100)
    paste_text("FA1020104")
    press_enter()
    time.sleep(4)
    click_relative(app_hwnd, 420, 100)
    clear_focused_text()
    capture(APP_TITLE, "12_articulo_cargado_vigente.png")

    # Abre la vista previa; no se envia ningun trabajo de impresion.
    click_relative(app_hwnd, 607, 725)
    time.sleep(1.2)
    capture("Vista previa de impresion", "13_vista_previa_vigente.png")
    close_window("Vista previa de impresion")

    # Documenta las alternativas de configuracion sin guardar cambios.
    click_relative(app_hwnd, 50, 542)
    click_relative(app_hwnd, 50, 626)
    capture(APP_TITLE, "14_configuracion_impresion_vigente.png")
    click_relative(app_hwnd, 50, 502)
    click_relative(app_hwnd, 50, 626)

    # Consulta de promociones vigente; no selecciona ni imprime registros.
    click_relative(app_hwnd, 316, 725)
    time.sleep(4)
    capture("Precios Especiales Vigentes", "15_precios_especiales_vigentes.png")
    close_window("Precios Especiales Vigentes")

    # Restablece la preferencia temporal de escaneo automatico.
    click_relative(app_hwnd, 50, 585)
    capture(APP_TITLE, "16_principal_restaurada.png")


if __name__ == "__main__":
    main()
