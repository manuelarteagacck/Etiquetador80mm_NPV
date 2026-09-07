import ctypes
from pathlib import Path
import sys
import tkinter as tk

import win32gui
from PIL import ImageGrab

PROJECT_ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(PROJECT_ROOT))

from app import LabelPrintPreviewWindow


OUTPUT = Path(__file__).resolve().parent / "preview_after_fix.png"


def find_window(title):
    matches = []

    def collect(hwnd, _):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd) == title:
            matches.append(hwnd)

    win32gui.EnumWindows(collect, None)
    if not matches:
        raise RuntimeError(f"No se encontro la ventana: {title}")
    return matches[0]


def main():
    ctypes.windll.user32.SetProcessDPIAware()
    root = tk.Tk()
    root.title("Raiz de prueba")
    root.geometry("1x1+0+0")
    root.update_idletasks()

    original_wait_window = LabelPrintPreviewWindow.wait_window
    LabelPrintPreviewWindow.wait_window = lambda *_args, **_kwargs: None
    try:
        preview = LabelPrintPreviewWindow(
            root,
            {
                "DESCRIPCION": "AGUA NATURAL CIEL 600ML",
                "PRECIO": "$200.00",
                "VIGENCIA": "Valido a partir de: 30/07/2026 Aplican TyC",
                "ARTICULO": "AG1010011",
                "UPC": "7501055307906",
            },
            individual=False,
        )
    finally:
        LabelPrintPreviewWindow.wait_window = original_wait_window

    def capture():
        try:
            hwnd = find_window("Vista previa de impresion")
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            ImageGrab.grab(
                bbox=(left, top, right, bottom),
                all_screens=True,
            ).save(OUTPUT)
            print(f"PREVIEW={OUTPUT}")
        finally:
            preview.destroy()
            root.destroy()

    root.after(800, capture)
    root.after(5000, root.destroy)
    root.mainloop()


if __name__ == "__main__":
    main()
