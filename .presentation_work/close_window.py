import sys
import time

import win32con
import win32gui


fragment = sys.argv[1].lower()
matches = []


def collect(hwnd, _):
    title = win32gui.GetWindowText(hwnd)
    if win32gui.IsWindowVisible(hwnd) and fragment in title.lower():
        matches.append(hwnd)


win32gui.EnumWindows(collect, None)
if not matches:
    raise RuntimeError(f"No se encontro una ventana con: {sys.argv[1]}")
for hwnd in matches:
    win32gui.PostMessage(hwnd, win32con.WM_CLOSE, 0, 0)
time.sleep(0.8)
