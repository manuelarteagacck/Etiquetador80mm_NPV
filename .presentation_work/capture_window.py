import ctypes
import sys
import time

import win32con
import win32gui
from PIL import ImageGrab


def find_window(title_fragment):
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


ctypes.windll.user32.SetProcessDPIAware()
fragment = sys.argv[1]
output = sys.argv[2]
wait_seconds = float(sys.argv[3]) if len(sys.argv) > 3 else 0.0
deadline = time.time() + wait_seconds
while True:
    try:
        hwnd, title = find_window(fragment)
        break
    except RuntimeError:
        if time.time() >= deadline:
            raise
        time.sleep(0.5)
win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
try:
    win32gui.SetForegroundWindow(hwnd)
except Exception:
    pass
time.sleep(0.6)
left, top, right, bottom = win32gui.GetWindowRect(hwnd)
image = ImageGrab.grab(bbox=(left, top, right, bottom), all_screens=True)
image.save(output)
print(title, (left, top, right, bottom), image.size, output)
