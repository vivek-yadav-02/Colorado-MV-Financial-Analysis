# -*- coding: utf-8 -*-
"""Click at window-relative coords in the Power BI window, wait, capture. Usage: click_capture.py OUT wait x y [x y ...]"""
import sys, time, ctypes
import win32gui, win32con, win32ui, win32api, win32process
from PIL import Image

ctypes.windll.user32.SetProcessDPIAware()
TITLE = "Colorado_MV_Sales"
OUT = sys.argv[1]
WAIT = int(sys.argv[2])
clicks = [(int(sys.argv[i]), int(sys.argv[i + 1])) for i in range(3, len(sys.argv) - 1, 2)]


def find():
    res = []
    def cb(h, _):
        if win32gui.IsWindowVisible(h) and TITLE.lower() in win32gui.GetWindowText(h).lower():
            res.append(h)
    win32gui.EnumWindows(cb, None)
    return res[0] if res else None


hwnd = find()
assert hwnd, "window not found"
win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
try:
    win32gui.SetForegroundWindow(hwnd)
except Exception:
    pass
time.sleep(1.5)
l, t, r, b_ = win32gui.GetWindowRect(hwnd)
print("rect", l, t, r, b_, flush=True)
for (x, y) in clicks:
    win32api.SetCursorPos((l + x, t + y))
    time.sleep(0.3)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    print("clicked", x, y, flush=True)
    time.sleep(1.0)
time.sleep(WAIT)

_, pid = win32process.GetWindowThreadProcessId(hwnd)
others = []
def cb2(h, _):
    if win32gui.IsWindowVisible(h) and win32gui.GetWindowText(h):
        try:
            _, p = win32process.GetWindowThreadProcessId(h)
        except Exception:
            p = 0
        if p == pid and h != hwnd:
            others.append((h, win32gui.GetWindowText(h)))
win32gui.EnumWindows(cb2, None)
print("other windows:", others, flush=True)


def capture(h, path):
    l, t, r, b_ = win32gui.GetWindowRect(h)
    w, hh = r - l, b_ - t
    hwndDC = win32gui.GetWindowDC(h)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfcDC, w, hh)
    saveDC.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(h, saveDC.GetSafeHdc(), 2)
    info = bmp.GetInfo()
    im = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), bmp.GetBitmapBits(True), "raw", "BGRX", 0, 1)
    win32gui.DeleteObject(bmp.GetHandle()); saveDC.DeleteDC(); mfcDC.DeleteDC(); win32gui.ReleaseDC(h, hwndDC)
    im.save(path)
    print("saved", path, im.size, flush=True)


capture(hwnd, OUT)
for i, (h, tt) in enumerate(others[:3]):
    try:
        capture(h, OUT.replace(".png", f"_dlg{i}.png")); print("dialog:", tt)
    except Exception as e:
        print("dlg fail", tt, e)
