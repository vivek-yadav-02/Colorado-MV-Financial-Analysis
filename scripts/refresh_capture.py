# -*- coding: utf-8 -*-
"""Wait for Power BI window, dismiss popup, click 'Refresh now' until the notification bars disappear,
collapse side panes, capture. Usage: refresh_capture.py OUT.png"""
import sys, time, ctypes
import win32gui, win32con, win32ui, win32api, win32process
from PIL import Image

ctypes.windll.user32.SetProcessDPIAware()
TITLE = "Colorado_MV_Sales"
OUT = sys.argv[1]


def find():
    res = []
    def cb(h, _):
        if win32gui.IsWindowVisible(h) and TITLE.lower() in win32gui.GetWindowText(h).lower():
            res.append(h)
    win32gui.EnumWindows(cb, None)
    return res[0] if res else None


def capture(h):
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
    return im


def click(hwnd, x, y):
    l, t, _, _ = win32gui.GetWindowRect(hwnd)
    win32api.SetCursorPos((l + x, t + y)); time.sleep(0.3)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    print("clicked", x, y, flush=True); time.sleep(1.0)


t0 = time.time()
hwnd = None
while time.time() - t0 < 240 and hwnd is None:
    hwnd = find(); time.sleep(3)
assert hwnd, "window not found"
print(f"window after {time.time()-t0:.0f}s", flush=True)
time.sleep(50)
win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
try: win32gui.SetForegroundWindow(hwnd)
except Exception: pass
time.sleep(2)

im = capture(hwnd)
if im.getpixel((1600, 150))[0] > 200 and im.getpixel((1600, 150))[1] < 200:   # not used; placeholder
    pass
# dismiss "Collaborate and share" popup if present (dark grey box around x 1415-1930, y 70-250)
if sum(im.getpixel((1700, 200))) < 400:
    click(hwnd, 1891, 108); time.sleep(1)

def data_loaded(im):
    # red columns of the Annual Sales chart appear around y=480 once data is loaded
    reds = 0
    for x in range(330, 1000, 10):
        r, g, bb = im.getpixel((x, 480))
        if r > 180 and g < 120 and bb < 120:
            reds += 1
    return reds


for attempt in range(4):
    im = capture(hwnd)
    n = data_loaded(im)
    print("red-bar probe", n, flush=True)
    if n >= 5:
        break
    click(hwnd, 766, 140)          # ribbon Home > Refresh
    time.sleep(30)

# collapse Visualizations and Data panes
click(hwnd, 1686, 233); time.sleep(1.5)
click(hwnd, 1910, 233); time.sleep(4)
im = capture(hwnd)
im.save(OUT); print("saved", OUT, im.size, flush=True)
