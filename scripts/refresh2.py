# -*- coding: utf-8 -*-
"""Find the yellow/grey notification bar in Power BI Desktop and click its 'Refresh now' button; verify data loaded."""
import sys, time, ctypes
import numpy as np
import win32gui, win32con, win32ui, win32api
from PIL import Image

ctypes.windll.user32.SetProcessDPIAware()
TITLE = "Colorado_MV_Sales"


def find():
    res = []
    def cb(h, _):
        if win32gui.IsWindowVisible(h) and TITLE.lower() in win32gui.GetWindowText(h).lower():
            l, t, r, b_ = win32gui.GetWindowRect(h)
            res.append(((r - l) * (b_ - t), h))
    win32gui.EnumWindows(cb, None)
    return max(res)[1] if res else None


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
    win32api.SetCursorPos((l + x, t + y)); time.sleep(0.25)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    print("clicked", x, y, flush=True)


def bar_rows(im):
    """rows between the ribbon and the canvas where x=300 is light (a notification bar)."""
    a = np.asarray(im).astype(int)
    ys = []
    for y in range(140, 320):
        if a[y, 300].sum() > 3 * 200 and a[y, 60].sum() > 3 * 200:
            ys.append(y)
    # group into bands
    bands = []
    for y in ys:
        if bands and y - bands[-1][-1] <= 2:
            bands[-1].append(y)
        else:
            bands.append([y])
    return [(b[0], b[-1]) for b in bands if b[-1] - b[0] >= 20]


def data_loaded(im):
    """true when the page canvas shows red bars (dashboard) or a table header — probe several page rows for saturated red."""
    a = np.asarray(im).astype(int)
    r, g, bb = a[:, :, 0], a[:, :, 1], a[:, :, 2]
    red = (r > 180) & (g < 120) & (bb < 120)
    return int(red[300:1000, 250:1200].sum())


hwnd = find(); assert hwnd
win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
try: win32gui.SetForegroundWindow(hwnd)
except Exception: pass
time.sleep(1.5)
for attempt in range(4):
    im = capture(hwnd)
    n = data_loaded(im)
    full_ribbon = sum(im.getpixel((300, 190))) > 3 * 200     # light at y=190 => full ribbon; dark => compact ribbon
    print("red px:", n, "full ribbon:", full_ribbon, flush=True)
    if n > 200:
        break
    if full_ribbon:
        click(hwnd, 766, 140)      # Home > Refresh (full ribbon)
    else:
        click(hwnd, 296, 121)      # Home > Refresh (compact ribbon)
    time.sleep(30)
im = capture(hwnd)
im.save(sys.argv[1] if len(sys.argv) > 1 else "refresh2.png")
print("final red px:", data_loaded(im), flush=True)
