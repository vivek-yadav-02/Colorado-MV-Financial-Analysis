# -*- coding: utf-8 -*-
"""Dismiss save prompts, maximize Power BI, load data via the 'Refresh now' bar, show the Filters pane, capture."""
import sys, time, ctypes
import numpy as np
import win32gui, win32con, win32ui, win32api
from PIL import Image

ctypes.windll.user32.SetProcessDPIAware()
OUT = sys.argv[1] if len(sys.argv) > 1 else "settle.png"
SHOW_FILTERS = "--filters" in sys.argv


def windows():
    res = []
    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            if t:
                res.append((h, t, win32gui.GetWindowRect(h)))
    win32gui.EnumWindows(cb, None)
    return res


def main_window():
    c = [(r[2] - r[0]) * (r[3] - r[1]) for h, t, r in windows() if "Colorado_MV_Sales" in t]
    ws = [(h, r) for h, t, r in windows() if "Colorado_MV_Sales" in t]
    return max(ws, key=lambda x: (x[1][2] - x[1][0]) * (x[1][3] - x[1][1]))[0] if ws else None


def click_abs(x, y):
    win32api.SetCursorPos((x, y)); time.sleep(0.25)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)


def dismiss_dialogs():
    for h, t, (l, tp, r, b) in windows():
        if t == "Microsoft Power BI Desktop" and r - l < 900:
            click_abs(l + int((r - l) * 0.862), tp + int((b - tp) * 0.80))   # Cancel button (relative position)
            print("dismissed save prompt", flush=True); time.sleep(1.5)


def capture(h):
    l, t, r, b_ = win32gui.GetWindowRect(h)
    w, hh = r - l, b_ - t
    hwndDC = win32gui.GetWindowDC(h); mfcDC = win32ui.CreateDCFromHandle(hwndDC); saveDC = mfcDC.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap(); bmp.CreateCompatibleBitmap(mfcDC, w, hh); saveDC.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(h, saveDC.GetSafeHdc(), 2); info = bmp.GetInfo()
    im = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), bmp.GetBitmapBits(True), "raw", "BGRX", 0, 1)
    win32gui.DeleteObject(bmp.GetHandle()); saveDC.DeleteDC(); mfcDC.DeleteDC(); win32gui.ReleaseDC(h, hwndDC)
    return im


def red_px(im):
    a = np.asarray(im).astype(int); r, g, b = a[:, :, 0], a[:, :, 1], a[:, :, 2]
    H, W = a.shape[:2]
    return int(((r > 180) & (g < 120) & (b < 120))[int(H*0.27):int(H*0.91), int(W*0.13):int(W*0.62)].sum())


def bar_row(im):
    """y of the notification bar: a light row band between ribbon and canvas at x=300 whose row below is dark canvas."""
    a = np.asarray(im).astype(int)
    W = im.width; x1, x0 = int(300 * W / 1938), int(60 * W / 1938)
    for y in range(int(150 * im.height / 1098), int(300 * im.height / 1098)):
        if a[y, x1].sum() > 600 and a[y, x0].sum() > 600 and a[y + int(30 * im.height / 1098), x1].sum() < 300:
            return y
    return None


dismiss_dialogs()
h = main_window(); assert h, "no main window"
win32gui.ShowWindow(h, win32con.SW_MAXIMIZE)
try: win32gui.SetForegroundWindow(h)
except Exception: pass
time.sleep(1.5)
l, t, r, b = win32gui.GetWindowRect(h)
SX, SY = (r - l) / 1938, (b - t) / 1098          # my reference coordinates come from a 1938x1098 maximized window
print("rect", (l, t, r, b), "scale", round(SX, 3), round(SY, 3), flush=True)
def click_ref(x, y): click_abs(l + int(x * SX), t + int(y * SY))

for attempt in range(4):
    im = capture(h)
    n = red_px(im); y = bar_row(im)
    print(f"attempt {attempt}: red px {n}, bar row {y}", flush=True)
    if n > 5000 or y is None:
        break
    click_abs(l + int(1380 * SX), t + y)             # 'Refresh now' button on the bar
    time.sleep(30)
    dismiss_dialogs()

if SHOW_FILTERS:
    click_ref(411, 72); time.sleep(2)                # View tab (Filters button located afterwards from the capture)
im = capture(h); im.save(OUT); print("saved", OUT, "red px", red_px(im), flush=True)
