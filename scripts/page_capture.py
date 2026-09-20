# -*- coding: utf-8 -*-
"""Switch Power BI Desktop to each report page via the tab strip, verify via the sidebar highlight, crop and save.
Usage: page_capture.py OUT_PREFIX"""
import sys, time, ctypes
import numpy as np
import win32gui, win32con, win32ui, win32api
from PIL import Image

ctypes.windll.user32.SetProcessDPIAware()
TITLE = "Colorado_MV_Sales"
PREFIX = sys.argv[1] if len(sys.argv) > 1 else "final"
PAGES = ["dashboard", "trend", "counties", "regions", "quarters", "data", "about"]
TAB_X = [313, 470, 588, 694, 800, 915, 1022]      # tab centres on the maximized 1938px-wide window
TAB_Y = 1037
NAV_Y0, NAV_STEP, NAV_H = 76, 36, 28
SIDEBAR = np.array([14, 23, 41])


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


def crop_page(im):
    a = np.asarray(im).astype(int)
    mask = (np.abs(a - SIDEBAR).sum(axis=2) <= 6)
    cols = mask.sum(axis=0)
    idx = np.where(cols > 0.3 * cols.max())[0]
    if len(idx) == 0:
        return None, None
    left = int(idx.min())
    colpix = a[:, left + 3, :]
    dark = colpix.sum(axis=1) < 3 * 128
    rows = np.where(mask[:, left + 3])[0]
    anchor = int(rows[len(rows) // 2])
    top = int(rows.min())            # sidebar colour starts exactly at the page top (canvas above it is a different dark)
    outspace = np.abs(colpix - np.array([11, 18, 32])).sum(axis=1) <= 8     # canvas colour outside the page
    bottom = anchor
    while bottom < len(colpix) - 1 and not outspace[bottom + 1] and dark[bottom + 1]: bottom += 1
    bottom += 1
    height = bottom - top
    width = round(height * 1280 / 720)
    scale = height / 720
    return im.crop((left, top, left + width, bottom)), scale


def active_index(page_img, scale):
    """find which nav row carries the red active bar at page x≈9."""
    a = np.asarray(page_img).astype(int)
    x = int(9.5 * scale)
    best, best_n = None, 0
    for i in range(len(PAGES)):
        y0 = int((NAV_Y0 + i * NAV_STEP - 6) * scale); y1 = int((NAV_Y0 + i * NAV_STEP - 6 + NAV_H) * scale)
        seg = a[y0:y1, x, :]
        n = int(((seg[:, 0] > 150) & (seg[:, 1] < 120) & (seg[:, 2] < 120)).sum())
        if n > best_n:
            best, best_n = i, n
    return best if best_n > 5 else None


hwnd = find()
assert hwnd, "window not found"
win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
try: win32gui.SetForegroundWindow(hwnd)
except Exception: pass
time.sleep(1.5)

for i, key in enumerate(PAGES):
    ok = False
    for attempt in range(4):
        click(hwnd, TAB_X[i], TAB_Y)
        time.sleep(5 if attempt == 0 else 7)
        im = capture(hwnd)
        if im.width < 1000:                      # a tooltip window matched the title; re-find the main window
            time.sleep(3); hwnd = find(); continue
        page, scale = crop_page(im)
        if page is None:
            print(key, "crop failed, retry", flush=True); continue
        idx = active_index(page, scale)
        print(f"{key}: attempt {attempt} active={idx}", flush=True)
        if idx == i:
            page.save(f"{PREFIX}_{key}.png", optimize=True)
            ok = True
            break
    if not ok:
        im.save(f"{PREFIX}_{key}_UNVERIFIED.png")
        print(key, "could not verify page switch", flush=True)
print("done", flush=True)
