# -*- coding: utf-8 -*-
"""Wait for the Power BI Desktop window whose title contains TITLE, wait for it to settle, capture it."""
import sys, time, ctypes, ctypes.wintypes as wt
import win32gui, win32con, win32ui, win32process
from PIL import Image

TITLE = sys.argv[1] if len(sys.argv) > 1 else "Colorado_MV_Sales"
OUT = sys.argv[2] if len(sys.argv) > 2 else "pbi_shot.png"
SETTLE = int(sys.argv[3]) if len(sys.argv) > 3 else 45      # seconds to wait after window appears
MAXWAIT = int(sys.argv[4]) if len(sys.argv) > 4 else 240

ctypes.windll.user32.SetProcessDPIAware()


def find_windows():
    res = []
    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            if t and TITLE.lower() in t.lower():
                res.append((h, t))
    win32gui.EnumWindows(cb, None)
    return res


def all_pbi_windows():
    res = []
    def cb(h, _):
        if win32gui.IsWindowVisible(h):
            t = win32gui.GetWindowText(h)
            if t:
                try:
                    _, pid = win32process.GetWindowThreadProcessId(h)
                except Exception:
                    pid = 0
                res.append((h, t, pid))
    win32gui.EnumWindows(cb, None)
    return res


t0 = time.time()
hwnd = None
while time.time() - t0 < MAXWAIT:
    w = find_windows()
    if w:
        hwnd, title = w[0]
        print(f"[{time.time()-t0:5.0f}s] found window: {title!r}", flush=True)
        break
    time.sleep(3)
if hwnd is None:
    print("window not found; visible windows:", flush=True)
    for h, t, pid in all_pbi_windows():
        print("  ", pid, t)
    sys.exit(1)

# wait for load (data import etc.)
time.sleep(SETTLE)

# list other top-level windows that might be dialogs from the same process
_, pid = win32process.GetWindowThreadProcessId(hwnd)
dialogs = [(h, t) for h, t, p in all_pbi_windows() if p == pid and h != hwnd]
print("other windows of the PBI process:", dialogs, flush=True)

try:
    win32gui.ShowWindow(hwnd, win32con.SW_MAXIMIZE)
except Exception as e:
    print("maximize failed", e)
try:
    win32gui.SetForegroundWindow(hwnd)
except Exception as e:
    print("foreground failed (ok)", e)
time.sleep(2)


def capture(h, path):
    l, t, r, b = win32gui.GetWindowRect(h)
    w, hh = r - l, b - t
    hwndDC = win32gui.GetWindowDC(h)
    mfcDC = win32ui.CreateDCFromHandle(hwndDC)
    saveDC = mfcDC.CreateCompatibleDC()
    bmp = win32ui.CreateBitmap()
    bmp.CreateCompatibleBitmap(mfcDC, w, hh)
    saveDC.SelectObject(bmp)
    ctypes.windll.user32.PrintWindow(h, saveDC.GetSafeHdc(), 2)
    info = bmp.GetInfo()
    data = bmp.GetBitmapBits(True)
    im = Image.frombuffer("RGB", (info["bmWidth"], info["bmHeight"]), data, "raw", "BGRX", 0, 1)
    win32gui.DeleteObject(bmp.GetHandle())
    saveDC.DeleteDC(); mfcDC.DeleteDC(); win32gui.ReleaseDC(h, hwndDC)
    im.save(path)
    print("saved", path, im.size, flush=True)


capture(hwnd, OUT)
for i, (h, t) in enumerate(dialogs[:3]):
    try:
        capture(h, OUT.replace(".png", f"_dlg{i}.png"))
        print("dialog title:", t)
    except Exception as e:
        print("dialog capture failed", t, e)
