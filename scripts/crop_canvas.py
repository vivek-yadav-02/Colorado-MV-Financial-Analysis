# -*- coding: utf-8 -*-
"""Crop the 16:9 report page out of a Power BI Desktop window capture using the sidebar as anchor."""
import sys
import numpy as np
from PIL import Image

src, dst = sys.argv[1], sys.argv[2]
im = Image.open(src).convert("RGB")
a = np.asarray(im).astype(int)
SIDEBAR = np.array([14, 23, 41])
PAGE_BG = np.array([11, 18, 32])
mask = (np.abs(a - SIDEBAR).sum(axis=2) <= 6)
cols = mask.sum(axis=0)
col_idx = np.where(cols > 0.3 * cols.max())[0]
left = int(col_idx.min())
# in a column just inside the sidebar, the page spans the contiguous run of rows that are NOT the outspace colour
colpix = a[:, left + 3, :]
notbg = colpix.sum(axis=1) < 3 * 128   # dark rows = inside the page (UI chrome is light)
rows = np.where(mask[:, left + 3])[0]
anchor = int(rows[len(rows) // 2])
top = anchor
while top > 0 and notbg[top - 1]:
    top -= 1
bottom = anchor
while bottom < len(notbg) - 1 and notbg[bottom + 1]:
    bottom += 1
bottom += 1
height = bottom - top
width = round(height * 1280 / 720)
box = (left, top, left + width, bottom)
print("crop box", box, "scale", round(height / 720, 3))
out = im.crop(box)
if out.width < 1280:
    out = out.resize((1280, 720), Image.LANCZOS)
out.save(dst, optimize=True)
print("saved", dst, out.size)
