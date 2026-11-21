#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Render selected PPTX pages to PNG via PowerPoint COM for visual QA."""
import os
import sys
from win32com import client

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PPTX = os.path.join(ROOT, "data_steward_course.pptx")
OUT = os.path.join(ROOT, "evidence", "render")
os.makedirs(OUT, exist_ok=True)

PAGES = [1, 3, 12, 14, 15, 18, 22, 28]  # cover, objects, dedup, flowchart, merge, quality, watermark, cert

app = client.Dispatch("PowerPoint.Application")
app.Visible = True
pres = app.Presentations.Open(PPTX, WithWindow=False)
try:
    for n in PAGES:
        path = os.path.join(OUT, f"page_{n:02d}.png")
        if os.path.exists(path):
            os.remove(path)
        pres.Slides(n).Export(path, "PNG", 1600, 900)
        print(f"page {n}: {os.path.getsize(path)} bytes")
finally:
    pres.Close()
    app.Quit()
