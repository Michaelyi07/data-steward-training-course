#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Export vector PDF via PowerPoint COM (pptx-to-pdf skill chain 1)."""
import os
import sys
from win32com import client

for _s in (sys.stdout, sys.stderr):
    _enc = getattr(_s, "encoding", "") or ""
    if _enc.lower().replace("-", "") not in ("utf8", "utf"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PPTX = os.path.join(ROOT, "data_steward_course.pptx")
PDF = os.path.join(ROOT, "data_steward_course.pdf")
if os.path.exists(PDF):
    os.remove(PDF)
app = client.Dispatch("PowerPoint.Application")
app.Visible = True
pres = app.Presentations.Open(PPTX, WithWindow=False)
try:
    pres.SaveAs(PDF, 32)  # 32 = ppSaveAsPDF
finally:
    pres.Close()
    app.Quit()
print("PDF_OK", os.path.getsize(PDF), "bytes")
