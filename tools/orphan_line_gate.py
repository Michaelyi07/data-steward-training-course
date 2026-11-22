#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Chinese orphan-character / orphan-punctuation line gate (method card §二 第4层).

Flags lines that are a single CJK char/alnum, or made only of punctuation.
Whitelist: flowchart branch labels such as 是/否 (intentionally standalone).
"""
import os
import sys

from pptx import Presentation

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PPTX = os.path.join(ROOT, "data_steward_course.pptx")
PUNCT = set("。，、；：？！）》」』】…·—-,.;:!?)]}\"'")
# Intentional standalone semantic labels (flowchart branches) — not orphan text.
WHITELIST = {"是", "否"}

prs = Presentation(PPTX)
hits = []
for si, slide in enumerate(prs.slides, 1):
    for shp in slide.shapes:
        if not shp.has_text_frame:
            continue
        for para in shp.text_frame.paragraphs:
            for raw in para.text.split("\x0b"):
                ln = raw.strip()
                if not ln:
                    continue
                if ln in WHITELIST:
                    continue
                single = len(ln) == 1 and (ln in PUNCT or ln.isalnum())
                only_punct = len(ln) <= 2 and all(c in PUNCT for c in ln)
                if single or only_punct:
                    hits.append((si, shp.shape_id, ln))

print(f"orphan-line gate: {len(hits)} hit(s)")
for si, sid, txt in hits:
    print(f"  slide {si} shape {sid}: {txt!r}")
sys.exit(1 if hits else 0)
