#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Acceptance verification for the data administrator course. Real checks, raw output.
Run: python tools/verify.py   (from repo root or anywhere)
"""
import json
import os
import re
import sys
import zipfile

for _s in (sys.stdout, sys.stderr):
    _enc = getattr(_s, "encoding", "") or ""
    if _enc.lower().replace("-", "") not in ("utf8", "utf"):
        try:
            _s.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
PPTX = os.path.join(ROOT, "data_steward_course.pptx")
NOTES = os.path.join(ROOT, "notes_per_page.json")
PLAN = os.path.join(ROOT, "lesson_plan.md")
DOCS = os.path.join(ROOT, "docs")
WF = os.path.join(ROOT, "wireframes")

results = []


def check(name, ok, detail=""):
    results.append((name, bool(ok), detail))
    print(f"[{'PASS' if ok else 'FAIL'}] {name}" + (f" — {detail}" if detail else ""))


def read(p):
    with open(p, encoding="utf-8") as f:
        return f.read()


# C1: 28 pages + note on every page (both JSON and embedded pptx notes)
from pptx import Presentation
prs = Presentation(PPTX)
check("C1-pages 25-30", 25 <= len(prs.slides) <= 30, f"slides={len(prs.slides)}")
notes = json.load(open(NOTES, encoding="utf-8"))
check("C1-notes json 28 entries", len(notes) == 28, f"count={len(notes)}")
embedded = 0
for sl in prs.slides:
    if sl.has_notes_slide and sl.notes_slide.notes_text_frame.text.strip():
        embedded += 1
check("C1-notes embedded in pptx", embedded == 28, f"with_notes={embedded}")
page_ids_ok = all(f"P{i:02d}" in notes for i in range(1, 29))
check("C1-page ids P01-P28 bound", page_ids_ok)

# C2: two duplicate classes with distinct rules in deck + docs
ppt_text = ""
with zipfile.ZipFile(PPTX) as z:
    for n in z.namelist():
        if n.startswith("ppt/slides/slide") and n.endswith(".xml"):
            ppt_text += z.read(n).decode("utf-8", "replace")
dedup_doc = read(os.path.join(DOCS, "dedup_merge_tables.md"))
check("C2-confirmed rule (USCC same)", "统一社会信用代码相同" in ppt_text and "确定重复" in ppt_text)
check("C2-suspected rule (name normalized)", "规范化" in ppt_text and "疑似重复" in ppt_text)
check("C2-manual review stated", "人工复核" in dedup_doc and "不得自动合并" in dedup_doc)
check("C2-30 day undo + over-window approval", "30" in dedup_doc and "超窗" in dedup_doc and "审批" in dedup_doc)
check("C2-3 drills with answer paths", dedup_doc.count("演练") >= 3)

# C3: 6 master data objects individually in deck + object sheets
obj_doc = read(os.path.join(DOCS, "object_sheets.md"))
objs = ["会员企业", "联系人", "能力卡", "需求卡", "商机", "活动参与"]
for o in objs:
    check(f"C3-object {o}", o in ppt_text and o in obj_doc)
check("C3-unique SoR rule", "唯一" in obj_doc and "SoR" in obj_doc)
check("C3-soft delete excluded from quality, kept in audit", "软删" in obj_doc and "审计" in obj_doc)

# C4: fictional data + desensitization scan across all deliverables
scan_targets = [PLAN, NOTES, os.path.join(DOCS, "synthetic_data_register.md")]
for d in ("dedup_merge_tables.md", "quality_report_sample.md", "masking_watermark_samples.md", "object_sheets.md"):
    scan_targets.append(os.path.join(DOCS, d))
all_text = "\n".join(read(p) for p in scan_targets if os.path.exists(p)) + ppt_text
check("C4-no client name", "会员空间平台" not in all_text and "S4-TRAIN" not in all_text)
# real 18-digit credit code pattern (real ones start with valid region code, not FICT)
real_uscc = re.findall(r"\b[1-9]\d{5}[A-Z0-9]{12}\b", all_text)
check("C4-no real-looking USCC", len(real_uscc) == 0, f"hits={real_uscc[:3]}")
real_phone = re.findall(r"(?<!\d)1[3-9]\d{9}(?!\d)", all_text)
check("C4-no real phone number", len(real_phone) == 0, f"hits={real_phone[:3]}")
check("C4-synthetic register exists", "虚构" in read(os.path.join(DOCS, "synthetic_data_register.md")))
check("C4-masked phone samples", "****5678" in read(os.path.join(DOCS, "masking_watermark_samples.md")))
check("C4-export approval + watermark", "审批" in read(os.path.join(DOCS, "masking_watermark_samples.md"))
      and "水印" in read(os.path.join(DOCS, "masking_watermark_samples.md")))

# wireframes
wfs = [f for f in os.listdir(WF) if f.endswith(".svg")]
check("wireframes 4 svg source files", len(wfs) == 4, f"count={len(wfs)}")

passed = sum(1 for _, ok, _ in results if ok)
total = len(results)
print(f"\nTOTAL: {passed}/{total} passed")
sys.exit(0 if passed == total else 1)
