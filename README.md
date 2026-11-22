# Data Administrator Training Course — Master Data, Dedup & Data Quality for data administrators, 28 slides, with a note on
every page. All sample data is fictional; no client name, real company, credit code
or phone number appears in any deliverable. Chinese-language delivery.

## Deliverables

| File | What it is |
|---|---|
| `data_steward_course.pptx` | 28-slide source deck (16:9), notes embedded on every slide |
| `data_steward_course.pdf` | Vector PDF exported from the deck |
| `lesson_plan.md` | Verbatim-level instructor plan: 8 chapters, objectives, question chains, activities, per-page lecture script, certification items |
| `notes_per_page.json` | page_id (P01–P28) → verbatim instructor note |
| `docs/object_sheets.md` | 6 master-data object sheets + ownership/SoR map |
| `docs/dedup_merge_tables.md` | Two-class dedup decision table, flowchart, merge worksheet, 3 drills |
| `docs/quality_report_sample.md` | 3-rate quality report sample (fictional) |
| `docs/synthetic_data_register.md` | All fictional entities registered |
| `docs/masking_watermark_samples.md` | Mobile masking / export approval / watermark samples |
| `wireframes/*.svg` | 4 wireframe source files (SVG, open directly in browser/GitHub) |
| `build_pptx.py`, `tools/` | Reproducible build + verification + QA + render scripts |
| `evidence/` | Real command outputs and rendered-page screenshots |

## Reproduce

```bash
python build_pptx.py                       # rebuild pptx + notes from lesson_plan.md
python -m pip install python-pptx          # if needed (v1.0.2 verified)
python tools/verify.py                     # 24 acceptance checks -> evidence/verify.txt
python <qa>/ai_flavor_gate.py data_steward_course.pptx
python <qa>/geometry_gate.py data_steward_course.pptx
# vector PDF (PowerPoint COM, Windows + pywin32):
D:/Program Files (x86)/Python313/python.exe tools/export_pdf.py
```

`<qa>` = `~/.workbuddy/skills/pptx-build-render-qa/scripts/`.

## Design & QA gates

- Layout: deep navy / slate / restrained orange accent, Microsoft YaHei, 16:9,
  one idea per page, CJK line-breaking (`eaLnBrk` / `hangingPunct`) on all paragraphs.
- AI-flavor gate: FAIL=0 WARN=0; geometry gate: exit 0 (no overlap/overflow);
  orphan-line gate: 0 hits (no lone character or punctuation on its own line;
  flowchart 是/否 labels whitelisted).
- No decorative short rule above page titles, no isolated accent blocks.
- Subject-matter gate: independent reviewer verified all 8 chapters against the
  source knowledge card and client brief — PASS after fixing a rate-denominator
  contradiction and a placeholder-polluted object table.
- Post-delivery visual review removed the title-rule decoration and fixed two
  orphan-line cases (P02 "金额", P03 full stop); see evidence/render_fix/.
- Desensitization: zero hits for client name / real 18-digit USCC / real mobile.

## Content boundaries

Owner role names, field-dictionary required items, name-normalization word list,
three-rate thresholds/expiry days, watermark fields, import rollback policy and
merge worksheet columns are illustrative and marked "[需按本题标定]" — they are not
invented as client facts. Client-specified rules (USCC match = confirmed duplicate,
normalized-name match = suspected duplicate, 30-day undo, `****`+last-4 masking,
AI never writes/sends/calculates amounts) are stated verbatim.

## Integration points (stubs)

This is a training deliverable; no live data platform is required. Metrics, IDs and
codes in examples are stubs aligned to the fictional data register. No external SDK
is used; nothing sends data anywhere.
