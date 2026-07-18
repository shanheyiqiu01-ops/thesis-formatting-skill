---
name: thesis-formatting-skill
description: Create, revise, and visually verify general Chinese academic Word documents using A4 SimSun/Times New Roman/SimHei typography, chapter-based figure/table/equation numbering, real Word SEQ fields, bookmarks, and REF cross-references. Use for theses, dissertations, journal or conference manuscripts, technical reports, project reports, paper figure-explanation reports, or DOCX cleanup where formatting, numbering, and citations must update automatically.
---

# Academic Word Format

## Core workflow

1. Use the `documents` skill for DOCX creation or editing and follow its render-and-verify workflow.
2. Check whether the user supplied an official university, conference, or journal template. Treat it as authoritative. Otherwise copy `assets/通用中文学术Word模板_含交叉引用.docx`.
3. Use `assets/NF-SWIPT_中文学术技术报告模板.docx` only when the task is specifically the NF-SWIPT project.
4. Read `references/format-spec.md` before making layout, numbering, or citation decisions.
5. Keep scientific claims bounded by the available model, simulation, or experiment evidence.
6. Update all Word fields before delivery: open the document in Word, select all with `Ctrl+A`, then press `F9`.
7. Render the final DOCX to PDF, inspect every page, and run:

```powershell
python scripts/audit_academic_docx.py "<document.docx>"
```

## Non-negotiable numbering

- Number figures, tables, and equations by chapter: `图2-1`, `表3-2`, `(4-3)`.
- Implement displayed numbers with Word `SEQ` fields. Set chapter resets with `\r 1` at the first item of each chapter and continue with `\n`.
- Enclose each numbered item in a unique bookmark such as `fig_2_1`, `tab_3_2`, or `eq_4_3`.
- Insert in-text figure, table, equation, and reference citations as `REF bookmark \h` fields.
- Number bibliography entries with `SEQ Reference`; bookmark each entry as `ref_1`, `ref_2`, and so on.
- Never type a cross-reference number manually. Initial visible field results may be populated for compatibility, but the field code must remain the source of truth.

## Typography and figures

- Use Chinese SimSun 12 pt and English/numerals Times New Roman 12 pt for body text, with fixed 20 pt line spacing.
- Use SimHei for Chinese headings and Times New Roman for Latin heading text.
- Use SimSun/Times New Roman 10.5 pt for captions.
- Keep figure-internal labels, ticks, and legends consistent at 9–10.5 pt. Word cannot repair fonts embedded in PNG/PDF; configure them in MATLAB before export.
- Preserve one six-part explanation after every result figure: axes/object, calculation basis, trend, method comparison, supported conclusion, and applicability boundary.

## Output discipline

- Choose the document structure from the actual task. Do not carry NF-SWIPT terminology, figures, methods, or parameter names into unrelated theses or reports.
- Keep a submission manuscript separate from an internal explanatory report. The manuscript follows the target venue; an explanatory report may retain fixed source-figure identifiers when it maps directly to another artifact.
- Deliver the DOCX and, when requested, a PDF preview. Do not leave temporary PDFs, rendered pages, or staging files in the project folder.
