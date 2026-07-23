---
name: thesis-formatting-skill
description: Create, draft, revise, format, and verify Chinese academic Word manuscripts and technical reports. Use for theses, journal or conference papers, paper templates, DOCX cleanup, equation-heavy manuscripts, chapter-based numbering, evidence-bounded result writing, real OMML math, standard three-line tables, Word SEQ/REF cross-references, pagination QA, or final artifact handoff.
---

# Chinese Academic Paper and Word Production

## Core workflow

1. Use the `documents` skill for DOCX creation or editing and follow its render-and-verify contract.
2. Treat an official university, conference, or journal template as authoritative. Otherwise copy `assets/通用中文学术Word模板_含交叉引用.docx`; use `assets/NF-SWIPT_中文学术技术报告模板.docx` only for that project.
3. Read `references/format-spec.md`. For a new manuscript or major revision, also read `references/paper-production-workflow.md`.
4. Establish the evidence boundary before drafting: identify source code, configuration, tables/CSV, figures, verified literature, allowed claims, prohibited claims, and whether rerunning experiments is authorized.
5. Build the chapter outline and map every main claim to a formula, table, figure, source, or explicit limitation. Do not start with long prose when the evidence chain is not closed.
6. Create or revise the DOCX using named Word styles, real OMML equations, standard three-line tables, and real `SEQ`/bookmark/`REF` fields. Read `references/word-math-and-tables.md` before handling equations or tables.
7. Update fields in Microsoft Word, repaginate, and record the actual Word page count. Do not force a page target by converting equations to text, shrinking the body below the governing template, or using unsafe margins.
8. Run structural QA:

```powershell
python scripts/audit_academic_docx.py "<document.docx>" --strict
```

9. Render the final DOCX to PDF/PNGs and inspect every page. If the required renderer is unavailable, follow the `documents` skill fallback: perform Word pagination plus structural OOXML audits and disclose that visual PNG QA was not completed.
10. Remove temporary PDFs, page images, staging scripts, lock files, and backups from the delivery folder. Preserve user files and unrelated Git changes.

## Non-negotiable mathematics

- Insert every mathematical variable or expression in prose, captions, and table cells as a Word OMML object. Examples include `P_r`, `P_EH`, `R_ub`, `rho`, `mu`, subscripts, superscripts, inequalities, and vectors.
- Use inline `<m:oMath>` inside prose and table cells; use `<m:oMathPara>` or a display paragraph for standalone equations.
- Do not leave LaTeX commands, underscore notation, or equation screenshots in the delivered DOCX.
- Keep ordinary technical acronyms as text when they function as names, such as SWIPT, PSO, MATLAB, IEEE, and AWGN. When an acronym is used as a mathematical quantity in an equation or constraint, insert it as OMML.
- Keep SI units upright and outside the variable object where practical. The micro prefix in `μW`, `μH`, and similar units is not a mathematical variable.
- Use MathML transformed with Microsoft Office `MML2OMML.XSL`, or equivalent deterministic OOXML construction. Use `scripts/omml_helpers.py` for reusable conversion helpers.

## Non-negotiable tables

- Use standard uncolored three-line tables for academic parameter and result tables.
- Keep only a top rule, a header-separating rule, and a bottom rule. Remove vertical borders, internal horizontal grids, cell shading, gradients, and decorative colors.
- Use explicit column widths and table geometry. Repeat the header row, vertically center cells, and allow rows to expand; never use a fixed row height that can clip text.
- Insert variables in headers and cells as OMML, not as plain underscore text.
- Prefer 10.5 pt table text unless an official template specifies otherwise. Reduce typography only after fixing widths, wrapping, redundant columns, and placement.

## Numbering and cross-references

- Number figures, tables, and equations by chapter: `图2-1`, `表3-2`, `(4-3)`.
- Generate numbers with Word `SEQ` fields. Reset the first item of a chapter with `\r 1`; continue with `\n`.
- Enclose each number in a unique bookmark such as `fig_2_1`, `tab_3_2`, `eq_4_3`, or `ref_1`.
- Insert every in-text citation to a figure, table, equation, or bibliography entry using `REF <bookmark> \h`. Never type a cross-reference number manually.
- Update all fields before delivery by opening the DOCX in Word, selecting all, and pressing `F9`.

## Evidence-bounded paper writing

- Separate verified facts, model assumptions, design settings, simulation outputs, estimates, and unknowns.
- State comparator, metric, operating condition, and evidence source for each claimed improvement.
- Do not use “global optimum”, “statistically significant”, “hardware verified”, or “universally superior” without direct supporting evidence.
- Call a theoretical capacity expression a capacity bound or proxy when it is not an implemented throughput measurement.
- Call an analytic ripple or error formula a proxy when no circuit/time-domain/experimental waveform verifies the physical quantity.
- Preserve one six-part explanation after every result figure: axes/object, calculation basis, main trend, method comparison, supported conclusion, and applicability boundary.

## Pagination and layout decisions

Apply fixes in this order:

1. Correct misplaced tables, captions, and keep-with-next settings.
2. Remove duplicated explanation and redundant columns.
3. Adjust figure size modestly while preserving legibility.
4. Tighten caption, equation, table, and reference spacing within the governing template.
5. Report a remaining page-count deviation honestly if further compression would violate the template or readability.

Never solve pagination by reverting OMML to text, using colored spreadsheet tables, distorting figures, or silently changing the required body style.

## Output and handoff

- Deliver only the requested final artifact unless previews are explicitly requested.
- Keep the submission manuscript separate from internal explanatory reports and source figures.
- Before moving a project, compare source and destination, copy to a staging location, verify relative paths and SHA-256 hashes, run available validation, switch directories only after verification, and delete the old tree only when the user explicitly requested migration.
- Never overwrite or reset unrelated user changes in a dirty worktree.
