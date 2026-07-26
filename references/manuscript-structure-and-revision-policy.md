# Manuscript Structure and Revision Policy

Use this reference when drafting or revising a Chinese conference/journal manuscript, especially when editing an existing DOCX.

## 1. Default conference-paper structure

Unless an official template overrides it, use this order:

1. Chinese title.
2. Author, affiliation, correspondence, and funding metadata required by the venue.
3. Abstract.
4. Keywords.
5. `1 引言`.
6. `2 系统模型` or the domain-equivalent theoretical/model section.
7. `3 方法、算法或优化问题`.
8. `4 仿真/实验结果与讨论`.
9. `5 结论`.
10. Acknowledgements when required.
11. `参考文献`.

### Introduction logic

Build the introduction in this sequence:

1. Research background and application need.
2. Prior work grouped by technical route, with verified citations.
3. Specific unresolved gap and why existing methods do not close it.
4. Bounded contributions that match the available model, data, figures, and validation.
5. A final organization paragraph, for example: Section 2 establishes the system model; Section 3 formulates and solves the method; Section 4 presents results and limitations; Section 5 concludes the paper.

Do not replace the organization paragraph with a vague sentence such as “the remainder of this paper is arranged as follows” without naming the sections and their roles.

### Model and method sections

- Define the system architecture and assumptions before performance formulas.
- Derive variables, units, constraints, and proxies in the same order used by the source code or experiment.
- Distinguish physical quantities, theoretical bounds, proxy indicators, simulation settings, and measured results.
- State the decision variables, objective, constraints, solver, validation method, and comparison baselines.

### Results section

For every principal figure, explain:

1. Axes and plotted object.
2. Calculation or experimental basis.
3. Main trend.
4. Method-to-method comparison.
5. Supported conclusion.
6. Applicability boundary or limitation.

### Conclusion

Summarize only supported findings. Do not add new claims, citations, equations, or unreported experiments.

## 2. First-use terminology and abbreviations

For a Chinese manuscript draft, keep the Chinese abstract in Chinese. Write the
technical term normally and do not append its English full term or abbreviation
in parentheses unless the user or the governing template explicitly requires
that form. This avoids crowding a Chinese abstract with English parentheticals
when the paper will later receive a separate English translation.

Treat the main text as the abbreviation-definition scope. At the first
main-text occurrence, write:

`中文全称（English full term，ABBREVIATION）`

Main-text examples:

- `同时无线信息与能量传输（simultaneous wireless information and power transfer，SWIPT）`
- `近场同时无线信息与能量传输（near-field simultaneous wireless information and power transfer，NF-SWIPT）`
- `基于功率分割的近场同时无线信息与能量传输（power-splitting near-field simultaneous wireless information and power transfer，PS-NF-SWIPT）`
- `粒子群优化（particle swarm optimization，PSO）`

After the first main-text definition, use the abbreviation consistently. When
an English abstract or full English manuscript is prepared later, define its
abbreviations independently according to the English template; do not reuse the
Chinese draft's abstract formatting by default.

Apply the same rule to TS, PS, EH, ID, BER, SNR, PTE, AWGN, and other non-universal abbreviations. Do not expand ordinary units or software/product names unnecessarily.

## 3. Literature-citation formatting

- Use the official venue style first.
- Under the user's default Chinese manuscript style, use true Word superscript formatting for bracketed numeric citations.
- Join adjacent multiple citations without punctuation: `[1][2][3]`.
- Do not use `[1],[2],[3]` or `[1]，[2]` unless the template explicitly requires separators.
- Keep the superscript group attached to the preceding statement and avoid an isolated citation at the beginning of a new line.
- For an existing DOCX, remove only comma separator runs. Do not replace the surrounding `REF` fields or their formatting.

## 4. Empty or deleted bibliography policy

This is a hard default:

- If the `参考文献` heading exists but entries are empty or deleted, leave the bibliography area blank for the author.
- Do not reconstruct entries from prior drafts, chat history, session logs, web search, Crossref, DOI guesses, or nearby citations.
- Do not create placeholder references that look real.
- Do not add `SEQ Reference` fields, new bibliography bookmarks, or static reference numbers unless the user explicitly requests and approves bibliography reconstruction.
- Do not flatten surviving in-text citation fields to static text merely to silence missing-bookmark errors.
- Report the number and location of surviving citations and the fact that the bibliography is empty.

Only restore or generate bibliography entries when all of the following are true:

1. The user explicitly asks for bibliography restoration or generation.
2. The source list is author-supplied or explicitly approved.
3. Every record is verified to the requested standard.
4. The restoration method preserves or intentionally rebuilds citation numbering and cross-references.

## 5. Versioned pinpoint revision

1. Hash and record the source DOCX.
2. Copy it to a staging location.
3. Inventory paragraphs, tables, figures, OMML nodes, fields, bookmarks, headers, footers, and media.
4. Apply only requested changes.
5. Save a sibling revision with the requested date prefix/suffix; never overwrite the prior dated draft.
6. Run structural QA and open the result in Word.
7. Render to PDF/PNG for page-by-page QA when possible. If rendering fails after bounded attempts, disclose the limitation rather than repeatedly launching Word or claiming visual validation.
8. Hash-verify the delivered copy and confirm that the previous version remains present.

## 6. Acceptance checklist

- Correct section order and explicit introduction organization paragraph.
- Chinese abstract contains no unrequested English full-term or abbreviation parentheticals.
- Main text defines each non-universal abbreviation at its first occurrence.
- Multiple citations contain no unwanted commas.
- Citation superscripts/fields are preserved.
- Empty bibliography remains empty unless explicit restoration was authorized.
- No fabricated references or metadata.
- OMML, figures, tables, and unrelated content are unchanged.
- Prior draft remains intact and the new dated file is distinguishable.
- Word opens the final DOCX; structural and visual QA claims match the checks actually completed.
