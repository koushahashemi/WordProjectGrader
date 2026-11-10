# Original rubric version 2 — 2.0.0-original-rubric

Separate `original_rubric` branch; main scoring is unchanged.
Each rubric item awards only 5, 3, 1 or 0. Detailed scores map 5→5,
4→3, 3→3, 2→1, 1→1, 0→0. Existing calibration, accepted alternatives,
partial-credit evidence, duplicate-error handling and batch features are retained.
Reports explain the band conversion; totals use the converted item scores.

New: formatted HTML report. "Save full report..." now defaults to an .html file that opens in
the browser: class summary table (colour-coded scores, added-sentence flag, links to each student),
then one section per student with marks lost and plain-language feedback first, a Copy feedback
button, and the full evidence checks collapsed underneath. Prints one student per page. Text and
JSON exports are still available from the same dialog.

New: hidden-instruction sentence check (instructor only, not scored). Flags text added after
the first body paragraph under "A Multifaceted Approach", either at the end of the paragraph or
as its own paragraph. Shown as Yes / No / Check in the results list (Yes rows highlighted), in the
detailed report, and as three new columns at the end of the grades CSV (flag, word count, text).
Never included in student feedback.

Fix: Title Page now reads the cover as displayed lines. Name, student number and set
separated by Shift+Enter line breaks (instead of separate paragraphs) count as separate
lines, and blank line breaks count as blank lines between them.

---

## BCIT Word Project Grader 3.1.0
### BSYS 1000 | Fall 2026 | Instructor Grading Tool

- Drag and drop one or multiple Word files anywhere in the app to grade them. Drops during grading are queued; non-Word and temporary lock files are filtered.
- BCIT course title, description, report headings, JSON metadata and export filenames.
- TOCs retain credit for automatic generation, title and structure even when entry typography is wrong (2 + 1 + 1 + 1 groups).
- Existing Word citations retain presence credit. APA 6 and first-source details are assessed separately; an empty placeholder no longer erases all credit.
- Captions without the Figure label are recognized through their Word fields. Label and bold/non-italic problems affect only the two-mark formatting group.
- The supplied examples receive 4/5 for the automatic TOC, 3/5 for citations with an empty first source, and 3/5 for correctly placed captions with presentation errors.
- The full-mark reference remains 60/60. Student submissions and identities are not bundled or committed.

Download BearsGrader.exe (BCIT Word Project Grader) and replace the previous executable. Python and a browser are not required. Exact rendered pagination remains an XML-only limitation.
