# Bears Assignment Grader 3.0.0

A native Windows desktop app that automatically grades the **Word Project 2026 / Together for Bears** assignment. Batch results show all 12 rubric criteria and a total out of 60. No browser, server, account, API key, or AI service is required.

## Windows executable

Open **Releases**, download **BearsGrader.exe** from version 3.0.0, and double-click it. Replace the previous executable. Python does not need to be installed. Windows may show a reputation warning for an unsigned application; verify the download came from this repository.

GitHub Actions builds and tests the executable on Windows, including an actual desktop self-test and the embedded assignment assets, before publishing tagged releases. If a build fails, no new executable is published; inspect the Actions tab.

## Use

1. Choose **Add Word files...** for multiple submissions, or **Add folder...** for the `.docx` files directly inside a folder.
2. Each file is scored automatically. Scroll the batch table horizontally for all criteria.
3. **Double-click a filename** to open that exact original submission in the default Word-document application. Enter also opens the selected file. Same-named files from different folders retain their individual paths. Moved or deleted files produce an error message rather than opening a different submission.
4. Select a submission for its rubric table. **Feedback on lost marks** explains deductions in ordinary language with references to the instructions/style guide; **Technical details** contains the audit trail and scoring groups.
5. Export the batch CSV, plain-language feedback, or the complete text/JSON report before closing the app. Results are held in memory. Source paths are included in JSON reports so the original files remain identifiable.

Files are read locally, never modified or uploaded. A bad file is reported individually without stopping the batch. No instructor score-entry step is required.

## Version 3 grading ownership

The same underlying typography problem is not deducted again under another criterion.

| Criterion | Responsibility |
| --- | --- |
| Headings | Five equal one-mark groups: Title style definition; Heading 1 definition; Heading 2 definition; Heading 3 definition; Heading 1–3 application to document text. A style group includes its required font, size and color; Title also includes its border. |
| Body Text | Normal style typography, body-text application/overrides, and original body-text retention. |
| Title Page | Applying Title style to the cover title, title-property feature, student-information lines, centering, spacing and separate cover page. Inherited Title-definition errors remain under Headings. |
| Branding Consistency | Page margins and document-wide layout. Does not repeat Title, Normal or heading typography. |

Inherited Normal/heading-definition failures appear as **INFO** when encountered in another criterion. They remain visible in the audit trail but do not deduct there. A local conflicting override still belongs to the criterion containing that text. Repeated instances of the same failed property/item within a criterion count once. The rubric table's **Lost items** count reflects distinct deductions or failed scoring groups, not the number of affected paragraphs.

## Scoring

- **General criteria:** start at 5 and deduct one mark per distinct failed item/property, down to 0. One simple error gives **4/5**, not an automatic jump to 3. Missing primary features, such as no TOC or no references section, can still receive 0. Citation source fields and APA style use this same distinct-error rule; the old 5/3/1/0-only bands are no longer used.
- **Headings:** five equal one-mark groups as listed above. Repeating a wrong heading assignment does not multiply the one-mark application-group deduction. Applying Title style to the cover is evaluated only under Title Page.
- **Figures and Captions:** presence of four captions (2), placement above figures (1), formatting/style (2). Repeated failures stay within their group. Missing captions affect presence; placement/style are evaluated on the captions that exist. The supplied guide does not prescribe caption alignment, so observed alignment is reported without an invented alignment deduction.
- **Page Numbers:** four categories—font/size/style; right alignment; title page counted as 1 with number hidden; PAGE and Document Title features. Start at 5, deduct 1 per failed category, and award 0 if all four fail.

A **page break is sufficient after the cover**; a section break is not required there. A footer-free cover is accepted when it is followed by the correct visible footer displaying page 2. The app checks a separate cover section's start/continuation metadata and its following page's active footer. An empty different-first-page footer on page 2 is not accepted. This exception does not remove the separate-section requirement before the bibliography.

## Retained instructor calibration

The instructor's full-mark reference continues to score **60/60**, using reusable rules rather than a filename, student identity or hash. Student examples are not committed or bundled; regression fixtures use synthetic documents.

Accepted alternatives are documented as **ACCEPTED**, including a properly styled typed cover title, a correct border applied to title paragraphs, approved half-inch list positions and 1.25-inch side margins, the Birth list introduction as a heading before the remaining real list, and the previously approved references-footer/numbering variants under their benchmark conditions. Existing references-section detection and required checks remain; a page break alone before references is not a new section.

Both **List of Figures** and **Table of Figures** headings are accepted, with Heading 1 or TOCHeading style. Text comparison ignores punctuation, case and spacing and allows bounded minor spelling differences, while retaining meaningful numeric differences such as a wrong year. Fonts such as Segoe UI italic/Bold are equivalent to their base family with the relevant effect. Document-title controls are recognized inline or around paragraphs. Measurements account for hanging indents, unit conversions and small rounding differences; font sizes are displayed in points.

Text/JSON reports retain a separate written-requirement diagnostic score, which disables accepted alternatives while using the version 3 deduction scheme. It is not the original rubric's coarse level-band score. Developers can request it directly using `grade(data, filename, calibrated=False)`.

## Limits

The app inspects saved OOXML rather than rendering in Microsoft Word. Exact physical pagination, visual balance, the identity of student-information lines and contextual sentence quality cannot be fully established. Affected results remain marked **Estimated**; unverified checks do not deduct marks. Grading one full-mark reference correctly does not validate every possible submission.

The extra sentence is checked separately for presence/length/sentence endings; the rubric allocates it no separate points. Overall-level thresholds in the supplied rubric conflict with the 60-point total, so the app reports points. `assets/Bears.docx` is the original source, not a model answer. The PDFs in `assets/` are the assignment references.

## Run or build from source

On Windows, download the repository ZIP, **Extract All**, and double-click **Install and Run.bat**. It finds Python 3.11–3.13 or installs Python 3.12 with `winget`, creates an isolated environment, installs dependencies and starts the native app. If `winget` is unavailable, it explains manual installation. Installation needs internet; grading works offline.

To build locally, install Python 3.12 and run **Build Windows EXE.bat**. The output is `dist/BearsGrader.exe`.

For development:

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
python -m pip install -r requirements-dev.txt
python -m pytest -q
python desktop.py
```

Tkinter ships with normal Windows Python. Linux needs a Python Tk installation and graphical display. `python desktop.py --self-test --output desktop-smoke.json` exercises the native UI and bundled baseline. GUI tests skip when no display is available on Linux.

Tests cover grading ownership, equal heading groups, distinct-error 4/5 scores, repeated formatting, cover/footer variants, exact original-path opening, source metadata, captions, batch continuation, feedback exports and input validation. Limits: 15 MB per upload, 60 MB expanded ZIP size and 2,000 ZIP entries. XML external entity resolution/network access are disabled; uploaded document code is never executed by grading. Opening a file by double-click is a separate explicit action in the default application.
