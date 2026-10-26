# Bears Assignment Grader — Windows desktop app

Automatically grades `.docx` submissions for the **Word Project 2026 / Together for Bears** assignment. This version replaces the original manual-scoring browser prototype with a native desktop app. No browser, web server, API key, or AI service is required.

## Windows executable (recommended)

Open this repository's **Releases** page and download **BearsGrader.exe** from the latest desktop release. Double-click it. Python does not need to be installed. An unsigned app may trigger Windows reputation checks; verify the download came from this repository. The executable is built on Windows by the repository's GitHub Actions workflow, tested before upload, and attached to tagged releases. If a build fails, no executable is published: inspect the **Actions** tab. This project does not contain a checked-in executable.

## Use the app

1. Click **Add Word files…** and select multiple submissions, or **Add folder…** to grade all `.docx` files directly inside a folder.
2. Grading starts automatically. Each file receives 12 criterion scores (0, 1, 3, or 5), plus a total out of 60 in the batch table. Scroll the table horizontally to see every criterion.
3. Select a submission to see its rubric table. Select a criterion to see the measured checks, deductions, expected values, actual values, and paragraph locations. Failed checks appear first.
4. **Export grades CSV…** saves the full score table for the batch. **Save full report…** saves every finding as text or JSON.

No instructor scoring step is required. Files are read locally and never modified or uploaded. Results remain in memory until exported; closing the app clears the batch. One invalid submission is reported as an error and does not interrupt the others. There is no invented score for documents that cannot be read. Adding the same file again creates another row so different same-named submissions are not silently discarded.

## What is automatically checked

- Title/Heading 1–3 and Normal style definitions, inheritance, theme fonts/colors, text-run overrides, original heading applications and body-text coverage.
- Title field, centered title-page lines, information-line count and blank-line separation.
- Heading TOC versus figure-list TOC field codes, saved results, entry styles, leader dots and titles.
- Four real caption sequence fields, original caption wording, duplicate original labels, caption style, image placement and original image binaries. Recompressed images may be reported as changed.
- The numbering definitions actually applied to the original lists, list text styles, bullet shape/color and indent (including level overrides).
- APA sixth-edition source metadata, the first citation's Olson book details and citation-source resolution.
- Word bibliography fields and saved entry formatting.
- The footers referenced by each section, PAGE/TITLE fields, alignment, typography and first-page numbering configuration.
- Section boundaries, separate references footer, upper-letter numbering starting at A, earlier decimal numbering and section margins.
- The required appended sentence's presence, word count and sentence-ending count, reported separately without invented point deductions.

## Scoring policy and practical limits

The app assigns scores automatically. For criteria whose rubric uses qualitative descriptions, the deterministic mapping is **5 = all measured checks pass; 3 = at least 75% pass; 1 = at least one passes; 0 = none pass**. A missing primary feature receives 0 for that criterion. Recognizable manually created contents/figure lists receive partial credit (3 when entry typography matches, otherwise 1); a manually typed title on a separate cover page receives 1. An automatic Word field is required for full credit. Checks are not individually worth points; the pass fraction selects a rubric level. Every criterion's report records its actual policy and the passed/total measured checks. Some criteria intentionally check both style definitions and their application. These thresholds are implementation policy because the supplied rubric does not specify numeric boundaries for “mostly” and “inconsistent.”

Citations use the rubric's specific mapping: APA 6 + updated Olson source = 5; updated source only = 3; APA 6 only = 1; neither = 0. Missing/unresolved citation fields cannot earn 5. A typed bibliography without a Word field can earn 3 if its Bibliography style meets the guide, or 1 otherwise.

**Estimated scores:** the engine inspects OOXML, not Microsoft Word's rendered pages. It cannot reliably verify physical page isolation, visual balance, exactly which Word gallery preset/commands were used, whether title-page lines truly identify the student, or whether the added sentence is context-appropriate. Those limitations are shown in the results; unverified checks do not deduct points. Full marks for measured checks are therefore not proof of full visual compliance. No claim of pixel-perfect Microsoft Word grading is made. Original-image byte comparison may flag a visually unchanged recompressed image. Changes to the first 100 characters of repeated body paragraphs can affect body-coverage matching. Teachers can inspect flagged details but do not have to enter scores.

The extra sentence and filename have no separate rubric point allocation. The rubric's overall-level thresholds conflict with its 60-point total; the app reports the point total rather than inventing overall levels. The bundled `Bears.docx` is the original source, not a model answer. Assignment PDFs remain in `assets/` for reference.

## Run from source on Windows

Download the repository ZIP (**Code → Download ZIP**), **Extract All**, then double-click **Install and Run.bat**. It finds Python 3.11–3.13 or installs Python 3.12 using `winget`, creates `.venv-windows`, installs dependencies, and starts the desktop app. If `winget` is unavailable, the script gives manual Python installation instructions. Internet access is required for initial installation, but grading works offline.

To build your own executable, install Python 3.12 and double-click **Build Windows EXE.bat**. The output is `dist/BearsGrader.exe`. The Windows Actions workflow also tests the packaged executable, including its embedded assignment assets and desktop UI, before publishing it.

## Develop / test

```bash
python -m venv .venv
# Windows: .venv\Scripts\activate
source .venv/bin/activate
python -m pip install -r requirements-dev.txt
python -m pytest -q
python desktop.py
```

Tkinter is bundled with standard Windows Python. On Linux, install your distribution's Python Tk package and use a graphical session. `python desktop.py --self-test --output desktop-smoke.json` exercises the actual native UI and bundled baseline (requires a display).

Tests include a compliant structural fixture, deductions for formatting and footer defects, all four citation rubric levels, real versus fake TOCs, batch continuation after invalid files, formula-safe CSV names, XML entity protection, and package expansion limits. The engine limits files to 15 MB, expanded packages to 60 MB, and ZIP entries to 2,000. No uploaded macros or document code are executed.
