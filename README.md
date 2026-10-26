# Bears Word Project Grader

A local teacher review app for the **Word Project 2026 / Together for Bears** assignment. Upload one or multiple `.docx` submissions, inspect OOXML evidence against the supplied rubric, assign 0/1/3/5 points per criterion, and export feedback, evidence JSON, or a batch CSV.

## Run

Python 3.11+ recommended. From this repository:

```bash
python -m venv .venv
source .venv/bin/activate
# Windows PowerShell: .venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
streamlit run app.py
```

Open the local address printed by Streamlit. No accounts, external services, Microsoft Word installation, or API keys are required. Student uploads remain in server memory; the app does not save them to disk. Run locally for local processing. Do not expose the server publicly without authentication and an appropriate deployment review.

## Grading approach

The 12 rubric rows total 60 points. The app extracts saved style definitions, direct formatting indicators, field instructions (including simple fields), numbering definitions, citation source metadata, footer fields, section settings, and margins. It compares the required added sentence with the original Bears paragraph and flags original paragraphs that no longer match exactly. Required citation edits can also cause text differences.

This is an **assisted grading tool**, not a fully automatic marker. Scores start unreviewed. Missing structural evidence may suggest zero, but every final score is a teacher decision. Review all rows before treating a grade as final. Draft exports are marked explicitly. Feedback and scores live in the current Streamlit session: download them before leaving. JSON exports are evidence records; importing/resuming them is not implemented.

Saved XML cannot prove a student's editing process. Style inheritance, theme resolution, applied numbering, first-page footer behavior, caption placement, heading hierarchy, exact pagination, and visual consistency require review in Microsoft Word. The app does not render a document preview. Context and exactly-one-sentence compliance require teacher judgment. Word counting treats internal apostrophes and hyphens as part of a word; verify edge cases against class policy. The extra sentence and filename checks carry no invented point allocation. The rubric's overall-level thresholds are inconsistent with its 60-point scale, so only points are reported. Optional background and callout styles are not mandatory.

The original `assets/Bears.docx` is the baseline, not a model answer. Assignment PDFs are bundled for reference. The student instruction prohibiting AI is assignment context; the grader itself calls no AI services and makes no AI-use determinations.

## Validate

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q
```

Malformed or oversized ZIP documents are rejected (15 MB upload, 60 MB expanded package, 2,000 entries). Uploaded content is not executed and XML external entity resolution/network access are disabled. This is a local instructor tool, not a hardened public upload service.
