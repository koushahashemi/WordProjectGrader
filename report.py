"""Readable HTML grading report: class summary, then one section per submission.

Instructor-facing. Each student section leads with marks lost and plain-language
feedback (the same wording as the feedback panel), then the hidden-instruction
sentence flag, then the full evidence checks collapsed underneath.
"""
import html
import re
from datetime import datetime

from app_info import APP_TITLE, COURSE_TEXT
from feedback import REFERENCES, feedback_items
from grader import CRITERIA

SHORT = ['Headings', 'Body', 'Title page', 'TOC', 'Captions', 'Figures list', 'Lists',
         'Citations', 'References', 'Footers', 'Sections', 'Branding']

e = html.escape


def student_name(filename):
    """'2171277-700986 - Nov 5 2026 801 PM_Amaru_BearsDocument.docx' -> 'Amaru'."""
    name = re.sub(r'\.docx$', '', filename, flags=re.I)
    name = re.sub(r'^\d+-\d+\s*-\s*.*?\b\d{3,4}\s*[AP]M[\s_-]*', '', name)   # Learning Hub download prefix
    name = re.sub(r'[_\s-]*Bears[_\s-]*Document$', '', name, flags=re.I)
    name = name.replace('_', ' ').strip()
    name = re.sub(r'(?<=[a-z])(?=[A-Z])', ' ', name)                          # NasirAhmad -> Nasir Ahmad
    return name or filename


def _anchor(i):
    return f's{i}'


def _score_class(score):
    return 'full' if score == 5 else 'mid' if score == 3 else 'low'


def _flag_badge(sentence):
    flag = (sentence or {}).get('flag', '')
    cls = {'Yes': 'flag-yes', 'No': 'flag-no', 'Check': 'flag-check'}.get(flag, 'flag-no')
    return f'<span class="badge {cls}">{e(flag or "-")}</span>'


def _check_label(c):
    if not c.get('scored', True):
        return 'info', 'Info'
    if c.get('accepted'):
        return 'ok', 'Accepted'
    if c['passed'] is None:
        return 'unk', 'Unverified'
    return ('ok', 'Pass') if c['passed'] else ('bad', 'Fail')


def _summary(results):
    graded = [r for r in results if 'error' not in r]
    avg = sum(r['total'] for r in graded) / len(graded) if graded else 0
    flagged = sum(1 for r in graded if r.get('sentence', {}).get('flag') == 'Yes')
    head = ''.join(f'<th title="{e(c)}">{e(s)}</th>' for c, s in zip(CRITERIA, SHORT))
    rows = []
    for i, r in enumerate(results):
        link = f'<a href="#{_anchor(i)}">{e(student_name(r["filename"]))}</a>'
        if 'error' in r:
            rows.append(f'<tr><td class="name">{link}</td><td colspan="{len(CRITERIA) + 2}" class="err">Could not grade: {e(r["error"])}</td></tr>')
            continue
        cells = ''.join(f'<td class="sc {_score_class(c["score"])}">{c["score"]}</td>' for c in r['criteria'])
        rows.append(f'<tr><td class="name">{link}</td><td class="total">{r["total"]}</td>{cells}'
                    f'<td>{_flag_badge(r.get("sentence"))}</td></tr>')
    stats = (f'<div class="stats"><div><b>{len(results)}</b><span>submissions</span></div>'
             f'<div><b>{avg:.1f}</b><span>average / 60</span></div>'
             f'<div><b>{flagged}</b><span>added sentence flagged</span></div></div>')
    return (f'<section class="summary"><h2>Class summary</h2>{stats}'
            f'<div class="tablewrap"><table class="grid"><thead><tr><th>Student</th><th>Total</th>{head}'
            f'<th>Added sentence</th></tr></thead><tbody>{"".join(rows)}</tbody></table></div>'
            f'<p class="legend"><span class="sc full">5</span> full marks <span class="sc mid">3</span> some issues '
            f'<span class="sc low">1 / 0</span> major issues. Click a name to jump to that student.</p></section>')


def _sentence_box(sentence):
    if not sentence:
        return ''
    flag = sentence.get('flag', 'Check')
    text = sentence.get('text', '')
    quote = f'<blockquote>“{e(text)}”</blockquote>' if text else ''
    return (f'<div class="sentence {("flag-yes" if flag == "Yes" else "flag-check" if flag == "Check" else "flag-no")}">'
            f'<div class="sentence-head">Hidden-instruction sentence: <b>{e(flag)}</b>'
            f'<span class="note">Instructor only · not scored · not in student feedback</span></div>'
            f'<p>{e(sentence.get("finding", ""))}</p>{quote}</div>')


def _lost(row):
    lost = 5 - row['score']
    items = feedback_items(row)
    band = ''
    if 'detailed_score' in row and row['detailed_score'] != row['score']:
        band = f'<span class="band">Detailed {row["detailed_score"]}/5 → rubric band {row["score"]}/5</span>'
    bullets = ''.join(f'<li>{e(m)}</li>' for m in items) or '<li>See the evidence checks below for this criterion.</li>'
    est = ' <span class="est">estimated</span>' if row.get('estimated') else ''
    return (f'<div class="crit {_score_class(row["score"])}"><div class="crit-head">'
            f'<span class="crit-name">{e(row["criterion"])}</span>'
            f'<span class="crit-score">{row["score"]}/5{est}</span>'
            f'<span class="crit-lost">−{lost} mark{"s" if lost != 1 else ""}</span></div>'
            f'<div class="ref">{e(REFERENCES.get(row["criterion"], ""))} {band}</div>'
            f'<ul>{bullets}</ul></div>')


def _evidence(row):
    trs = []
    order = {'bad': 0, 'unk': 1, 'info': 2, 'ok': 3}
    for c in sorted(row['checks'], key=lambda c: order[_check_label(c)[0]]):
        cls, label = _check_label(c)
        found = e(str(c.get('found', '')))
        if c.get('accepted'):
            found += f'<br><i>Accepted alternative: {e(c["accepted"])}</i>'
        trs.append(f'<tr><td><span class="st {cls}">{label}</span></td><td>{e(c["name"])}'
                   f'<div class="loc">{e(c.get("location", ""))}</div></td>'
                   f'<td>{e(str(c.get("expected", "")))}</td><td>{found}</td></tr>')
    failed = sum(1 for c in row['checks'] if _check_label(c)[0] == 'bad')
    return (f'<details><summary>{e(row["criterion"])} — {row["score"]}/5'
            f'<span class="muted"> · {len(row["checks"])} checks, {failed} failed</span></summary>'
            f'<table class="ev"><thead><tr><th>Result</th><th>Check</th><th>Expected</th><th>Found</th></tr></thead>'
            f'<tbody>{"".join(trs)}</tbody></table></details>')


def _student(i, r):
    name = e(student_name(r['filename']))
    top = f'<a class="top" href="#top">↑ Summary</a>'
    if 'error' in r:
        return (f'<section class="student" id="{_anchor(i)}"><div class="s-head"><div><h2>{name}</h2>'
                f'<div class="file">{e(r["filename"])}</div></div>{top}</div>'
                f'<p class="err">This document could not be graded: {e(r["error"])}</p></section>')
    lost_rows = [row for row in r['criteria'] if row['score'] < 5]
    full_rows = [row for row in r['criteria'] if row['score'] == 5]
    lost_html = ''.join(_lost(row) for row in lost_rows) or '<p class="allfull">No marks lost.</p>'
    full_html = (f'<p class="fulllist"><b>Full marks:</b> {e(", ".join(row["criterion"] for row in full_rows))}</p>'
                 if full_rows else '')
    est = '<span class="est">estimated</span>' if r.get('estimated') else ''
    return (f'<section class="student" id="{_anchor(i)}">'
            f'<div class="s-head"><div><h2>{name}</h2><div class="file">{e(r["filename"])}</div></div>'
            f'<div class="s-total"><b>{r["total"]}</b>/60 {est}</div></div>'
            f'{_sentence_box(r.get("sentence"))}'
            f'<div class="s-bar"><h3>Marks lost</h3>'
            f'<button class="copy" onclick="copyFeedback(this)">Copy feedback</button>{top}</div>'
            f'<div class="feedback">{lost_html}</div>{full_html}'
            f'<h3>Evidence checks</h3><div class="evidence">{"".join(_evidence(row) for row in r["criteria"])}</div>'
            f'</section>')


CSS = """
:root{--ink:#1d2327;--muted:#5f6b72;--line:#dde3e6;--bg:#f6f8f9;--card:#fff;--teal:#00796B;
--full:#e3f3ee;--full-ink:#0b5d4b;--mid:#fff3d6;--mid-ink:#7a5300;--low:#fde4e1;--low-ink:#9b1c12;}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:14px/1.5 "Segoe UI",system-ui,-apple-system,sans-serif}
header{background:var(--teal);color:#fff;padding:20px 28px}
header h1{margin:0;font-size:22px} header p{margin:2px 0 0;opacity:.9}
main{max-width:1100px;margin:0 auto;padding:20px 16px 60px}
section{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:20px 22px;margin:0 0 20px}
h2{margin:0 0 4px;font-size:19px} h3{margin:18px 0 8px;font-size:15px;color:var(--teal)}
.stats{display:flex;gap:12px;margin:10px 0 14px;flex-wrap:wrap}
.stats div{background:var(--bg);border-radius:8px;padding:8px 14px;min-width:120px}
.stats b{display:block;font-size:20px} .stats span{color:var(--muted);font-size:12px}
.tablewrap{overflow-x:auto}
table{border-collapse:collapse;width:100%}
.grid th,.grid td{border-bottom:1px solid var(--line);padding:6px 6px;text-align:center;white-space:nowrap}
.grid th{font-size:12px;color:var(--muted);font-weight:600}
.grid td.name{text-align:left;font-weight:600} .grid td.total{font-weight:700}
.sc{display:inline-block;min-width:26px;border-radius:5px;padding:1px 6px;font-weight:600}
td.sc{display:table-cell}
.full{background:var(--full);color:var(--full-ink)} .mid{background:var(--mid);color:var(--mid-ink)} .low{background:var(--low);color:var(--low-ink)}
.legend{color:var(--muted);font-size:12px;margin:10px 0 0} .legend .sc{margin:0 4px 0 10px}
.badge{border-radius:10px;padding:1px 9px;font-size:12px;font-weight:600}
.flag-yes{background:#ffe9a8;color:#6b4a00} .flag-no{background:#eef1f2;color:var(--muted)} .flag-check{background:#e3ecfb;color:#1f4f9a}
a{color:var(--teal)}
.s-head{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;border-bottom:1px solid var(--line);padding-bottom:10px}
.file{color:var(--muted);font-size:12px;word-break:break-all}
.s-total{font-size:14px;white-space:nowrap} .s-total b{font-size:30px;color:var(--teal)}
.est{font-size:11px;color:var(--muted);border:1px solid var(--line);border-radius:8px;padding:0 6px;margin-left:4px}
.sentence{border-radius:8px;padding:10px 14px;margin:14px 0 4px;border:1px solid var(--line)}
.sentence.flag-yes{background:#fff6d6;border-color:#f0d47a;color:var(--ink)}
.sentence.flag-no{background:var(--bg);color:var(--ink)} .sentence.flag-check{background:#eef4fd;color:var(--ink)}
.sentence p{margin:4px 0} .sentence .note{float:right;font-size:11px;color:var(--muted)}
blockquote{margin:6px 0 0;padding:6px 12px;border-left:3px solid #e0b93a;background:#fff;font-style:italic}
.s-bar{display:flex;align-items:center;gap:10px} .s-bar h3{flex:1}
.copy{border:1px solid var(--teal);background:#fff;color:var(--teal);border-radius:6px;padding:4px 10px;cursor:pointer;font:inherit;font-size:12px}
.top{font-size:12px;white-space:nowrap}
.crit{border-left:4px solid;border-radius:6px;padding:8px 12px;margin:0 0 10px;background:var(--card)}
.crit.mid{border-color:#e0a800;background:#fffbf0;color:var(--ink)} .crit.low{border-color:#d0453a;background:#fff7f6;color:var(--ink)}
.crit-head{display:flex;gap:12px;align-items:baseline;flex-wrap:wrap}
.crit-name{font-weight:700;flex:1} .crit-score{font-weight:700} .crit-lost{color:var(--low-ink);font-size:12px;font-weight:600}
.ref{color:var(--muted);font-size:12px} .band{margin-left:8px;font-style:italic}
.crit ul{margin:6px 0 2px;padding-left:20px} .crit li{margin:3px 0}
.allfull{background:var(--full);color:var(--full-ink);padding:8px 12px;border-radius:6px}
.fulllist{color:var(--muted);font-size:13px}
details{border:1px solid var(--line);border-radius:6px;margin:0 0 6px}
summary{cursor:pointer;padding:6px 10px;font-weight:600} .muted{color:var(--muted);font-weight:400}
.ev{font-size:12px} .ev th,.ev td{border-top:1px solid var(--line);padding:5px 8px;text-align:left;vertical-align:top}
.ev th{color:var(--muted)} .loc{color:var(--muted);font-size:11px}
.st{border-radius:4px;padding:0 6px;font-weight:600;white-space:nowrap}
.st.ok{background:var(--full);color:var(--full-ink)} .st.bad{background:var(--low);color:var(--low-ink)}
.st.unk{background:#e3ecfb;color:#1f4f9a} .st.info{background:#eef1f2;color:var(--muted)}
.err{color:var(--low-ink)}
.limits{color:var(--muted);font-size:12px}
@media print{body{background:#fff} header{color:#000;background:#fff;border-bottom:2px solid var(--teal)}
section{border:none;padding:0} .student{page-break-before:always} .copy,.top{display:none}
details{border:none} details>summary{list-style:none} .evidence{display:none}}
@media (max-width:640px){header{padding:16px} section{padding:14px} .s-head{flex-direction:column}}
"""

JS = """
function copyFeedback(btn){
  const sec=btn.closest('section');
  const name=sec.querySelector('h2').innerText, total=sec.querySelector('.s-total b').innerText;
  const parts=[name+' - Total: '+total+'/60',''];
  sec.querySelectorAll('.feedback .crit').forEach(c=>{
    parts.push(c.querySelector('.crit-name').innerText+': '+c.querySelector('.crit-score').innerText.replace('estimated','').trim());
    c.querySelectorAll('li').forEach(li=>parts.push('  - '+li.innerText));
    parts.push('');
  });
  if(!sec.querySelector('.feedback .crit')) parts.push('No marks lost.');
  const text=parts.join('\\n');
  const done=()=>{const t=btn.innerText;btn.innerText='Copied';setTimeout(()=>btn.innerText=t,1500);};
  if(navigator.clipboard&&navigator.clipboard.writeText){navigator.clipboard.writeText(text).then(done,()=>fallback(text,done));}
  else fallback(text,done);
}
function fallback(text,done){const a=document.createElement('textarea');a.value=text;document.body.appendChild(a);a.select();
  try{document.execCommand('copy');done();}catch(e){} a.remove();}
"""


def html_report(results, generated=None):
    """Return a self-contained HTML report for a batch of graded results."""
    generated = generated or datetime.now()
    version = next((r.get('version') for r in results if 'version' in r), '')
    limits = next((r.get('limitations') for r in results if r.get('limitations')), [])
    students = ''.join(_student(i, r) for i, r in enumerate(results))
    limits_html = ''.join(f'<li>{e(x)}</li>' for x in limits)
    return (f'<!doctype html><html lang="en"><head><meta charset="utf-8">'
            f'<meta name="viewport" content="width=device-width,initial-scale=1">'
            f'<title>Word Project Grading Report</title><style>{CSS}</style></head><body>'
            f'<header id="top"><h1>{e(APP_TITLE)} — Grading Report</h1>'
            f'<p>{e(COURSE_TEXT)} · Generated {generated:%b %d, %Y %I:%M %p} · Engine {e(str(version))}</p></header>'
            f'<main>{_summary(results)}{students}'
            f'<section class="limits"><h3>Limitations</h3><ul>{limits_html}</ul></section></main>'
            f'<script>{JS}</script></body></html>')
