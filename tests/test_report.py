"""Formatted HTML report."""
import html
from test_automatic import compliant, mutate, NS
from autograder import grade
from feedback import feedback_text
from grader import BASE
from report import html_report, student_name


def test_student_name_cleanup():
    assert student_name('2171277-700986 - Nov 5 2026 801 PM_Amaru_BearsDocument.docx')=='Amaru'
    assert student_name('NasirAhmad_BearsDocument.docx')=='Nasir Ahmad'
    assert student_name('JordanLee_BearsDocument.docx')=='Jordan Lee'


def test_report_has_summary_feedback_and_errors(compliant):
    good=grade(compliant,'JordanLee_BearsDocument.docx')
    base=grade(BASE.read_bytes(),'Starter_BearsDocument.docx')
    bad={'filename':'<b>Broken</b>.docx','error':'Not a Word file'}
    page=html_report([good,base,bad])
    assert page.startswith('<!doctype html>') and 'Class summary' in page
    assert 'Jordan Lee' in page and 'No marks lost.' in page
    # every lost-mark message from the feedback panel appears in the report
    for line in feedback_text(base).splitlines():
        if line.startswith('• ') and 'not met.' not in line:
            assert html.escape(line[2:]) in page
    assert 'Could not grade: Not a Word file' in page
    assert '<b>Broken</b>' not in page and '&lt;b&gt;Broken&lt;/b&gt;' in page   # filenames are escaped


def test_report_shows_added_sentence_flag(compliant):
    def change(roots):
        for p in roots['word/document.xml'].findall('.//w:p',NS):
            ts=p.findall('.//w:t',NS)
            if ''.join(t.text or '' for t in ts).strip().startswith('Effective bear conservation requires'):
                ts[-1].text+=' Bears need safe corridors.';break
    page=html_report([grade(mutate(compliant,change),'x.docx')])
    assert 'Hidden-instruction sentence: <b>Yes</b>' in page and 'Bears need safe corridors.' in page
