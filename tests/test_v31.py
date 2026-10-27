import os
from pathlib import Path
from types import SimpleNamespace
import sys
import time
import tkinter as tk
from unittest.mock import patch
import pytest
from tkinterdnd2 import TkinterDnD, COPY, REFUSE_DROP
from autograder import grade, NS, W, CRITERIA
from grader import BASE
from test_automatic import compliant, mutate, setprop, paragraph, node


def test_many_toc_typography_errors_lose_only_one_mark(compliant):
    def change(roots):
        styles=roots['word/styles.xml']
        for sid,size in [('TOC1','24'),('TOC2','22'),('TOC3','20')]:
            s=styles.xpath('./w:style[@w:styleId=$s]',namespaces=NS,s=sid)[0]
            setprop(s,'rPr/rFonts',ascii='Cambria',hAnsi='Cambria');setprop(s,'rPr/sz',val=size)
        # Include entries at every TOC level so all three styles are assessed.
        first=roots['word/document.xml'].xpath('//w:fldSimple[starts-with(@w:instr,"TOC \\o")]',namespaces=NS)[0]
        p=first.getparent();body=p.getparent();index=list(body).index(p)
        cached=list(first);code=first.get(W+'instr');p.remove(first)
        for kind in ('begin','instruction','separate'):
            run=node('r');p.append(run)
            if kind=='instruction':child=node('instrText');child.text=code;run.append(child)
            else:run.append(node('fldChar',fldCharType=kind))
        for run in cached:p.append(run)
        q2=paragraph('Bear Life Cycle 5','TOC2');q3=paragraph('Anatomy 6','TOC3')
        end=node('r');end.append(node('fldChar',fldCharType='end'));q3.append(end)
        body.insert(index+1,q2);body.insert(index+2,q3)
    result=grade(mutate(compliant,change),'toc.docx');row=result['criteria'][3]
    assert row['score']==4
    assert sum(c['passed'] is False and c['name'].endswith('style font') for c in row['checks'])==3
    assert sum(c['passed'] is False and c['name'].endswith('style size') for c in row['checks'])==2
    assert [g['weight'] for g in row['scoring_groups']]==[2,1,1,1]
    assert [g['passed'] for g in row['scoring_groups']]==[True,True,False,True]


def test_missing_automatic_toc_still_gets_other_group_credit(compliant):
    def change(roots):
        for f in roots['word/document.xml'].xpath('//w:fldSimple',namespaces=NS):
            if f.get(W+'instr','').startswith('TOC \\o'):
                p=f.getparent()
                for child in list(f):p.append(child)
                p.remove(f)
    row=grade(mutate(compliant,change),'manual.docx')['criteria'][3]
    assert 0<row['score']<=3
    assert not row['scoring_groups'][0]['passed']


def test_missing_figure_word_preserves_presence_and_placement(compliant):
    def change(roots):
        for p in roots['word/document.xml'].xpath('//w:p[w:pPr/w:pStyle/@w:val="Caption"]',namespaces=NS):
            for t in p.findall('.//w:t',NS):t.text=(t.text or '').replace('Figure ','')
    row=grade(mutate(compliant,change),'labels.docx')['criteria'][4]
    assert row['score']==3
    assert [g['passed'] for g in row['scoring_groups']]==[True,True,False]
    assert any(c['name']=='Figure label shown' and not c['passed'] for c in row['checks'])


def test_bold_caption_is_only_a_formatting_group_deduction(compliant):
    def change(roots):
        s=roots['word/styles.xml'].xpath('./w:style[@w:styleId="Caption"]',namespaces=NS)[0]
        setprop(s,'rPr/b',val='1')
    row=grade(mutate(compliant,change),'bold.docx')['criteria'][4]
    assert row['score']==3
    assert [g['passed'] for g in row['scoring_groups']]==[True,True,False]


def test_citation_presence_and_style_credit_survives_empty_first_source(compliant):
    def change(roots):
        source=roots['customXml/item1.xml'].find('b:Source',NS)
        for c in list(source):
            if c.tag!='{'+NS['b']+'}Tag':source.remove(c)
    row=grade(mutate(compliant,change),'sources.docx')['criteria'][7]
    assert row['score']==3
    assert [g['weight'] for g in row['scoring_groups']]==[2,1,1,1]
    assert [g['passed'] for g in row['scoring_groups']]==[True,True,False,False]


def test_wrong_citation_style_only_loses_one_mark(compliant):
    def change(roots):
        source=roots['customXml/item1.xml'];source.set('StyleName','MLA');source.set('Version','7');source.set('SelectedStyle','MLA.xsl')
    assert grade(mutate(compliant,change),'style.docx')['criteria'][7]['score']==4


native_display=bool(os.environ.get('DISPLAY')) or sys.platform=='win32'

@pytest.fixture
def native_app():
    if not native_display:pytest.skip('Native drop test requires a desktop display')
    if tk.TkVersion>=9 and sys.platform!='win32':pytest.skip('tkdnd requires Tcl/Tk 8.6; use the Tk 8.6 validation environment')
    from desktop import GraderApp
    root=TkinterDnD.Tk();root.withdraw();app=GraderApp(root);root.update()
    try:yield root,app
    finally:root.destroy()


def tcl_files(root,paths):return root.tk.call('format','%s',root.tk.call('list',*[str(p) for p in paths]))


def test_native_drop_targets_registered_and_course_branding(native_app):
    root,app=native_app
    assert app.drop_enabled and root.TkdndVersion
    for widget in [root,app.batch,app.feedback,app.add]:
        assert root.tk.call('bind',str(widget),'<<Drop>>')
    assert 'BCIT Word Project Grader' in root.title()
    assert 'BSYS 1000' in root.title() and 'Fall 2026' in root.title()


def test_native_drop_batch_handles_spaces_unicode_and_filters_non_word(native_app,tmp_path,compliant):
    root,app=native_app
    files=[tmp_path/'Student One.docx',tmp_path/'Étudiant 二.docx']
    for p in files:p.write_bytes(compliant)
    invalid=tmp_path/'ignore.txt';invalid.write_text('ignore')
    lock=tmp_path/'~$Student.docx';lock.write_bytes(compliant)
    folder=tmp_path/'folder';folder.mkdir()
    data=tcl_files(root,[*files,invalid,lock,folder])
    assert app.on_file_drop(SimpleNamespace(data=data))==COPY
    deadline=time.monotonic()+30
    while app.busy and time.monotonic()<deadline:root.update();time.sleep(.02)
    assert not app.busy and len(app.results)==2
    assert {r['source_path'] for r in app.results}=={str(p.resolve()) for p in files}
    assert all('error' not in r for r in app.results)


def test_drop_while_busy_is_queued_and_graded(native_app,tmp_path,compliant):
    root,app=native_app
    first=tmp_path/'first.docx';second=tmp_path/'second.docx'
    first.write_bytes(compliant);second.write_bytes(compliant)
    app.start([first])
    assert app.on_file_drop(SimpleNamespace(data=tcl_files(root,[second])))==COPY
    assert app.pending_paths==[str(second.resolve())]
    deadline=time.monotonic()+30
    while (app.busy or app.pending_paths) and time.monotonic()<deadline:root.update();time.sleep(.02)
    assert not app.busy and not app.pending_paths
    assert [r['filename'] for r in app.results]==['first.docx','second.docx']


def test_non_word_drop_is_refused_without_grading(native_app,tmp_path):
    root,app=native_app
    p=tmp_path/'ignore.pdf';p.write_bytes(b'pdf')
    assert app.on_file_drop(SimpleNamespace(data=tcl_files(root,[p])))==REFUSE_DROP
    assert not app.results and not app.busy


@pytest.mark.parametrize('label',['Fig.','Picture','FigureX'])
def test_incorrect_caption_label_still_preserves_real_caption_credit(compliant,label):
    def change(roots):
        for p in roots['word/document.xml'].xpath('//w:p[w:pPr/w:pStyle/@w:val="Caption"]',namespaces=NS):
            for t in p.findall('.//w:t',NS):t.text=(t.text or '').replace('Figure',label)
    row=grade(mutate(compliant,change),'labels.docx')['criteria'][4]
    assert row['score']==3
    assert [g['passed'] for g in row['scoring_groups']]==[True,True,False]


def test_missing_space_after_figure_is_not_a_label_error(compliant):
    def change(roots):
        for p in roots['word/document.xml'].xpath('//w:p[w:pPr/w:pStyle/@w:val="Caption"]',namespaces=NS):
            for t in p.findall('.//w:t',NS):t.text=(t.text or '').replace('Figure ','Figure')
    assert grade(mutate(compliant,change),'spaces.docx')['criteria'][4]['score']==5


def test_course_branding_in_reports(compliant):
    from autograder import detail_text
    from feedback import feedback_text
    result=grade(compliant,'report.docx')
    assert result['application']=='BCIT Word Project Grader'
    assert result['course']=='BSYS 1000' and result['term']=='Fall 2026'
    assert 'BCIT Word Project Grader' in detail_text(result)
    assert 'BSYS 1000' in detail_text(result)
