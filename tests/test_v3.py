from pathlib import Path
from types import SimpleNamespace
import os
import sys
import tkinter as tk
from unittest.mock import patch
import pytest
from autograder import grade, grade_paths, NS, W, CRITERIA
from grader import BASE
from feedback import feedback_text
from test_automatic import compliant, mutate, setprop, node, ensure, paragraph


def scores(result):return {r['criterion']:r['score'] for r in result['criteria']}


@pytest.mark.parametrize('sid',['Title','Heading1','Heading2','Heading3'])
def test_one_bad_heading_style_is_four_only_under_headings(compliant,sid):
    def change(roots):
        style=roots['word/styles.xml'].xpath('./w:style[@w:styleId=$s]',namespaces=NS,s=sid)[0]
        setprop(style,'rPr/sz',val='42');setprop(style,'rPr/color',val='000000')
    result=grade(mutate(compliant,change),'heading.docx');s=scores(result)
    assert s['Headings']==4
    assert s['Title Page']==5 and s['Branding Consistency']==5
    assert all(s[c]==5 for c in CRITERIA if c!='Headings')
    assert len(result['criteria'][0]['scoring_groups'])==5
    assert [g['weight'] for g in result['criteria'][0]['scoring_groups']]==[1]*5


def test_heading_application_is_one_group_even_if_repeated(compliant):
    def change(roots):
        for p in roots['word/document.xml'].xpath('//w:p[w:pPr/w:pStyle/@w:val="Heading2" or w:pPr/w:pStyle/@w:val="Heading3"]',namespaces=NS):setprop(p,'pPr/pStyle',val='Normal')
    result=grade(mutate(compliant,change),'application.docx')
    assert scores(result)['Headings']==4
    assert [g['passed'] for g in result['criteria'][0]['scoring_groups']]==[True]*4+[False]


def test_title_application_is_not_a_heading_deduction(compliant):
    def change(roots):
        first=roots['word/document.xml'].find('w:body/w:p',NS);setprop(first,'pPr/pStyle',val='Normal')
    s=scores(grade(mutate(compliant,change),'cover.docx'))
    assert s['Headings']==5 and s['Branding Consistency']==5
    assert s['Title Page']==4


def test_title_direct_override_only_affects_cover(compliant):
    def change(roots):
        first=roots['word/document.xml'].find('w:body/w:p',NS)
        for r in first.findall('.//w:r',NS):setprop(r,'rPr/sz',val='22')
    s=scores(grade(mutate(compliant,change),'override.docx'))
    assert s['Headings']==5 and s['Title Page']==4 and s['Branding Consistency']==5


def test_normal_font_failure_is_not_duplicated_in_inheriting_lists(compliant):
    def change(roots):
        normal=roots['word/styles.xml'].xpath('./w:style[@w:styleId="Normal"]',namespaces=NS)[0]
        setprop(normal,'rPr/rFonts',ascii='Arial',hAnsi='Arial')
    result=grade(mutate(compliant,change),'normal.docx');s=scores(result)
    assert s['Body Text']==4
    # Individual criteria still avoid duplicate inheritance penalties. The overall
    # brand assessment can reflect a font problem affecting all visible body text.
    assert all(s[c]==5 for c in CRITERIA if c not in ('Body Text','Branding Consistency'))
    assert s['Branding Consistency']==4
    assert len(result['criteria'][1]['deductions'])==1
    assert 'Bulleted Lists:' not in feedback_text(result)


def test_repeated_direct_font_errors_count_once(compliant):
    def change(roots):
        for p in roots['word/document.xml'].xpath('//w:p',namespaces=NS):
            if ''.join(p.xpath('.//w:t/text()',namespaces=NS)).startswith('Bears play a vital ecological role'):
                for r in p.findall('.//w:r',NS):setprop(r,'rPr/rFonts',ascii='Arial',hAnsi='Arial')
    result=grade(mutate(compliant,change),'many.docx')
    assert scores(result)['Body Text']==4
    assert result['criteria'][1]['deductions']==['format:Normal:font']


def test_distinct_body_font_and_color_errors_cost_two(compliant):
    def change(roots):
        style=roots['word/styles.xml'].xpath('./w:style[@w:styleId="Normal"]',namespaces=NS)[0]
        setprop(style,'rPr/rFonts',ascii='Arial',hAnsi='Arial');setprop(style,'rPr/color',val='123456')
    result=grade(mutate(compliant,change),'two.docx')
    assert scores(result)['Body Text']==3
    assert len(result['criteria'][1]['deductions'])==2


def test_toc_title_using_heading_style_does_not_repeat_definition_error(compliant):
    def change(roots):
        doc=roots['word/document.xml']
        p=next(p for p in doc.findall('.//w:p',NS) if ''.join(p.xpath('.//w:t/text()',namespaces=NS))=='Contents')
        setprop(p,'pPr/pStyle',val='Heading1')
        style=roots['word/styles.xml'].xpath('./w:style[@w:styleId="Heading1"]',namespaces=NS)[0];setprop(style,'rPr/sz',val='42')
    s=scores(grade(mutate(compliant,change),'toc.docx'))
    assert s['Headings']==4 and s['Table of Contents (TOC)']==5


def test_layout_margin_problem_is_owned_by_branding(compliant):
    def change(roots):
        for section in roots['word/document.xml'].findall('.//w:sectPr',NS):setprop(section,'pgMar',left='2880')
    result=grade(mutate(compliant,change),'layout.docx');s=scores(result)
    assert s['Branding Consistency']==4
    assert all(s[c]==5 for c in CRITERIA if c!='Branding Consistency')
    assert result['criteria'][11]['overall_assessment']['domains']['Margins and layout']==.75


def split_cover(roots,start='2'):
    """Insert a one-page, footer-free cover section before the Contents page."""
    body=roots['word/document.xml'].find('w:body',NS)
    contents=next(p for p in body if ''.join(p.xpath('.//w:t/text()',namespaces=NS))=='Contents')
    end=paragraph();sec=ensure(end,'pPr/sectPr');setprop(sec,'pgMar',top='1440',bottom='1440',left='1440',right='1440');setprop(sec,'pgNumType',start='1',fmt='decimal')
    body.insert(list(body).index(contents),end)
    sections=body.findall('.//w:sectPr',NS)
    setprop(sections[1],'pgNumType',start=start,fmt='decimal')
    flag=sections[1].find('w:titlePg',NS)
    if flag is not None:sections[1].remove(flag)


@pytest.mark.parametrize('start',['2',None])
def test_footer_free_cover_followed_by_page_two_is_accepted(compliant,start):
    def change(roots):
        split_cover(roots)
        if start is None:
            sections=roots['word/document.xml'].findall('.//w:sectPr',NS)
            sections[1].find('w:pgNumType',NS).attrib.pop(W+'start',None)
    result=grade(mutate(compliant,change),'cover.docx')
    assert scores(result)['Title Page']==5 and scores(result)['Page Numbers']==5


def test_footer_free_cover_wrong_restart_loses_one_counting_mark(compliant):
    def change(roots):split_cover(roots,start='1')
    result=grade(mutate(compliant,change),'wrong.docx')
    assert scores(result)['Page Numbers']==4
    groups=result['criteria'][9]['scoring_groups']
    assert [g['passed'] for g in groups]==[True,True,False,True]


def test_page_break_without_cover_section_break_is_accepted(compliant):
    # The normal fixture already uses an explicit pageBreakBefore after its cover.
    result=grade(compliant,'pagebreak.docx')
    assert scores(result)['Title Page']==5
    assert scores(result)['Page Numbers']==5


def test_batch_retains_exact_original_paths_including_same_names(tmp_path,compliant):
    first=tmp_path/'a'/'Student.docx';second=tmp_path/'b'/'Student.docx'
    for p in [first,second]:p.parent.mkdir();p.write_bytes(compliant)
    results=grade_paths([first,second])
    assert results[0]['source_path']==str(first.resolve())
    assert results[1]['source_path']==str(second.resolve())
    assert results[0]['source_path']!=results[1]['source_path']


@pytest.mark.skipif(sys.platform!='win32' and not os.environ.get('DISPLAY'),reason='Requires desktop display')
def test_double_click_filename_opens_exact_original_and_handles_missing_file(tmp_path,compliant):
    from desktop import GraderApp
    p=tmp_path/'A name with spaces.docx';p.write_bytes(compliant)
    root=tk.Tk();root.withdraw()
    try:
        app=GraderApp(root);app.add_result(grade_paths([p])[0]);root.update()
        event=SimpleNamespace(x=10,y=40)
        with patch.object(app.batch,'identify_row',return_value='0'),patch.object(app.batch,'identify_column',return_value='#1'),patch('desktop.launch_document') as launch:
            app.open_submission(event);launch.assert_called_once_with(str(p.resolve()))
        with patch.object(app.batch,'identify_row',return_value='0'),patch.object(app.batch,'identify_column',return_value='#2'),patch('desktop.launch_document') as launch:
            app.open_submission(event);launch.assert_not_called()
        p.unlink()
        with patch('desktop.messagebox.showerror') as error:
            app.open_submission();error.assert_called_once()
    finally:root.destroy()


def test_page_two_must_have_an_actual_visible_footer(compliant):
    def change(roots):
        split_cover(roots)
        main=roots['word/document.xml'].findall('.//w:sectPr',NS)[1]
        ensure(main,'titlePg')  # Empty separate first-page footer hides the page-2 footer.
    result=grade(mutate(compliant,change),'missing-page2.docx')
    assert scores(result)['Page Numbers']<5
    assert not result['criteria'][9]['scoring_groups'][3]['passed']


def test_footer_inherited_normal_font_is_not_a_second_deduction(compliant):
    def change(roots):
        styles=roots['word/styles.xml']
        normal=styles.xpath('./w:style[@w:styleId="Normal"]',namespaces=NS)[0]
        setprop(normal,'rPr/rFonts',ascii='Arial',hAnsi='Arial')
        footer=styles.xpath('./w:style[@w:styleId="Footer"]',namespaces=NS)[0]
        fonts=footer.find('w:rPr/w:rFonts',NS);fonts.getparent().remove(fonts)
    s=scores(grade(mutate(compliant,change),'inherit.docx'))
    assert s['Body Text']==4 and s['Page Numbers']==5


def test_typed_cover_wrong_style_is_only_one_application_error(compliant):
    def change(roots):
        p=roots['word/document.xml'].find('w:body/w:p',NS)
        field=p.find('w:fldSimple',NS)
        for child in list(field):p.append(child)
        p.remove(field);setprop(p,'pPr/pStyle',val='Normal')
    result=grade(mutate(compliant,change),'typed-cover.docx')
    assert scores(result)['Title Page']==4 and scores(result)['Headings']==5


def test_missing_all_lists_does_not_receive_four_marks(compliant):
    def change(roots):
        styles=roots['word/styles.xml']
        for sid in ('ListBullet','ListParagraph'):
            for style in styles.xpath('./w:style[@w:styleId=$s]',namespaces=NS,s=sid):
                for num in style.findall('w:pPr/w:numPr',NS):num.getparent().remove(num)
        for num in roots['word/document.xml'].findall('.//w:numPr',NS):num.getparent().remove(num)
    assert scores(grade(mutate(compliant,change),'no-lists.docx'))['Bulleted Lists']==0
