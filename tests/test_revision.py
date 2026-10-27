import re
from pathlib import Path
import pytest
from lxml import etree as E
from autograder import grade, Document, NS, W, CRITERIA
from matching import same_text, twips, near, font_family
from feedback import feedback_text
from test_automatic import compliant, mutate, row, setprop, ensure, node


def test_text_comparison_ignores_punctuation_spaces_and_small_typos():
    assert same_text('Life Cycle of a Polar Bear.','Life Cycle of a Polar Bear')
    assert same_text('Alfred A.Knopf','Alfred A. Knopf')
    assert same_text('Brown bear in its natural enviroment.','Brown bear in its natural environment.')
    assert not same_text('1957','1956')
    assert not same_text('Penguins in Antarctica','Brown bear in its natural environment')


@pytest.mark.parametrize('value',['1440','1in','2.54cm','25.4mm','72pt'])
def test_word_measurements_convert_to_twips(value):assert twips(value)==pytest.approx(1440)


def test_rounding_tolerance_is_small():
    assert near('1437',1440)
    assert not near('1800',1440)
    assert not near('2880',1440)


@pytest.mark.parametrize('name',['Segoe UI italic','Segoe UI Bold','SegoeUI-BoldItalic','Segoe UI Semibold'])
def test_mac_face_names_are_equivalent(name):assert font_family(name)[0]=='Segoe UI'


def test_figures_heading_accepts_toc_heading(compliant):
    def change(roots):
        for p in roots['word/document.xml'].findall('.//w:p',NS):
            if ''.join(p.xpath('.//w:t/text()',namespaces=NS))=='List of Figures':setprop(p,'pPr/pStyle',val='TOCHeading')
    assert row(grade(mutate(compliant,change),'sample.docx'),CRITERIA[5])['score']==5


def test_rounded_marker_position_and_margin_alternative(compliant):
    def change(roots):
        for sty in roots['word/styles.xml'].xpath('./w:style[@w:styleId="ListBullet"]',namespaces=NS):setprop(sty,'pPr/ind',left='717',hanging='357')
        for sec in roots['word/document.xml'].findall('.//w:sectPr',NS):setprop(sec,'pgMar',top='1437',bottom='1440',left='1800',right='1800')
    result=grade(mutate(compliant,change),'measurements.docx')
    assert row(result,CRITERIA[6])['score']==5
    assert row(result,CRITERIA[11])['score']==5
    assert row(result,CRITERIA[11])['accepted_variants']>0


def test_unaccepted_margins_still_lose_marks(compliant):
    def change(roots):
        for sec in roots['word/document.xml'].findall('.//w:sectPr',NS):setprop(sec,'pgMar',left='2880')
    assert row(grade(mutate(compliant,change),'bad.docx'),CRITERIA[11])['score']<5


def test_block_level_title_control_is_recognized(compliant):
    def change(roots):
        for name,footer in roots.items():
            if not name.startswith('word/footer'):continue
            p=footer.find('w:p',NS)
            for field in p.findall('w:fldSimple',NS):
                if field.get(W+'instr')=='TITLE':
                    for r in list(field):p.append(r)
                    p.remove(field)
            control=node('sdt');setprop(control,'sdtPr/dataBinding',xpath='/p:coreProperties[1]/d:title[1]',prefixMappings="xmlns:d='http://purl.org/dc/elements/1.1/'")
            ensure(control,'sdtContent').append(p);footer.append(control)
    result=grade(mutate(compliant,change),'title-control.docx')
    assert row(result,CRITERIA[9])['score']==5


@pytest.mark.parametrize('failed,expected',[(0,5),(1,4),(2,3),(3,2),(4,0)])
def test_page_number_scoring_uses_four_groups(compliant,failed,expected):
    def change(roots):
        doc=roots['word/document.xml']
        if failed>=3:
            first=doc.findall('.//w:sectPr',NS)[0];first.remove(first.find('w:titlePg',NS))
        for name,footer in roots.items():
            if not name.startswith('word/footer'):continue
            for p in footer.findall('.//w:p',NS):
                if failed>=1:
                    for r in p.findall('.//w:r',NS):setprop(r,'rPr/sz',val='22')
                if failed>=2:setprop(p,'pPr/jc',val='left')
                if failed>=4:
                    for field in list(p.findall('w:fldSimple',NS)):
                        if field.get(W+'instr')=='TITLE':
                            for r in list(field):p.append(r)
                            p.remove(field)
    result=grade(mutate(compliant,change),'footer.docx');footer=row(result,CRITERIA[9])
    assert len(footer['scoring_groups'])==4
    assert footer['score']==expected
    assert sum(not g['passed'] for g in footer['scoring_groups'])==failed


def test_font_sizes_are_reported_in_points(compliant):
    def change(roots):
        p=next(p for p in roots['word/document.xml'].findall('.//w:p',NS) if ''.join(p.xpath('.//w:t/text()',namespaces=NS))=='A Multifaceted Approach')
        for r in p.findall('.//w:r',NS):setprop(r,'rPr/sz',val='28')
    result=grade(mutate(compliant,change),'size.docx')
    failures=[c for c in row(result,CRITERIA[0])['checks'] if not c['passed']]
    assert any('16 pt' in c['expected'] and '14 pt' in c['found'] for c in failures)
    assert not any("'size': '32'" in c['expected'] for c in failures)


def test_caption_groups_and_punctuation(compliant):
    def change(roots):
        for t in roots['word/document.xml'].findall('.//w:t',NS):
            if t.text and 'Life Cycle of a Polar Bear' in t.text:t.text=t.text+'.'
    result=grade(mutate(compliant,change),'caption.docx');captions=row(result,CRITERIA[4])
    assert captions['score']==5
    assert [g['weight'] for g in captions['scoring_groups']]==[2,1,2]


def test_caption_formatting_errors_cost_only_formatting_group(compliant):
    def change(roots):
        for p in roots['word/document.xml'].xpath('//w:p[w:pPr/w:pStyle/@w:val="Caption"]',namespaces=NS):
            for r in p.findall('.//w:r',NS):setprop(r,'rPr/sz',val='22')
    captions=row(grade(mutate(compliant,change),'caption.docx'),CRITERIA[4])
    assert captions['score']==3
    assert [g['passed'] for g in captions['scoring_groups']]==[True,True,False]


def test_caption_below_image_costs_placement_group(compliant):
    def change(roots):
        body=roots['word/document.xml'].find('w:body',NS)
        p=next(p for p in body if p.find('w:pPr/w:pStyle',NS) is not None and p.find('w:pPr/w:pStyle',NS).get(W+'val')=='Caption')
        i=list(body).index(p);image=body[i+1];body.remove(image);body.insert(i,image)
    captions=row(grade(mutate(compliant,change),'caption.docx'),CRITERIA[4])
    assert captions['score']==4
    assert [g['passed'] for g in captions['scoring_groups']]==[True,False,True]


def test_missing_caption_loses_presence_not_placement_twice(compliant):
    def change(roots):
        doc=roots['word/document.xml']
        p=next(p for p in doc.findall('.//w:p',NS) if p.find('w:pPr/w:pStyle',NS) is not None and p.find('w:pPr/w:pStyle',NS).get(W+'val')=='Caption')
        p.getparent().remove(p)
    captions=row(grade(mutate(compliant,change),'caption.docx'),CRITERIA[4])
    assert captions['score']==3
    assert [g['passed'] for g in captions['scoring_groups']]==[False,True,True]


def test_mac_font_face_formatting_passes(compliant):
    def change(roots):
        for p in roots['word/document.xml'].xpath('//w:p[w:pPr/w:pStyle/@w:val="Caption"]',namespaces=NS):
            for r in p.findall('.//w:r',NS):setprop(r,'rPr/rFonts',ascii='Segoe UI italic',hAnsi='Segoe UI italic')
    assert row(grade(mutate(compliant,change),'mac.docx'),CRITERIA[4])['score']==5


def test_missing_publisher_space_is_not_a_deduction(compliant):
    def change(roots):roots['customXml/item1.xml'].find('b:Source/b:Publisher',NS).text='Alfred A.Knopf'
    assert row(grade(mutate(compliant,change),'source.docx'),CRITERIA[7])['score']==5


def test_feedback_explains_linkage_without_xml_ids(compliant):
    def change(roots):
        last=roots['word/document.xml'].findall('.//w:sectPr',NS)[-1]
        last.remove(last.find('w:footerReference',NS))
    result=grade(mutate(compliant,change),'section.docx')
    content=feedback_text(result,CRITERIA[10])
    assert 'Link to Previous' in content and 'Instruction 13' in content
    assert 'rId' not in content and 'Distinct explicit' not in content
    assert 'Headings:' not in content


def test_feedback_reports_only_lost_items_and_deduplicates_footer(compliant):
    def change(roots):
        for name,footer in roots.items():
            if name.startswith('word/footer'):
                for r in footer.findall('.//w:r',NS):setprop(r,'rPr/sz',val='22')
    result=grade(mutate(compliant,change),'font.docx')
    content=feedback_text(result)
    assert '4/5' in content and 'Instruction 12' in content and '11 pt' in content
    assert content.count('Footer text does not match')==1
    assert 'Figures and Captions:' not in content
