from copy import deepcopy
import pytest
from autograder import grade, NS, W, CRITERIA
from branding import assess_brand
from feedback import feedback_text
from test_automatic import compliant, mutate, setprop


def brand(result):return result['criteria'][11]


def test_full_adherence_gets_full_brand_score(compliant):
    result=grade(compliant,'complete.docx')
    assert brand(result)['score']==5
    assert all(v==1 for v in brand(result)['overall_assessment']['domains'].values())


def uncoordinated(roots):
    for sid in ['Title','Heading1','Heading2','Heading3','Normal']:
        style=roots['word/styles.xml'].xpath('./w:style[@w:styleId=$s]',namespaces=NS,s=sid)[0]
        setprop(style,'rPr/rFonts',ascii='Calibri',hAnsi='Calibri');setprop(style,'rPr/color',val='4F81BD');setprop(style,'rPr/sz',val='22')
    for p in roots['word/document.xml'].findall('.//w:p',NS):
        style=p.find('w:pPr/w:pStyle',NS)
        if style is not None and style.get(W+'val') in ('Heading1','Heading2','Heading3'):style.set(W+'val','Normal')
    for sid in ['Caption','ListBullet','TOC1','TOC2','TOC3','TableofFigures','Footer']:
        for style in roots['word/styles.xml'].xpath('./w:style[@w:styleId=$s]',namespaces=NS,s=sid):
            setprop(style,'rPr/rFonts',ascii='Calibri',hAnsi='Calibri');setprop(style,'rPr/color',val='4F81BD');setprop(style,'rPr/sz',val='24')


def test_no_coordinated_style_system_gets_zero_even_with_good_margins(compliant):
    result=grade(mutate(compliant,uncoordinated),'uncoordinated.docx')
    assert brand(result)['score']==0
    assessment=brand(result)['overall_assessment']
    assert assessment['zero_benchmark_pattern']
    assert assessment['domains']['Margins and layout']==1
    content=feedback_text(result,CRITERIA[11])
    assert 'overall style-guide assessment' in content
    assert 'coordinated brand' in content


def test_directly_branded_body_alone_does_not_establish_brand(compliant):
    def change(roots):
        uncoordinated(roots)
        for p in roots['word/document.xml'].findall('.//w:p',NS):
            if ''.join(p.xpath('.//w:t/text()',namespaces=NS)).startswith('Bears play a vital ecological role'):
                for r in p.findall('.//w:r',NS):
                    setprop(r,'rPr/rFonts',ascii='Segoe UI',hAnsi='Segoe UI');setprop(r,'rPr/color',val='4E4E4E');setprop(r,'rPr/sz',val='22')
    assert brand(grade(mutate(compliant,change),'direct.docx'))['score']==0


def test_branding_does_not_average_other_rubric_marks(compliant):
    from autograder import Document
    result=grade(compliant,'complete.docx');doc=Document(compliant)
    rows=deepcopy(result['criteria'][:11])
    margins=[c for c in result['criteria'][11]['checks'] if ' margin' in c['name']]
    for row in rows:row['score']=0;row['strict_score']=0
    assert assess_brand(doc,rows,margins)['score']==5


def test_one_regional_typography_issue_does_not_automatically_lose_brand_mark(compliant):
    def change(roots):
        style=roots['word/styles.xml'].xpath('./w:style[@w:styleId="Heading1"]',namespaces=NS)[0]
        setprop(style,'rPr/sz',val='42')
    result=grade(mutate(compliant,change),'isolated.docx')
    assert result['criteria'][0]['score']==4 and brand(result)['score']==5


def test_evidence_is_informational_not_extra_itemized_deductions(compliant):
    result=grade(mutate(compliant,uncoordinated),'sample.docx');row=brand(result)
    assert [c['name'] for c in row['checks'] if c['scored'] and c['passed'] is False]==['Overall style-guide adherence']
    assert all(not c['scored'] for c in row['checks'] if c['name'].startswith('Overall evidence:'))


def test_original_student_identity_and_filename_do_not_control_brand_score(compliant):
    data=mutate(compliant,uncoordinated)
    assert brand(grade(data,'AnyOtherName.docx'))['score']==brand(grade(data,'Reference.docx'))['score']==0


def test_widespread_direct_body_formatting_affects_overall_brand(compliant):
    def change(roots):
        for p in roots['word/document.xml'].findall('.//w:p',NS):
            if ''.join(p.xpath('.//w:t/text()',namespaces=NS)).startswith('Bears play a vital ecological role'):
                for r in p.findall('.//w:r',NS):setprop(r,'rPr/rFonts',ascii='Arial',hAnsi='Arial')
    row=brand(grade(mutate(compliant,change),'visible.docx'))
    assert row['score']<5
    assert row['overall_assessment']['visible_body_adherence'][0]<.2
