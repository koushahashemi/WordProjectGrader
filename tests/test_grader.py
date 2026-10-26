from io import BytesIO
from zipfile import ZipFile
from pathlib import Path
import pytest
from grader import analyze, fields, xml, BASE
from streamlit.testing.v1 import AppTest


def altered(replacements):
    result=BytesIO()
    with ZipFile(BASE) as source, ZipFile(result,'w') as dest:
        for name in source.namelist():
            data=source.read(name)
            if name=='word/document.xml':
                root=xml(data)
                ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                for paragraph in root.xpath('//w:p',namespaces=ns):
                    nodes=paragraph.xpath('.//w:t',namespaces=ns)
                    text=''.join(n.text or '' for n in nodes)
                    for a,b in replacements:
                        if text==a and nodes:
                            nodes[0].text=b
                            for n in nodes[1:]: n.text=''
                from lxml import etree
                data=etree.tostring(root)
            dest.writestr(name,data)
    return result.getvalue()


def test_original_is_not_model_answer():
    r=analyze(BASE.read_bytes(),'Bears.docx')
    assert len(r['findings'])==12
    by={f['criterion']:f for f in r['findings']}
    assert by['Figures and Captions']['suggested']==0
    assert by['Section Break and Page Numbering']['suggested']==0
    assert r['sentence']['word_count']==0
    assert r['changed_original_paragraphs']==[]
    assert 'Placeholder' in by['Citations (In-text)']['found']


def test_added_sentence_and_length():
    original='Effective bear conservation requires a multifaceted approach. Safe and tested strategies must be used.'
    r=analyze(altered([(original,original+' Protecting habitat helps bears survive.')]),'Student_BearsDocument.docx')
    assert r['sentence']['word_count']==5
    assert r['sentence']['text']=='Protecting habitat helps bears survive.'
    r=analyze(altered([(original,original+' '+' '.join(['bear']*26)+'.')]),'long.docx')
    assert r['sentence']['status']=='Needs attention'


def test_invalid_and_missing_parts():
    with pytest.raises(ValueError): analyze(b'not a zip','bad.docx')
    b=BytesIO()
    with ZipFile(b,'w') as z: z.writestr('foo.xml','<foo/>')
    with pytest.raises(ValueError): analyze(b.getvalue(),'bad.docx')


def test_simple_fields():
    root=xml(b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:fldSimple w:instr="PAGE"/></w:document>')
    assert 'PAGE' in fields(root)


def test_ui_starts():
    app=AppTest.from_file(str(Path(__file__).parents[1]/'app.py')).run()
    assert not app.exception
    assert 'Bears' in app.title[0].value


def test_ui_review_and_exports():
    import hashlib
    raw=BASE.read_bytes();name='Test_BearsDocument.docx'
    key=hashlib.sha256(raw).hexdigest()[:24]+hashlib.sha256(name.encode()).hexdigest()[:8]
    # Exercise the real review UI with the uploader's output replaced by a local fixture.
    source=(Path(__file__).parents[1]/'app.py').read_text()
    source=source.replace("uploads=st.file_uploader('Student Word documents',type=['docx'],accept_multiple_files=True,help='Maximum 15 MB per document; expanded XML packages limited to 60 MB.')", "class Upload:\n    name='Test_BearsDocument.docx'\n    def getvalue(self):\n        return (Path(__file__).parent/'assets'/'Bears.docx').read_bytes()\nuploads=[Upload(), Upload()]\nuploads[1].name='Other_BearsDocument.docx'")
    # AppTest.from_string has no repository-relative __file__, so use absolute asset paths.
    source=source.replace("Path(__file__).parent/'assets'",repr(str(BASE.parent)))
    source=source.replace("sorted(("+repr(str(BASE.parent))+").iterdir())","sorted(Path("+repr(str(BASE.parent))+").iterdir())")
    source=source.replace("("+repr(str(BASE.parent))+"/'Bears.docx')","(Path("+repr(str(BASE))+"))")
    app=AppTest.from_string(source).run()
    assert not app.exception
    app.selectbox(key=f'score-{key}-0').set_value(5).run()
    assert not app.exception
    assert app.metric[0].value=='5 / 60'
    assert app.metric[1].value=='1 / 12'

    app.selectbox[0].set_value(1).run()
    assert not app.exception
    assert app.metric[0].value=='0 / 60'
    app.selectbox[0].set_value(0).run()
    assert not app.exception
    assert app.metric[0].value=='5 / 60'
    assert app.selectbox(key=f'score-{key}-0').value==5
