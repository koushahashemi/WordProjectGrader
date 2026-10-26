from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
import pytest
from grader import package, xml, fields
from autograder import grade


def test_invalid_archive_and_missing_document():
    with pytest.raises(ValueError):grade(b'not a zip','bad.docx')
    b=BytesIO()
    with ZipFile(b,'w') as z:z.writestr('foo.xml','<foo/>')
    with pytest.raises(ValueError):grade(b.getvalue(),'bad.docx')


def test_malformed_xml_is_a_document_error():
    b=BytesIO()
    with ZipFile(b,'w') as z:
        z.writestr('word/document.xml','<broken')
        z.writestr('word/styles.xml','<styles/>')
    with pytest.raises(ValueError):grade(b.getvalue(),'bad.docx')


def test_package_expansion_limit():
    b=BytesIO()
    with ZipFile(b,'w',compression=ZIP_DEFLATED) as z:z.writestr('big.xml',b'a'*60_000_001)
    with pytest.raises(ValueError,match='too large'):package(b.getvalue())


def test_external_entities_are_not_resolved(tmp_path):
    secret=tmp_path/'test.txt';secret.write_text('must-not-be-loaded')
    root=xml(f'<!DOCTYPE doc [<!ENTITY ext SYSTEM "{secret.as_uri()}">]><doc>&ext;</doc>'.encode())
    assert 'must-not-be-loaded' not in ''.join(root.itertext())


def test_simple_fields():
    root=xml(b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:fldSimple w:instr="PAGE"/></w:document>')
    assert 'PAGE' in fields(root)
