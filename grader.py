"""Safe shared OOXML package readers and assignment constants."""
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, BadZipFile
from lxml import etree as E

NS = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'b':'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'}
W = '{'+NS['w']+'}'
CRITERIA = ['Headings', 'Body Text', 'Title Page', 'Table of Contents (TOC)',
 'Figures and Captions', 'List of Figures', 'Bulleted Lists', 'Citations (In-text)',
 'Bibliography / References', 'Page Numbers', 'Section Break and Page Numbering',
 'Branding Consistency']
BASE = Path(__file__).parent/'assets'/'Bears.docx'

def xml(data):
    return E.fromstring(data, parser=E.XMLParser(resolve_entities=False, no_network=True))


def package(data):
    try:
        with ZipFile(BytesIO(data)) as z:
            infos=z.infolist()
            if len(infos)>2000 or sum(i.file_size for i in infos)>60_000_000:
                raise ValueError('Document is too large to inspect safely (60 MB expanded limit).')
            if any(i.flag_bits & 1 for i in infos):
                raise ValueError('Password-protected documents are unsupported.')
            return {i.filename:z.read(i) for i in infos if i.filename.endswith('.xml')}
    except (BadZipFile, RuntimeError, E.XMLSyntaxError) as exc:
        raise ValueError('Upload a valid, unencrypted .docx file.') from exc


def texts(root):
    return [] if root is None else [''.join(p.xpath('.//w:t/text()',namespaces=NS)) for p in root.xpath('.//w:body/w:p',namespaces=NS)]


def fields(root):
    # Field instructions may span several runs; simple fields use an attribute.
    return ' '.join(root.xpath('//w:instrText/text() | //w:fldSimple/@w:instr',namespaces=NS))
