"""Read-only, assignment-specific OOXML evidence extraction. No Word automation."""
from dataclasses import dataclass, asdict
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile, BadZipFile
import re
from lxml import etree as E

NS = {'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
      'b':'http://schemas.openxmlformats.org/officeDocument/2006/bibliography'}
W = '{'+NS['w']+'}'
CRITERIA = ['Headings', 'Body Text', 'Title Page', 'Table of Contents (TOC)',
 'Figures and Captions', 'List of Figures', 'Bulleted Lists', 'Citations (In-text)',
 'Bibliography / References', 'Page Numbers', 'Section Break and Page Numbering',
 'Branding Consistency']
BASE = Path(__file__).parent/'assets'/'Bears.docx'

@dataclass
class Finding:
    criterion: str
    status: str
    expected: str
    found: str
    suggested: int | None = None


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


def analyze(data: bytes, filename: str):
    parts=package(data)
    try:
        doc=xml(parts['word/document.xml']); styles=xml(parts['word/styles.xml'])
        roots={k:xml(v) for k,v in parts.items()}
    except (KeyError,E.XMLSyntaxError) as exc:
        raise ValueError('Document is missing valid Word document or style XML.') from exc
    ps=doc.xpath('//w:body/w:p',namespaces=NS)
    content=texts(doc); instructions=fields(doc)
    allfields=' '.join(fields(v) for k,v in roots.items() if k.startswith('word/'))
    def xp(node,path): return node.xpath(path,namespaces=NS)
    def style(sid):
        matches=xp(styles,f'//w:style[@w:styleId="{sid}"]')
        return matches[0] if matches else None
    def style_report(sid,size,color,italic=False):
        s=style(sid)
        if s is None: return f'{sid}: style missing'
        checks={'font':xp(s,'./w:rPr/w:rFonts/@w:ascii'), 'size':xp(s,'./w:rPr/w:sz/@w:val'), 'color':xp(s,'./w:rPr/w:color/@w:val')}
        ok=checks['font']==['Segoe UI'] and checks['size']==[str(size*2)] and [v.upper() for v in checks['color']]==[color]
        if italic: ok=ok and bool(xp(s,'./w:rPr/w:i[not(@w:val="0") and not(@w:val="false")]'))
        return f'{sid}: '+('matches explicit style settings' if ok else f'found {checks}; expected Segoe UI, {size} pt, #{color}'+(', italic' if italic else ''))
    out=[]
    def add(name,expected,found,missing=False):
        out.append(Finding(name,'Missing evidence' if missing else 'Review required',expected,found,0 if missing else None))
    reports=[style_report(s,n,'00796B') for s,n in [('Title',26),('Heading1',18),('Heading2',16),('Heading3',14)]]
    title=style('Title'); border=xp(title,'./w:pPr/w:pBdr/w:bottom/@w:sz') if title is not None else []
    add(CRITERIA[0],'Modify Title and Heading 1–3 styles; teal title bottom border, 1 pt; use appropriate hierarchy.', '\n'.join(reports)+f'\nTitle border size (eighth points): {border or "not explicit"}. Heading hierarchy needs review.')
    normal=sum(not xp(p,'./w:pPr/w:pStyle/@w:val') or xp(p,'./w:pPr/w:pStyle/@w:val')==['Normal'] for p in ps if xp(p,'.//w:t'))
    overrides=sum(bool(xp(p,'.//w:rPr/w:rFonts | .//w:rPr/w:sz | .//w:rPr/w:color')) for p in ps)
    add(CRITERIA[1],'Normal style: Segoe UI, 11 pt, #4E4E4E; body paragraphs use Normal.',style_report('Normal',11,'4E4E4E')+f'\n{normal} nonempty top-level paragraphs use Normal/default; {overrides} paragraphs contain direct font/size/color formatting. Review which are body text.')
    add(CRITERIA[2],'Separate manual title page: TITLE/DOCPROPERTY Title field, centered name, student number, set, blank lines.',f'Title field evidence: {bool(re.search(r"\bTITLE\b|DOCPROPERTY\s+\"?Title",instructions,re.I))}. Opening text: '+ ' | '.join(content[:10]))
    toc=bool(re.search(r'\bTOC\b',instructions,re.I))
    add(CRITERIA[3],'Automatic Word TOC, branded styles and leader dots; no unrelated content on its pages.',f'Field instructions: {instructions[:1500]}\nPage isolation and Automatic Table 1/2 appearance need visual review.',not toc)
    seq=len(re.findall(r'\bSEQ\s+Figure\b',instructions,re.I))
    captions=[t for t in content if re.match(r'Figure\s+\d',t,re.I)]
    add(CRITERIA[4],'Four real Figure SEQ captions above original figures, original titles retained; modified Caption style.',f'{seq} Figure sequence fields. Caption-like text: {captions}\n'+style_report('Caption',10,'4E4E4E',True)+'\nCheck image order and caption placement.',seq==0)
    figure_toc=bool(re.search(r'\bTOC\b[^\n]*?\\[ca]\s+"?Figure',instructions,re.I))
    add(CRITERIA[5],'Word-generated figure list, Heading 1 title, Segoe UI 11 gray, dots, dedicated page.',f'Figure-list field detected: {figure_toc}; List of Figures title: {any(t.strip()=="List of Figures" for t in content)}. Review field switches and page layout.',not figure_toc)
    numbering=roots.get('word/numbering.xml')
    bullets=xp(numbering,'//w:lvl[w:numFmt/@w:val="bullet"]/w:lvlText/@w:val') if numbering is not None else []
    list_details=[]
    if numbering is not None:
        for p in ps:
            pstyle=xp(p,'./w:pPr/w:pStyle/@w:val')
            stnode=style(pstyle[0]) if pstyle else None
            nums=xp(p,'./w:pPr/w:numPr/w:numId/@w:val')
            levels=xp(p,'./w:pPr/w:numPr/w:ilvl/@w:val')
            if not nums and stnode is not None:
                nums=xp(stnode,'./w:pPr/w:numPr/w:numId/@w:val')
                levels=xp(stnode,'./w:pPr/w:numPr/w:ilvl/@w:val')
            if not nums: continue
            num=xp(numbering,f'//w:num[@w:numId="{nums[0]}"]/w:abstractNumId/@w:val')
            if not num: continue
            lvl=xp(numbering,f'//w:abstractNum[@w:abstractNumId="{num[0]}"]/w:lvl[@w:ilvl="{levels[0] if levels else "0"}"]')
            if not lvl: continue
            lvl=lvl[0]
            label=''.join(xp(p,'.//w:t/text()'))[:85]
            list_details.append(f'{label}: style={pstyle}, marker={xp(lvl,"./w:lvlText/@w:val")}, format={xp(lvl,"./w:numFmt/@w:val")}, left indent={xp(lvl,"./w:pPr/w:ind/@w:left")}, marker color={xp(lvl,"./w:rPr/w:color/@w:val")}')
    add(CRITERIA[6],'Modify applied list style; coral circle/square bullets, 0.25-inch indent; Segoe UI 11 gray text.',f'Available bullet definitions: {bullets}.\n'+ '\n'.join(list_details)+ '\nVerify the numbering definition actually applied to each list, style inheritance, marker color and 360-twip indent.')
    sources=[v for k,v in roots.items() if k.startswith('customXml/') and v.tag=='{'+NS['b']+'}Sources']
    first=re.search(r'\bCITATION\s+(\S+)',instructions,re.I)
    # Display source metadata, not package contents or student identifiers.
    summaries=[]
    for r in sources:
        summaries.append('Bibliography style: '+str(r.get('SelectedStyle') or r.get('StyleName') or 'not declared'))
        for s in r.xpath('.//b:Source',namespaces=NS):
            tag=s.findtext('b:Tag',namespaces=NS)
            if first and tag==first.group(1): summaries.append('First citation source: '+' '.join(s.itertext()))
    add(CRITERIA[7],'APA 6; first citation source is book: Olson, S. F. (1956), The Singing Wilderness, New York, Alfred A. Knopf.',f'First citation tag: {first.group(1) if first else "none"}. '+'\n'.join(summaries)+'\nVerify exact source fields and rendered APA 6 citations.',first is None)
    bib=bool(re.search(r'\bBIBLIOGRAPHY\b',instructions,re.I))
    add(CRITERIA[8],'Word BIBLIOGRAPHY field on own page, formatted to guide.',f'Bibliography field detected: {bib}; page isolation requires visual review.',not bib)
    footers=[(k,v) for k,v in roots.items() if re.match(r'word/footer\d+\.xml$',k)]
    footer_summary=[]
    for k,v in footers:
        footer_summary.append(f'{k}: fields {fields(v)}; alignment {xp(v,"//w:jc/@w:val")}')
    sections=xp(doc,'//w:sectPr')
    add(CRITERIA[9],'Right-aligned PAGE and TITLE fields; Segoe UI 9 gray; title page counts as 1 but number hidden.', '\n'.join(footer_summary)+f'\nDifferent-first-page flags: {len(xp(doc,"//w:sectPr/w:titlePg"))}. Inspect first-page footer separately.',not footers)
    details=[]
    for i,s in enumerate(sections):
        details.append(f'Section {i+1}: break={xp(s,"./w:type/@w:val") or ["default nextPage"]}, numbering={[(x.get(W+"fmt"),x.get(W+"start")) for x in xp(s,"./w:pgNumType")]}, explicit footer references={len(xp(s,"./w:footerReference"))}')
    add(CRITERIA[10],'Next-page section before references; independent footer; upperLetter numbering starting 1 (A); earlier numbering retained.','\n'.join(details),len(sections)<2)
    margins=[{side:s.find('w:pgMar',NS).get(W+side) for side in ['top','bottom','left','right']} for s in sections if s.find('w:pgMar',NS) is not None]
    add(CRITERIA[11],'One-inch margins (1440 twips); consistent brand typography, colors, spacing and layout.',f'Section margins: {margins}. Theme fonts/colors and visual consistency need review. Light sand background and callouts are optional.')
    sentence={'status':'Review required','text':'Heading not found.','word_count':None}
    for i,t in enumerate(content):
        if t.strip()=='A Multifaceted Approach':
            following=next((x for x in content[i+1:] if x.strip()),'')
            baseline='Effective bear conservation requires a multifaceted approach. Safe and tested strategies must be used.'
            extra=following[len(baseline):].strip() if following.startswith(baseline) else None
            count=len(re.findall(r"\b[\w]+(?:[’'-][\w]+)*\b",extra or ''))
            sentence={'status':'Review required' if extra and count<=25 else 'Needs attention','text':extra if extra is not None else following,'word_count':count if extra is not None else None,'note':'Check exactly one context-appropriate sentence was appended. No rubric points assigned. Baseline paragraph changed.' if extra is None else 'Check exactly one context-appropriate sentence. No rubric points assigned.'}
            break
    baseline_parts=package(BASE.read_bytes())
    original=texts(xml(baseline_parts['word/document.xml']))
    removed=[t for t in original if t.strip() and t not in content and not t.startswith('Figure ')]
    return {'filename':filename,'findings':[asdict(f) for f in out], 'sentence':sentence,'changed_original_paragraphs':removed,'paragraph_count':len(ps),'field_instructions':allfields}
