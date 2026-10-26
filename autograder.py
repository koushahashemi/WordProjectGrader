"""Deterministic Bears-assignment scoring, with a check-level audit trail."""
from collections import Counter
from dataclasses import dataclass, asdict
from difflib import SequenceMatcher
from pathlib import Path
import csv
import io
import re
from lxml import etree as ET
from grader import BASE, CRITERIA, NS, W, package, xml

VERSION = '2.0.1'
A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
TEAL, GRAY, CORAL = '00796B', '4E4E4E', 'FF6F61'


def norm(text):
    return re.sub(r'\s+', ' ', text).strip()


def text(node):
    return norm(''.join(node.xpath('.//w:t/text()', namespaces=NS)))


def val(node, path, default=None, attr='val'):
    found = node.find(path, NS) if node is not None else None
    return found.get(W+attr, default) if found is not None else default


def enabled(node):
    return node is not None and node.get(W+'val', '1') not in ('0', 'false', 'off')


@dataclass
class Check:
    name: str
    passed: bool | None
    expected: str
    found: str
    location: str = ''


class Document:
    def __init__(self, data):
        if len(data) > 15_000_000:
            raise ValueError('File exceeds the 15 MB upload limit.')
        try:
            self.parts = package(data)
            self.roots = {k: xml(v) for k, v in self.parts.items()}
            self.doc = self.roots['word/document.xml']
            self.styles = self.roots['word/styles.xml']
        except (KeyError, ET.XMLSyntaxError) as exc:
            raise ValueError('Missing or malformed Word document/styles XML.') from exc
        self.ps = self.doc.xpath('//w:body//w:p', namespaces=NS)
        self.index = {p: i for i, p in enumerate(self.ps)}
        self.byid = {s.get(W+'styleId'): s for s in self.styles.findall('w:style', NS)}
        self.semantic = {}
        for sid, s in self.byid.items():
            for name in (sid, val(s, 'w:name', '')):
                self.semantic[re.sub(r'[\s_-]', '', name).casefold()] = sid
        self.default_style = next((sid for sid, s in self.byid.items() if s.get(W+'default') == '1' and s.get(W+'type') == 'paragraph'), 'Normal')
        self.theme = self.roots.get('word/theme/theme1.xml')
        self.defaults = self.styles.find('w:docDefaults', NS)
        self.field_list = self.parse_fields(self.doc)
        self.sections = self.doc.xpath('//w:sectPr', namespaces=NS)

    def sid(self, name):
        return self.semantic.get(re.sub(r'[\s_-]', '', name).casefold(), name)

    def pstyle(self, p):
        return val(p, 'w:pPr/w:pStyle', self.default_style)

    def chain(self, sid):
        result, seen = [], set()
        while sid in self.byid and sid not in seen:
            seen.add(sid)
            node = self.byid[sid]
            result.insert(0, node)
            sid = val(node, 'w:basedOn')
        return result

    def layers(self, sid, p=None, run=None, kind='rPr'):
        layers = []
        if self.defaults is not None:
            layers.append(self.defaults.find(f'w:{kind}Default/w:{kind}', NS))
        layers += [s.find('w:'+kind, NS) for s in self.chain(sid)]
        if run is not None:
            rid = val(run, 'w:rPr/w:rStyle')
            layers += [s.find('w:'+kind, NS) for s in self.chain(rid)]
        if p is not None and kind == 'pPr':
            layers.append(p.find('w:pPr', NS))
        if run is not None and kind == 'rPr':
            layers.append(run.find('w:rPr', NS))
        return [n for n in layers if n is not None]

    def prop(self, sid, prop, attr='val', p=None, run=None, kind='rPr'):
        result = None
        for layer in self.layers(sid, p, run, kind):
            node = layer.find('w:'+prop, NS)
            if node is not None and W+attr in node.attrib:
                result = node.get(W+attr)
        return result

    def color(self, node):
        if node is None:
            return None
        color = node.get(W+'color') or node.get(W+'val')
        theme = node.get(W+'themeColor')
        if theme and self.theme is not None:
            mapping = {'text1':'dk1', 'text2':'dk2', 'background1':'lt1', 'background2':'lt2'}
            entry = self.theme.find('.//'+A+'clrScheme/'+A+mapping.get(theme, theme))
            if entry is not None and len(entry):
                color = entry[0].get('lastClr') or entry[0].get('val')
                try:
                    rgb = [int(color[i:i+2], 16) for i in (0,2,4)]
                    shade, tint = node.get(W+'themeShade'), node.get(W+'themeTint')
                    if shade: rgb = [round(c*int(shade,16)/255) for c in rgb]
                    if tint: rgb = [round(c+(255-c)*int(tint,16)/255) for c in rgb]
                    color = ''.join(f'{c:02X}' for c in rgb)
                except (TypeError, ValueError):
                    return None
        return color.upper() if color else None

    def format(self, sid, p=None, run=None):
        result = {'font':None, 'size':None, 'color':None, 'italic':False}
        for layer in self.layers(sid, p, run):
            fonts = layer.find('w:rFonts', NS)
            if fonts is not None:
                # Theme declarations override the fallback font in Word.
                theme = fonts.get(W+'asciiTheme') or fonts.get(W+'hAnsiTheme')
                font = fonts.get(W+'ascii') or fonts.get(W+'hAnsi')
                if theme and self.theme is not None:
                    family = 'majorFont' if theme.startswith('major') else 'minorFont'
                    latin = self.theme.find('.//'+A+family+'/'+A+'latin')
                    font = latin.get('typeface') if latin is not None else None
                if theme or font: result['font'] = font
            size = layer.find('w:sz', NS)
            if size is not None: result['size'] = size.get(W+'val')
            color = layer.find('w:color', NS)
            if color is not None: result['color'] = self.color(color)
            italic = layer.find('w:i', NS)
            if italic is not None:
                # OOXML style italics toggle inherited state; direct italics set it.
                is_style = layer.getparent() is not None and layer.getparent().tag == W+'style'
                if is_style and enabled(italic): result['italic'] = not result['italic']
                elif not is_style: result['italic'] = enabled(italic)
        return result

    def parse_fields(self, root):
        records, stack = [], []
        for node in root.iter():
            if not isinstance(node.tag, str): continue
            p = next((a for a in node.iterancestors() if a.tag == W+'p'), None)
            if node.tag == W+'fldSimple':
                records.append({'code':node.get(W+'instr',''), 'p':p, 'node':node, 'ps':set([p]) if p is not None else set()})
            elif node.tag == W+'fldChar':
                typ = node.get(W+'fldCharType')
                if typ == 'begin':
                    record = {'code':'', 'p':p, 'node':node, 'ps':set(), 'reading':True}
                    stack.append(record); records.append(record)
                elif typ == 'separate' and stack: stack[-1]['reading'] = False
                elif typ == 'end' and stack: stack.pop()
            elif node.tag == W+'instrText' and stack and stack[-1]['reading']:
                stack[-1]['code'] += node.text or ''
            if p is not None:
                for record in stack: record['ps'].add(p)
        return records

    def fields(self, pattern):
        return [f for f in self.field_list if re.search(pattern, f['code'], re.I)]

    def location(self, p):
        return f'Paragraph {self.index.get(p, -1)+1}: {text(p)[:100]}' if p is not None else 'Document'

    def page_boundary_before(self, p):
        if p not in self.index: return False
        index = self.index[p]
        for q in reversed(self.ps[:index+1]):
            if q is p:
                sid = self.pstyle(q)
                if self.prop(sid, 'pageBreakBefore', p=q, kind='pPr') not in (None,'0','false'): return True
                if any(enabled(n) for n in [q.find('w:pPr/w:pageBreakBefore',NS)]): return True
            else:
                if q.find('w:pPr/w:sectPr', NS) is not None:
                    typ=val(q,'w:pPr/w:sectPr/w:type','nextPage')
                    if typ in ('nextPage','oddPage','evenPage'): return True
                if q.xpath('.//w:br[@w:type="page"]',namespaces=NS): return True
                if text(q): return False
        return index == 0

    def list_level(self, p):
        sid = self.pstyle(p)
        num = self.prop(sid,'numPr/w:numId',p=p,kind='pPr')
        level = self.prop(sid,'numPr/w:ilvl',p=p,kind='pPr') or '0'
        root = self.roots.get('word/numbering.xml')
        if root is None or num in (None,'0'): return None
        # Variable matching avoids interpolating document-controlled text in XPath.
        numbers = root.xpath('./w:num[@w:numId=$n]', namespaces=NS, n=num)
        if not numbers: return None
        over = numbers[0].xpath('./w:lvlOverride[@w:ilvl=$l]/w:lvl',namespaces=NS,l=level)
        if over: return over[0]
        aid = val(numbers[0],'w:abstractNumId')
        nodes = root.xpath('./w:abstractNum[@w:abstractNumId=$a]/w:lvl[@w:ilvl=$l]',namespaces=NS,a=aid,l=level)
        return nodes[0] if nodes else None


class Scorer:
    def __init__(self, doc):
        self.d = doc
        self.rows = []

    def check(self, name, passed, expected, found, p=None):
        return Check(name, passed, expected, str(found), self.d.location(p) if p is not None else 'Document')

    def formatting(self, sid, size, color, p=None, italic=None):
        sid = self.d.sid(sid)
        expected = {'font':'Segoe UI','size':str(size*2),'color':color}
        if italic is not None: expected['italic'] = italic
        f = self.d.format(sid, p)
        checks = [self.check(f'{sid} style {k}', f[k] == v, str(v), f[k], p) for k,v in expected.items()]
        if p is not None:
            mismatches = []
            for run in p.findall('.//w:r', NS):
                if not text(run): continue
                actual = self.d.format(self.d.pstyle(p), p, run)
                bad = {k:actual[k] for k,v in expected.items() if actual[k] != v}
                if bad: mismatches.append(f'{text(run)[:60]}: {bad}')
            checks.append(self.check('Displayed run formatting', not mismatches, str(expected), '; '.join(mismatches[:8]) or 'All text runs match', p))
        return checks

    def row(self, criterion, checks, score=None, policy=''):
        known = [c for c in checks if c.passed is not None]
        ratio = sum(c.passed for c in known)/len(known) if known else 0
        if score is None: score = 5 if ratio == 1 else 3 if ratio >= .75 else 1 if ratio > 0 else 0
        unknown = any(c.passed is None for c in checks)
        self.rows.append({'criterion':criterion,'score':score,'maximum':5,'estimated':unknown,
                          'checks':[asdict(c) for c in checks],
                          'measured_passed':sum(c.passed for c in known),'measured_total':len(known),
                          'policy':policy or '5: all measured checks pass; 3: at least 75%; 1: some pass; 0: none pass. Unverified checks do not deduct points and mark the score estimated.'})


def grade(data, filename):
    d = Document(data); s = Scorer(d); ck = s.check
    original = Document(BASE.read_bytes())
    expected_headings = [(text(p),original.pstyle(p)) for p in original.ps if original.pstyle(p) in [original.sid(x) for x in ('Heading1','Heading2','Heading3')]]
    heading_checks = []
    for sid,size in [('Title',26),('Heading1',18),('Heading2',16),('Heading3',14)]: heading_checks += s.formatting(sid,size,TEAL)
    borders = [layer.find('w:pBdr/w:bottom',NS) for layer in d.layers(d.sid('Title'),kind='pPr')]
    border = next((n for n in reversed(borders) if n is not None),None)
    heading_checks.append(ck('Title bottom border',border is not None and border.get(W+'sz')=='8' and d.color(border)==TEAL and border.get(W+'val') not in ('nil','none'), '1 pt, #00796B bottom border',ET.tostring(border).decode() if border is not None else 'Missing'))
    for heading,sid in expected_headings:
        matches = [p for p in d.ps if text(p)==heading and d.pstyle(p) not in (d.sid('TOC1'),d.sid('TOC2'),d.sid('TOC3'))]
        ok = any(d.pstyle(p)==d.sid(sid) for p in matches)
        heading_checks.append(ck('Heading style application',ok,f'{heading}: {sid}', [d.pstyle(p) for p in matches] or 'Heading missing',matches[0] if matches else None))
        if matches and ok:
            level=int(sid[-1]); heading_checks += s.formatting(sid,{1:18,2:16,3:14}[level],TEAL,matches[0])[-1:]
    s.row(CRITERIA[0],heading_checks)

    # Match baseline body prose by text similarity to allow updated citation results.
    baseline_body = [text(p) for p in original.ps if original.pstyle(p)==original.sid('Normal') and len(text(p))>100 and not text(p).startswith('Figure ')]
    generated = set()
    for f in d.fields(r'^\s*(?:TOC|BIBLIOGRAPHY)\b'): generated.update(f['ps'])
    body_ps=[]
    for p in d.ps:
        t=text(p)
        if not t or p in generated: continue
        if any(t[:100]==b[:100] or SequenceMatcher(None,t,b).ratio()>.85 for b in set(baseline_body)):
            body_ps.append(p)
    body_checks=s.formatting('Normal',11,GRAY)
    body_checks.append(ck('Body paragraphs present',bool(body_ps),'Original body text retained',f'{len(body_ps)} matching paragraphs'))
    for p in body_ps:
        body_checks.append(ck('Body paragraph style',d.pstyle(p)==d.sid('Normal'),'Normal',d.pstyle(p),p))
        body_checks += s.formatting('Normal',11,GRAY,p)[-1:]
    original_counts=Counter(baseline_body)
    retained=sum(min(count,sum(text(p)==t or text(p).startswith(t[:100]) for p in body_ps)) for t,count in original_counts.items())
    body_checks.append(ck('Body text coverage',retained>=len(baseline_body),'All original body paragraphs retained',f'{retained}/{len(baseline_body)}'))
    s.row(CRITERIA[1],body_checks)

    titlefields=d.fields(r'^\s*(?:TITLE\b|DOCPROPERTY\s+"?Title\b)')
    intro=next((p for p in d.ps if text(p)=='Introduction to Bears'),None)
    cutoff=d.index.get(intro,len(d.ps))
    titlefield=next((f for f in titlefields if d.index.get(f['p'],10**9)<cutoff),None)
    titleps=[]
    if titlefield:
        start=d.index[titlefield['p']]
        for p in d.ps[start:cutoff]:
            if p is not titlefield['p'] and d.page_boundary_before(p): break
            titleps.append(p)
    if not titlefield:
        for p in d.ps[:cutoff]:
            if titleps and d.page_boundary_before(p): break
            titleps.append(p)
    nonempty=[p for p in titleps if text(p)]
    checks=[ck('Document Title field',titlefield is not None,'TITLE or DOCPROPERTY Title on first page',titlefield['code'] if titlefield else 'Missing')]
    if titlefield:
        checks.append(ck('Title style',d.pstyle(titlefield['p'])==d.sid('Title'),'Title',d.pstyle(titlefield['p']),titlefield['p']))
        checks += s.formatting('Title',26,TEAL,titlefield['p'])[-1:]
    checks.append(ck('Student information lines',len(nonempty)>=4,'Title plus name, student number and set',f'{len(nonempty)} nonempty paragraphs; content labels are not required'))
    for p in nonempty:
        checks.append(ck('Title page centered',d.prop(d.pstyle(p),'jc',p=p,kind='pPr')=='center','center',d.prop(d.pstyle(p),'jc',p=p,kind='pPr'),p))
        if titlefield and p is not titlefield['p']:
            checks.append(ck('Student information style',d.pstyle(p)==d.sid('Normal'),'Normal',d.pstyle(p),p))
    checks.append(ck('Blank lines between information',len(nonempty)>=4 and all(any(not text(q) for q in titleps[titleps.index(a)+1:titleps.index(b)]) for a,b in zip(nonempty[:3],nonempty[1:4])),'Blank paragraph between title/name/number/set',f'{sum(not text(p) for p in titleps)} blank paragraphs'))
    checks.append(ck('Title page is physically separate',None,'Only title and student information on rendered page 1','Word pagination is not evaluated by the XML engine.'))
    s.row(CRITERIA[2],checks,(1 if nonempty and any(d.page_boundary_before(p) for p in d.ps[1:cutoff+1]) else 0) if not titlefield else None, 'Without the required document-title field, a recognizable cover receives 1; no cover receives 0.' if not titlefield else '')

    tocs=d.fields(r'^\s*TOC\b')
    figures_lists=[f for f in tocs if re.search(r'\\[ca]\s+"?Figure\b',f['code'],re.I)]
    main_tocs=[f for f in tocs if f not in figures_lists]
    for criterion,found,isfig in [(CRITERIA[3],main_tocs,False),(CRITERIA[5],figures_lists,True)]:
        checks=[ck('Automatic Word field',bool(found),'TOC field with figure switch' if isfig else 'TOC heading field', [f['code'] for f in found] or 'Missing')]
        if found:
            f=found[0]; first=f['p']
            checks.append(ck('Field has saved results',bool(text(first)) or any(text(p) for p in f['ps']),'Nonempty generated entries','Saved text present' if any(text(p) for p in f['ps']) else 'No cached results',first))
            expected_sids=['TableofFigures'] if isfig else ['TOC1','TOC2','TOC3']
            has_entry_style=any(d.pstyle(p) in [d.sid(sid) for sid in expected_sids] for p in f['ps'])
            checks.append(ck('Generated entry styles applied',has_entry_style,str(expected_sids),[d.pstyle(p) for p in f['ps']]))
            for sid in expected_sids:
                used=[p for p in f['ps'] if d.pstyle(p)==d.sid(sid)]
                if used:
                    checks+=s.formatting(sid,11,GRAY)
                    layers=d.layers(d.sid(sid),kind='pPr')
                    checks.append(ck('Leader dots',any(layer.xpath('./w:tabs/w:tab[@w:leader="dot"]',namespaces=NS) for layer in layers) or any(p.xpath('./w:pPr/w:tabs/w:tab[@w:leader="dot"]',namespaces=NS) for p in used),'Dot leader on generated entry tab','Dot leader detected' if any(layer.xpath('./w:tabs/w:tab[@w:leader="dot"]',namespaces=NS) for layer in layers) else 'Inspect applied tabs'))
            if isfig:
                titles=[p for p in d.ps if text(p)=='List of Figures' and d.index[p]<=d.index.get(first,0)]
                checks.append(ck('List of Figures title',bool(titles),'List of Figures heading before entries','Present' if titles else 'Missing'))
                if titles: checks.append(ck('List title style',d.pstyle(titles[-1])==d.sid('Heading1'),'Heading1',d.pstyle(titles[-1]),titles[-1]))
            else:
                titles=[p for p in d.ps[:d.index.get(first,0)+1] if text(p).casefold() in ('contents','table of contents')]
                checks.append(ck('TOC title',bool(titles),'Contents or Table of Contents','Present' if titles else 'Missing'))
                if titles: checks+=s.formatting(d.pstyle(titles[-1]),18,TEAL,titles[-1])
            checks.append(ck('Dedicated rendered pages',None,'No unrelated content on TOC/figure-list pages','XML does not reliably encode rendered pagination or identify the exact Word gallery preset.'))
        fallback_score=None
        if not found:
            names=('list of figures',) if isfig else ('contents','table of contents')
            titlep=next((p for p in d.ps if text(p).casefold() in names),None)
            if titlep is not None:
                after=d.ps[d.index[titlep]+1:]
                entries=[]
                for p in after:
                    if text(p)=='Introduction to Bears' or (entries and d.page_boundary_before(p)): break
                    if text(p): entries.append(p)
                expected_entries=('Figure ' if isfig else 'Introduction to Bears')
                plausible=[p for p in entries if (expected_entries in text(p) or re.search(r'[.\t ]\s*\d+$',text(p)))]
                checks.append(ck('Manually created list content',bool(plausible),'Recognizable entries beneath list title',f'{len(plausible)} recognizable entries',titlep))
                if plausible:
                    good=all(all(c.passed for c in s.formatting(d.pstyle(p),11,GRAY,p)) for p in plausible)
                    checks.append(ck('Manual entry formatting',good,'Segoe UI 11 pt gray','Matches' if good else 'Formatting differs'))
                    fallback_score=3 if good else 1
                else:fallback_score=1
            else:fallback_score=0
        s.row(criterion,checks,fallback_score,('Manual list: 3 if recognizable entries meet brand typography; 1 if a list title exists but entries/formatting are incomplete; 0 if absent. An automatic field is required for 5.' if not found else ''))

    seq=d.fields(r'^\s*SEQ\s+Figure\b')
    captions=s.formatting('Caption',10,GRAY,italic=True)
    from zipfile import ZipFile
    from io import BytesIO
    import hashlib
    with ZipFile(BASE) as baseline_zip, ZipFile(BytesIO(data)) as submission_zip:
        original_images=Counter(hashlib.sha256(baseline_zip.read(n)).hexdigest() for n in baseline_zip.namelist() if n.startswith('word/media/'))
        submitted_images=Counter(hashlib.sha256(submission_zip.read(n)).hexdigest() for n in submission_zip.namelist() if n.startswith('word/media/'))
    captions.append(ck('Original figure images retained',not (original_images-submitted_images),'All original image binaries retained','All retained' if not (original_images-submitted_images) else 'One or more original images differ or are missing; image recompression can also change bytes.'))
    captions.append(ck('Four caption fields',len(seq)==4,'Exactly four SEQ Figure fields',len(seq)))
    expected_titles=['Brown bear in its natural environment.','Grizzly bear with her cubs.','Life Cycle of a Polar Bear','Polar bear resting on Arctic ice.']
    for title in expected_titles:
        matches=[f for f in seq if title.rstrip('.') in text(f['p'])]
        manual_duplicates=[p for p in d.ps if title.rstrip('.') in text(p) and p not in generated and not any(f['p'] is p for f in seq)]
        captions.append(ck('Original typed figure label replaced',not manual_duplicates,'No duplicate typed original caption',f'{len(manual_duplicates)} duplicate label paragraphs'))
        captions.append(ck('Original caption text',bool(matches),title,'Found in caption field paragraph' if matches else 'Missing or changed'))
        for f in matches[:1]:
            p=f['p']; captions.append(ck('Caption style applied',d.pstyle(p)==d.sid('Caption'),'Caption',d.pstyle(p),p))
            captions+=s.formatting('Caption',10,GRAY,p,True)[-1:]
            index=d.index.get(p,-1)
            nextp=next((q for q in d.ps[index+1:] if text(q) or q.xpath('.//w:drawing | .//w:pict',namespaces=NS)),None)
            image=nextp is not None and bool(nextp.xpath('.//w:drawing | .//w:pict',namespaces=NS))
            captions.append(ck('Caption above image',image,'Next content paragraph contains figure','Image below' if image else 'Image not immediately below',p))
    s.row(CRITERIA[4],captions,0 if not seq else None)

    list_checks=[]
    baseline_lists=[text(p) for p in original.ps if original.list_level(p) is not None and text(p)]
    for label in baseline_lists:
        p=next((p for p in d.ps if text(p)==label),None)
        lvl=d.list_level(p) if p is not None else None
        list_checks.append(ck('Applied list numbering',lvl is not None,'Real Word list for '+label, 'Present' if lvl is not None else 'Missing or typed bullet',p))
        if lvl is None: continue
        sid=d.pstyle(p)
        list_checks+=s.formatting(sid,11,GRAY,p)
        fmt=val(lvl,'w:numFmt'); marker=val(lvl,'w:lvlText')
        if fmt=='bullet':
            list_checks.append(ck('Bullet shape',marker in ('●','■','•','\uf0b7','\uf0a7'),'Solid circle or square',marker,p))
            col=lvl.find('w:rPr/w:color',NS)
            color=d.color(col) if col is not None else d.format(sid)['color']
            list_checks.append(ck('Bullet color',color==CORAL,'#FF6F61',color,p))
        else:
            list_checks.append(ck('Numbering format',fmt=='decimal','Decimal numbered list',fmt,p))
        indent=val(lvl,'w:pPr/w:ind',attr='left') or val(lvl,'w:pPr/w:ind',attr='start')
        applied=d.prop(sid,'ind',attr='left',p=p,kind='pPr') or d.prop(sid,'ind',attr='start',p=p,kind='pPr') or indent
        list_checks.append(ck('List indentation',applied=='360','0.25 inch / 360 twips',applied,p))
    s.row(CRITERIA[6],list_checks,0 if not list_checks else None)

    sources=[r for k,r in d.roots.items() if k.startswith('customXml/') and r.tag=='{'+NS['b']+'}Sources']
    citation_fields=d.fields(r'^\s*CITATION\b')
    apa=any(re.search(r'APA.*(?:sixth|6)|(?:sixth|6).*APA',str(r.attrib),re.I) or (r.get('StyleName','').upper()=='APA' and r.get('Version')=='6') for r in sources)
    tag=None
    if citation_fields:
        match=re.match(r'\s*CITATION\s+(?:"([^"]+)"|(\S+))',citation_fields[0]['code'],re.I)
        if match: tag=match.group(1) or match.group(2)
    source=next((q for r in sources for q in r.findall('b:Source',NS) if q.findtext('b:Tag',namespaces=NS)==tag),None)
    expected_source={'SourceType':'Book','Title':'The Singing Wilderness','Year':'1956','City':'New York','Publisher':'Alfred A. Knopf'}
    source_checks=[]
    for name,want in expected_source.items():
        actual=source.findtext('b:'+name,namespaces=NS) if source is not None else None
        source_checks.append(ck('First source '+name,norm(actual or '').rstrip('.').casefold()==want.rstrip('.').casefold(),want,actual))
    persons=source.xpath('.//b:Person',namespaces=NS) if source is not None else []
    author_ok=any(p.findtext('b:Last',namespaces=NS)=='Olson' and re.sub(r'[^A-Za-z]','',(p.findtext('b:First',default='',namespaces=NS)+' '+p.findtext('b:Middle',default='',namespaces=NS))).upper()=='SF' for p in persons)
    source_checks.append(ck('First source author',author_ok,'Olson, S. F.', '; '.join(' '.join(p.itertext()) for p in persons) or 'Missing'))
    updated=all(c.passed for c in source_checks)
    checks=[ck('Citation fields',bool(citation_fields),'Real CITATION fields',len(citation_fields)),ck('APA 6 bibliography setting',apa,'APA sixth edition', [r.attrib for r in sources])]+source_checks
    source_tags={q.findtext('b:Tag',namespaces=NS) for r in sources for q in r.findall('b:Source',NS)}
    unresolved=[]
    for f in citation_fields:
        m=re.match(r'\s*CITATION\s+(?:"([^"]+)"|(\S+))',f['code'],re.I)
        if not m or (m.group(1) or m.group(2)) not in source_tags: unresolved.append(f['code'])
    checks.append(ck('Citation sources resolved',not unresolved,'Every citation tag resolves to a saved source',unresolved or 'All resolved'))
    score=5 if apa and updated and citation_fields and not unresolved else 3 if updated else 1 if apa else 0
    s.row(CRITERIA[7],checks,score,'Rubric: 5 = APA 6 and first source updated; 3 = first source updated without APA 6; 1 = APA 6 without source update; 0 = neither. Missing or unresolved citations cannot receive 5.')

    bibliography=d.fields(r'^\s*BIBLIOGRAPHY\b')
    bibchecks=[ck('Word bibliography field',bool(bibliography),'BIBLIOGRAPHY field',[f['code'] for f in bibliography] or 'Missing')]
    if bibliography:
        f=bibliography[0]
        entries=[p for p in f['ps'] if text(p)]
        bibchecks.append(ck('Bibliography saved entries',bool(entries),'Nonempty reference entries',len(entries)))
        for p in sorted(entries,key=lambda p:d.index.get(p,0)):
            bibchecks+=s.formatting(d.pstyle(p),11,GRAY,p)
        bibchecks.append(ck('References on separate rendered page',None,'Own page after main text','Physical pagination is not evaluated. Section-break structure is scored separately.'))
    else:
        entries=[p for p in d.ps if d.pstyle(p)==d.sid('Bibliography')]
        if entries:
            bibchecks+=s.formatting('Bibliography',11,GRAY)
    if bibliography: bibscore=None
    elif entries: bibscore=3 if all(c.passed for c in bibchecks[1:]) else 1
    else: bibscore=0
    s.row(CRITERIA[8],bibchecks,bibscore)

    # Only referenced footers count: unused package parts cannot earn marks.
    # package() keeps XML parts, including .rels through a separate safe read below.
    from zipfile import ZipFile
    from io import BytesIO
    with ZipFile(BytesIO(data)) as z:
        rels=xml(z.read('word/_rels/document.xml.rels')) if 'word/_rels/document.xml.rels' in z.namelist() else None
    relationships={r.get('Id'):r for r in rels} if rels is not None else {}
    def footer(ref):
        rel=relationships.get(ref.get(R+'id'))
        if rel is None or rel.get('TargetMode')=='External': return None
        import posixpath
        path=posixpath.normpath(posixpath.join('word',rel.get('Target','')))
        return d.roots.get(path)
    checks=[]
    inherited={}
    default_ids=[]
    for i,section in enumerate(d.sections):
        for ref in section.findall('w:footerReference',NS): inherited[ref.get(W+'type')]=ref
        ref=inherited.get('default'); fr=footer(ref) if ref is not None else None
        default_ids.append(ref.get(R+'id') if ref is not None else None)
        fs=d.parse_fields(fr) if fr is not None else []
        page=[f for f in fs if re.match(r'\s*PAGE\b',f['code'],re.I)]
        title=[f for f in fs if re.match(r'\s*(TITLE\b|DOCPROPERTY\s+"?Title\b)',f['code'],re.I)]
        checks.append(ck(f'Section {i+1} footer fields',bool(page and title),'PAGE and document title fields',[f['code'] for f in fs]))
        for p in fr.findall('.//w:p',NS) if fr is not None else []:
            if not text(p) and not d.parse_fields(p): continue
            checks.append(ck('Footer right aligned',d.prop(d.pstyle(p),'jc',p=p,kind='pPr')=='right','right',d.prop(d.pstyle(p),'jc',p=p,kind='pPr')))
            checks+=s.formatting(d.pstyle(p),9,GRAY,p)
            # Text between field results should include a separator.
        if i==0:
            first_ref=inherited.get('first'); first_footer=footer(first_ref) if first_ref is not None else None
            firstfs=d.parse_fields(first_footer) if first_footer is not None else []
            hidden=enabled(section.find('w:titlePg',NS)) and not any(re.match(r'\s*PAGE\b',f['code'],re.I) for f in firstfs)
            checks.append(ck('Title-page number hidden',hidden,'Different first page with no PAGE field','Hidden' if hidden else 'Not hidden'))
            start=val(section,'w:pgNumType',attr='start')
            checks.append(ck('Main page numbering starts at 1',start in (None,'1'),'1 or Word default',start or 'default 1'))
    s.row(CRITERIA[9],checks,0 if not any(default_ids) else None)

    checks=[ck('Separate references section',len(d.sections)>=2,'At least two sections',len(d.sections))]
    if len(d.sections)>=2:
        last=d.sections[-1]; before=d.sections[-2]
        checks.append(ck('New section starts next page',val(last,'w:type','nextPage')=='nextPage','nextPage',val(last,'w:type','nextPage')))
        refs=last.findall('w:footerReference',NS)
        own=next((r for r in refs if r.get(W+'type')=='default'),None)
        independent=own is not None and len(default_ids)>=2 and default_ids[-1] != default_ids[-2]
        checks.append(ck('References footer unlinked',independent,'Distinct explicit default footer reference',default_ids))
        checks.append(ck('References letter numbering',val(last,'w:pgNumType',attr='fmt')=='upperLetter','upperLetter',val(last,'w:pgNumType',attr='fmt')))
        checks.append(ck('References numbering starts at A',val(last,'w:pgNumType',attr='start')=='1','1 (displayed A)',val(last,'w:pgNumType',attr='start')))
        checks.append(ck('Main section decimal numbering',all(val(sec,'w:pgNumType',default='decimal',attr='fmt')=='decimal' for sec in d.sections[:-1]),'Previous sections retain decimal numbering',[val(sec,'w:pgNumType',default='decimal',attr='fmt') for sec in d.sections[:-1]]))
        if bibliography:
            p=bibliography[0]['p']; idx=d.index.get(p,-1)
            boundaries=[d.index.get(sec.getparent().getparent(),-1) for sec in d.sections[:-1]]
            checks.append(ck('Bibliography after last section break',idx>max(boundaries),'Bibliography field in final section',f'Bibliography paragraph {idx+1}; section-end paragraphs {[b+1 for b in boundaries]}',p))
    s.row(CRITERIA[10],checks,0 if len(d.sections)<2 else None)

    checks=[]
    for i,sec in enumerate(d.sections):
        for side in ('top','bottom','left','right'):
            actual=val(sec,'w:pgMar',attr=side)
            checks.append(ck(f'Section {i+1} {side} margin',actual=='1440','1 inch (1440 twips)',actual))
    checks+=s.formatting('Normal',11,GRAY)
    checks+=s.formatting('Title',26,TEAL)
    for sid,size in [('Heading1',18),('Heading2',16),('Heading3',14)]: checks+=s.formatting(sid,size,TEAL)
    checks.append(ck('Visual consistency and spacing',None,'Consistent brand layout; optional background/callouts are not required','No fixed paragraph-spacing values are specified. XML checks cannot evaluate visual balance.'))
    s.row(CRITERIA[11],checks)

    extra=''; sentence_note='Required heading or first body paragraph not found.'; count=None
    target=next((i for i,p in enumerate(d.ps) if text(p)=='A Multifaceted Approach'),None)
    if target is not None:
        paragraph=next((p for p in d.ps[target+1:] if text(p)),None)
        base='Effective bear conservation requires a multifaceted approach. Safe and tested strategies must be used.'
        if paragraph is not None and text(paragraph).startswith(base):
            extra=text(paragraph)[len(base):].strip()
            count=len(re.findall(r"\b\w+(?:[’'-]\w+)*\b",extra))
            one=len(re.findall(r'[.!?]+(?:[”\"\u2019]*)\s*(?:$|(?=[A-Z]))',extra))==1
            sentence_note='Length and single-sentence checks pass.' if extra and count<=25 and one else 'Missing addition, more than 25 words, or not exactly one sentence.'
        elif paragraph is not None:
            extra=text(paragraph); sentence_note='Original paragraph changed; appended sentence could not be isolated.'
    sentence={'text':extra,'word_count':count,'finding':sentence_note,'context':'Context appropriateness is not determined by deterministic checks. No separate points are assigned by the rubric.'}
    s.rows.sort(key=lambda row:CRITERIA.index(row['criterion']))
    return {'filename':filename,'version':VERSION,'total':sum(r['score'] for r in s.rows),'maximum':60,
            'estimated':any(r['estimated'] for r in s.rows),'criteria':s.rows,'sentence':sentence,
            'limitations':['Scores use the documented deterministic rubric mapping; qualitative partial-credit boundaries are implementation policy, not extra instructions from the rubric.', 'Page isolation, exact rendering, identity of student-information lines and sentence context cannot be established fully from OOXML. Scores involving those checks are marked estimated.']}


def grade_paths(paths):
    """A bad submission must not interrupt grading the remaining batch."""
    results=[]
    for path in paths:
        path=Path(path)
        try:
            if path.stat().st_size>15_000_000: raise ValueError('File exceeds 15 MB.')
            result=grade(path.read_bytes(),path.name)
        except (ValueError,OSError,ET.LxmlError) as exc:
            result={'filename':path.name,'error':str(exc)}
        results.append(result)
    return results


def csv_report(results):
    output=io.StringIO(newline=''); writer=csv.writer(output)
    writer.writerow(['Submission',*CRITERIA,'Total','Maximum','Status'])
    for result in results:
        name=result['filename']
        if name.lstrip().startswith(('=','+','-','@')): name="'"+name
        if 'error' in result: writer.writerow([name,*['']*12,'','','Error: '+result['error']]); continue
        writer.writerow([name,*[r['score'] for r in result['criteria']],result['total'],60,'Estimated' if result['estimated'] else 'Scored'])
    return output.getvalue()


def detail_text(result):
    if 'error' in result: return result['filename']+'\n'+result['error']
    lines=[f'{result["filename"]}: {result["total"]}/60'+(' (estimated)' if result['estimated'] else ''),f'Engine {result["version"]}','']
    for row in result['criteria']:
        lines += [f'{row["criterion"]}: {row["score"]}/5',row['policy']]
        for c in row['checks']:
            label='PASS' if c['passed'] else 'UNVERIFIED' if c['passed'] is None else 'FAIL'
            lines += [f'  [{label}] {c["name"]} — {c["location"]}',f'    Expected: {c["expected"]}',f'    Found: {c["found"]}']
        lines.append('')
    lines += ['Additional sentence (no separate rubric points):',str(result['sentence']),*result['limitations']]
    return '\n'.join(lines)
