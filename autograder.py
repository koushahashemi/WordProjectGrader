"""Deterministic Bears-assignment scoring, with a check-level audit trail."""
from collections import Counter
from dataclasses import dataclass, asdict, field
from difflib import SequenceMatcher
from pathlib import Path
import csv
import io
import re
from lxml import etree as ET
from grader import BASE, CRITERIA, NS, W, package, xml
from branding import assess_brand
from app_info import APP_TITLE, COURSE_NAME, TERM, COURSE_TEXT
from matching import canonical, same_text, near, twips, inches, points, font_family, describe_format

VERSION = '3.1.0'
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
    accepted: str = ''
    scored: bool = True
    issues: list[str] = field(default_factory=list)


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
        result = {'font':None, 'size':None, 'color':None, 'italic':False, 'bold':False}
        face_bold=False
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
                if theme or font:
                    result['font'] = font
                    face_bold=bool(re.search(r'bold|demi',font or '',re.I))
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
            bold=layer.find('w:b',NS)
            if bold is not None:
                is_style=layer.getparent() is not None and layer.getparent().tag==W+'style'
                if is_style and enabled(bold):result['bold']=not result['bold']
                elif not is_style:result['bold']=enabled(bold)
        result['bold']=result['bold'] or face_bold
        result['font'],face_italic=font_family(result['font'])
        result['italic']=result['italic'] or face_italic
        return result

    def format_origin(self,sid,prop,run=None):
        tag={'font':'rFonts','size':'sz','color':'color','italic':'i','bold':'b'}[prop]
        origin=self.sid('Normal')
        for style in self.chain(sid):
            if style.find('w:rPr/w:'+tag,NS) is not None:origin=style.get(W+'styleId')
        if run is not None:
            for style in self.chain(val(run,'w:rPr/w:rStyle')):
                if style.find('w:rPr/w:'+tag,NS) is not None:origin=style.get(W+'styleId')
            if run.find('w:rPr/w:'+tag,NS) is not None:origin='direct'
        return origin

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

    def title_controls(self,root):
        result=[]
        for control in root.findall('.//w:sdt',NS):
            binding=control.find('w:sdtPr/w:dataBinding',NS)
            if binding is None: continue
            path=binding.get(W+'xpath','')
            if not re.search(r'(?:^|/)(?:[\w]+:)?title(?:\[\d+\])?$',path,re.I):continue
            mappings=binding.get(W+'prefixMappings','')
            if 'http://purl.org/dc/elements/1.1/' not in mappings:continue
            p=next((a for a in control.iterancestors() if a.tag==W+'p'),None)
            if p is None:p=control.find('.//w:sdtContent//w:p',NS)
            if p is not None and text(control):result.append({'code':'Bound document Title content control','p':p,'node':control,'ps':{p}})
        return result

    def right_aligned(self,p):
        sid=self.pstyle(p)
        if self.prop(sid,'jc',p=p,kind='pPr')=='right':return True
        tabs={}
        for layer in self.layers(sid,p=p,kind='pPr'):
            for tab in layer.findall('w:tabs/w:tab',NS):
                pos=tab.get(W+'pos')
                if tab.get(W+'val')=='clear':tabs.pop(pos,None)
                else:tabs[pos]=tab.get(W+'val')
        # A leading tab using an inherited right tab stop is valid right alignment.
        first=next((n for n in p.iter() if n.tag in (W+'t',W+'tab') and (n.tag==W+'tab' or n.text)),None)
        return first is not None and first.tag==W+'tab' and 'right' in tabs.values()

    def fields(self, pattern):
        return [f for f in self.field_list if re.search(pattern, f['code'], re.I)]

    def location(self, p):
        return (f'Paragraph {self.index[p]+1}: {text(p)[:100]}' if p in self.index else f'Footer/content paragraph: {text(p)[:100]}') if p is not None else 'Document'

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
    def __init__(self, doc, calibrated=True):
        self.d = doc
        self.calibrated = calibrated
        self.rows = []
        self.area = None

    def check(self, name, passed, expected, found, p=None):
        return Check(name, passed, expected, str(found), self.d.location(p) if p is not None else 'Document')

    def formatting(self, sid, size, color, p=None, italic=None, direct_only=False, bold=None):
        sid = self.d.sid(sid)
        expected = {'font':'Segoe UI','size':str(size*2),'color':color}
        if italic is not None: expected['italic'] = italic
        if bold is not None:expected['bold']=bold
        f = self.d.format(sid, p)
        def owner(origin,prop):
            references={self.d.sid('Normal'):({'font':'Segoe UI','size':'22','color':GRAY},CRITERIA[1])}
            for name,n in [('Title',26),('Heading1',18),('Heading2',16),('Heading3',14)]:
                references[self.d.sid(name)]=({'font':'Segoe UI','size':str(n*2),'color':TEAL},CRITERIA[0])
            if origin in references:
                required,area=references[origin]
                if prop in required and self.d.format(origin)[prop]!=required[prop]:return area
            return self.area
        checks=[]
        for k,v in expected.items():
            origin=self.d.format_origin(sid,k)
            c=self.check(f'{sid} style {k}',f[k]==v,points(v) if k=='size' else str(v),points(f[k]) if k=='size' else f[k],p)
            c.issues=[f'format:{origin}:{k}']
            c.scored=owner(origin,k)==self.area
            if not c.scored:c.expected+=' (graded under '+owner(origin,k)+')'
            checks.append(c)
        if p is not None:
            mismatches=[];informational=[];issues=set()
            inherited={n.get(W+'styleId') for n in self.d.chain(self.d.pstyle(p))}|{self.d.sid('Normal')}
            for run in p.findall('.//w:r',NS):
                if not text(run):continue
                actual=self.d.format(self.d.pstyle(p),p,run)
                run_expected=expected
                left=self.d.prop(self.d.pstyle(p),'ind',attr='left',p=p,kind='pPr')
                right=self.d.prop(self.d.pstyle(p),'ind',attr='right',p=p,kind='pPr')
                if sid==self.d.sid('Normal') and near(left,720) and near(right,720) and actual['color']==CORAL:
                    run_expected={**expected,'color':CORAL}
                bad={k:actual[k] for k,v in run_expected.items() if actual[k]!=v}
                active={}
                for k in bad:
                    origin=self.d.format_origin(self.d.pstyle(p),k,run)
                    if direct_only and origin in inherited:continue
                    if owner(origin,k)!=self.area:continue
                    active[k]=bad[k]
                    root=sid if origin=='direct' else origin
                    issues.add(f'format:{root}:{k}')
                if active:mismatches.append(f'{text(run)[:60]}: {describe_format(active)}')
                elif bad:informational.append(f'{text(run)[:60]}: {describe_format(bad)} (inherited formatting graded elsewhere)')
            c=self.check('Displayed run formatting',not mismatches and not informational,describe_format(expected),'; '.join((mismatches or informational)[:8]) or 'All text runs match',p)
            c.issues=sorted(issues);c.scored=not mismatches and not informational or bool(issues)
            checks.append(c)
        return checks

    def accept(self,check,reason):
        if self.calibrated and check.passed is False:check.accepted=reason
        return check

    def row(self, criterion, checks, score=None, policy=''):
        known=[c for c in checks if c.scored and c.passed is not None]
        def failures(calibrated):
            return sorted({issue for c in known if c.passed is False and not (calibrated and c.accepted) for issue in (c.issues or [c.name])})
        raw=failures(False);deductions=failures(True)
        raw_any_pass=any(c.passed for c in known)
        any_pass=any(c.passed or c.accepted for c in known)
        strict_score=score if score is not None else (0 if known and not raw_any_pass else max(0,5-len(raw)))
        total=score if score is not None else (0 if known and not any_pass else max(0,5-len(deductions)))
        self.rows.append({'criterion':criterion,'score':total,'strict_score':strict_score,'maximum':5,
                          'estimated':any(c.scored and c.passed is None for c in checks),
                          'checks':[asdict(c) for c in checks],
                          'accepted_variants':sum(bool(c.accepted) for c in known),
                          'measured_passed':sum(bool(c.passed or c.accepted) for c in known),'measured_total':len(known),
                          'deductions':deductions,
                          'policy':('Teacher benchmark: documented accepted variants earn credit; original mismatches stay visible. ' if any(c.accepted for c in checks) else '')+(policy or 'Start at 5; deduct one mark per distinct failed item or formatting property. Repeated inherited failures count once. Missing primary features can receive 0. Informational checks owned by another criterion do not deduct marks.')})

    def grouped_row(self,criterion,groups,policy,page_numbers=False):
        summaries=[];checks=[]
        for name,weight,evidence in groups:
            known=[c for c in evidence if c.scored and c.passed is not None]
            owned_elsewhere=not known and any(not c.scored for c in evidence)
            passed=(bool(known) or owned_elsewhere) and all(c.passed or c.accepted for c in known)
            raw_passed=(bool(known) or owned_elsewhere) and all(c.passed for c in known)
            summaries.append({'name':name,'weight':weight,'passed':passed,'raw_passed':raw_passed,'earned':weight if passed else 0,'checks':[c.name for c in evidence]})
            checks+=evidence
        failed=sum(not g['passed'] for g in summaries)
        raw_failed=sum(not g['raw_passed'] for g in summaries)
        score=(0 if failed==4 else 5-failed) if page_numbers else sum(g['earned'] for g in summaries)
        strict=(0 if raw_failed==4 else 5-raw_failed) if page_numbers else sum(g['weight'] for g in summaries if g['raw_passed'])
        self.row(criterion,checks,score,policy)
        self.rows[-1]['strict_score']=strict
        self.rows[-1]['scoring_groups']=summaries


def grade(data, filename, calibrated=True):
    d = Document(data); s = Scorer(d,calibrated); ck = s.check
    original = Document(BASE.read_bytes())
    expected_headings = [(text(p),original.pstyle(p)) for p in original.ps if original.pstyle(p) in [original.sid(x) for x in ('Heading1','Heading2','Heading3')]]
    s.area=CRITERIA[0]
    heading_checks = []
    heading_groups=[]
    for sid,size in [('Title',26),('Heading1',18),('Heading2',16),('Heading3',14)]:
        evidence=s.formatting(sid,size,TEAL)
        heading_checks+=evidence
        heading_groups.append((sid+' style formatting',1,evidence))
    borders = [layer.find('w:pBdr/w:bottom',NS) for layer in d.layers(d.sid('Title'),kind='pPr')]
    border = next((n for n in reversed(borders) if n is not None),None)
    heading_checks.append(ck('Title bottom border',border is not None and border.get(W+'sz')=='8' and d.color(border)==TEAL and border.get(W+'val') not in ('nil','none'), '1 pt, #00796B bottom border',ET.tostring(border).decode() if border is not None else 'Missing'))
    applied_titles=[p for p in d.ps if d.pstyle(p)==d.sid('Title') and text(p)]
    def correct_border(n):
        return n is not None and n.get(W+'sz')=='8' and d.color(n)==TEAL and n.get(W+'val') not in ('nil','none')
    if applied_titles and all(correct_border(p.find('w:pPr/w:pBdr/w:bottom',NS)) for p in applied_titles):
        s.accept(heading_checks[-1],'Teacher benchmark accepts the correct border applied to the title paragraph even when the saved Title-style border differs.')
    heading_groups[0][2].append(heading_checks[-1])
    application=[]
    for heading,sid in expected_headings:
        matches = [p for p in d.ps if same_text(text(p),heading) and d.pstyle(p) not in (d.sid('TOC1'),d.sid('TOC2'),d.sid('TOC3'))]
        ok = any(d.pstyle(p)==d.sid(sid) for p in matches)
        application.append(ck('Heading style application',ok,f'{heading}: {sid}', [d.pstyle(p) for p in matches] or 'Heading missing',matches[0] if matches else None))
        if matches and ok:
            level=int(sid[-1]); application += s.formatting(sid,{1:18,2:16,3:14}[level],TEAL,matches[0],direct_only=True)[-1:]
    # Direct formatting on heading text belongs to application, not to style definitions.
    heading_groups.append(('Heading 1–3 application to text',1,application))
    s.grouped_row(CRITERIA[0],heading_groups,'Five equal one-mark groups: Title, Heading 1, Heading 2, Heading 3 style definitions, and Heading 1–3 application to text. Title-style application to the cover is graded only under Title Page.')

    # Match baseline body prose by text similarity to allow updated citation results.
    s.area=CRITERIA[1]
    baseline_body = [text(p) for p in original.ps if original.pstyle(p)==original.sid('Normal') and len(text(p))>100 and not text(p).startswith('Figure ')]
    generated = set()
    for f in d.fields(r'^\s*(?:TOC|BIBLIOGRAPHY)\b'): generated.update(f['ps'])
    body_ps=[]
    for p in d.ps:
        t=text(p)
        if not t or p in generated: continue
        if any(t[:100]==b[:100] or SequenceMatcher(None,t,b,autojunk=False).ratio()>.85 for b in set(baseline_body)):
            body_ps.append(p)
    body_checks=s.formatting('Normal',11,GRAY)
    body_checks.append(ck('Body paragraphs present',bool(body_ps),'Original body text retained',f'{len(body_ps)} matching paragraphs'))
    for p in body_ps:
        body_checks.append(ck('Body paragraph style',d.pstyle(p)==d.sid('Normal'),'Normal',d.pstyle(p),p))
        body_checks += s.formatting('Normal',11,GRAY,p)[-1:]
    original_counts=Counter(baseline_body)
    remaining=[text(p) for p in body_ps]
    retained=0
    for t in baseline_body:
        match=next((i for i,value in enumerate(remaining) if value==t or value.startswith(t[:100]) or SequenceMatcher(None,value,t,autojunk=False).ratio()>.85),None)
        if match is not None:retained+=1;remaining.pop(match)
    body_checks.append(ck('Body text coverage',retained>=len(baseline_body),'All original body paragraphs retained',f'{retained}/{len(baseline_body)}'))
    s.row(CRITERIA[1],body_checks,0 if not body_ps else None)

    s.area=CRITERIA[2]
    titlefields=d.fields(r'^\s*(?:TITLE\b|DOCPROPERTY\s+"?Title\b)')+d.title_controls(d.doc)
    intro=next((p for p in d.ps if same_text(text(p),'Introduction to Bears')),None)
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
        checks += s.formatting('Title',26,TEAL,titlefield['p'],direct_only=True)[-1:]
    checks.append(ck('Student information lines',len(nonempty)>=4,'Title plus name, student number and set',f'{len(nonempty)} nonempty paragraphs; content labels are not required'))
    for p in nonempty:
        checks.append(ck('Title page centered',d.prop(d.pstyle(p),'jc',p=p,kind='pPr')=='center','center',d.prop(d.pstyle(p),'jc',p=p,kind='pPr'),p))
        if titlefield and p is not titlefield['p']:
            checks.append(ck('Student information style',d.pstyle(p)==d.sid('Normal'),'Normal',d.pstyle(p),p))
    checks.append(ck('Blank lines between information',len(nonempty)>=4 and all(any(not text(q) for q in titleps[titleps.index(a)+1:titleps.index(b)]) for a,b in zip(nonempty[:3],nonempty[1:4])),'Blank paragraph between title/name/number/set',f'{sum(not text(p) for p in titleps)} blank paragraphs'))
    if not titlefield and nonempty:
        cover=nonempty[0]
        checks.append(ck('Typed cover title style',d.pstyle(cover)==d.sid('Title'),'Title',d.pstyle(cover),cover))
        checks+=s.formatting('Title',26,TEAL,cover,direct_only=True)[-1:]
        if len(nonempty)>=4 and any(same_text(text(cover),text(p)) for p in original.ps if original.pstyle(p)==original.sid('Title')) and any(d.page_boundary_before(p) for p in d.ps[1:cutoff+1]):
            s.accept(checks[0],'Teacher benchmark accepts a typed assignment title on a separate cover with student-information lines; applying Title style is evaluated separately.')
    checks.append(ck('Title page is physically separate',None,'Only title and student information on rendered page 1','Word pagination is not evaluated by the XML engine.'))
    has_cover=bool(nonempty) and (titlefield is not None or any(d.page_boundary_before(p) for p in d.ps[1:cutoff+1]))
    s.row(CRITERIA[2],checks,0 if not has_cover else None)

    tocs=d.fields(r'^\s*TOC\b')
    figures_lists=[f for f in tocs if re.search(r'\\[ca]\s+"?Figure\b',f['code'],re.I)]
    main_tocs=[f for f in tocs if f not in figures_lists]
    for criterion,found,isfig in [(CRITERIA[3],main_tocs,False),(CRITERIA[5],figures_lists,True)]:
        s.area=criterion
        toc_title_evidence=[]
        plausible=[]
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
                titles=[p for p in d.ps if any(same_text(text(p),t) for t in ('list of figures','table of figures')) and d.index[p]<=d.index.get(first,0)]
                checks.append(ck('List of Figures title',bool(titles),'List of Figures heading before entries','Present' if titles else 'Missing'))
                if titles: checks.append(ck('List title style',d.pstyle(titles[-1]) in (d.sid('Heading1'),d.sid('TOCHeading')),'Heading1 or TOCHeading',d.pstyle(titles[-1]),titles[-1]))
            else:
                titles=[p for p in d.ps[:d.index.get(first,0)+1] if text(p).casefold() in ('contents','table of contents')]
                toc_title_evidence.append(ck('TOC title',bool(titles),'Contents or Table of Contents','Present' if titles else 'Missing'))
                checks+=toc_title_evidence
                if titles:
                    sid=d.pstyle(titles[-1])
                    shared=sid in [d.sid(x) for x in ('Title','Heading1','Heading2','Heading3')]
                    title_format=s.formatting(sid,18,TEAL,titles[-1],direct_only=shared)[-1:] if shared else s.formatting(sid,18,TEAL,titles[-1])
                    checks+=title_format;toc_title_evidence+=title_format
            checks.append(ck('Dedicated rendered pages',None,'No unrelated content on TOC/figure-list pages','XML does not reliably encode rendered pagination or identify the exact Word gallery preset.'))
        fallback_score=None
        if not found:
            names=('list of figures','table of figures') if isfig else ('contents','table of contents')
            titlep=next((p for p in d.ps if any(same_text(text(p),t) for t in names)),None)
            if titlep is not None:
                after=d.ps[d.index[titlep]+1:]
                entries=[]
                for p in after:
                    if same_text(text(p),'Introduction to Bears') or (entries and d.page_boundary_before(p)): break
                    if text(p): entries.append(p)
                expected_entries=('Figure ' if isfig else 'Introduction to Bears')
                plausible=[p for p in entries if (expected_entries in text(p) or re.search(r'[.\t ]\s*\d+$',text(p)))]
                checks.append(ck('Manually created list content',bool(plausible),'Recognizable entries beneath list title',f'{len(plausible)} recognizable entries',titlep))
                if plausible:
                    good=all(all(c.passed for c in s.formatting(d.pstyle(p),11,GRAY,p)) for p in plausible)
                    checks.append(ck('Manual entry formatting',good,'Segoe UI 11 pt gray','Matches' if good else 'Formatting differs'))
                    fallback_score=None if good else 1
                else:fallback_score=1
            else:fallback_score=0
        if isfig:
            s.row(criterion,checks,fallback_score,('Manual list: one mark is lost for a missing automatic field when entries are otherwise correct; 1 if entries are incomplete; 0 if absent.' if not found else ''))
        else:
            if not found:
                titles=[p for p in d.ps if text(p).casefold() in ('contents','table of contents')]
                toc_title_evidence=[ck('TOC title',bool(titles),'Contents or Table of Contents','Present' if titles else 'Missing')]
                if titles:toc_title_evidence+=s.formatting(d.pstyle(titles[0]),18,TEAL,titles[0])
            feature=[c for c in checks if c.name in ('Automatic Word field','Field has saved results','Dedicated rendered pages')]
            title_ids={id(c) for c in toc_title_evidence}
            entry=[c for c in checks if id(c) not in title_ids and (' style ' in c.name or c.name=='Manual entry formatting')]
            structure=[c for c in checks if c.name in ('Generated entry styles applied','Leader dots','Manually created list content')]
            if not entry:entry=[ck('TOC entry typography',False,'Segoe UI, 11 pt, Dark Gray','No recognizable entries to format')]
            if not structure:structure=[ck('TOC entry structure',False,'Entries with correct structure and leader dots','No recognizable entries')]
            s.grouped_row(criterion,[('Automatic TOC with saved entries',2,feature),('TOC title and formatting',1,toc_title_evidence),('TOC entry typography',1,entry),('Entry structure and leader dots',1,structure)],'Four groups: real automatic TOC with saved entries (2), title/formatting (1), entry typography (1), and entry structure/leader dots (1). Repeated font/size errors across TOC levels affect only typography; unverified pagination does not deduct marks.')

    s.area=CRITERIA[4]
    seq=d.fields(r'^\s*SEQ\s+Figure\b')
    presence=[ck('Four caption fields',len(seq)==4,'Four real Word figure captions',f'{len(seq)} caption fields')]
    figure_count=len(d.doc.xpath('//w:drawing | //w:pict',namespaces=NS))
    presence.append(ck('Four figures present',figure_count>=4,'At least four figures',f'{figure_count} figures'))
    expected_titles=['Brown bear in its natural environment.','Grizzly bear with her cubs.','Life Cycle of a Polar Bear','Polar bear resting on Arctic ice.']
    placement=[];formatting=s.formatting('Caption',10,GRAY,italic=True,bold=False)
    matched=[]
    def caption_content(value):
        return re.sub(r'^\s*(?:Figure\s*)?\d+\s*[:.\-–]?\s*','',value,flags=re.I)
    for title in expected_titles:
        matches=[f for f in seq if same_text(caption_content(text(f['p'])),title) or canonical(text(f['p'])).endswith(canonical(title))]
        presence.append(ck('Original caption text',bool(matches),title,'Found (punctuation/spacing/minor spelling differences ignored)' if matches else 'Missing or substantially changed'))
        for f in matches[:1]:
            p=f['p'];matched.append(p)
            applied=any(n.get(W+'styleId')==d.sid('Caption') for n in d.chain(d.pstyle(p)))
            formatting.append(ck('Caption style applied',applied,'Caption or a style based on Caption',d.pstyle(p),p))
            formatting+=s.formatting('Caption',10,GRAY,p,True,bold=False)[-1:]
            formatting.append(ck('Figure label shown',bool(re.match(r'^\s*Figure(?:\b|(?=\d))',text(p),re.I)),'Figure label before caption number','Present' if re.match(r'^\s*Figure(?:\b|(?=\d))',text(p),re.I) else 'Missing Figure label',p))
            alignment=d.prop(d.pstyle(p),'jc',p=p,kind='pPr') or 'left (default)'
            formatting.append(ck('Caption alignment',True,'No specific alignment is prescribed by the supplied guide',alignment,p))
            index=d.index.get(p,-1)
            nextp=next((q for q in d.ps[index+1:] if text(q) or q.xpath('.//w:drawing | .//w:pict',namespaces=NS)),None)
            image=nextp is not None and bool(nextp.xpath('.//w:drawing | .//w:pict',namespaces=NS))
            # An inline picture may share the caption paragraph; check XML order.
            if p.xpath('.//w:drawing | .//w:pict',namespaces=NS):
                nodes=list(p.iter());image=nodes.index(f['node'])<min(nodes.index(n) for n in p.xpath('.//w:drawing | .//w:pict',namespaces=NS))
            placement.append(ck('Caption above image',image,'Caption directly above its figure','Figure follows caption' if image else 'Figure does not immediately follow caption',p))
    if not matched:
        placement.append(ck('Caption above image',False,'Captions above figures','No recognizable captions'))
        formatting.append(ck('Caption text present',False,'Captions to evaluate','No recognizable captions'))
    s.grouped_row(CRITERIA[4],[('All four figures have captions',2,presence),('Captions above figures',1,placement),('Caption formatting and style',2,formatting)],'Three groups: captions present (2 marks), above figures (1 mark), formatting/style (2 marks). Each group earns its weight when its measured checks pass; repeated failures within a group do not stack. Placement/style are assessed on present captions, avoiding a second deduction solely for an absent caption.')

    s.area=CRITERIA[6]
    list_checks=[]
    baseline_lists=[text(p) for p in original.ps if original.list_level(p) is not None and text(p)]
    for label in baseline_lists:
        p=next((p for p in d.ps if same_text(text(p),label)),None)
        lvl=d.list_level(p) if p is not None else None
        list_checks.append(ck('Applied list numbering',lvl is not None,'Real Word list for '+label, 'Present' if lvl is not None else 'Missing or typed bullet',p))
        if lvl is None:
            if p is not None and label.startswith('Birth:') and d.pstyle(p)==d.sid('Heading1'):
                following=next((q for q in d.ps[d.index[p]+1:] if text(q)),None)
                if following is not None and d.list_level(following) is not None:
                    s.accept(list_checks[-1],'Teacher benchmark accepts the Birth introductory line as a heading before the remaining real list items.')
            continue
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
        indents={}
        for layer in [lvl.find('w:pPr',NS),*d.layers(sid,p=p,kind='pPr')]:
            ind=layer.find('w:ind',NS) if layer is not None else None
            if ind is not None:indents.update({k.split('}')[-1]:v for k,v in ind.attrib.items()})
        left=twips(indents.get('start',indents.get('left')))
        hanging=twips(indents.get('hanging','0'))
        first=twips(indents.get('firstLine','0'))
        bullet=left-hanging+first if None not in (left,hanging,first) else None
        list_checks.append(ck('List indentation',near(bullet,360),'Bullet position: 0.25 in from left margin',f'Bullet: {inches(bullet)}; text: {inches(left)}; hanging indent: {inches(hanging)}',p))
        if near(left,720) and near(bullet,720):s.accept(list_checks[-1],'Teacher benchmark accepts a half-inch list position as an alternative to the quarter-inch specification.')
    s.row(CRITERIA[6],list_checks,0 if not list_checks else None)

    s.area=CRITERIA[7]
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
        source_checks.append(ck('First source '+name,same_text(actual,want,typos=name not in ('Year','SourceType')),want,actual))
    persons=source.xpath('.//b:Person',namespaces=NS) if source is not None else []
    author_ok=any(same_text(p.findtext('b:Last',namespaces=NS),'Olson') and ''.join(part[0] for part in re.findall(r'[A-Za-z]+',p.findtext('b:First',default='',namespaces=NS)+' '+p.findtext('b:Middle',default='',namespaces=NS))).upper()=='SF' for p in persons)
    source_checks.append(ck('First source author',author_ok,'Olson, S. F.', '; '.join(' '.join(p.itertext()) for p in persons) or 'Missing'))
    updated=all(c.passed for c in source_checks)
    checks=[ck('Citation fields',bool(citation_fields),'Real CITATION fields',len(citation_fields)),ck('APA 6 bibliography setting',apa,'APA sixth edition', [r.attrib for r in sources])]+source_checks
    source_tags={q.findtext('b:Tag',namespaces=NS) for r in sources for q in r.findall('b:Source',NS)}
    unresolved=[]
    for f in citation_fields:
        m=re.match(r'\s*CITATION\s+(?:"([^"]+)"|(\S+))',f['code'],re.I)
        if not m or (m.group(1) or m.group(2)) not in source_tags: unresolved.append(f['code'])
    checks.append(ck('Citation sources resolved',not unresolved,'Every citation tag resolves to a saved source',unresolved or 'All resolved'))
    s.grouped_row(CRITERIA[7],[('Word in-text citations present',2,[checks[0]]),('APA 6 citation style',1,[checks[1]]),('First source book details',1,source_checks[:-1]),('First source author and source links',1,[source_checks[-1],checks[-1]])],'Credit is retained for existing Word citations (2), APA 6 style (1), first-source book details (1), and author/source links (1). Wrong style or incomplete source data does not erase presence credit.')
    if not citation_fields:
        s.rows[-1]['score']=0;s.rows[-1]['strict_score']=0

    s.area=CRITERIA[8]
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
    elif entries: bibscore=None
    else: bibscore=0
    s.row(CRITERIA[8],bibchecks,bibscore)

    s.area=CRITERIA[9]
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
    cover_only=False
    separate_cover=False
    if titleps and len(d.sections)>1:
        first_end=d.index.get(d.sections[0].getparent().getparent(),-1)
        cover_last=max(d.index[p] for p in titleps)
        following=next((d.index[p] for p in d.ps[cover_last+1:] if text(p)),len(d.ps))
        cover_only=first_end>=cover_last and first_end<following
    for i,section in enumerate(d.sections):
        for ref in section.findall('w:footerReference',NS): inherited[ref.get(W+'type')]=ref
        ref=inherited.get('default'); fr=footer(ref) if ref is not None else None
        default_ids.append(ref.get(R+'id') if ref is not None else None)
        fs=(d.parse_fields(fr)+d.title_controls(fr)) if fr is not None else []
        empty_default=fr is None or (not text(fr) and not fs)
        cover_without_footer=i==0 and cover_only and empty_default
        if cover_without_footer:separate_cover=True
        if i==1 and separate_cover and enabled(section.find('w:titlePg',NS)):
            active=inherited.get('first')
            fr=footer(active) if active is not None else None
            fs=(d.parse_fields(fr)+d.title_controls(fr)) if fr is not None else []
            if fr is None or not text(fr):
                checks.append(ck('Displayed run formatting',False,'Segoe UI, 9 pt, Dark Gray on page 2','The following page has an empty first-page footer'))
                checks.append(ck('Footer right aligned',False,'Right-aligned footer on page 2','The following page footer is missing'))
        page=[f for f in fs if re.match(r'\s*PAGE\b',f['code'],re.I)]
        title=[f for f in fs if re.match(r'\s*(TITLE\b|DOCPROPERTY\s+"?Title\b|Bound document Title)',f['code'],re.I)]
        if not cover_without_footer:
            checks.append(ck(f'Section {i+1} footer fields',bool(page and title),'PAGE and document title fields',[f['code'] for f in fs]))
        for p in fr.findall('.//w:p',NS) if fr is not None else []:
            if not text(p) and not d.parse_fields(p): continue
            checks.append(ck('Footer right aligned',d.right_aligned(p),'Right alignment or leading right tab',d.prop(d.pstyle(p),'jc',p=p,kind='pPr') or ('Right tab' if d.right_aligned(p) else 'Left/default')))
            checks+=s.formatting(d.pstyle(p),9,GRAY,p)[-1:]
            # Text between field results should include a separator.
        if i==0:
            first_ref=inherited.get('first'); first_footer=footer(first_ref) if first_ref is not None else None
            firstfs=d.parse_fields(first_footer) if first_footer is not None else []
            hidden=(cover_without_footer or enabled(section.find('w:titlePg',NS))) and not any(re.match(r'\s*PAGE\b',f['code'],re.I) for f in firstfs)
            checks.append(ck('Title-page number hidden',hidden,'Different first page with no PAGE field','Hidden' if hidden else 'Not hidden'))
            start=val(section,'w:pgNumType',attr='start')
            next_start=val(d.sections[1],'w:pgNumType',attr='start') if len(d.sections)>1 else None
            correct_count=start in (None,'1') and (not cover_without_footer or next_start in (None,'2'))
            checks.append(ck('Main page numbering starts at 1',correct_count,'Title counts as 1; a footer-free cover may be followed by page 2',f'Cover start: {start or "default 1"}; next section: {next_start or "continues numbering"}'))
    footer_groups=[('Font size and style',1,[]),('Right alignment',1,[]),('Title page counted as one, number hidden',1,[]),('Page number and Document Title field',1,[])]
    for c in checks:
        group=0 if c.name=='Displayed run formatting' else 1 if c.name=='Footer right aligned' else 2 if c.name in ('Title-page number hidden','Main page numbering starts at 1') else 3
        footer_groups[group][2].append(c)
    for name,weight,evidence in footer_groups:
        if not evidence:evidence.append(ck(name,False,'Required footer feature','Footer missing'))
    s.grouped_row(CRITERIA[9],footer_groups,'Start at 5. Deduct 1 per failed category: font/size/style, right alignment, title-page counting/hidden number, PAGE and Document Title fields. Repeated instances count once. If all four categories fail, award 0.',page_numbers=True)

    s.area=CRITERIA[10]
    checks=[ck('Separate references section',len(d.sections)>=2,'At least two sections',len(d.sections))]
    if len(d.sections)>=2:
        last=d.sections[-1]; before=d.sections[-2]
        checks.append(ck('New section starts next page',val(last,'w:type','nextPage')=='nextPage','nextPage',val(last,'w:type','nextPage')))
        refs=last.findall('w:footerReference',NS)
        own=next((r for r in refs if r.get(W+'type')=='default'),None)
        independent=own is not None and len(default_ids)>=2 and default_ids[-1] != default_ids[-2]
        own_first=next((r for r in refs if r.get(W+'type')=='first'),None)
        previous_first=None
        for sec in d.sections[:-1]:
            for r in sec.findall('w:footerReference',NS):
                if r.get(W+'type')=='first':previous_first=r.get(R+'id')
        if enabled(last.find('w:titlePg',NS)) and own_first is not None and own_first.get(R+'id')!=previous_first:
            first_footer=footer(own_first)
            independent=independent or (first_footer is not None and bool(re.search(r'\bPAGE\b',' '.join(f['code'] for f in d.parse_fields(first_footer)),re.I)))
        checks.append(ck('References footer unlinked',independent,'Distinct explicit default footer reference',default_ids))
        if not refs and bibliography and enabled(last.find('w:titlePg',NS)) and all(default_ids) and len(set(default_ids))>1:
            s.accept(checks[-1],'Teacher benchmark accepts an inherited references footer when the bibliography is in a separate final section and independent footers already exist earlier. The footer is still linked; this exception does not assert otherwise.')
        checks.append(ck('References letter numbering',val(last,'w:pgNumType',attr='fmt')=='upperLetter','upperLetter',val(last,'w:pgNumType',attr='fmt')))
        checks.append(ck('References numbering starts at A',val(last,'w:pgNumType',attr='start')=='1','1 (displayed A)',val(last,'w:pgNumType',attr='start')))
        if bibliography and val(last,'w:pgNumType',attr='fmt') is None and val(last,'w:pgNumType',attr='start') is None:
            s.accept(checks[-2],'Teacher benchmark accepts references without explicit A/B/C numbering metadata. This is an instructor calibration exception, not proof of alphabetic numbering.')
            s.accept(checks[-1],'Teacher benchmark does not require an explicit A restart when the references section has no numbering override.')
        checks.append(ck('Main section decimal numbering',all(val(sec,'w:pgNumType',default='decimal',attr='fmt')=='decimal' for sec in d.sections[:-1]),'Previous sections retain decimal numbering',[val(sec,'w:pgNumType',default='decimal',attr='fmt') for sec in d.sections[:-1]]))
        if bibliography:
            p=bibliography[0]['p']; idx=d.index.get(p,-1)
            boundaries=[d.index.get(sec.getparent().getparent(),-1) for sec in d.sections[:-1]]
            checks.append(ck('Bibliography after last section break',idx>max(boundaries),'Bibliography field in final section',f'Bibliography paragraph {idx+1}; section-end paragraphs {[b+1 for b in boundaries]}',p))
    s.row(CRITERIA[10],checks,0 if len(d.sections)<2 else None)

    s.area=CRITERIA[11]
    checks=[]
    for i,sec in enumerate(d.sections):
        for side in ('top','bottom','left','right'):
            actual=val(sec,'w:pgMar',attr=side)
            c=ck(f'Section {i+1} {side} margin',near(actual,1440),'1 in',inches(actual))
            c.issues=['margin:'+side];checks.append(c)
            if side in ('left','right') and near(actual,1800):s.accept(c,'Instructor accepts 1.25 in side margins, as in the supplied margin reference.')
    margin_evidence=[asdict(c) for c in checks]
    overall=assess_brand(d,s.rows,margin_evidence,calibrated)
    strict=assess_brand(d,s.rows,margin_evidence,False,overall['visible_body_adherence'])
    for c in checks:c.scored=False
    for name,fraction in overall['domains'].items():
        c=ck('Overall evidence: '+name,fraction==1,'Coordinated adherence to the Brand Style Guide',f'{fraction:.0%} of measured evidence meets the guide')
        c.scored=False;checks.append(c)
    checks.append(ck('Overall style-guide adherence',overall['score']==5,'Consistent overall adherence to the Brand Style Guide',overall['summary']))
    checks.append(ck('Rendered appearance',None,'Consistent visual layout','Physical page balance and exact rendered appearance remain unverified.'))
    s.row(CRITERIA[11],checks,overall['score'],overall['policy'])
    s.rows[-1]['strict_score']=strict['score']
    s.rows[-1]['overall_assessment']=overall
    s.rows[-1]['accepted_variants']=sum(bool(c.accepted) for c in checks)

    extra=''; sentence_note='Required heading or first body paragraph not found.'; count=None
    target=next((i for i,p in enumerate(d.ps) if same_text(text(p),'A Multifaceted Approach')),None)
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
    return {'filename':filename,'application':APP_TITLE,'course':COURSE_NAME,'term':TERM,'version':VERSION,'total':sum(r['score'] for r in s.rows),'maximum':60,'strict_total':sum(r['strict_score'] for r in s.rows),
            'scoring_mode':'Teacher calibrated' if calibrated else 'Written requirements',
            'estimated':any(r['estimated'] for r in s.rows),'criteria':s.rows,'sentence':sentence,
            'limitations':['Scores use the documented deterministic rubric mapping; qualitative partial-credit boundaries are implementation policy, not extra instructions from the rubric.', 'Page isolation, exact rendering, identity of student-information lines and sentence context cannot be established fully from OOXML. Scores involving those checks are marked estimated.']}


def grade_paths(paths):
    """A bad submission must not interrupt grading the remaining batch."""
    results=[]
    for path in paths:
        path=Path(path).resolve()
        try:
            if path.stat().st_size>15_000_000: raise ValueError('File exceeds 15 MB.')
            result=grade(path.read_bytes(),path.name)
        except (ValueError,OSError,ET.LxmlError) as exc:
            result={'filename':path.name,'error':str(exc)}
        result['source_path']=str(path)
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
    lines=[APP_TITLE,COURSE_TEXT,f'{result["filename"]}: {result["total"]}/60'+(' (estimated)' if result['estimated'] else ''),f'Engine {result["version"]}',f'Scoring: {result.get("scoring_mode","Automatic")}; written-requirement diagnostic total: {result.get("strict_total",result["total"])} / 60','']
    for row in result['criteria']:
        lines += [f'{row["criterion"]}: {row["score"]}/5',row['policy']]
        for c in row['checks']:
            label='INFO' if not c.get('scored',True) else 'ACCEPTED' if c.get('accepted') else 'PASS' if c['passed'] else 'UNVERIFIED' if c['passed'] is None else 'FAIL'
            lines += [f'  [{label}] {c["name"]} — {c["location"]}',f'    Expected: {c["expected"]}',f'    Found: {c["found"]}'+(f'\n    Accepted alternative: {c["accepted"]}' if c.get('accepted') else '')]
        lines.append('')
    lines += ['Additional sentence (no separate rubric points):',str(result['sentence']),*result['limitations']]
    return '\n'.join(lines)
