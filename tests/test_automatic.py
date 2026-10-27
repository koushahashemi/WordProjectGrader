from copy import deepcopy
from io import BytesIO
import json
from pathlib import Path
from zipfile import ZipFile
import pytest
from lxml import etree as E
from autograder import Document, grade, grade_paths, csv_report, CRITERIA, W, NS
from grader import BASE, package, xml


def node(tag, **attrs):
    n=E.Element(W+tag)
    for k,v in attrs.items(): n.set(W+k,str(v))
    return n


def ensure(parent,path):
    for name in path.split('/'):
        child=parent.find('w:'+name,NS)
        if child is None:child=node(name);parent.append(child)
        parent=child
    return parent


def setprop(parent,path,**attrs):
    n=ensure(parent,path)
    for k,v in attrs.items():n.set(W+k,str(v))
    return n


def paragraph(value='',style='Normal',field=None,center=False,page=False):
    p=node('p');setprop(p,'pPr/pStyle',val=style)
    if center:setprop(p,'pPr/jc',val='center')
    if page:ensure(p,'pPr/pageBreakBefore')
    r=node('r');p.append(r)
    if field:
        simple=node('fldSimple',instr=field);p.remove(r);p.append(simple);simple.append(r)
    t=node('t');t.text=value;r.append(t)
    return p


def repack(roots,original=BASE.read_bytes(),extra=None):
    dest=BytesIO()
    with ZipFile(BytesIO(original)) as z,ZipFile(dest,'w') as out:
        for name in z.namelist():out.writestr(name,E.tostring(roots[name]) if name in roots else z.read(name))
        for name,root in (extra or {}).items():
            if name not in z.namelist():out.writestr(name,E.tostring(root))
    return dest.getvalue()


@pytest.fixture
def compliant():
    roots={k:xml(v) for k,v in package(BASE.read_bytes()).items()}
    styles=roots['word/styles.xml']
    specs={'Normal':(11,'4E4E4E'),'Title':(26,'00796B'),'Heading1':(18,'00796B'),'Heading2':(16,'00796B'),'Heading3':(14,'00796B'),'Caption':(10,'4E4E4E'),'TOC1':(11,'4E4E4E'),'TOC2':(11,'4E4E4E'),'TOC3':(11,'4E4E4E'),'TOCHeading':(18,'00796B'),'TableofFigures':(11,'4E4E4E'),'Bibliography':(11,'4E4E4E'),'Footer':(9,'4E4E4E')}
    for sid,(size,color) in specs.items():
        matches=styles.xpath('./w:style[@w:styleId=$s]',namespaces=NS,s=sid)
        s=matches[0] if matches else node('style',styleId=sid,type='paragraph')
        if not matches:styles.append(s);setprop(s,'name',val=sid);setprop(s,'basedOn',val='Normal')
        rp=ensure(s,'rPr')
        for child in list(rp):rp.remove(child)
        setprop(s,'rPr/rFonts',ascii='Segoe UI',hAnsi='Segoe UI')
        setprop(s,'rPr/sz',val=size*2);setprop(s,'rPr/color',val=color)
        if sid=='Caption':ensure(s,'rPr/i')
        if sid in ('TOC1','TOC2','TOC3','TableofFigures'):setprop(s,'pPr/tabs/tab',val='right',leader='dot',pos='9000')
        if sid=='Title':
            border=setprop(s,'pPr/pBdr/bottom',val='single',sz='8',color='00796B')
            for attr in ('themeColor','themeShade','themeTint'):border.attrib.pop(W+attr,None)
    doc=roots['word/document.xml'];body=doc.find('w:body',NS)
    original=list(body)
    for c in list(body):body.remove(c)
    body.append(paragraph('Bears: An In-Depth Exploration','Title','TITLE',True))
    for value in ['Jordan Lee','12345678','Set A']:
        body.append(paragraph('',center=True));body.append(paragraph(value,center=True))
    body.append(paragraph('Contents','TOCHeading',page=True))
    body.append(paragraph('Introduction to Bears\t4','TOC1','TOC \\o "1-3" \\h \\z \\u'))
    body.append(paragraph('List of Figures','Heading1',page=True))
    body.append(paragraph('Figure 1: Brown bear in its natural environment.\t4','TableofFigures','TOC \\c "Figure" \\h'))
    # Original text and headings remain; replace originals' figure labels with real captions.
    images=[]
    for p in original:
        if p.tag==W+'p' and p.xpath('.//w:drawing',namespaces=NS):images.append(deepcopy(p))
    fig=0
    for p in original:
        if p.tag==W+'sectPr':continue
        if p.xpath('.//w:drawing',namespaces=NS):continue
        value=''.join(p.xpath('.//w:t/text()',namespaces=NS))
        if value.startswith('Figure '):
            fig+=1;body.append(paragraph(value,'Caption','SEQ Figure \\* ARABIC'));body.append(images[fig-1]);continue
        # Normalize direct run formatting, preserving fields and text.
        for rp in p.xpath('.//w:rPr',namespaces=NS):
            for c in list(rp):
                if c.tag in [W+x for x in ('rFonts','sz','color','i','rStyle')]:rp.remove(c)
        if value=='Introduction to Bears':ensure(p,'pPr/pageBreakBefore')
        body.append(p)
    section=node('sectPr');setprop(section,'pgMar',top='1440',bottom='1440',left='1440',right='1440');ensure(section,'titlePg')
    setprop(section,'pgNumType',fmt='decimal',start='1')
    ref=node('footerReference',type='default');ref.set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id','rIdAutoFooter1');section.append(ref)
    end=paragraph();ensure(end,'pPr').append(section);body.append(end)
    body.append(paragraph('References','Heading1'))
    body.append(paragraph('Olson, S. F. (1956). The Singing Wilderness. New York: Alfred A. Knopf.','Bibliography','BIBLIOGRAPHY'))
    last=deepcopy(section);last.remove(last.find('w:titlePg',NS));setprop(last,'pgNumType',fmt='upperLetter',start='1')
    last.find('w:footerReference',NS).set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id','rIdAutoFooter2');body.append(last)
    for lvl in roots['word/numbering.xml'].findall('.//w:lvl',NS):
        if lvl.find('w:numFmt',NS) is not None:
            setprop(lvl,'pPr/ind',left='720',hanging='360');setprop(lvl,'rPr/color',val='FF6F61')
            if lvl.find('w:numFmt',NS).get(W+'val')=='bullet':setprop(lvl,'lvlText',val='●')
    # The ListBullet style inherits the modified Normal style.
    for sty in styles.xpath('./w:style[@w:styleId="ListBullet"]',namespaces=NS):
        setprop(sty,'pPr/ind',left='720')
    source=roots['customXml/item1.xml'];source.set('SelectedStyle','APASixthEditionOfficeOnline.xsl');source.set('StyleName','APA');source.set('Version','6')
    b='{'+NS['b']+'}'
    placeholder=source.find('b:Source',NS)
    for key,value in {'SourceType':'Book','Title':'The Singing Wilderness','Year':'1956','City':'New York','Publisher':'Alfred A. Knopf'}.items():
        c=E.SubElement(placeholder,b+key);c.text=value
    author=E.SubElement(placeholder,b+'Author');author=E.SubElement(author,b+'Author');names=E.SubElement(author,b+'NameList');person=E.SubElement(names,b+'Person')
    for key,value in {'Last':'Olson','First':'S.','Middle':'F.'}.items():E.SubElement(person,b+key).text=value
    with ZipFile(BASE) as z:rels=xml(z.read('word/_rels/document.xml.rels'))
    extra={}
    for i in (1,2):
        rel=E.SubElement(rels,'{http://schemas.openxmlformats.org/package/2006/relationships}Relationship',Id=f'rIdAutoFooter{i}',Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',Target=f'footerAuto{i}.xml')
        footer=node('ftr');p=paragraph('Bears: An In-Depth Exploration ','Footer','TITLE');setprop(p,'pPr/jc',val='right');page=E.SubElement(p,W+'fldSimple',{W+'instr':'PAGE'});run=E.SubElement(page,W+'r');E.SubElement(run,W+'t').text='1';footer.append(p);extra[f'word/footerAuto{i}.xml']=footer
    roots['word/_rels/document.xml.rels']=rels
    # Ensure XML relationship part replaced, not omitted by package()'s XML filter.
    return repack(roots,extra=extra)


def row(result,name):return next(r for r in result['criteria'] if r['criterion']==name)


def test_baseline_auto_scores_without_manual_input():
    r=grade(BASE.read_bytes(),'Bears.docx')
    assert [q['criterion'] for q in r['criteria']]==CRITERIA
    assert all(q['score'] in (0,1,3,5) for q in r['criteria'])
    assert r['total']==sum(q['score'] for q in r['criteria'])
    for name in ['Title Page','Table of Contents (TOC)','Figures and Captions','List of Figures','Bibliography / References','Page Numbers','Section Break and Page Numbering']:assert row(r,name)['score']==0


def test_compliant_structure_scores_full_with_explicit_layout_limits(compliant):
    r=grade(compliant,'JordanLee_BearsDocument.docx')
    failures={x['criterion']:[c for c in x['checks'] if c['passed'] is False] for x in r['criteria'] if x['score']!=5}
    assert r['total']==60, json.dumps(failures,indent=2)
    assert r['estimated']
    assert any(c['passed'] is None for x in r['criteria'] for c in x['checks'])


def mutate(data,fn):
    roots={k:xml(v) for k,v in package(data).items()};fn(roots)
    return repack(roots,original=data)


def test_direct_formatting_cannot_replace_modified_styles(compliant):
    def change(roots):
        normal=roots['word/styles.xml'].xpath('./w:style[@w:styleId="Normal"]',namespaces=NS)[0]
        setprop(normal,'rPr/rFonts',ascii='Arial',hAnsi='Arial')
    r=grade(mutate(compliant,change),'bad.docx')
    assert row(r,'Body Text')['score']<5
    assert any(c['passed'] is False and 'font' in c['name'] for c in row(r,'Body Text')['checks'])


def test_only_applied_bullet_definitions_are_scored(compliant):
    def change(roots):
        num=roots['word/numbering.xml']
        for c in num.findall('.//w:lvl/w:rPr/w:color',NS):c.set(W+'val','000000')
    r=grade(mutate(compliant,change),'bad.docx')
    assert row(r,'Bulleted Lists')['score']<5
    assert any(c['passed'] is False and c['name']=='Bullet color' for c in row(r,'Bulleted Lists')['checks'])


def test_figure_list_is_not_main_toc(compliant):
    def change(roots):
        for f in roots['word/document.xml'].xpath('//w:fldSimple',namespaces=NS):
            if f.get(W+'instr','').startswith('TOC \\o'):f.set(W+'instr','HYPERLINK "example"')
    r=grade(mutate(compliant,change),'bad.docx')
    # The heading list remains as a manually created TOC: it can earn partial, not full, credit.
    assert row(r,'Table of Contents (TOC)')['score']==3
    assert row(r,'Table of Contents (TOC)')['checks'][0]['passed'] is False
    assert row(r,'List of Figures')['score']==5


@pytest.mark.parametrize('apa,updated,expected',[(True,True,5),(False,True,3),(True,False,1),(False,False,0)])
def test_exact_citation_rubric_levels(compliant,apa,updated,expected):
    def change(roots):
        sources=roots['customXml/item1.xml']
        if not apa:sources.set('SelectedStyle','MLASeventhEditionOfficeOnline.xsl');sources.set('StyleName','MLA');sources.set('Version','7')
        if not updated:sources.find('b:Source/b:Year',NS).text='2000'
    assert row(grade(mutate(compliant,change),'test.docx'),'Citations (In-text)')['score']==expected


def test_linked_footer_and_wrong_letters_lose_marks(compliant):
    def change(roots):
        sections=roots['word/document.xml'].findall('.//w:sectPr',NS)
        section=sections[-1];section.remove(section.find('w:footerReference',NS));setprop(section,'pgNumType',fmt='decimal',start='2')
    r=grade(mutate(compliant,change),'test.docx')
    assert row(r,'Section Break and Page Numbering')['score']<5
    assert any(c['name']=='References footer unlinked' and not c['passed'] for c in row(r,'Section Break and Page Numbering')['checks'])


def test_batch_continues_past_bad_file_and_csv_neutralizes_names(tmp_path,compliant):
    a=tmp_path/'good.docx';a.write_bytes(compliant)
    b=tmp_path/'bad.docx';b.write_bytes(b'not a Word file')
    results=grade_paths([b,a])
    assert 'error' in results[0] and results[1]['total']==60
    results[1]['filename']='=formula.docx'
    csv=csv_report(results)
    assert "'=formula.docx" in csv and 'Error:' in csv


def test_style_inheritance_and_theme_font_resolution(compliant):
    def change(roots):
        styles=roots['word/styles.xml']
        caption=styles.xpath('./w:style[@w:styleId="Caption"]',namespaces=NS)[0]
        caption.find('w:rPr',NS).remove(caption.find('w:rPr/w:rFonts',NS))
    d=Document(mutate(compliant,change))
    assert d.format('Caption')['font']=='Segoe UI'


def test_field_instructions_split_across_runs():
    root=xml(b'<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>TO</w:instrText></w:r><w:r><w:instrText>C \\o "1-3"</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Entry</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p></w:body></w:document>')
    d=Document(BASE.read_bytes())
    assert d.parse_fields(root)[0]['code'].startswith('TOC')


def test_oversized_file_rejected():
    with pytest.raises(ValueError,match='15 MB'):grade(b'x'*15_000_001,'huge.docx')


def test_manual_toc_gets_partial_credit_not_missing(compliant):
    def change(roots):
        for f in roots['word/document.xml'].xpath('//w:fldSimple',namespaces=NS):
            if f.get(W+'instr','').startswith('TOC \\o'):
                parent=f.getparent()
                for child in list(f):parent.append(child)
                parent.remove(f)
    r=grade(mutate(compliant,change),'typed.docx')
    assert row(r,'Table of Contents (TOC)')['score']==3


def test_typed_title_cover_gets_rubric_partial_credit(compliant):
    def change(roots):
        f=roots['word/document.xml'].xpath('//w:fldSimple[@w:instr="TITLE"]',namespaces=NS)[0]
        parent=f.getparent()
        for child in list(f):parent.append(child)
        parent.remove(f)
    data=mutate(compliant,change)
    assert row(grade(data,'typed.docx',calibrated=False),'Title Page')['score']==1
    assert row(grade(data,'typed.docx'),'Title Page')['score']==5


def test_full_author_name_is_equivalent_to_initials(compliant):
    def change(roots):roots['customXml/item1.xml'].find('b:Source/b:Author/b:Author/b:NameList/b:Person/b:First',NS).text='Sigurd'
    r=grade(mutate(compliant,change),'full-name.docx')
    assert row(r,'Citations (In-text)')['score']==5


def test_table_of_figures_is_a_valid_title(compliant):
    def change(roots):
        for t in roots['word/document.xml'].findall('.//w:t',NS):
            if t.text=='List of Figures':t.text='Table of Figures'
    assert row(grade(mutate(compliant,change),'figures.docx'),'List of Figures')['score']==5


def test_optional_coral_callouts_are_not_body_errors(compliant):
    def change(roots):
        for p in roots['word/document.xml'].findall('.//w:body/w:p',NS):
            value=''.join(p.xpath('.//w:t/text()',namespaces=NS))
            if value.startswith('Bears play a vital ecological role'):
                setprop(p,'pPr/ind',left='720',right='720')
                for r in p.findall('.//w:r',NS):setprop(r,'rPr/color',val='FF6F61');ensure(r,'rPr/i')
                break
    r=grade(mutate(compliant,change),'callout.docx')
    assert row(r,'Body Text')['score']==5


def test_coral_body_text_without_callout_indents_still_loses_marks(compliant):
    def change(roots):
        for p in roots['word/document.xml'].findall('.//w:body/w:p',NS):
            if ''.join(p.xpath('.//w:t/text()',namespaces=NS)).startswith('Bears play a vital ecological role'):
                for r in p.findall('.//w:r',NS):setprop(r,'rPr/color',val='FF6F61')
                break
    assert row(grade(mutate(compliant,change),'bad.docx'),'Body Text')['score']<5


def test_body_coverage_handles_updated_opening_citation(compliant):
    def change(roots):
        for p in roots['word/document.xml'].findall('.//w:body/w:p',NS):
            if ''.join(p.xpath('.//w:t/text()',namespaces=NS)).startswith('“The great brown bear'):
                # Changing cached citation text must not look like deletion of the paragraph.
                for t in p.findall('.//w:t',NS):
                    if 'Placeholder' in (t.text or ''):t.text=t.text.replace('Placeholder','Olson, 1956')
    assert row(grade(mutate(compliant,change),'citation.docx'),'Body Text')['score']==5


def test_property_content_control_footer_and_right_tab(compliant):
    def change(roots):
        styles=roots['word/styles.xml'];style=styles.xpath('./w:style[@w:styleId="Footer"]',namespaces=NS)[0]
        setprop(style,'rPr/sz',val='22');setprop(style,'pPr/tabs/tab',val='right',pos='9360')
        for k,root in roots.items():
            if not k.startswith('word/footer'):continue
            p=root.find('w:p',NS)
            if p is None:continue
            jc=p.find('w:pPr/w:jc',NS)
            if jc is not None:jc.getparent().remove(jc)
            lead=node('r');lead.append(node('tab'));p.insert(1,lead)
            for field in list(p.findall('w:fldSimple',NS)):
                if field.get(W+'instr')=='TITLE':
                    control=node('sdt');setprop(control,'sdtPr/dataBinding',xpath='/cp:coreProperties[1]/dc:title[1]',prefixMappings="xmlns:dc='http://purl.org/dc/elements/1.1/'")
                    contents=ensure(control,'sdtContent')
                    for r in list(field):contents.append(r)
                    p.replace(field,control)
            for r in p.findall('.//w:r',NS):setprop(r,'rPr/sz',val='18')
    r=grade(mutate(compliant,change),'footer.docx')
    assert row(r,'Page Numbers')['score']==5


def test_applied_title_border_is_explicit_calibration_exception(compliant):
    def change(roots):
        title=roots['word/styles.xml'].xpath('./w:style[@w:styleId="Title"]',namespaces=NS)[0]
        setprop(title,'pPr/pBdr/bottom',val='single',sz='8',color='4F81BD')
        for p in roots['word/document.xml'].xpath('//w:p[w:pPr/w:pStyle/@w:val="Title"]',namespaces=NS):setprop(p,'pPr/pBdr/bottom',val='single',sz='8',color='00796B')
    data=mutate(compliant,change);r=grade(data,'border.docx')
    assert row(r,'Headings')['score']==5
    assert row(r,'Headings')['strict_score']<5
    assert any(c['accepted'] and c['passed'] is False for c in row(r,'Headings')['checks'])
    assert row(grade(data,'border.docx',calibrated=False),'Headings')['score']<5


def test_quarter_and_half_inch_lists_both_earn_calibrated_credit(compliant):
    def change(roots):
        for indent in roots['word/numbering.xml'].findall('.//w:lvl/w:pPr/w:ind',NS):indent.set(W+'left','720');indent.set(W+'hanging','0')
        for sty in roots['word/styles.xml'].xpath('./w:style[@w:styleId="ListBullet"]',namespaces=NS):setprop(sty,'pPr/ind',left='720')
    r=grade(mutate(compliant,change),'lists.docx')
    assert row(r,'Bulleted Lists')['score']==5
    assert row(r,'Bulleted Lists')['accepted_variants']>0


def test_wrong_list_indent_still_deducts(compliant):
    def change(roots):
        for indent in roots['word/numbering.xml'].findall('.//w:lvl/w:pPr/w:ind',NS):indent.set(W+'left','1440')
        for sty in roots['word/styles.xml'].xpath('./w:style[@w:styleId="ListBullet"]',namespaces=NS):setprop(sty,'pPr/ind',left='1440')
    assert row(grade(mutate(compliant,change),'lists.docx'),'Bulleted Lists')['score']<5


def test_unlinked_first_page_reference_footer_is_recognized(compliant):
    def change(roots):
        sections=roots['word/document.xml'].findall('.//w:sectPr',NS)
        last=sections[-1];ensure(last,'titlePg');ref=last.find('w:footerReference',NS);ref.set(W+'type','first')
    r=grade(mutate(compliant,change),'first-footer.docx')
    assert next(c for c in row(r,'Section Break and Page Numbering')['checks'] if c['name']=='References footer unlinked')['passed']


def test_reference_calibration_exceptions_stay_visible(compliant):
    def change(roots):
        doc=roots['word/document.xml'];body=doc.find('w:body',NS)
        sections=doc.findall('.//w:sectPr',NS)
        first=deepcopy(sections[0]);new=paragraph();ensure(new,'pPr').append(first)
        body.insert(8,new)
        # Independent earlier main-section footer; final references inherit it.
        sections[0].find('w:footerReference',NS).set('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id','rIdAutoFooter2')
        last=sections[-1];last.remove(last.find('w:footerReference',NS));last.remove(last.find('w:pgNumType',NS));ensure(last,'titlePg')
    data=mutate(compliant,change);r=grade(data,'references.docx')
    section=row(r,'Section Break and Page Numbering')
    assert section['score']==5 and section['strict_score']<5
    assert section['accepted_variants']==3
    assert row(grade(data,'references.docx',calibrated=False),'Section Break and Page Numbering')['score']<5
