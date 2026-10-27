"""Holistic adherence to the assignment's style guide, independent of point totals."""
import math
from grader import CRITERIA, NS
from matching import near


def assess_brand(document,rows,margins,accept_alternatives=True,visible_body=None):
    specs=[('Title',26,'00796B'),('Heading1',18,'00796B'),('Heading2',16,'00796B'),('Heading3',14,'00796B'),('Normal',11,'4E4E4E')]
    formats=[(document.format(document.sid(sid)),size,color) for sid,size,color in specs]
    typography=sum(f['font']=='Segoe UI' for f,_,_ in formats)+sum(f['size']==str(size*2) for f,size,_ in formats)
    typography/=10
    colors=sum(f['color']==color for f,_,color in formats)/5
    coordinated=sum(f['font']=='Segoe UI' and f['size']==str(size*2) and f['color']==color for f,size,color in formats)/5
    byname={r['criterion']:r for r in rows}
    def ok(c):return bool(c['passed'] or (accept_alternatives and c.get('accepted')))
    def ratio(checks):return sum(ok(c) for c in checks)/len(checks) if checks else 0
    headings=[c for c in byname[CRITERIA[0]]['checks'] if c['name']=='Heading style application']
    bodies=[c for c in byname[CRITERIA[1]]['checks'] if c['name']=='Body paragraph style']
    heading_use=ratio(headings)
    application=(heading_use+ratio(bodies))/2
    # Aggregate distinct visual/style properties within each supporting component.
    # Functional citation/source errors and other rubric point totals are not inputs.
    support=[]
    for name in [CRITERIA[3],CRITERIA[4],CRITERIA[5],CRITERIA[6],CRITERIA[9]]:
        checks=byname[name]['checks'];buckets={}
        for c in checks:
            key=c['name']
            if not (' style ' in key or key in ('Displayed run formatting','Caption style applied','List title style','Bullet color','Bullet shape','List indentation','Footer right aligned')):continue
            if c['passed'] is None:continue
            keys=c.get('issues') or [key]
            for issue in keys:buckets.setdefault(issue,[]).append(ok(c))
        support.append(sum(all(values) for values in buckets.values())/len(buckets) if buckets else 0)
    supporting=sum(support)/len(support)
    sides={side:[] for side in ('top','bottom','left','right')}
    for c in margins:
        for side in sides:
            if c['name'].endswith(side+' margin'):sides[side].append(ok(c))
    layout=sum(bool(values) and all(values) for values in sides.values())/4
    # Include visible long-form body text, not only stored style definitions.
    if visible_body is None:
        body_totals=[0,0,0];characters=0
        for p in document.ps:
            if document.pstyle(p)!=document.sid('Normal'):continue
            if len(''.join(p.xpath('.//w:t/text()',namespaces=NS)))<100:continue
            callout=near(document.prop(document.pstyle(p),'ind',attr='left',p=p,kind='pPr'),720) and near(document.prop(document.pstyle(p),'ind',attr='right',p=p,kind='pPr'),720)
            for run in p.findall('.//w:r',NS):
                count=len(''.join(run.xpath('.//w:t/text()',namespaces=NS)))
                if not count:continue
                f=document.format(document.pstyle(p),p,run);characters+=count
                body_totals[0]+=count*(f['font']=='Segoe UI')
                body_totals[1]+=count*(f['size']=='22')
                body_totals[2]+=count*(f['color']=='4E4E4E' or (callout and f['color']=='FF6F61'))
        visible_body=[v/characters for v in body_totals] if characters else [0,0,0]
    typography=(typography+(visible_body[0]+visible_body[1])/2)/2
    colors=(colors+visible_body[2])/2
    domains={'Typography system':typography,'Brand colors':colors,'Style application':application,
             'Lists, captions, contents and footers':supporting,'Margins and layout':layout}
    # Zero benchmark: direct formatting of body prose alone does not establish a
    # coordinated brand when the named style system, headings and supporting
    # elements overwhelmingly fail the guide. No identity or document hash is used.
    zero_anchor=coordinated<=.2 and heading_use<.2 and supporting<.5
    mean=sum(domains.values())/len(domains)
    score=min(5,math.floor(mean*5+.5))
    layout_cap=0 if layout==0 else 1+math.floor(4*layout)
    if heading_use<.5 or min(visible_body)<.8:score=min(score,4)
    if sum(v<.8 for v in visible_body)>=2:score=min(score,3)
    score=0 if zero_anchor else min(score,layout_cap)
    levels={5:'Consistent overall adherence',4:'Mostly consistent with limited weaknesses',3:'Mixed adherence across the document',
            2:'Limited coordinated adherence',1:'Only isolated elements meet the guide',0:'No meaningful coordinated brand system'}
    summary=levels[score]+'.'
    if score<5:
        weaknesses=[name.lower() for name,value in domains.items() if value<.9]
        if weaknesses:summary+=' Areas needing attention: '+', '.join(weaknesses)+'.'
    if zero_anchor:
        summary+=' The named styles, heading application and supporting elements do not establish the required coordinated brand. Correct margins or directly formatted body text alone are insufficient.'
    return {'score':score,'summary':summary,'domains':domains,'coordinated_style_fraction':coordinated,
            'visible_body_adherence':visible_body,'heading_application_fraction':heading_use,'zero_benchmark_pattern':zero_anchor,
            'policy':'Overall assessment of typography, colors, style application, supporting elements and layout. Uses raw style-guide evidence, not an average or sum of rubric scores. Five equally weighted domain proportions are rounded to a 0–5 band; margin compliance bounds the band. An absent coordinated style system with negligible heading/supporting adoption receives 0. Widespread heading-application or visible-body deviations limit full-credit bands; isolated formatting defects do not automatically cause another deduction.'}
