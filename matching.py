"""Presentation-independent text, font and measurement comparisons."""
import re
import unicodedata


def canonical(value):
    return ''.join(c for c in unicodedata.normalize('NFKC',str(value or '')).casefold() if c.isalnum())


def same_text(actual,expected,typos=True):
    a,b=canonical(actual),canonical(expected)
    if not a or not b:return False
    if a==b:return True
    # Numeric changes are content changes (years, list numbers), not spelling variations.
    if not typos or re.findall(r'\d+',a)!=re.findall(r'\d+',b):return False
    limit=2 if min(len(a),len(b))>=20 else 1
    if min(len(a),len(b))<5 or abs(len(a)-len(b))>limit:return False
    # Bounded Levenshtein comparison; avoid treating a different title as a typo.
    previous=list(range(len(b)+1))
    for i,x in enumerate(a,1):
        current=[i]
        for j,y in enumerate(b,1):current.append(min(current[-1]+1,previous[j]+1,previous[j-1]+(x!=y)))
        if min(current)>limit:return False
        previous=current
    return previous[-1]<=limit and previous[-1]/max(len(a),len(b))<=.08


def twips(value):
    if value is None:return None
    m=re.fullmatch(r'\s*([+-]?\d+(?:\.\d+)?)\s*(in|cm|mm|pt|pc)?\s*',str(value),re.I)
    if not m:return None
    factors={None:1,'in':1440,'cm':1440/2.54,'mm':1440/25.4,'pt':20,'pc':240}
    return float(m[1])*factors[m[2].lower() if m[2] else None]


def near(value,target,tolerance=8):
    actual=twips(value)
    return actual is not None and abs(actual-target)<=tolerance


def inches(value):
    number=twips(value)
    return 'not specified' if number is None else f'{number/1440:.3f}'.rstrip('0').rstrip('.')+' in'


def points(value):
    try:return f'{float(value)/2:g} pt'
    except (TypeError,ValueError):return 'not specified'


def font_family(name):
    name=str(name or '')
    italic=bool(re.search(r'italic|oblique',name,re.I))
    family=re.sub(r'(?:[\s_-]*(?:semi[ -]?bold|demi[ -]?bold|bold|italic|oblique|regular|normal|light|semilight|black|heavy|medium))+$','',name,flags=re.I).strip()
    if canonical(family)=='segoeui':family='Segoe UI'
    return family,italic


def describe_format(values):
    return ', '.join(f'{k}: {points(v) if k=="size" else "#"+v if k=="color" and v else v}' for k,v in values.items())
