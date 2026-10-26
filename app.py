import csv
import hashlib
import io
import json
import re
from pathlib import Path
import streamlit as st
from grader import analyze, CRITERIA

st.set_page_config(page_title='Bears • Word Project Grader', page_icon='🐻', layout='wide')
st.markdown('''<style>:root {--primary:#00796b;} .stApp {background:#fafaf7;} h1,h2,h3 {color:#00796b;} [data-testid="stMetric"] {background:white;border:1px solid #ddd;border-radius:12px;padding:16px;} </style>''',unsafe_allow_html=True)
st.title('🐻 Bears · Word Project Grader')
st.caption('Word Project 2026 · Together for Bears · Teacher review workspace')
st.write('Upload submissions, inspect Word evidence, and finalize the 12-criterion rubric. Scores use the assignment’s 0, 1, 3, or 5 point levels.')
with st.sidebar:
    st.header('Assignment reference')
    st.write('**60 points · 12 criteria**')
    st.caption('Built for this assignment only. No AI service or API key is used. Files are processed in memory and are not written to disk by this app.')
    for p in sorted((Path(__file__).parent/'assets').iterdir()):
        st.download_button(p.name,p.read_bytes(),file_name=p.name,key='asset-'+p.name)
    st.info('The extra sentence has no separate point allocation. Overall level thresholds in the supplied rubric need instructor clarification; this app reports points out of 60.')
    st.caption('Document XML can show saved features, but cannot prove which Word commands a student clicked. Exact pagination and contextual quality require teacher review.')

uploads=st.file_uploader('Student Word documents',type=['docx'],accept_multiple_files=True,help='Maximum 15 MB per document; expanded XML packages limited to 60 MB.')
if not uploads:
    st.info('Start by uploading a student submission. The original Bears document and assignment references are already included.')
    st.stop()

valid=[]
for upload in uploads:
    raw=upload.getvalue()
    key=hashlib.sha256(raw).hexdigest()[:24]+hashlib.sha256(upload.name.encode()).hexdigest()[:8]
    if len(raw)>15_000_000:
        st.error(f'{upload.name}: file exceeds 15 MB.'); continue
    if 'result-'+key not in st.session_state:
        try: st.session_state['result-'+key]=analyze(raw,upload.name)
        except ValueError as exc:
            st.error(f'{upload.name}: {exc}'); continue
    valid.append((key,st.session_state['result-'+key]))
if not valid: st.stop()
# Widget keys are transient when switching submissions; keep a separate durable session store.
st.session_state.setdefault('reviews', {})
def remember(widget_key):
    st.session_state['reviews'][widget_key]=st.session_state[widget_key]
def restore(widget_key,default):
    if widget_key not in st.session_state:
        st.session_state[widget_key]=st.session_state['reviews'].get(widget_key,default)

selected=st.selectbox('Review submission',range(len(valid)),format_func=lambda i:f'{i+1}. {valid[i][1]["filename"]}')
key,result=valid[selected]
a,b=st.columns([3,1])
with a:
    st.subheader(result['filename'])
    st.caption(f'{result["paragraph_count"]} top-level paragraphs inspected · Saved XML evidence, not a rendered Word preview')
with b:
    if not re.fullmatch(r'.+_BearsDocument\.docx',result['filename'],re.I): st.warning('Filename should be YourFullName_BearsDocument.docx.')

st.warning('Suggested zeroes indicate absent structural evidence, not a final grade. Confirm every criterion. Matching appearance alone does not establish correct style or field usage.')
for i,f in enumerate(result['findings']):
    with st.expander(f'{i+1}. {f["criterion"]} · {f["status"]}',expanded=i==0):
        left,right=st.columns([3,1])
        with left:
            st.markdown('**Expected**');st.write(f['expected'])
            st.markdown('**Found in document**');st.text(f['found'])
        with right:
            options=['Unreviewed',0,1,3,5]
            score_key=f'score-{key}-{i}'
            restore(score_key,'Unreviewed')
            st.selectbox('Teacher score',options,key=score_key,on_change=remember,args=(score_key,),help='5: Level 4 · 3: Level 3 · 1: Level 2 · 0: Level 1')
            if f['suggested'] is not None: st.caption(f'Suggested: {f["suggested"]}/5. Confirm using the evidence.')
        comment_key=f'comment-{key}-{i}'
        restore(comment_key,'')
        st.text_area('Feedback',key=comment_key,on_change=remember,args=(comment_key,),placeholder='Explain any deductions or record your review.')

with st.expander('Additional sentence and original-text comparison',expanded=True):
    s=result['sentence']
    st.write('**Required:** one appropriate sentence, at most 25 words, appended to the first body paragraph under “A Multifaceted Approach.”')
    st.write('Detected addition / paragraph:');st.text(s['text'] or '(No added text detected)')
    st.write(f'Word count: {s["word_count"] if s["word_count"] is not None else "Cannot isolate addition"} · {s["status"]}')
    st.caption(s.get('note','Teacher must assess context and sentence count. No separate points assigned.'))
    restore('sentence-'+key,'')
    st.text_area('Sentence review notes',key='sentence-'+key,on_change=remember,args=('sentence-'+key,))
    st.caption('Original paragraphs no longer matched exactly (includes legitimate citation edits and other required changes; not automatically errors):')
    for t in result['changed_original_paragraphs']: st.text(t)

restore('overall-'+key,'')
st.text_area('Overall feedback',key='overall-'+key,on_change=remember,args=('overall-'+key,))

def reviewed(k,r):
    rows=[]
    for i,f in enumerate(r['findings']):
        val=st.session_state['reviews'].get(f'score-{k}-{i}','Unreviewed')
        rows.append({**f,'score':None if val=='Unreviewed' else val,'feedback':st.session_state['reviews'].get(f'comment-{k}-{i}','')})
    complete=all(x['score'] is not None for x in rows)
    return {'filename':r['filename'],'complete':complete,'total':sum(x['score'] or 0 for x in rows),'maximum':60,'criteria':rows,'sentence':r['sentence'],'sentence_notes':st.session_state['reviews'].get('sentence-'+k,''),'overall_feedback':st.session_state['reviews'].get('overall-'+k,'')}

report=reviewed(key,result)
c1,c2=st.columns(2)
c1.metric('Reviewed points' if not report['complete'] else 'Final score',f'{report["total"]} / 60')
c2.metric('Criteria reviewed',f'{sum(x["score"] is not None for x in report["criteria"])} / 12')
lines=[f'Bears Word Project — {report["filename"]}',f'Status: {"Final" if report["complete"] else "DRAFT — incomplete review"}',f'Points: {report["total"]}/60','']
for x in report['criteria']:
    lines.extend([f'{x["criterion"]}: {x["score"] if x["score"] is not None else "Unreviewed"}/5',x['feedback'],''])
lines.extend(['Sentence review: '+report['sentence_notes'],'Overall feedback: '+report['overall_feedback']])
st.download_button('Download feedback (.txt)','\n'.join(lines),file_name=Path(result['filename']).stem+'_feedback.txt',mime='text/plain')
st.download_button('Download evidence and review (.json)',json.dumps(report,indent=2),file_name=Path(result['filename']).stem+'_review.json',mime='application/json')
output=io.StringIO();writer=csv.writer(output)
def safe(v):
    s=str(v)
    return "'"+s if s.lstrip().startswith(('=','+','-','@')) else s
writer.writerow(['Submission','Status',*CRITERIA,'Points','Maximum','Feedback'])
for k,r in valid:
    review=reviewed(k,r)
    writer.writerow([safe(r['filename']),'Final' if review['complete'] else 'Draft',*[x['score'] if x['score'] is not None else '' for x in review['criteria']],review['total'],60,safe(review['overall_feedback'])])
st.download_button('Download batch grade sheet (.csv)',output.getvalue(),file_name='bears_grades.csv',mime='text/csv')
st.caption('Download your reviews before closing this session. Refreshing or restarting can clear uploads, scores, and comments.')
