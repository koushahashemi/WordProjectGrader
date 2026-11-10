"""Hidden-instruction check: sentence added after 'A Multifaceted Approach'."""
from test_automatic import compliant, mutate, paragraph, NS
from autograder import grade, csv_report, detail_text, SENTENCE_BASE
from feedback import feedback_text


def _body_paragraph(roots):
    return next(p for p in roots['word/document.xml'].findall('.//w:p',NS)
                if ''.join(p.xpath('.//w:t/text()',namespaces=NS)).strip()==SENTENCE_BASE)


def _append(extra):
    def change(roots):
        t=_body_paragraph(roots).findall('.//w:t',NS)[-1]
        t.text=t.text+' '+extra
    return change


def test_unchanged_paragraph_is_not_flagged(compliant):
    s=grade(compliant,'clean.docx')['sentence']
    assert s['flag']=='No' and s['text']==''


def test_sentence_at_end_of_paragraph_is_flagged(compliant):
    added='Protecting habitat corridors also helps bears travel safely between feeding areas.'
    result=grade(mutate(compliant,_append(added)),'added.docx')
    s=result['sentence']
    assert s['flag']=='Yes' and s['text']==added and s['word_count']==11
    assert 'one sentence' in s['finding'] and 'Over' not in s['finding']
    assert 'HIDDEN-INSTRUCTION SENTENCE: Yes' in detail_text(result)
    assert 'Yes,11,'+added in csv_report([result])


def test_sentence_as_separate_paragraph_is_flagged(compliant):
    added='Community education reduces conflict between people and bears.'
    def change(roots):
        p=_body_paragraph(roots); p.addnext(paragraph(added))
    s=grade(mutate(compliant,change),'separate.docx')['sentence']
    assert s['flag']=='Yes' and s['text']==added and 'separate paragraph' in s['placement']


def test_long_sentence_reports_word_limit(compliant):
    added=' '.join(['word']*30)+'.'
    s=grade(mutate(compliant,_append(added)),'long.docx')['sentence']
    assert s['flag']=='Yes' and s['word_count']==30 and 'Over the 25-word limit' in s['finding']


def test_rewritten_paragraph_needs_manual_check(compliant):
    def change(roots):
        p=_body_paragraph(roots)
        ts=p.findall('.//w:t',NS); ts[0].text='Bears need many kinds of help.'
        for t in ts[1:]: t.text=''
    s=grade(mutate(compliant,change),'rewritten.docx')['sentence']
    assert s['flag']=='Check'


def test_flag_never_appears_in_student_feedback(compliant):
    result=grade(mutate(compliant,_append('Bears matter to everyone.')),'fb.docx')
    text=feedback_text(result).lower()
    assert 'sentence' not in text and 'hidden' not in text
