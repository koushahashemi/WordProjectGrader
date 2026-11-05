import pytest
from autograder import grade, grade_detailed, ORIGINAL_RUBRIC_BANDS, csv_report, detail_text
from feedback import feedback_text
from grader import BASE
from test_automatic import compliant, mutate, setprop


@pytest.mark.parametrize('detailed,band', [(5,5),(4,3),(3,3),(2,1),(1,1),(0,0)])
def test_exact_requested_mapping(detailed,band):
    assert ORIGINAL_RUBRIC_BANDS[detailed] == band


@pytest.mark.parametrize('calibrated', [True,False])
def test_all_rows_diagnostics_and_totals_use_bands(calibrated):
    data=BASE.read_bytes()
    original=grade_detailed(data,'base.docx',calibrated)
    result=grade(data,'base.docx',calibrated)
    for before,after in zip(original['criteria'],result['criteria']):
        assert after['score']==ORIGINAL_RUBRIC_BANDS[before['score']]
        assert after['strict_score']==ORIGINAL_RUBRIC_BANDS[before['strict_score']]
        assert after['detailed_score']==before['score']
        assert after['checks']==before['checks']
        assert after.get('scoring_groups')==before.get('scoring_groups')
    assert result['total']==sum(r['score'] for r in result['criteria'])==9
    assert result['strict_total']==sum(r['strict_score'] for r in result['criteria'])
    assert '5, 3, 1, 0' in result['scoring_mode']
    assert 'original rubric' in detail_text(result)
    assert 'Original rubric band' in feedback_text(result)
    assert ',9,60,' in csv_report([result])


def test_full_compliance_still_gets_sixty(compliant):
    result=grade(compliant,'full.docx')
    assert result['total']==60
    assert all(r['score']==5 for r in result['criteria'])


def test_one_typography_error_maps_four_to_three_without_changing_evidence(compliant):
    def change(roots):
        from autograder import NS
        style=roots['word/styles.xml'].xpath('./w:style[@w:styleId="TOC1"]',namespaces=NS)[0]
        setprop(style,'rPr/rFonts',ascii='Cambria',hAnsi='Cambria')
    result=grade(mutate(compliant,change),'partial.docx')
    row=result['criteria'][3]
    assert row['detailed_score']==4 and row['score']==3
    assert [g['passed'] for g in row['scoring_groups']]==[True,True,False,True]
    assert '4/5 detailed assessment → 3/5' in feedback_text(result)
