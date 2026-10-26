import csv
import io
import json
import os
from pathlib import Path
import sys
import time
import tkinter as tk
from unittest.mock import patch
import pytest
from autograder import grade
from desktop import GraderApp
from grader import BASE

pytestmark=pytest.mark.skipif(sys.platform!='win32' and not os.environ.get('DISPLAY'),reason='Desktop test requires a graphical display')


def test_native_batch_scores_deductions_and_exports(tmp_path):
    root=tk.Tk();root.withdraw()
    try:
        app=GraderApp(root)
        bad=tmp_path/'invalid.docx';bad.write_bytes(b'not Word')
        app.start([BASE,bad,BASE])
        deadline=time.monotonic()+20
        while app.busy and time.monotonic()<deadline:
            root.update();time.sleep(.02)
        assert not app.busy
        assert len(app.batch.get_children())==3
        assert app.results[0]['total']==4
        assert 'error' in app.results[1]
        assert app.results[2]['total']==4
        app.batch.selection_set('0');root.update();app.select_submission();root.update()
        assert len(app.rubric.get_children())==12
        # TOC is missing: click the criterion and see concrete automatic deductions.
        app.rubric.selection_set('3');root.update();app.select_criterion()
        details=app.details.get('1.0','end')
        assert '0/5' in details and '[FAIL]' in details and 'Missing' in details
        for name,method in [('grades.csv',app.export_csv),('report.txt',app.export_report),('report.json',app.export_report)]:
            target=tmp_path/name
            with patch('desktop.filedialog.asksaveasfilename',return_value=str(target)):method()
            assert target.exists()
        rows=list(csv.reader(io.StringIO((tmp_path/'grades.csv').read_text(encoding='utf-8-sig'))))
        assert len(rows)==4 and rows[1][-3]=='4'
        assert len(json.loads((tmp_path/'report.json').read_text()))==3
        app.batch.selection_set('1');root.update();app.select_submission();assert 'unable to grade' in app.selected.get()
        app.clear_batch();assert not app.results and not app.batch.get_children()
    finally:root.destroy()
