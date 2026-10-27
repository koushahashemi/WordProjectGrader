"""Native Windows desktop UI. No browser, server, account, or API key."""
import argparse
import json
from pathlib import Path
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from tkinter.scrolledtext import ScrolledText
from autograder import VERSION, grade, grade_paths, csv_report, detail_text
from grader import BASE, CRITERIA
from feedback import feedback_text

SHORT = ['Headings','Body','Title page','TOC','Captions','Figures list','Lists','Citations','References','Footers','Sections','Branding']


class GraderApp:
    def __init__(self, root):
        self.root=root
        self.results=[]
        self.events=queue.Queue()
        self.busy=False
        root.title(f'Bears Assignment Grader {VERSION}')
        root.geometry('1280x840'); root.minsize(900,650)
        root.configure(background='#FAF3E0')
        family='Segoe UI' if 'Segoe UI' in tkfont.families(root) else 'DejaVu Sans'
        for name in ('TkDefaultFont','TkTextFont','TkMenuFont','TkHeadingFont'):
            tkfont.nametofont(name).configure(family=family,size=10)
        style=ttk.Style(root)
        if 'vista' in style.theme_names(): style.theme_use('vista')
        style.configure('Treeview',rowheight=27)
        style.configure('.',font=(family,10))
        style.configure('Title.TLabel',font=(family,21,'bold'),foreground='#00796B')
        style.configure('Sub.TLabel',font=(family,10))
        outer=ttk.Frame(root,padding=16); outer.pack(fill='both',expand=True)
        ttk.Label(outer,text='Bears Assignment Grader',style='Title.TLabel').pack(anchor='w')
        ttk.Label(outer,text='Teacher-calibrated automatic scoring | Word Project 2026 | 12 criteria / 60 points',style='Sub.TLabel').pack(anchor='w',pady=(2,10))
        tools=ttk.Frame(outer);tools.pack(fill='x')
        self.add=ttk.Button(tools,text='Add Word files...',command=self.choose_files);self.add.pack(side='left',padx=(0,6))
        self.folder=ttk.Button(tools,text='Add folder...',command=self.choose_folder);self.folder.pack(side='left',padx=6)
        self.clear=ttk.Button(tools,text='Clear batch',command=self.clear_batch);self.clear.pack(side='left',padx=6)
        ttk.Button(tools,text='Save feedback...',command=self.export_feedback).pack(side='right',padx=6)
        ttk.Button(tools,text='Export grades CSV...',command=self.export_csv).pack(side='right',padx=6)
        ttk.Button(tools,text='Save full report...',command=self.export_report).pack(side='right',padx=6)
        self.progress=ttk.Progressbar(outer,mode='determinate');self.progress.pack(fill='x',pady=(10,4))
        self.status=tk.StringVar(value='Select multiple .docx files or a folder. Grading runs automatically; original documents are never changed.')
        ttk.Label(outer,textvariable=self.status,wraplength=1150).pack(anchor='w',pady=(0,8))
        panes=ttk.Panedwindow(outer,orient='vertical');panes.pack(fill='both',expand=True)
        batch=ttk.Frame(panes);panes.add(batch,weight=2)
        columns=['filename','total','status',*SHORT]
        self.batch=ttk.Treeview(batch,columns=columns,show='headings',selectmode='browse',height=7)
        for col in columns:
            self.batch.heading(col,text='Submission' if col=='filename' else 'Total /60' if col=='total' else col.title())
            self.batch.column(col,width=230 if col=='filename' else 105 if col=='status' else 82,minwidth=65,stretch=False,anchor='w' if col in ('filename','status') else 'center')
        self.batch.tag_configure('error',foreground='#B3261E')
        self.batch.grid(row=0,column=0,sticky='nsew')
        vs=ttk.Scrollbar(batch,orient='vertical',command=self.batch.yview);vs.grid(row=0,column=1,sticky='ns')
        self.batch.configure(yscrollcommand=vs.set)
        hs=ttk.Scrollbar(batch,orient='horizontal',command=self.batch.xview);hs.grid(row=1,column=0,sticky='ew')
        self.batch.configure(xscrollcommand=hs.set)
        batch.columnconfigure(0,weight=1);batch.rowconfigure(0,weight=1)
        self.batch.bind('<<TreeviewSelect>>',self.select_submission)
        review=ttk.Frame(panes,padding=(0,12,0,0));panes.add(review,weight=3)
        self.selected=tk.StringVar(value='Select a submission to see its rubric scores and deduction details.')
        ttk.Label(review,textvariable=self.selected,font=(family,11,'bold')).pack(anchor='w')
        ttk.Label(review,text='Feedback summarizes all lost marks. Select a criterion to inspect its technical details and scoring breakdown.',wraplength=1150).pack(anchor='w',pady=(4,8))
        split=ttk.Panedwindow(review,orient='horizontal');split.pack(fill='both',expand=True)
        rubric=ttk.Frame(split);split.add(rubric,weight=1)
        self.rubric=ttk.Treeview(rubric,columns=['criterion','score','failed','unknown'],show='headings',height=12,selectmode='browse')
        for col,label,width in [('criterion','Criterion',230),('score','/5',45),('failed','Failed',50),('unknown','Unverified',75)]:
            self.rubric.heading(col,text=label);self.rubric.column(col,width=width,minwidth=40,anchor='w' if col=='criterion' else 'center')
        self.rubric.pack(fill='both',expand=True)
        self.rubric.tag_configure('deduction',foreground='#A12B1F')
        self.rubric.bind('<<TreeviewSelect>>',self.select_criterion)
        self.views=ttk.Notebook(split)
        self.feedback=ScrolledText(self.views,wrap='word',font=(family,10),padx=12,pady=10,state='disabled')
        self.details=ScrolledText(self.views,wrap='word',font=(family,10),padx=12,pady=10,state='disabled')
        self.views.add(self.feedback,text='Feedback on lost marks')
        self.views.add(self.details,text='Technical details')
        split.add(self.views,weight=2)
        ttk.Label(outer,text='Teacher-calibrated scores. ACCEPTED notes explain benchmark alternatives. Estimated = XML cannot verify every layout requirement.',wraplength=1150).pack(anchor='w',pady=(10,0))
        root.after(100,self.poll)

    def show_text(self,content):
        self.details.configure(state='normal');self.details.delete('1.0','end');self.details.insert('1.0',content);self.details.configure(state='disabled')

    def show_feedback(self,result=None,criterion=None):
        self.feedback.configure(state='normal');self.feedback.delete('1.0','end')
        if result:self.feedback.insert('1.0',feedback_text(result,criterion))
        self.feedback.configure(state='disabled')

    def choose_files(self):
        paths=filedialog.askopenfilenames(title='Select student submissions',filetypes=[('Word documents','*.docx')])
        if paths:self.start(paths)

    def choose_folder(self):
        folder=filedialog.askdirectory(title='Select folder of student submissions')
        if folder:
            paths=sorted(p for p in Path(folder).iterdir() if p.is_file() and p.suffix.lower()=='.docx' and not p.name.startswith('~$'))
            if paths:self.start(paths)
            else:messagebox.showinfo('No submissions','No .docx files were found directly inside this folder.')

    def start(self,paths):
        if self.busy:return
        self.busy=True
        for button in (self.add,self.folder,self.clear):button.configure(state='disabled')
        self.progress.configure(maximum=len(paths),value=0)
        self.status.set(f'Grading {len(paths)} document(s)...')
        def worker():
            try:
                for index,path in enumerate(paths):
                    try:result=grade_paths([path])[0]
                    except Exception as exc:
                        # Keep the worker alive and report an engine error without a fabricated grade.
                        result={'filename':Path(path).name,'error':f'Unexpected grading error: {type(exc).__name__}: {exc}'}
                    self.events.put(('result',result,index+1,len(paths)))
            finally:self.events.put(('done',))
        threading.Thread(target=worker,daemon=True).start()

    def add_result(self,result):
        idx=str(len(self.results));self.results.append(result)
        if 'error' in result:
            values=[result['filename'],'-','Error',*['-']*12];tags=('error',)
        else:
            values=[result['filename'],result['total'],'Estimated' if result['estimated'] else 'Scored',*[r['score'] for r in result['criteria']]];tags=()
        self.batch.insert('','end',iid=idx,values=values,tags=tags)
        if not self.batch.selection():self.batch.selection_set(idx)

    def poll(self):
        try:
            while True:
                event=self.events.get_nowait()
                if event[0]=='result':
                    self.add_result(event[1]);self.progress.configure(value=event[2]);self.status.set(f'Graded {event[2]} of {event[3]}: {event[1]["filename"]}')
                else:
                    self.busy=False
                    for button in (self.add,self.folder,self.clear):button.configure(state='normal')
                    errors=sum('error' in r for r in self.results)
                    self.status.set(f'Batch complete: {len(self.results)-errors} scored, {errors} error(s). Select a row to inspect deductions, or export the table.')
        except queue.Empty:pass
        self.root.after(100,self.poll)

    def select_submission(self,event=None):
        self.rubric.delete(*self.rubric.get_children())
        chosen=self.batch.selection()
        if not chosen:return
        result=self.results[int(chosen[0])]
        if 'error' in result:
            self.selected.set(result['filename']+' - unable to grade');self.show_text(result['error']);self.show_feedback(result);return
        self.selected.set(f'{result["filename"]} - {result["total"]}/60'+(' (estimated)' if result['estimated'] else ''))
        for i,row in enumerate(result['criteria']):
            self.rubric.insert('','end',iid=str(i),values=[row['criterion'],row['score'],sum(c['passed'] is False and not c.get('accepted') for c in row['checks']),sum(c['passed'] is None for c in row['checks'])],tags=('deduction',) if row['score']<5 else ())
        self.show_text(detail_text(result))
        self.show_feedback(result)
        failed=next((i for i,r in enumerate(result['criteria']) if r['score']<5),0)
        self.rubric.selection_set(str(failed))

    def select_criterion(self,event=None):
        selected=self.batch.selection();chosen=self.rubric.selection()
        if not selected or not chosen:return
        result=self.results[int(selected[0])]
        if 'error' in result:return
        row=result['criteria'][int(chosen[0])]
        lines=[f'{row["criterion"]}: {row["score"]}/5',f'Written-requirement diagnostic: {row.get("strict_score",row["score"])}/5',f'Measured checks passed: {row["measured_passed"]}/{row["measured_total"]}',row['policy'],'']
        for group in row.get('scoring_groups',[]):
            summary=('Pass' if group['passed'] else 'Fail: one deduction') if row['criterion']==CRITERIA[9] else f"{group['earned']}/{group['weight']} marks"
            lines.append(group['name']+': '+summary)
        lines.append('')
        # Failed checks first, then limitations, then passing evidence.
        checks=sorted(row['checks'],key=lambda c:0 if c['passed'] is False else 1 if c['passed'] is None else 2)
        for c in checks:
            label='ACCEPTED' if c.get('accepted') else 'PASS' if c['passed'] else 'UNVERIFIED' if c['passed'] is None else 'FAIL'
            lines += [f'[{label}] {c["name"]}',c['location'],f'Expected: {c["expected"]}',f'Found: {c["found"]}',('Accepted alternative: '+c['accepted']) if c.get('accepted') else '', '']
        lines+=['ADDITIONAL SENTENCE (no separate rubric points)',json.dumps(result['sentence'],indent=2,ensure_ascii=False)]
        self.show_text('\n'.join(lines))
        self.show_feedback(result)

    def clear_batch(self):
        if self.busy:return
        self.results.clear();self.batch.delete(*self.batch.get_children());self.rubric.delete(*self.rubric.get_children());self.show_text('');self.show_feedback();self.status.set('Batch cleared. Add documents to grade.');self.selected.set('Select a submission.');self.progress.configure(value=0)

    def export_feedback(self):
        if not self.results:messagebox.showinfo('No grades','Add documents first.');return
        path=filedialog.asksaveasfilename(defaultextension='.txt',initialfile='Bears_Feedback.txt',filetypes=[('Plain-language feedback','*.txt')])
        if path:
            try:Path(path).write_text('\n\n'.join(feedback_text(r) for r in self.results),encoding='utf-8')
            except OSError as exc:messagebox.showerror('Export failed',str(exc))

    def export_csv(self):
        if not self.results:messagebox.showinfo('No grades','Add documents first.');return
        path=filedialog.asksaveasfilename(defaultextension='.csv',initialfile='Bears_Grades.csv',filetypes=[('CSV grade table','*.csv')])
        if path:
            try:Path(path).write_text(csv_report(self.results),encoding='utf-8-sig')
            except OSError as exc:messagebox.showerror('Export failed',str(exc))

    def export_report(self):
        if not self.results:messagebox.showinfo('No grades','Add documents first.');return
        path=filedialog.asksaveasfilename(defaultextension='.txt',initialfile='Bears_Grading_Report.txt',filetypes=[('Detailed text report','*.txt'),('Machine-readable JSON','*.json')])
        if path:
            try:
                content=json.dumps(self.results,indent=2,ensure_ascii=False) if Path(path).suffix.lower()=='.json' else '\n\n'+'\n\n'.join(detail_text(r) for r in self.results)
                Path(path).write_text(content,encoding='utf-8')
            except OSError as exc:messagebox.showerror('Export failed',str(exc))


def main():
    parser=argparse.ArgumentParser(description='Bears Assignment Grader')
    parser.add_argument('--self-test',action='store_true')
    parser.add_argument('--output',type=Path)
    args=parser.parse_args()
    if args.self_test:
        result=grade(BASE.read_bytes(),'Bears.docx')
        assert len(result['criteria'])==12 and result['maximum']==60
        root=tk.Tk();root.withdraw();app=GraderApp(root);app.add_result(result);root.update();app.select_submission();root.update();root.destroy()
        if args.output:args.output.write_text(json.dumps({'ok':True,'total':result['total'],'criteria':12}),encoding='utf-8')
        return
    root=tk.Tk();GraderApp(root);root.mainloop()


if __name__=='__main__':main()
