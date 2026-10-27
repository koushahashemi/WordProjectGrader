"""Native Windows desktop UI. No browser, server, account, or API key."""
import argparse
import gc
import time
from types import SimpleNamespace
import os
import subprocess
import json
from pathlib import Path
import queue
import sys
import threading
from tkinterdnd2 import TkinterDnD, DND_FILES, COPY, REFUSE_DROP
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter import font as tkfont
from tkinter.scrolledtext import ScrolledText
from autograder import VERSION, grade, grade_paths, csv_report, detail_text
from grader import BASE, CRITERIA
from feedback import feedback_text

from app_info import APP_TITLE, COURSE_TEXT, APP_DESCRIPTION


SHORT = ['Headings','Body','Title page','TOC','Captions','Figures list','Lists','Citations','References','Footers','Sections','Branding']


def launch_document(path):
    path=Path(path).resolve()
    if not path.is_file():raise FileNotFoundError('The original submission has been moved or deleted: '+str(path))
    if path.suffix.lower()!='.docx':raise ValueError('Only submitted Word documents can be opened.')
    if sys.platform=='win32':os.startfile(str(path))
    else:subprocess.Popen(['open' if sys.platform=='darwin' else 'xdg-open',str(path)],stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)


class GraderApp:
    def __init__(self, root):
        # Tk interpreters and variables must be reclaimed on the GUI thread.
        # Automatic cyclic GC could otherwise run inside the grading worker.
        gc.disable()
        self.root=root
        self.results=[]
        self.events=queue.Queue()
        self.busy=False
        self.pending_paths=[]
        self.closed=False
        root.title(f'{APP_TITLE} {VERSION} - BSYS 1000 - Fall 2026')
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
        ttk.Label(outer,text=APP_TITLE,style='Title.TLabel').pack(anchor='w')
        ttk.Label(outer,text=COURSE_TEXT,style='Sub.TLabel').pack(anchor='w',pady=(2,10))
        ttk.Label(outer,text=APP_DESCRIPTION,wraplength=1150).pack(anchor='w',pady=(0,8))
        tools=ttk.Frame(outer);tools.pack(fill='x')
        self.add=ttk.Button(tools,text='Add Word files...',command=self.choose_files);self.add.pack(side='left',padx=(0,6))
        self.folder=ttk.Button(tools,text='Add folder...',command=self.choose_folder);self.folder.pack(side='left',padx=6)
        self.clear=ttk.Button(tools,text='Clear batch',command=self.clear_batch);self.clear.pack(side='left',padx=6)
        ttk.Button(tools,text='Save feedback...',command=self.export_feedback).pack(side='right',padx=6)
        ttk.Button(tools,text='Export grades CSV...',command=self.export_csv).pack(side='right',padx=6)
        ttk.Button(tools,text='Save full report...',command=self.export_report).pack(side='right',padx=6)
        self.progress=ttk.Progressbar(outer,mode='determinate');self.progress.pack(fill='x',pady=(10,4))
        self.status=tk.StringVar(value='Drop Word files anywhere in this window, or use Add Word files. Double-click a filename to open its original.')
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
        self.batch.bind('<Double-1>',self.open_submission)
        self.batch.bind('<Return>',self.open_submission)
        review=ttk.Frame(panes,padding=(0,12,0,0));panes.add(review,weight=3)
        self.selected=tk.StringVar(value='Select a submission to see its rubric scores and deduction details.')
        ttk.Label(review,textvariable=self.selected,font=(family,11,'bold')).pack(anchor='w')
        ttk.Label(review,text='Feedback summarizes all lost marks. Select a criterion to inspect its technical details and scoring breakdown.',wraplength=1150).pack(anchor='w',pady=(4,8))
        split=ttk.Panedwindow(review,orient='horizontal');split.pack(fill='both',expand=True)
        rubric=ttk.Frame(split);split.add(rubric,weight=1)
        self.rubric=ttk.Treeview(rubric,columns=['criterion','score','failed','unknown'],show='headings',height=12,selectmode='browse')
        for col,label,width in [('criterion','Criterion',230),('score','/5',45),('failed','Lost items',70),('unknown','Unverified',75)]:
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
        self.drop_enabled=False
        if hasattr(root,'TkdndVersion'):self.enable_file_drop()
        root.bind('<Destroy>',self.on_window_destroy,add='+')
        self.poll_id=root.after(100,self.poll)

    def on_window_destroy(self,event):
        if event.widget is not self.root:return
        self.closed=True
        try:self.root.after_cancel(self.poll_id)
        except tk.TclError:pass
        # Dispose Tcl-backed variables on the GUI thread, never during worker GC.
        self.status=None;self.selected=None

    def enable_file_drop(self):
        # Register descendants too, so a file can be dropped over a table, text,
        # heading or button instead of only an otherwise empty root background.
        def register(widget):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind('<<Drop>>',self.on_file_drop)
            for child in widget.winfo_children():register(child)
        register(self.root);self.drop_enabled=True

    def on_file_drop(self,event):
        try:paths=self.root.tk.splitlist(event.data)
        except tk.TclError:
            self.status.set('The dropped file list could not be read. Use Add Word files instead.');return REFUSE_DROP
        valid=[str(Path(p).resolve()) for p in paths if Path(p).is_file() and Path(p).suffix.lower()=='.docx' and not Path(p).name.startswith('~$')]
        if not valid:
            self.status.set('Drop one or more .docx files. Folders, Word temporary files and other file types are ignored.');return REFUSE_DROP
        self.start(valid)
        if len(valid)<len(paths):self.status.set(f'Added or queued {len(valid)} Word file(s); ignored {len(paths)-len(valid)} unsupported item(s).')
        return COPY

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
        paths=list(paths)
        if not paths:return
        if self.busy:
            self.pending_paths.extend(paths)
            self.status.set(f'Queued {len(self.pending_paths)} more document(s) to grade after the current batch.')
            return
        gc.collect()  # GUI-thread collection before starting background work.
        self.busy=True
        for button in (self.add,self.folder,self.clear):button.configure(state='disabled')
        self.progress.configure(maximum=len(paths),value=0)
        self.status.set(f'Grading {len(paths)} document(s)...')
        events=self.events
        def worker():
            try:
                for index,path in enumerate(paths):
                    try:result=grade_paths([path])[0]
                    except Exception as exc:
                        # Keep the worker alive and report an engine error without a fabricated grade.
                        result={'filename':Path(path).name,'source_path':str(Path(path).resolve()),'error':f'Unexpected grading error: {type(exc).__name__}: {exc}'}
                    events.put(('result',result,index+1,len(paths)))
            finally:events.put(('done',))
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
        if self.closed:return
        try:
            while True:
                event=self.events.get_nowait()
                if event[0]=='result':
                    self.add_result(event[1]);self.progress.configure(value=event[2]);self.status.set(f'Graded {event[2]} of {event[3]}: {event[1]["filename"]}')
                else:
                    self.busy=False
                    gc.collect()
                    for button in (self.add,self.folder,self.clear):button.configure(state='normal')
                    errors=sum('error' in r for r in self.results)
                    self.status.set(f'Batch complete: {len(self.results)-errors} scored, {errors} error(s). Select a row to inspect deductions, or export the table.')
                    if self.pending_paths:
                        pending=self.pending_paths;self.pending_paths=[];self.start(pending)
        except queue.Empty:pass
        self.poll_id=self.root.after(100,self.poll)

    def open_submission(self,event=None):
        if event is not None and hasattr(event,'y') and getattr(event,'keysym','')!='Return':
            item=self.batch.identify_row(event.y)
            if not item or self.batch.identify_column(event.x)!='#1':return
            self.batch.selection_set(item)
        else:
            items=self.batch.selection()
            if not items:return
            item=items[0]
        path=self.results[int(item)].get('source_path')
        if not path:
            messagebox.showinfo('Original file unavailable','Add the original document again to open it from this list.')
            return
        try:launch_document(path)
        except (OSError,ValueError) as exc:messagebox.showerror('Cannot open submission',str(exc))
        return 'break'

    def select_submission(self,event=None):
        self.rubric.delete(*self.rubric.get_children())
        chosen=self.batch.selection()
        if not chosen:return
        result=self.results[int(chosen[0])]
        if 'error' in result:
            self.selected.set(result['filename']+' - unable to grade');self.show_text(result['error']);self.show_feedback(result);return
        self.selected.set(f'{result["filename"]} - {result["total"]}/60'+(' (estimated)' if result['estimated'] else ''))
        for i,row in enumerate(result['criteria']):
            self.rubric.insert('','end',iid=str(i),values=[row['criterion'],row['score'],sum(not g['passed'] for g in row['scoring_groups']) if row.get('scoring_groups') else len(row.get('deductions',[])),sum(c.get('scored',True) and c['passed'] is None for c in row['checks'])],tags=('deduction',) if row['score']<5 else ())
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
        lines=[f'{row["criterion"]}: {row["score"]}/5',f'Written-requirement diagnostic: {row.get("strict_score",row["score"])}/5',f'Evidence checks passed: {row["measured_passed"]}/{row["measured_total"]}',row['policy'],'']
        for group in row.get('scoring_groups',[]):
            summary=('Pass' if group['passed'] else 'Fail: one deduction') if row['criterion']==CRITERIA[9] else f"{group['earned']}/{group['weight']} marks"
            lines.append(group['name']+': '+summary)
        lines.append('')
        # Failed checks first, then limitations, then passing evidence.
        checks=sorted(row['checks'],key=lambda c:0 if c['passed'] is False else 1 if c['passed'] is None else 2)
        for c in checks:
            label='INFO' if not c.get('scored',True) else 'ACCEPTED' if c.get('accepted') else 'PASS' if c['passed'] else 'UNVERIFIED' if c['passed'] is None else 'FAIL'
            lines += [f'[{label}] {c["name"]}',c['location'],f'Expected: {c["expected"]}',f'Found: {c["found"]}',('Accepted alternative: '+c['accepted']) if c.get('accepted') else '', '']
        lines+=['ADDITIONAL SENTENCE (no separate rubric points)',json.dumps(result['sentence'],indent=2,ensure_ascii=False)]
        self.show_text('\n'.join(lines))
        self.show_feedback(result)

    def clear_batch(self):
        if self.busy:return
        self.pending_paths.clear();self.results.clear();self.batch.delete(*self.batch.get_children());self.rubric.delete(*self.rubric.get_children());self.show_text('');self.show_feedback();self.status.set('Batch cleared. Add documents to grade.');self.selected.set('Select a submission.');self.progress.configure(value=0)

    def export_feedback(self):
        if not self.results:messagebox.showinfo('No grades','Add documents first.');return
        path=filedialog.asksaveasfilename(defaultextension='.txt',initialfile='BCIT_BSYS1000_Fall2026_Feedback.txt',filetypes=[('Plain-language feedback','*.txt')])
        if path:
            try:Path(path).write_text('\n\n'.join(feedback_text(r) for r in self.results),encoding='utf-8')
            except OSError as exc:messagebox.showerror('Export failed',str(exc))

    def export_csv(self):
        if not self.results:messagebox.showinfo('No grades','Add documents first.');return
        path=filedialog.asksaveasfilename(defaultextension='.csv',initialfile='BCIT_BSYS1000_Fall2026_Grades.csv',filetypes=[('CSV grade table','*.csv')])
        if path:
            try:Path(path).write_text(csv_report(self.results),encoding='utf-8-sig')
            except OSError as exc:messagebox.showerror('Export failed',str(exc))

    def export_report(self):
        if not self.results:messagebox.showinfo('No grades','Add documents first.');return
        path=filedialog.asksaveasfilename(defaultextension='.txt',initialfile='BCIT_BSYS1000_Fall2026_Report.txt',filetypes=[('Detailed text report','*.txt'),('Machine-readable JSON','*.json')])
        if path:
            try:
                content=json.dumps(self.results,indent=2,ensure_ascii=False) if Path(path).suffix.lower()=='.json' else '\n\n'+'\n\n'.join(detail_text(r) for r in self.results)
                Path(path).write_text(content,encoding='utf-8')
            except OSError as exc:messagebox.showerror('Export failed',str(exc))


def main():
    parser=argparse.ArgumentParser(description=APP_TITLE+' - BSYS 1000 Fall 2026')
    parser.add_argument('--self-test',action='store_true')
    parser.add_argument('--output',type=Path)
    args=parser.parse_args()
    gc_was_enabled=gc.isenabled()
    if args.self_test:
        root=TkinterDnD.Tk();root.withdraw();app=GraderApp(root)
        assert app.drop_enabled
        data=root.tk.call('format','%s',root.tk.call('list',str(BASE)))
        assert app.on_file_drop(SimpleNamespace(data=data))==COPY
        deadline=time.monotonic()+90
        while app.busy and time.monotonic()<deadline:
            root.update();time.sleep(.02)
        assert not app.busy and len(app.results)==1
        result=app.results[0]
        assert 'error' not in result and len(result['criteria'])==12 and result['maximum']==60
        app.select_submission();root.update();root.destroy();app=None;root=None;gc.collect()
        if gc_was_enabled:gc.enable()
        if args.output:args.output.write_text(json.dumps({'ok':True,'total':result['total'],'criteria':12,'drag_drop':True}),encoding='utf-8')
        return
    root=TkinterDnD.Tk();app=GraderApp(root)
    try:root.mainloop()
    finally:
        app=None;root=None;gc.collect()
        if gc_was_enabled:gc.enable()


if __name__=='__main__':main()
