"""Desktop window (tkinter)."""
from __future__ import annotations

import os
import queue
import subprocess
import sys
import time
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from . import __version__
from .content import PAPERS
from .content.drills import DRILLS
from .core import APP_LABEL, grade_task, grade_tasks, all_file_names
from .monitor import APP_NAME, Monitor
from .report import build as build_report
from .session import (default_root, fmt_secs, load_history, load_settings, make_record, new_attempt_folder,
                      prepare_folder, save_attempt, save_settings, weak_topics)
from .shortcuts import make_quiz

APP_TITLE = "ISRO Skill Test Practice"
SECTION_MIN = {"word": 26, "excel": 26, "ppt": 23}
OK, BAD = "#1a7f37", "#c62828"


def open_path(p):
    p = str(p)
    try:
        if sys.platform == "win32":
            os.startfile(p)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", p])
        else:
            subprocess.Popen(["xdg-open", p])
    except Exception as e:
        messagebox.showerror(APP_TITLE, f"Could not open {p}\n{e}")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_TITLE}  v{__version__}")
        self.geometry("1000x720")
        self.minsize(760, 560)
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        elif "clam" in style.theme_names():
            style.theme_use("clam")
        style.configure("H1.TLabel", font=("Segoe UI", 18, "bold"))
        style.configure("H2.TLabel", font=("Segoe UI", 12, "bold"))
        style.configure("Muted.TLabel", foreground="#5b6475")
        style.configure("Timer.TLabel", font=("Consolas", 22, "bold"))
        style.configure("Big.TButton", font=("Segoe UI", 11, "bold"), padding=6)
        self.frame = None
        self.show_home()

    def switch(self, frame):
        if self.frame is not None:
            self.frame.destroy()
        self.frame = frame
        frame.pack(fill="both", expand=True)

    def show_home(self):
        self.attributes("-topmost", False)
        self.switch(Home(self))


# =================================================================== home

class Home(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app, padding=18)
        self.app = app
        ttk.Label(self, text=APP_TITLE, style="H1.TLabel").pack(anchor="w")
        ttk.Label(self, text="Skill Test (on Computer) for Assistant / UDC / JPA / Stenographer – "
                             "MS Word 35 + MS Excel 35 + MS PowerPoint 30 marks, 75 minutes",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 12))

        grid = ttk.Frame(self)
        grid.pack(fill="both", expand=True)
        grid.columnconfigure(0, weight=1)
        grid.columnconfigure(1, weight=1)
        names = [p.title for p in PAPERS]

        # full mock
        f = ttk.LabelFrame(grid, text="  Full Mock Test  ", padding=12)
        f.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Label(f, text="All three sections, 75 minutes, marked out of 100.\nWork is checked after you submit.",
                  justify="left").pack(anchor="w")
        self.mock = ttk.Combobox(f, values=names, state="readonly", width=22)
        self.mock.current(0)
        self.mock.pack(anchor="w", pady=8)
        ttk.Button(f, text="Start Full Mock (75 min)", style="Big.TButton", command=self.start_mock).pack(anchor="w")

        # section practice
        f = ttk.LabelFrame(grid, text="  Section Practice  ", padding=12)
        f.grid(row=0, column=1, sticky="nsew", padx=6, pady=6)
        ttk.Label(f, text="One section of a mock with a shorter timer.\nYou can check each task as you go.",
                  justify="left").pack(anchor="w")
        row = ttk.Frame(f)
        row.pack(anchor="w", pady=8)
        self.sec_paper = ttk.Combobox(row, values=names, state="readonly", width=14)
        self.sec_paper.current(0)
        self.sec_paper.pack(side="left")
        self.sec_app = ttk.Combobox(row, values=["MS Word", "MS Excel", "MS PowerPoint"], state="readonly", width=14)
        self.sec_app.current(0)
        self.sec_app.pack(side="left", padx=6)
        ttk.Button(f, text="Start Section", style="Big.TButton", command=self.start_section).pack(anchor="w")

        # drills
        f = ttk.LabelFrame(grid, text="  Topic Drills  ", padding=12)
        f.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Label(f, text=f"{len(DRILLS)} short tasks, one syllabus topic each\n(Mail Merge, VLOOKUP, Subtotal, Slide Master…).",
                  justify="left").pack(anchor="w")
        ttk.Button(f, text="Choose Drills", style="Big.TButton", command=lambda: app.switch(DrillPicker(app))).pack(anchor="w", pady=8)

        # quiz
        f = ttk.LabelFrame(grid, text="  Shortcut Key Quiz  ", padding=12)
        f.grid(row=1, column=1, sticky="nsew", padx=6, pady=6)
        ttk.Label(f, text="20 quick questions on keyboard shortcuts.", justify="left").pack(anchor="w")
        self.quiz_app = ttk.Combobox(f, values=["All programs", "Word", "Excel", "PowerPoint"], state="readonly", width=16)
        self.quiz_app.current(0)
        self.quiz_app.pack(anchor="w", pady=8)
        ttk.Button(f, text="Start Quiz", style="Big.TButton", command=self.start_quiz).pack(anchor="w")

        # bottom bar
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=(12, 0))
        ttk.Button(bar, text="My Progress / History", command=lambda: app.switch(HistoryView(app))).pack(side="left")
        ttk.Button(bar, text="How it works", command=self.help).pack(side="left", padx=6)
        self.root_lbl = ttk.Label(bar, text=f"Test files are kept in: {default_root()}", style="Muted.TLabel")
        self.root_lbl.pack(side="left", padx=12)
        ttk.Button(bar, text="Change…", command=self.change_root).pack(side="left")

    def start_mock(self):
        p = PAPERS[self.mock.current()]
        self.app.switch(TestView(self.app, p.tasks, f"{p.title} – Full Test", p.minutes * 60, mode="Full mock",
                                 allow_check=False))

    def start_section(self):
        p = PAPERS[self.sec_paper.current()]
        app = ["word", "excel", "ppt"][self.sec_app.current()]
        tasks = p.section(app)
        self.app.switch(TestView(self.app, tasks, f"{p.title} – {APP_LABEL[app]}", SECTION_MIN[app] * 60,
                                 mode="Section practice", allow_check=True))

    def start_quiz(self):
        app = [None, "Word", "Excel", "PowerPoint"][self.quiz_app.current()]
        self.app.switch(QuizView(self.app, make_quiz(20, app)))

    def change_root(self):
        d = filedialog.askdirectory(title="Choose where test folders are created", initialdir=str(default_root()))
        if d:
            s = load_settings()
            s["root"] = d
            save_settings(s)
            self.root_lbl.config(text=f"Test files are kept in: {d}")

    def help(self):
        messagebox.showinfo(APP_TITLE, HELP_TEXT)


HELP_TEXT = """How it works

1. Pick a test. A new test folder is created and opened for you. Starter files (if any) and a Pictures folder are put in it.
2. Read each task in this window and do it in the real MS Word / Excel / PowerPoint.
3. Save every file in the test folder with the EXACT name asked (use .docx / .xlsx / .pptx unless told otherwise).
4. While you work, the app records what happens: which program is open, files created and saved, formulas typed, slides added, and how long you spend in each program.
5. Press Submit (or wait for the timer). Save and close your files first.
6. The app opens your files and checks each point. For every mistake it shows what it found, what was expected and how to do it, plus a step-by-step model answer.

Tip: keep this window on the right side of the screen with 'Keep on top' ticked, and Office on the left."""


# =================================================================== drill picker

class DrillPicker(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app, padding=14)
        self.app = app
        ttk.Label(self, text="Topic Drills", style="H1.TLabel").pack(anchor="w")
        ttk.Label(self, text="Select one or more drills (Ctrl+click), then Start. You can check each one instantly.",
                  style="Muted.TLabel").pack(anchor="w", pady=(0, 8))
        cols = ("topic", "mins")
        self.tree = ttk.Treeview(self, columns=cols, show="tree headings", selectmode="extended")
        self.tree.heading("#0", text="Drill")
        self.tree.heading("topic", text="Syllabus topic")
        self.tree.heading("mins", text="Minutes")
        self.tree.column("#0", width=380)
        self.tree.column("topic", width=300)
        self.tree.column("mins", width=70, anchor="center")
        self.items = {}
        for app_key in ("word", "excel", "ppt"):
            parent = self.tree.insert("", "end", text=APP_LABEL[app_key], open=True)
            for d in DRILLS:
                if d.app == app_key:
                    iid = self.tree.insert(parent, "end", text=f"{d.id}  {d.title}", values=(d.topic, d.minutes))
                    self.items[iid] = d
        self.tree.pack(fill="both", expand=True)
        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=8)
        ttk.Button(bar, text="← Home", command=app.show_home).pack(side="left")
        ttk.Button(bar, text="Start selected drills", style="Big.TButton", command=self.start).pack(side="right")

    def start(self):
        chosen = [self.items[i] for i in self.tree.selection() if i in self.items]
        if not chosen:
            messagebox.showinfo(APP_TITLE, "Select at least one drill first.")
            return
        mins = sum(d.minutes for d in chosen)
        self.app.switch(TestView(self.app, chosen, "Topic Drills", mins * 60, mode="Topic drills", allow_check=True,
                                 show_hints=True))


# =================================================================== test view

class TestView(ttk.Frame):
    def __init__(self, app: App, tasks, title, seconds, mode, allow_check, show_hints=False):
        super().__init__(app, padding=10)
        self.app = app
        self.tasks = tasks
        self.title_ = title
        self.mode = mode
        self.allowed = seconds
        self.allow_check = allow_check
        self.show_hints = show_hints
        self.started = time.monotonic()
        self.warned = set()
        self.done = False
        self.events_q = queue.Queue()

        self.folder = new_attempt_folder(title)
        prepare_folder(self.folder, tasks)
        self.known = all_file_names(tasks)
        self.monitor = Monitor(self.folder, on_event=self.events_q.put)

        # top bar
        top = ttk.Frame(self)
        top.pack(fill="x")
        ttk.Label(top, text=title, style="H2.TLabel").pack(side="left")
        ttk.Button(top, text="Submit test", style="Big.TButton", command=self.submit).pack(side="right")
        self.timer = ttk.Label(top, text=fmt_secs(seconds), style="Timer.TLabel")
        self.timer.pack(side="right", padx=16)
        self.topmost = tk.BooleanVar(value=False)
        ttk.Checkbutton(top, text="Keep on top", variable=self.topmost,
                        command=lambda: app.attributes("-topmost", self.topmost.get())).pack(side="right")

        info = ttk.Frame(self)
        info.pack(fill="x", pady=(4, 6))
        ttk.Label(info, text=f"Test folder: {self.folder}", style="Muted.TLabel").pack(side="left")
        ttk.Button(info, text="Open test folder", command=lambda: open_path(self.folder)).pack(side="right")

        body = ttk.PanedWindow(self, orient="horizontal")
        body.pack(fill="both", expand=True)
        left = ttk.Frame(body)
        self.tlist = tk.Listbox(left, width=34, activestyle="none", font=("Segoe UI", 10), exportselection=False)
        for t in tasks:
            self.tlist.insert("end", f"{APP_LABEL[t.app].replace('MS ', '')} · {t.title}  [{t.marks:g}]")
        self.tlist.pack(fill="both", expand=True)
        self.tlist.bind("<<ListboxSelect>>", lambda e: self.show_task())
        self.sec_lbl = ttk.Label(left, text="", justify="left")
        self.sec_lbl.pack(anchor="w", pady=4)
        body.add(left, weight=1)

        right = ttk.Frame(body)
        self.text = ScrolledText(right, wrap="word", font=("Segoe UI", 11), padx=10, pady=8, height=18)
        self.text.tag_configure("h", font=("Segoe UI", 13, "bold"))
        self.text.tag_configure("file", font=("Segoe UI", 11, "bold"), foreground="#0b4f9c")
        self.text.tag_configure("quote", font=("Consolas", 10), foreground="#333", lmargin1=24, lmargin2=24)
        self.text.tag_configure("hint", foreground="#7a5b00")
        self.text.pack(fill="both", expand=True)
        btns = ttk.Frame(right)
        btns.pack(fill="x", pady=4)
        self.b_starter = ttk.Button(btns, text="Open starter file", command=self.open_starter)
        self.b_starter.pack(side="left")
        if allow_check:
            ttk.Button(btns, text="Check this task now", command=self.check_one).pack(side="left", padx=6)
        if show_hints:
            ttk.Button(btns, text="Show model answer", command=self.hint).pack(side="left")
        ttk.Button(btns, text="Next →", command=lambda: self.move(1)).pack(side="right")
        ttk.Button(btns, text="← Previous", command=lambda: self.move(-1)).pack(side="right", padx=4)
        body.add(right, weight=3)

        logf = ttk.LabelFrame(self, text="  Live activity (tracked automatically)  ", padding=4)
        logf.pack(fill="x", pady=(6, 0))
        self.log = tk.Listbox(logf, height=6, font=("Consolas", 9), activestyle="none")
        self.log.pack(fill="x")
        self.times = ttk.Label(logf, text="", style="Muted.TLabel")
        self.times.pack(anchor="w")

        self.tlist.selection_set(0)
        self.show_task()
        self.monitor.start()
        open_path(self.folder)
        self.tick()

    # ---------- task display
    def current(self):
        sel = self.tlist.curselection()
        return self.tasks[sel[0]] if sel else self.tasks[0]

    def move(self, d):
        i = (self.tlist.curselection() or (0,))[0] + d
        if 0 <= i < len(self.tasks):
            self.tlist.selection_clear(0, "end")
            self.tlist.selection_set(i)
            self.tlist.see(i)
            self.show_task()

    def show_task(self):
        t = self.current()
        tx = self.text
        tx.config(state="normal")
        tx.delete("1.0", "end")
        tx.insert("end", f"{APP_LABEL[t.app]} – {t.title}   ({t.marks:g} marks)\n", "h")
        tx.insert("end", f"Save your work as: {t.file}\n", "file")
        if t.starters:
            tx.insert("end", "Starter file(s) in the test folder: " + ", ".join(t.starters) + "\n", "file")
        tx.insert("end", "\n")
        n = 0
        for line in t.instructions:
            if line.startswith("    ") or line.startswith("   "):
                tx.insert("end", line.strip() + "\n", "quote")
            else:
                n += 1
                tx.insert("end", f"{n}. {line}\n")
        tx.insert("end", "\nSave inside the test folder. Pictures are in the 'Pictures' sub-folder.\n", "hint")
        tx.config(state="disabled")
        self.b_starter.state(["!disabled"] if t.starters else ["disabled"])

    def open_starter(self):
        t = self.current()
        for name in t.starters:
            open_path(self.folder / name)

    def hint(self):
        t = self.current()
        messagebox.showinfo("Model answer", f"{t.title}\n\n" + "\n".join(f"{i}. {s}" for i, s in enumerate(t.solution, 1)))

    def check_one(self):
        t = self.current()
        res = grade_task(t, self.folder, self.known)
        ResultPopup(self.app, res)

    # ---------- timer & log
    def tick(self):
        if self.done:
            return
        while True:
            try:
                ev = self.events_q.get_nowait()
            except queue.Empty:
                break
            self.log.insert("end", f"{ev['time']}  {APP_NAME.get(ev['app'], ev['app']):<10} {ev['msg']}")
            self.log.see("end")
            if self.log.size() > 300:
                self.log.delete(0)
        used = time.monotonic() - self.started
        left = max(0, self.allowed - used)
        self.timer.config(text=fmt_secs(left), foreground=BAD if left <= 300 else "")
        s = self.monitor.app_seconds
        self.times.config(text="Time in programs –  Word " + fmt_secs(s["word"]) + "   Excel " + fmt_secs(s["excel"]) +
                          "   PowerPoint " + fmt_secs(s["ppt"]) +
                          ("" if self.monitor.com_ok or sys.platform != "win32" else
                           "   (detailed Office tracking unavailable)"))
        for mins in (10, 5, 1):
            if left <= mins * 60 and mins not in self.warned and self.allowed > mins * 60 + 30:
                self.warned.add(mins)
                self.app.bell()
                self.monitor.log("test", f"{mins} minute(s) left")
                self.flash(f"Only {mins} minute{'s' if mins > 1 else ''} left – save your files!")
        if left <= 0:
            self.monitor.log("test", "Time over – submitted automatically")
            self.finish(auto=True)
            return
        self.after(1000, self.tick)

    def flash(self, msg):
        w = tk.Toplevel(self)
        w.title("Time warning")
        w.attributes("-topmost", True)
        ttk.Label(w, text=msg, font=("Segoe UI", 14, "bold"), foreground=BAD, padding=20).pack()
        ttk.Button(w, text="OK", command=w.destroy).pack(pady=(0, 12))
        w.after(8000, lambda: w.winfo_exists() and w.destroy())

    # ---------- submit
    def submit(self):
        unsaved = [u.split(":", 1)[1] for u in self.monitor.unsaved]
        if unsaved:
            if not messagebox.askyesno(APP_TITLE, "These files still have unsaved changes in Office:\n\n  " +
                                       "\n  ".join(unsaved) + "\n\nUnsaved work will NOT be checked.\n"
                                       "Submit anyway? (No = go back and save)"):
                return
        elif not messagebox.askyesno(APP_TITLE, "Submit the test now?\n\nMake sure all files are saved in the test folder."):
            return
        self.finish()

    def finish(self, auto=False):
        if self.done:
            return
        self.done = True
        self.monitor.stop()
        used = min(time.monotonic() - self.started, self.allowed)
        wait = ttk.Label(self, text="Checking your files…", style="H2.TLabel")
        wait.place(relx=0.5, rely=0.5, anchor="center")
        self.update_idletasks()
        results = grade_tasks(self.tasks, self.folder)
        rec = make_record(self.mode, self.title_, self.folder, results, used, self.allowed,
                          self.monitor.app_seconds, self.monitor.events)
        try:
            save_attempt(rec)
        except Exception:
            pass
        report = self.folder / "Result_Report.html"
        try:
            build_report(rec, report)
        except Exception:
            report = None
        if auto:
            messagebox.showinfo(APP_TITLE, "Time is over. Your test has been submitted and checked.")
        self.app.attributes("-topmost", False)
        self.app.switch(ResultView(self.app, rec, report))


# =================================================================== results

def fill_result_tree(tree, results):
    """results: list of task result dicts. Returns {iid: (kind, data)}."""
    meta = {}
    for t in results:
        pid = tree.insert("", "end", text=f"{APP_LABEL[t['app']]} – {t['title']}", values=(f"{t['got']:g} / {t['marks']:g}",),
                          open=True)
        meta[pid] = ("task", t)
        for c in t["checks"]:
            if c["desc"] == "Note":
                iid = tree.insert(pid, "end", text=f"ⓘ {c['found']}", values=("",), tags=("note",))
            else:
                iid = tree.insert(pid, "end", text=("✔ " if c["ok"] else "✘ ") + c["desc"],
                                  values=(f"{c['got']:g} / {c['marks']:g}",), tags=("ok" if c["ok"] else "bad",))
            meta[iid] = ("check", c)
    tree.tag_configure("ok", foreground=OK)
    tree.tag_configure("bad", foreground=BAD)
    tree.tag_configure("note", foreground="#7a5b00")
    return meta


def show_detail(text, kind, data):
    text.config(state="normal")
    text.delete("1.0", "end")
    if kind == "task":
        text.insert("end", f"{data['title']}  –  {data['got']:g} / {data['marks']:g}\n", "h")
        text.insert("end", f"File: {data['file']}\n\n")
        text.insert("end", "Model answer (step by step):\n", "b")
        for i, s in enumerate(data.get("solution") or [], 1):
            text.insert("end", f"{i}. {s}\n")
    else:
        c = data
        text.insert("end", c["desc"] + "\n", "h")
        text.insert("end", ("✔ Correct" if c["ok"] else "✘ Not correct") + f"   ({c['got']:g} / {c['marks']:g} marks)\n\n",
                    "ok" if c["ok"] else "bad")
        text.insert("end", "What we found: ", "b")
        text.insert("end", c["found"] + "\n")
        if c.get("expected"):
            text.insert("end", "Expected: ", "b")
            text.insert("end", c["expected"] + "\n")
        if c.get("how"):
            text.insert("end", "\nHow to do it:\n", "b")
            text.insert("end", c["how"] + "\n", "how")
        if c.get("topic"):
            text.insert("end", f"\nSyllabus topic: {c['topic']}\n", "muted")
    text.config(state="disabled")


def make_detail_text(parent):
    t = ScrolledText(parent, wrap="word", font=("Segoe UI", 11), padx=10, pady=8, height=10)
    t.tag_configure("h", font=("Segoe UI", 12, "bold"))
    t.tag_configure("b", font=("Segoe UI", 11, "bold"))
    t.tag_configure("ok", foreground=OK, font=("Segoe UI", 11, "bold"))
    t.tag_configure("bad", foreground=BAD, font=("Segoe UI", 11, "bold"))
    t.tag_configure("how", background="#fff8e1")
    t.tag_configure("muted", foreground="#5b6475")
    t.config(state="disabled")
    return t


class ResultView(ttk.Frame):
    def __init__(self, app: App, rec, report):
        super().__init__(app, padding=12)
        self.app = app
        ttk.Label(self, text=f"{rec['title']} – Result", style="H1.TLabel").pack(anchor="w")
        summ = ttk.Frame(self)
        summ.pack(fill="x", pady=6)
        ttk.Label(summ, text=f"Total: {rec['score']:g} / {rec['out_of']:g}", font=("Segoe UI", 16, "bold")).pack(side="left")
        for k in ("word", "excel", "ppt"):
            if k in rec["sections"]:
                g, m = rec["sections"][k]
                ttk.Label(summ, text=f"   {APP_LABEL[k]}: {g:g}/{m:g}", font=("Segoe UI", 12)).pack(side="left")
        ttk.Label(self, text=f"Time used {fmt_secs(rec['seconds_used'])} of {fmt_secs(rec['seconds_allowed'])}   ·   "
                             f"Time in Word {fmt_secs(rec['app_seconds'].get('word', 0))}, "
                             f"Excel {fmt_secs(rec['app_seconds'].get('excel', 0))}, "
                             f"PowerPoint {fmt_secs(rec['app_seconds'].get('ppt', 0))}",
                  style="Muted.TLabel").pack(anchor="w")
        ttk.Label(self, text="Click any line to see what was found, what was expected and how to do it. "
                             "Click a task name for its model answer.", style="Muted.TLabel").pack(anchor="w", pady=(2, 6))

        pane = ttk.PanedWindow(self, orient="vertical")
        pane.pack(fill="both", expand=True)
        tf = ttk.Frame(pane)
        self.tree = ttk.Treeview(tf, columns=("marks",), show="tree headings")
        self.tree.heading("#0", text="Task / check")
        self.tree.heading("marks", text="Marks")
        self.tree.column("marks", width=90, anchor="center", stretch=False)
        sb = ttk.Scrollbar(tf, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        pane.add(tf, weight=3)
        self.detail = make_detail_text(pane)
        pane.add(self.detail, weight=2)
        self.meta = fill_result_tree(self.tree, rec["results"])
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        bar = ttk.Frame(self)
        bar.pack(fill="x", pady=8)
        ttk.Button(bar, text="← Home", command=app.show_home).pack(side="left")
        if report:
            ttk.Button(bar, text="Open full report (printable)", command=lambda: webbrowser.open(Path(report).as_uri())).pack(side="right")
        ttk.Button(bar, text="Open test folder", command=lambda: open_path(rec["folder"])).pack(side="right", padx=6)

        first_bad = next((i for i, (k, d) in self.meta.items() if k == "check" and not d["ok"] and d["desc"] != "Note"), None)
        if first_bad:
            self.tree.selection_set(first_bad)
            self.tree.see(first_bad)

    def on_select(self, _):
        sel = self.tree.selection()
        if sel and sel[0] in self.meta:
            show_detail(self.detail, *self.meta[sel[0]])


class ResultPopup(tk.Toplevel):
    """Instant check of one task (practice and drills)."""
    def __init__(self, app, res):
        super().__init__(app)
        self.title(f"Check: {res.title}")
        self.geometry("720x520")
        self.attributes("-topmost", True)
        ttk.Label(self, text=f"{res.title}: {res.got:g} / {res.marks:g}", style="H2.TLabel", padding=8).pack(anchor="w")
        d = res.to_dict()
        pane = ttk.PanedWindow(self, orient="vertical")
        pane.pack(fill="both", expand=True, padx=8, pady=4)
        tree = ttk.Treeview(pane, columns=("marks",), show="tree")
        tree.column("marks", width=80, anchor="center", stretch=False)
        pane.add(tree, weight=3)
        detail = make_detail_text(pane)
        pane.add(detail, weight=2)
        meta = fill_result_tree(tree, [d])
        tree.bind("<<TreeviewSelect>>", lambda e: tree.selection() and tree.selection()[0] in meta and
                  show_detail(detail, *meta[tree.selection()[0]]))
        ttk.Label(self, text="Fix the mistakes, save, and press 'Check this task now' again.", style="Muted.TLabel",
                  padding=6).pack(anchor="w")
        ttk.Button(self, text="Close", command=self.destroy).pack(pady=(0, 8))


# =================================================================== quiz

class QuizView(ttk.Frame):
    def __init__(self, app: App, questions):
        super().__init__(app, padding=20)
        self.app = app
        self.qs = questions
        self.i = 0
        self.score = 0
        self.answers = []
        ttk.Label(self, text="Shortcut Key Quiz", style="H1.TLabel").pack(anchor="w")
        self.progress = ttk.Label(self, text="", style="Muted.TLabel")
        self.progress.pack(anchor="w")
        self.q = ttk.Label(self, text="", font=("Segoe UI", 15), wraplength=800, padding=(0, 20))
        self.q.pack(anchor="w")
        self.btns = []
        for _ in range(4):
            b = tk.Button(self, text="", font=("Segoe UI", 13), width=28, relief="groove", bg="#f3f5fa")
            b.pack(anchor="w", pady=4)
            self.btns.append(b)
        self.feedback = ttk.Label(self, text="", font=("Segoe UI", 12, "bold"))
        self.feedback.pack(anchor="w", pady=10)
        bar = ttk.Frame(self)
        bar.pack(fill="x", side="bottom")
        ttk.Button(bar, text="← Home", command=app.show_home).pack(side="left")
        self.next_b = ttk.Button(bar, text="Next →", command=self.next)
        self.next_b.pack(side="right")
        self.show()

    def show(self):
        q = self.qs[self.i]
        self.progress.config(text=f"Question {self.i + 1} of {len(self.qs)}   ·   Score {self.score}")
        self.q.config(text=f"Which shortcut does this?\n{q['q']}")
        for b, opt in zip(self.btns, q["options"]):
            b.config(text=opt, state="normal", bg="#f3f5fa", command=lambda o=opt: self.pick(o))
        self.feedback.config(text="")
        self.next_b.state(["disabled"])

    def pick(self, opt):
        q = self.qs[self.i]
        ok = opt == q["answer"]
        self.score += ok
        self.answers.append((q, opt, ok))
        for b in self.btns:
            b.config(state="disabled")
            if b.cget("text") == q["answer"]:
                b.config(bg="#c8e6c9")
            elif b.cget("text") == opt:
                b.config(bg="#ffcdd2")
        self.feedback.config(text="Correct!" if ok else f"Not correct. The answer is {q['answer']}.",
                             foreground=OK if ok else BAD)
        self.next_b.state(["!disabled"])

    def next(self):
        self.i += 1
        if self.i < len(self.qs):
            self.show()
            return
        for w in self.winfo_children():
            w.destroy()
        ttk.Label(self, text=f"Quiz finished: {self.score} / {len(self.qs)}", style="H1.TLabel").pack(anchor="w")
        t = ScrolledText(self, wrap="word", font=("Segoe UI", 11), height=20)
        t.tag_configure("bad", foreground=BAD)
        for q, opt, ok in self.answers:
            line = f"{'✔' if ok else '✘'} {q['q']} → {q['answer']}" + ("" if ok else f"   (you chose {opt})") + "\n"
            t.insert("end", line, () if ok else ("bad",))
        t.config(state="disabled")
        t.pack(fill="both", expand=True, pady=10)
        bar = ttk.Frame(self)
        bar.pack(fill="x")
        ttk.Button(bar, text="← Home", command=self.app.show_home).pack(side="left")
        ttk.Button(bar, text="New quiz", command=lambda: self.app.switch(QuizView(self.app, make_quiz(20)))).pack(side="right")
        from .session import save_attempt as _save
        try:
            _save({"when": time.strftime("%Y-%m-%dT%H:%M:%S"), "mode": "Shortcut quiz", "title": "Shortcut Key Quiz",
                   "folder": "", "score": self.score, "out_of": len(self.qs), "sections": {}, "seconds_used": 0,
                   "seconds_allowed": 0, "app_seconds": {}, "results": [], "events": []})
        except Exception:
            pass


# =================================================================== history

class HistoryView(ttk.Frame):
    def __init__(self, app: App):
        super().__init__(app, padding=14)
        self.app = app
        ttk.Label(self, text="My Progress", style="H1.TLabel").pack(anchor="w")
        hist = load_history()
        cols = ("when", "mode", "score", "word", "excel", "ppt", "time")
        tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        for c, label, w in (("when", "Date", 140), ("mode", "Test", 260), ("score", "Score", 80), ("word", "Word", 70),
                            ("excel", "Excel", 70), ("ppt", "PowerPoint", 80), ("time", "Time used", 80)):
            tree.heading(c, text=label)
            tree.column(c, width=w, anchor="w" if c in ("when", "mode") else "center")
        self.rows = {}
        for rec in reversed(hist):
            s = rec.get("sections", {})
            fmt = lambda k: f"{s[k][0]:g}/{s[k][1]:g}" if k in s else "–"
            iid = tree.insert("", "end", values=(rec["when"].replace("T", " ")[:16], f"{rec['title']} ({rec['mode']})",
                                                 f"{rec['score']:g}/{rec['out_of']:g}", fmt("word"), fmt("excel"), fmt("ppt"),
                                                 fmt_secs(rec.get("seconds_used", 0))))
            self.rows[iid] = rec
        tree.pack(fill="both", expand=True)
        tree.bind("<Double-1>", lambda e: self.open(tree))
        ttk.Label(self, text="Double-click a test to open its full report.", style="Muted.TLabel").pack(anchor="w")

        ttk.Label(self, text="Weak areas (topics you most often get wrong)", style="H2.TLabel").pack(anchor="w", pady=(12, 4))
        weak = weak_topics(hist)
        wt = ttk.Treeview(self, columns=("app", "topic", "rate"), show="headings", height=8)
        for c, label, w in (("app", "Program", 120), ("topic", "Syllabus topic", 380), ("rate", "Correct", 120)):
            wt.heading(c, text=label)
            wt.column(c, width=w)
        for app_, topic, ok, n, rate in weak:
            wt.insert("", "end", values=(APP_LABEL.get(app_, app_), topic, f"{ok}/{n}  ({rate:.0%})"))
        wt.pack(fill="x")
        if not weak:
            ttk.Label(self, text="Take a test first – your weak areas will show here.", style="Muted.TLabel").pack(anchor="w")
        ttk.Button(self, text="← Home", command=app.show_home).pack(anchor="w", pady=8)

    def open(self, tree):
        sel = tree.selection()
        if not sel:
            return
        rec = self.rows[sel[0]]
        rep = Path(rec.get("folder") or ".") / "Result_Report.html"
        if rec.get("folder") and rep.exists():
            webbrowser.open(rep.as_uri())
        elif rec.get("results"):
            self.app.switch(ResultView(self.app, rec, None))


def main():
    app = App()
    app.mainloop()
