"""Desktop window (customtkinter, clean light look)."""
from __future__ import annotations

import os
import queue
import random
import subprocess
import sys
import time
import tkinter as tk
import webbrowser
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from . import __version__
from .content import PAPERS, build_paper
from .content.drills import DRILLS
from .core import APP_LABEL, all_file_names, grade_task, grade_tasks
from .monitor import APP_NAME, Monitor, arrange_split
from .report import build as build_report, build_instructions
from .session import (DEFAULT_STUDENT, current_student, default_root, delete_attempt, export_results, fmt_secs,
                      load_history, load_settings, make_problem_report, make_record, new_attempt_folder,
                      next_attempt_number, prepare_folder, remove_student, save_attempt, save_settings, set_student,
                      stats, students, weak_topics)
from .shortcuts import make_quiz
from . import updater
import threading

APP_TITLE = "ISRO Skill Test Practice"
SECTION_MIN = {"word": 26, "excel": 26, "ppt": 23}

# ---------------------------------------------------------------- look
NAVY, NAVY_DARK, NAVY_SOFT = "#0B3D91", "#082B66", "#E8EEF9"
SAFFRON, SAFFRON_DARK = "#FF9933", "#E67E10"
BG, CARD, INK, MUTED, LINE = "#F3F5F9", "#FFFFFF", "#1D2433", "#5B6475", "#E3E7EF"
OK, OK_SOFT, BAD, BAD_SOFT = "#1A7F37", "#E6F4EA", "#C62828", "#FDECEC"
APP_COL = {"word": "#2B579A", "excel": "#217346", "ppt": "#C43E1C"}
APP_SHORT = {"word": "Word", "excel": "Excel", "ppt": "PowerPoint"}
FAMILY = "Segoe UI" if sys.platform == "win32" else "DejaVu Sans"
MONO = "Consolas" if sys.platform == "win32" else "DejaVu Sans Mono"


def F(size=13, weight="normal"):
    return ctk.CTkFont(family=FAMILY, size=size, weight=weight)


def card(parent, **kw):
    kw.setdefault("fg_color", CARD)
    kw.setdefault("corner_radius", 14)
    kw.setdefault("border_width", 1)
    kw.setdefault("border_color", LINE)
    return ctk.CTkFrame(parent, **kw)


def btn(parent, text, command, kind="primary", **kw):
    styles = {
        "primary": dict(fg_color=NAVY, hover_color=NAVY_DARK, text_color="white"),
        "accent": dict(fg_color=SAFFRON, hover_color=SAFFRON_DARK, text_color="#1D1300"),
        "ghost": dict(fg_color=CARD, hover_color=NAVY_SOFT, text_color=NAVY, border_width=1, border_color="#C9D3E6"),
        "danger": dict(fg_color=CARD, hover_color=BAD_SOFT, text_color=BAD, border_width=1, border_color="#EBC4C4"),
    }
    kw.setdefault("height", 38)
    kw.setdefault("corner_radius", 10)
    kw.setdefault("font", F(13, "bold"))
    return ctk.CTkButton(parent, text=text, command=command, **styles[kind], **kw)


def label(parent, text, size=13, weight="normal", color=INK, **kw):
    return ctk.CTkLabel(parent, text=text, font=F(size, weight), text_color=color, **kw)


def pill(parent, text, fg, color="white", size=12, **kw):
    return ctk.CTkLabel(parent, text=f"  {text}  ", font=F(size, "bold"), fg_color=fg, text_color=color,
                        corner_radius=10, height=24, **kw)


def option(parent, values, width=200):
    m = ctk.CTkOptionMenu(parent, values=values, width=width, height=36, corner_radius=10, font=F(13),
                          fg_color=NAVY_SOFT, button_color="#C9D6EE", button_hover_color="#B6C7E6",
                          text_color=INK, dropdown_font=F(13))
    m.set(values[0])
    return m


def textbox(parent, height=200, mono=False, size=14):
    t = ctk.CTkTextbox(parent, height=height, wrap="word", corner_radius=10, border_width=0, fg_color=CARD,
                       text_color=INK, font=ctk.CTkFont(family=MONO if mono else FAMILY, size=size))
    return t


def responsive(view, limit=900):
    """Call view.layout(compact) when the window crosses the 'narrow' width (e.g. split screen)."""
    width = view.winfo_width() / max(ctk.ScalingTracker.get_window_scaling(view), 0.1)
    if width < 50:
        return
    compact = width < limit
    if compact != view.compact:
        view.layout(compact)


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


def style_tree():
    s = ttk.Style()
    if "clam" in s.theme_names():
        s.theme_use("clam")
    s.configure("Nice.Treeview", background=CARD, fieldbackground=CARD, foreground=INK, rowheight=32,
                font=(FAMILY, 11), borderwidth=0, relief="flat")
    s.configure("Nice.Treeview.Heading", background="#EEF2FA", foreground=MUTED, font=(FAMILY, 10, "bold"),
                relief="flat", padding=(8, 6))
    s.map("Nice.Treeview", background=[("selected", "#DCE6F7")], foreground=[("selected", INK)])
    s.map("Nice.Treeview.Heading", background=[("active", "#E2E8F5")])
    s.layout("Nice.Treeview", [("Nice.Treeview.treearea", {"sticky": "nswe"})])
    s.configure("Vertical.TScrollbar", background="#D5DBE7", troughcolor=CARD, borderwidth=0, arrowsize=12)


# =================================================================== app

class App(ctk.CTk):
    def __init__(self):
        ctk.set_appearance_mode("light")
        super().__init__(fg_color=BG)
        self.title(f"{APP_TITLE}  v{__version__}" + (f" (build {updater.BUILD})" if updater.BUILD else ""))
        scale = max(ctk.ScalingTracker.get_window_scaling(self), 0.1)      # Windows display scaling, e.g. 125%
        sw, sh = int(self.winfo_screenwidth() / scale), int(self.winfo_screenheight() / scale)
        self.geometry(f"{min(1180, sw - 40)}x{min(780, sh - 90)}")
        self.minsize(420, 480)                      # small enough for half-screen / Windows Snap
        style_tree()

        head = ctk.CTkFrame(self, fg_color=NAVY, corner_radius=0, height=64)
        head.pack(fill="x")
        head.pack_propagate(False)
        logo = ctk.CTkLabel(head, text="◉", font=F(26, "bold"), text_color=SAFFRON)
        logo.pack(side="left", padx=(20, 8))
        box = ctk.CTkFrame(head, fg_color="transparent")
        box.pack(side="left")
        label(box, APP_TITLE, 18, "bold", "white").pack(anchor="w")
        label(box, "Skill Test for Assistant / UDC / JPA / Stenographer", 11, color="#C9D6EE").pack(anchor="w")
        self.head_right = label(head, "Word 35  ·  Excel 35  ·  PowerPoint 30  ·  75 min", 12, color="#C9D6EE")
        self.head_right.pack(side="right", padx=20)
        self.bind("<Configure>", self._head_fit, add="+")
        ctk.CTkFrame(self, fg_color=SAFFRON, corner_radius=0, height=4).pack(fill="x")

        self.body = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self.body.pack(fill="both", expand=True)
        self.frame = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.update_info = None
        self._upd = None
        updater.cleanup_leftovers()
        self.show_home()
        self.after(1500, lambda: self.check_updates(quiet=True))

    # ---------- updates
    def check_updates(self, quiet=False):
        def work():
            try:
                self._upd = ("ok", updater.latest())
            except Exception as e:
                self._upd = ("error", str(e))
        self._upd = None
        threading.Thread(target=work, daemon=True).start()
        self._wait_updates(quiet, 0)

    def _wait_updates(self, quiet, n):
        if self._upd is None:
            if n < 60:
                self.after(250, lambda: self._wait_updates(quiet, n + 1))
            elif not quiet:
                messagebox.showinfo(APP_TITLE, "Could not reach GitHub. Check the internet connection and try again.")
            return
        kind, info = self._upd
        if kind == "error":
            if not quiet:
                messagebox.showinfo(APP_TITLE, "Could not check for updates (no internet?).\n\n" + info[:200])
            return
        if not updater.frozen():
            if not quiet:
                webbrowser.open(updater.PAGE)
            return
        if updater.update_available(info):
            self.update_info = info
            if isinstance(self.frame, Home):
                self.frame.show_update(info)
        elif not quiet:
            messagebox.showinfo(APP_TITLE, f"You have the latest version (build {updater.BUILD}).")

    def run_update(self):
        info = self.update_info
        if not info:
            return
        if isinstance(self.frame, TestView) and not self.frame.done:
            messagebox.showinfo(APP_TITLE, "Finish or submit the test first, then update.")
            return
        box = ctk.CTkToplevel(self, fg_color=BG)
        box.title("Updating")
        box.attributes("-topmost", True)
        box.geometry("420x170")
        label(box, f"Downloading {info['tag']}…", 15, "bold", NAVY).pack(anchor="w", padx=20, pady=(20, 8))
        bar = ctk.CTkProgressBar(box, height=12, corner_radius=6, progress_color=SAFFRON, fg_color="#E6EAF2")
        bar.set(0)
        bar.pack(fill="x", padx=20)
        note = label(box, "The app will close and open again by itself.", 12, color=MUTED)
        note.pack(anchor="w", padx=20, pady=10)
        state = {"done": 0, "total": info.get("size") or 1, "result": None}

        def work():
            try:
                state["result"] = ("ok", updater.download(info, lambda d, t: state.update(done=d, total=t or state["total"])))
            except Exception as e:
                state["result"] = ("error", str(e))
        threading.Thread(target=work, daemon=True).start()

        def poll():
            bar.set(min(1.0, state["done"] / max(1, state["total"])))
            if state["result"] is None:
                self.after(200, poll)
                return
            kind, val = state["result"]
            if kind == "error":
                box.destroy()
                messagebox.showerror(APP_TITLE, f"Update failed:\n{val}\n\nYou can also download it from:\n{updater.PAGE}")
                return
            note.configure(text="Installing… the app will reopen in a few seconds.")
            box.update()
            updater.install_and_restart(val)
            self.after(600, self.destroy)
        poll()

    def _head_fit(self, e=None):
        if e is not None and e.widget is not self:
            return
        narrow = self.winfo_width() / max(ctk.ScalingTracker.get_window_scaling(self), 0.1) < 900
        if narrow and self.head_right.winfo_manager():
            self.head_right.pack_forget()
        elif not narrow and not self.head_right.winfo_manager():
            self.head_right.pack(side="right", padx=20)

    def on_close(self):
        f = self.frame
        if isinstance(f, TestView) and not f.done:
            if not messagebox.askyesno(APP_TITLE, "A test is running.\n\nSubmit it now and close the app?"):
                return
            f.finish(quiet=True)
        self.destroy()

    def switch(self, cls, *args, **kw):
        if self.frame is not None:
            self.frame.destroy()
        self.frame = cls(self.body, self, *args, **kw)
        self.frame.pack(fill="both", expand=True, padx=20, pady=16)

    def show_home(self):
        self.attributes("-topmost", False)
        self.switch(Home)

    # one place that knows how to start any kind of test (used by Home, Retake, Fix & recheck and History)
    def start(self, spec, recheck=None):
        spec = dict(spec or {})
        kind = spec.get("kind")
        strict = bool(load_settings().get("strict"))
        if kind == "mock":
            p = next(p for p in PAPERS if p.id == spec["paper"])
            title, minutes, mode = f"{p.title} - Full Test", p.minutes, "Full mock"
        elif kind == "section":
            p = next(p for p in PAPERS if p.id == spec["paper"])
            title, minutes, mode = f"{p.title} - {APP_SHORT[spec['app']]}", SECTION_MIN[spec["app"]], "Section practice"
        elif kind == "drills":
            chosen = [d for d in DRILLS if d.id in spec["ids"]]
            title = "Drill " + chosen[0].id if len(chosen) == 1 else f"Drills ({len(chosen)})"
            minutes, mode = sum(d.minutes for d in chosen), "Topic drills"
        else:
            messagebox.showinfo(APP_TITLE, "This attempt was made with an older version and cannot be retaken "
                                           "automatically. Start it again from the Home screen.")
            return
        if recheck:
            # same folder, same numbers: fix the mistakes and mark again
            seed = spec.get("seed", 0)
            folder, attempt = recheck["folder"], recheck.get("attempt", 1)
            mode, minutes, check, hints = "Fix & recheck", max(15, minutes // 2), True, True
        else:
            n = next_attempt_number(title)
            seed = 0 if n == 1 else random.randint(1, 999_999)       # attempt 1 = original paper, retakes = new numbers
            folder, attempt = None, None
            check = hints = kind != "mock" and not strict
        spec["seed"] = seed
        if kind == "mock":
            tasks = build_paper(spec["paper"], seed).tasks
        elif kind == "section":
            tasks = build_paper(spec["paper"], seed).section(spec["app"])
        else:
            tasks = chosen
        self.switch(TestView, tasks, title, minutes * 60, mode, check, spec, show_hints=hints, folder=folder,
                    attempt=attempt)


# =================================================================== home

class Home(ctk.CTkScrollableFrame):
    def __init__(self, parent, app: App):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        names = [p.title for p in PAPERS]

        # update banner (filled in when a newer version is found)
        self.upd = ctk.CTkFrame(self, fg_color="#FFF1E0", corner_radius=12, border_width=1, border_color="#F5C58A")
        if app.update_info:
            self.after(10, lambda: self.show_update(app.update_info))

        # student chooser
        who = ctk.CTkFrame(self, fg_color="transparent")
        who.pack(fill="x", pady=(0, 10))
        label(who, "Student", 13, "bold", MUTED).pack(side="left")
        self.student = option(who, students() + ["+ Add student…"], 220)
        self.student.set(current_student())
        self.student.configure(command=self.pick_student)
        self.student.pack(side="left", padx=10)
        if len(students()) > 1:
            btn(who, "Remove", self.remove_student, "danger", width=90, height=32).pack(side="left")
        self.who_note = label(who, "Each student has separate folders, scores and weak areas.", 12, color=MUTED)

        # stats strip
        st = stats(load_history())
        strip = self.strip = ctk.CTkFrame(self, fg_color="transparent")
        items = ([("Tests taken", f"{st['count']}"), ("Best score", f"{st['best']:.0f}%"),
                  ("Last score", f"{st['last']:.0f}%"), ("Average", f"{st['avg']:.0f}%")] if st else
                 [("Welcome!", "Start with Mock Test 1"), ("Format", "Word · Excel · PowerPoint"),
                  ("Duration", "75 minutes"), ("Total", "100 marks")])
        for i, (k, v) in enumerate(items):
            c = card(strip)
            c.grid(row=0, column=i, sticky="nsew", padx=(0 if i == 0 else 8, 0))
            strip.grid_columnconfigure(i, weight=1)
            label(c, k, 12, color=MUTED).pack(anchor="w", padx=16, pady=(12, 0))
            label(c, v, 20 if st else 15, "bold", NAVY).pack(anchor="w", padx=16, pady=(0, 12))

        grid = self.grid_ = ctk.CTkFrame(self, fg_color="transparent")
        self.cards = []

        c = self.mode_card(grid, 0, 0, "★", NAVY, "Full Mock Test",
                           "All three sections in 75 minutes, marked out of 100.\nYour work is checked after you submit.")
        self.mock = option(c.controls, names, 150)
        self.mock.pack(side="left")
        btn(c.controls, "Start full mock", self.start_mock, "accent", width=150).pack(side="right")

        c = self.mode_card(grid, 0, 1, "◎", APP_COL["excel"], "Section Practice",
                           "Only Word, Excel or PowerPoint, with a shorter timer.\nCheck each task as you go.")
        self.sec_paper = option(c.controls, names, 128)
        self.sec_paper.pack(side="left")
        self.sec_app = option(c.controls, ["Word", "Excel", "PowerPoint"], 118)
        self.sec_app.pack(side="left", padx=6)
        btn(c.controls, "Start", self.start_section, width=80).pack(side="right")

        c = self.mode_card(grid, 1, 0, "✎", APP_COL["ppt"], "Topic Drills",
                           f"{len(DRILLS)} short tasks, one syllabus topic each: Mail Merge,\nVLOOKUP, Subtotal, Slide Master and more.")
        btn(c.controls, "Choose drills", lambda: app.switch(DrillPicker), width=150).pack(side="right")

        c = self.mode_card(grid, 1, 1, "⌨", APP_COL["word"], "Shortcut Key Quiz",
                           "20 quick questions on keyboard shortcuts\nfor Word, Excel and PowerPoint.")
        self.quiz_app = option(c.controls, ["All programs", "Word", "Excel", "PowerPoint"], 150)
        self.quiz_app.pack(side="left")
        btn(c.controls, "Start quiz", self.start_quiz, width=110).pack(side="right")

        bar = self.bar = ctk.CTkFrame(self, fg_color="transparent")
        self.bar_a = ctk.CTkFrame(bar, fg_color="transparent")
        self.bar_b = ctk.CTkFrame(bar, fg_color="transparent")
        btn(self.bar_a, "My progress & retakes", lambda: app.switch(HistoryView), "ghost", width=190).pack(side="left")
        btn(self.bar_a, "How it works", self.help, "ghost", width=120).pack(side="left", padx=8)
        self.strict = tk.BooleanVar(value=bool(load_settings().get("strict")))
        ctk.CTkSwitch(self.bar_a, text="Strict exam mode", variable=self.strict, font=F(12), progress_color=NAVY,
                      command=self.toggle_strict).pack(side="left", padx=6)
        btn(self.bar_b, "Change", self.change_root, "ghost", width=80).pack(side="right")
        btn(self.bar_b, "Check for updates", lambda: app.check_updates(), "ghost", width=150).pack(side="left")
        self.root_lbl = label(self.bar_b, f"Test folders: {default_root()}", 12, color=MUTED)
        self.root_lbl.pack(side="right", padx=8)

        self.compact = None
        self.layout(False)
        self.bind("<Configure>", lambda e: responsive(self))

    def layout(self, compact):
        self.compact = compact
        for w in (self.who_note, self.strip, self.grid_, self.bar, self.bar_a, self.bar_b):
            w.pack_forget()
        if not compact:
            self.who_note.pack(side="left", padx=12)
            self.strip.pack(fill="x", pady=(0, 14))
        self.grid_.pack(fill="both", expand=True)
        for i in (0, 1):
            self.grid_.grid_columnconfigure(i, weight=0 if (compact and i) else 1, uniform="" if compact else "c")
        for i, k in enumerate(self.cards):
            r, c = (i, 0) if compact else divmod(i, 2)
            k.grid(row=r, column=c, sticky="nsew", padx=(0, 0) if compact else ((0, 8) if c == 0 else (8, 0)), pady=8)
            k.desc.configure(text=k.desc_text.replace("\n", " ") if compact else k.desc_text,
                             wraplength=340 if compact else 0)
        self.bar.pack(fill="x", pady=(8, 0))
        if compact:
            self.bar_a.pack(side="top", anchor="w")
            self.bar_b.pack(side="top", fill="x", pady=(8, 0))
        else:
            self.bar_a.pack(side="left")
            self.bar_b.pack(side="right")

    def show_update(self, info):
        for w in self.upd.winfo_children():
            w.destroy()
        label(self.upd, f"New version {info['tag']} is available.", 14, "bold", "#7A3E00").pack(side="left", padx=16, pady=10)
        btn(self.upd, "Update now", self.app.run_update, "accent", width=120, height=34).pack(side="right", padx=(6, 12))
        notes = (info.get("notes") or "").split("\n---")[0]
        lines = [l.strip("-• ").strip() for l in notes.splitlines()
                 if l.strip() and not l.startswith(("Co-Authored-By", "Claude-Session"))][:15] or ["Improvements and fixes."]
        btn(self.upd, "What's new", lambda: Popup(self.app, f"What's new in {info['tag']}", lines=lines),
            "ghost", width=110, height=34).pack(side="right")
        if not self.upd.winfo_manager():
            self.upd.pack(fill="x", pady=(0, 10), before=self.winfo_children()[1])

    def toggle_strict(self):
        s = load_settings()
        s["strict"] = self.strict.get()
        save_settings(s)
        if self.strict.get():
            messagebox.showinfo(APP_TITLE, "Strict exam mode is ON.\n\nSection practice and drills will not show "
                                           "'Check task' or model answers until you submit, like the real exam.")

    def mode_card(self, grid, r, c, icon, colour, title, desc):
        k = card(grid)
        self.cards.append(k)
        top = ctk.CTkFrame(k, fg_color="transparent")
        top.pack(fill="x", padx=20, pady=(18, 6))
        ctk.CTkLabel(top, text=icon, width=44, height=44, corner_radius=22, fg_color=colour, text_color="white",
                     font=F(20, "bold")).pack(side="left")
        label(top, title, 18, "bold").pack(side="left", padx=12)
        k.desc_text = desc
        k.desc = label(k, desc, 13, color=MUTED, justify="left")
        k.desc.pack(anchor="w", padx=20)
        k.controls = ctk.CTkFrame(k, fg_color="transparent")
        k.controls.pack(side="bottom", fill="x", padx=20, pady=(26, 18))
        return k

    def pick_student(self, value):
        if value == "+ Add student…":
            name = ctk.CTkInputDialog(title="Add student", text="Student name:").get_input()
            name = (name or "").strip()
            if not name:
                self.student.set(current_student())
                return
            set_student(name)
        else:
            set_student(value)
        self.app.show_home()

    def remove_student(self):
        name = current_student()
        if messagebox.askyesno(APP_TITLE, f"Remove '{name}' from the student list?\n\nTheir saved scores and folders "
                                          "are kept; add the same name again to see them."):
            remove_student(name)
            self.app.show_home()

    def start_mock(self):
        p = PAPERS[[x.title for x in PAPERS].index(self.mock.get())]
        self.app.start({"kind": "mock", "paper": p.id})

    def start_section(self):
        p = PAPERS[[x.title for x in PAPERS].index(self.sec_paper.get())]
        app = {"Word": "word", "Excel": "excel", "PowerPoint": "ppt"}[self.sec_app.get()]
        self.app.start({"kind": "section", "paper": p.id, "app": app})

    def start_quiz(self):
        app = {"All programs": None}.get(self.quiz_app.get(), self.quiz_app.get())
        self.app.switch(QuizView, make_quiz(20, app), app)

    def change_root(self):
        d = filedialog.askdirectory(title="Choose where test folders are created", initialdir=str(default_root()))
        if d:
            s = load_settings()
            s["root"] = d
            save_settings(s)
            self.root_lbl.configure(text=f"Test folders: {d}")

    def help(self):
        messagebox.showinfo(APP_TITLE, HELP_TEXT)


HELP_TEXT = """How it works

1. Pick a test. A folder is made for it (e.g. Mock Test 1 - Full Test \\ Attempt 1) and opened for you, with any starter files and a Pictures folder.
2. Read each task here and do it in the real MS Word / Excel / PowerPoint.
3. Save every file in that folder with the EXACT name asked.
4. While you work, the app records which program is open, files created and saved, formulas typed, slides added, and the time you spend in each program.
5. Press Submit (or let the timer run out). Save your files first.
6. Each point is marked. For every mistake you see what was found, what was expected and how to do it, plus a model answer.

Retaking a test makes a new 'Attempt' folder with NEW names and numbers, so you cannot just remember the answers. Your old work and score are kept. 'Fix & recheck' lets you correct the same files and mark them again.

Strict exam mode hides 'Check task' and model answers until you submit. Copying the passages you must type is blocked.

'Instructions in Hindi' opens the question paper in English and Hindi in your browser.

Tip: tick 'Keep on top' and place this window beside Office."""


# =================================================================== drill picker

class DrillPicker(ctk.CTkFrame):
    def __init__(self, parent, app: App):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x")
        label(top, "Topic Drills", 22, "bold").pack(side="left")
        self.filter = ctk.CTkSegmentedButton(top, values=["All", "Word", "Excel", "PowerPoint"], command=self.render,
                                             font=F(13, "bold"), height=36, corner_radius=10, fg_color="#E6EAF2",
                                             selected_color=SAFFRON, selected_hover_color=SAFFRON_DARK,
                                             unselected_color="#E6EAF2", unselected_hover_color="#D5DDEC",
                                             text_color=INK)
        self.filter.set("All")
        self.filter.pack(side="right")
        label(self, "Tick one or more drills and press Start. Each drill can be checked instantly and has a model answer.",
              13, color=MUTED).pack(anchor="w", pady=(2, 10))

        self.vars = {d.id: tk.BooleanVar(value=False) for d in DRILLS}
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(side="bottom", fill="x", pady=(12, 0))
        btn(bar, "← Home", app.show_home, "ghost", width=110).pack(side="left")
        btn(bar, "Start selected", self.start, "accent", width=170).pack(side="right")
        self.summary = label(bar, "", 13, color=MUTED)
        self.summary.pack(side="right", padx=14)
        self.list = ctk.CTkScrollableFrame(self, fg_color=CARD, corner_radius=14, border_width=1, border_color=LINE)
        self.list.pack(fill="both", expand=True)

        self.render()

    def render(self, *_):
        for w in self.list.winfo_children():
            w.destroy()
        want = {"All": None, "Word": "word", "Excel": "excel", "PowerPoint": "ppt"}[self.filter.get()]
        for app_key in ("word", "excel", "ppt"):
            if want and app_key != want:
                continue
            h = ctk.CTkFrame(self.list, fg_color="transparent")
            h.pack(fill="x", padx=10, pady=(12, 4))
            pill(h, APP_LABEL[app_key], APP_COL[app_key]).pack(side="left")
            for d in DRILLS:
                if d.app != app_key:
                    continue
                row = ctk.CTkFrame(self.list, fg_color="transparent")
                row.pack(fill="x", padx=14, pady=2)
                ctk.CTkCheckBox(row, text=d.title, variable=self.vars[d.id], font=F(13), text_color=INK,
                                fg_color=NAVY, hover_color=NAVY_DARK, border_color="#AEB9CE", corner_radius=6,
                                command=self.update_summary).pack(side="left")
                label(row, f"{d.minutes} min", 12, color=MUTED).pack(side="right", padx=8)
                label(row, d.topic, 12, color=MUTED).pack(side="right", padx=8)
        self.update_summary()

    def chosen(self):
        return [d for d in DRILLS if self.vars[d.id].get()]

    def update_summary(self):
        c = self.chosen()
        self.summary.configure(text=f"{len(c)} selected · {sum(d.minutes for d in c)} min" if c else "Nothing selected")

    def start(self):
        c = self.chosen()
        if not c:
            messagebox.showinfo(APP_TITLE, "Tick at least one drill first.")
            return
        self.app.start({"kind": "drills", "ids": [d.id for d in c]})


# =================================================================== test view

class TaskItem(ctk.CTkFrame):
    def __init__(self, parent, task, on_click):
        super().__init__(parent, fg_color=CARD, corner_radius=10, border_width=1, border_color=LINE)
        self.task = task
        bar = ctk.CTkFrame(self, fg_color=APP_COL[task.app], width=5, height=36, corner_radius=3)
        bar.pack(side="left", padx=(8, 10), pady=10)
        box = ctk.CTkFrame(self, fg_color="transparent")
        box.pack(side="left", fill="x", expand=True, pady=6)
        self.l1 = label(box, APP_SHORT[task.app].upper(), 10, "bold", APP_COL[task.app])
        self.l1.pack(anchor="w")
        self.l2 = label(box, task.title, 13, "bold")
        self.l2.pack(anchor="w")
        self.l3 = label(self, f"{task.marks:g}", 13, "bold", MUTED)
        self.l3.pack(side="right", padx=12)
        for w in (self, bar, box, self.l1, self.l2, self.l3):
            w.bind("<Button-1>", lambda e: on_click(self))

    def select(self, on):
        self.configure(fg_color=NAVY_SOFT if on else CARD, border_color=NAVY if on else LINE)


class TestView(ctk.CTkFrame):
    def __init__(self, parent, app: App, tasks, title, seconds, mode, allow_check, spec, show_hints=False,
                 folder=None, attempt=None):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.tasks = tasks
        self.title_ = title
        self.mode = mode
        self.spec = spec
        self.allowed = seconds
        self.started = time.monotonic()
        self.warned = set()
        self.done = False
        self.events_q = queue.Queue()
        self.cur = 0
        self.mini = None
        self.allow_check = allow_check

        if folder:                                   # Fix & recheck: keep working in the same folder
            self.folder, self.attempt = Path(folder), attempt
        else:
            self.folder, self.attempt = new_attempt_folder(title)
        prepare_folder(self.folder, tasks)
        self.passages = [l.strip() for t in tasks for l in t.instructions if l.startswith("   ") and len(l.strip()) > 40]
        self.clip_warned = False
        self.known = all_file_names(tasks)
        self.monitor = Monitor(self.folder, on_event=self.events_q.put)

        # ---- top bar (wide: one line; narrow window: two lines)
        top = card(self)
        top.pack(fill="x")
        self.t_left = ctk.CTkFrame(top, fg_color="transparent")
        row = ctk.CTkFrame(self.t_left, fg_color="transparent")
        row.pack(anchor="w")
        self.t_title = label(row, title, 18, "bold")
        self.t_title.pack(side="left")
        self.t_pills = [pill(row, f"Attempt {self.attempt}", NAVY_SOFT, NAVY), pill(row, mode, "#FFF1E0", "#9A4D00")]
        rel = self.folder.relative_to(default_root()) if default_root() in self.folder.parents else self.folder
        self.flink = ctk.CTkLabel(self.t_left, text=f"Folder: …\\{rel}  (click to open)", font=F(11),
                                  text_color=MUTED, cursor="hand2")
        self.flink.bind("<Button-1>", lambda e: open_path(self.folder))

        self.t_ctrl = ctk.CTkFrame(top, fg_color="transparent")
        btn(self.t_ctrl, "Submit", self.submit, "accent", width=100, height=44).pack(side="right", padx=(6, 0))
        self.timer = ctk.CTkLabel(self.t_ctrl, text=fmt_secs(seconds), font=ctk.CTkFont(family=MONO, size=24, weight="bold"),
                                  fg_color=NAVY, text_color="white", corner_radius=10, width=110, height=44)
        self.timer.pack(side="right", padx=6)
        btn(self.t_ctrl, "Split screen", self.split_screen, "ghost", width=104, height=44).pack(side="right", padx=4)
        btn(self.t_ctrl, "Mini bar", self.open_mini, "ghost", width=84, height=44).pack(side="right", padx=4)
        self.topmost = tk.BooleanVar(value=False)
        self.top_switch = ctk.CTkSwitch(self.t_ctrl, text="Keep on top", variable=self.topmost, font=F(12),
                                        progress_color=NAVY, command=lambda: app.attributes("-topmost", self.topmost.get()))

        # ---- body
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, pady=10)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        self.side = ctk.CTkScrollableFrame(body, width=250, fg_color="transparent", label_text="Tasks",
                                           label_font=F(13, "bold"), label_fg_color="transparent", label_text_color=MUTED)
        self.items = []
        for i, t in enumerate(tasks):
            it = TaskItem(self.side, t, self.pick)
            it.pack(fill="x", pady=4, padx=2)
            self.items.append(it)

        main = card(body)
        main.grid(row=0, column=1, sticky="nsew")
        # task chooser used instead of the side list when the window is narrow
        self.task_menu = option(main, [f"{i + 1}. {APP_SHORT[t.app]} – {t.title}" for i, t in enumerate(tasks)], 300)
        self.task_menu.configure(command=lambda v: self.pick(self.items[int(v.split(".")[0]) - 1]))
        self.head = ctk.CTkFrame(main, fg_color="transparent")
        self.head.pack(fill="x", padx=20, pady=(16, 4))
        self.h_app = pill(self.head, "", NAVY)
        self.h_app.pack(side="left")
        self.h_title = label(self.head, "", 20, "bold")
        self.h_title.pack(side="left", padx=10)
        self.h_marks = pill(self.head, "", NAVY_SOFT, NAVY)
        self.h_marks.pack(side="right")
        self.h_file = label(main, "", 13, "bold", NAVY)
        self.h_file.pack(anchor="w", padx=20)
        self.h_start = label(main, "", 12, color=MUTED)
        self.h_start.pack(anchor="w", padx=20)
        ctk.CTkFrame(main, fg_color=LINE, height=1).pack(fill="x", padx=20, pady=8)

        btns = ctk.CTkFrame(main, fg_color="transparent")
        btns.pack(side="bottom", fill="x", padx=16, pady=12)
        self.b_tools = ctk.CTkFrame(btns, fg_color="transparent")
        self.b_acts = ctk.CTkFrame(btns, fg_color="transparent")
        btn(self.b_tools, "Open folder", lambda: open_path(self.folder), "ghost", width=104).pack(side="left")
        btn(self.b_tools, "In Hindi", self.hindi, "ghost", width=90).pack(side="left", padx=6)
        self.b_starter = btn(self.b_tools, "Starter file", self.open_starter, "ghost", width=104)
        self.b_starter.pack(side="left")
        btn(self.b_acts, "→", lambda: self.move(1), "ghost", width=44).pack(side="right")
        btn(self.b_acts, "←", lambda: self.move(-1), "ghost", width=44).pack(side="right", padx=6)
        if show_hints:
            btn(self.b_acts, "Model answer", self.hint, "ghost", width=120).pack(side="right")
        if allow_check:
            btn(self.b_acts, "Check task", self.check_one, width=104).pack(side="right", padx=6)

        self.text = textbox(main, size=14)
        self.text.pack(fill="both", expand=True, padx=14)
        self.text.tag_config("quote", foreground="#334", lmargin1=30, lmargin2=30, background="#F6F8FC")
        self.text.tag_config("num", foreground=NAVY)
        self.text.tag_config("hint", foreground="#7A5B00")
        self.block_copy(self.text)

        # ---- activity
        act = card(self)
        act.pack(fill="x")
        ah = ctk.CTkFrame(act, fg_color="transparent")
        ah.pack(fill="x", padx=16, pady=(10, 4))
        label(ah, "Live activity", 13, "bold").pack(side="left")
        self.chips = {}
        for k in ("ppt", "excel", "word"):
            c = pill(ah, f"{APP_SHORT[k]} 0:00", APP_COL[k], size=11)
            c.pack(side="right", padx=3)
            self.chips[k] = c
        self.log = textbox(act, height=90, mono=True, size=11)
        self.log.configure(state="disabled")
        self.body_ = body

        self.compact = None
        self.layout(False)
        self.bind("<Configure>", self.on_resize)
        app.bind("<Left>", lambda e: self.move(-1))           # ← → change task (main window and mini bar)
        app.bind("<Right>", lambda e: self.move(1))

        self.pick(self.items[0])
        self.monitor.start()
        if mode == "Fix & recheck":
            self.monitor.log("test", "Fix & recheck: correct your files and press 'Check task' or Submit")
        open_path(self.folder)
        self.tick()

    # ---------- adapt to the window size (works in split screen / half screen)
    def on_resize(self, _=None):
        width = self.winfo_width() / max(ctk.ScalingTracker.get_window_scaling(self.app), 0.1)
        compact = width < 900
        if compact != self.compact:
            self.layout(compact)

    def layout(self, compact):
        self.compact = compact
        for w in (self.t_left, self.t_ctrl, self.flink, self.top_switch, self.task_menu, self.b_tools,
                  self.b_acts, self.log, *self.t_pills):
            w.pack_forget()
        self.side.grid_forget()
        if compact:
            self.t_left.pack(side="top", fill="x", padx=14, pady=(10, 0))
            self.t_ctrl.pack(side="top", fill="x", padx=14, pady=8)
            self.t_title.configure(font=F(15, "bold"))
            self.h_title.configure(font=F(16, "bold"), wraplength=260, justify="left")
            self.h_start.configure(wraplength=360, justify="left")
            self.task_menu.pack(fill="x", padx=14, pady=(12, 0), before=self.head)
            self.b_acts.pack(side="top", fill="x")
            self.b_tools.pack(side="top", fill="x", pady=(8, 0))
        else:
            self.t_left.pack(side="left", padx=16, pady=10)
            for p in self.t_pills:
                p.pack(side="left", padx=(10, 0))
            self.flink.pack(anchor="w")
            self.t_ctrl.pack(side="right", padx=16, pady=10)
            self.top_switch.pack(side="right", padx=8)
            self.t_title.configure(font=F(18, "bold"))
            self.h_title.configure(font=F(20, "bold"), wraplength=0)
            self.h_start.configure(wraplength=0)
            self.side.grid(row=0, column=0, sticky="nsw", padx=(0, 12))
            self.b_tools.pack(side="left")
            self.b_acts.pack(side="right")
            self.log.pack(fill="x", padx=12, pady=(0, 10))

    def destroy(self):
        for k in ("<Left>", "<Right>"):
            self.app.unbind(k)
        super().destroy()

    def open_mini(self):
        """Small floating question box; the main window goes to the taskbar."""
        if self.mini is None or not self.mini.winfo_exists():
            self.mini = MiniBar(self)
        self.mini.deiconify()
        self.mini.lift()
        self.app.attributes("-topmost", False)
        self.app.iconify()
        self.monitor.log("test", "Mini bar opened (Office can stay full screen)")

    def close_mini(self):
        if self.mini is not None:
            try:
                self.mini.destroy()
            except tk.TclError:
                pass
            self.mini = None
        self.app.deiconify()
        self.app.lift()

    def split_screen(self):
        """Office on the left, this window on the right (always on top)."""
        moved = arrange_split(self.app, app_share=0.42)
        self.topmost.set(True)
        self.app.attributes("-topmost", True)
        if moved == 0 and sys.platform == "win32":
            Popup(self.app, "Split screen", "Open Word, Excel or PowerPoint first, then press Split screen again.\n"
                  "Tip: you can also press Windows key + Left arrow in Office.", auto_close=7000)

    # ---------- no copying of the passages to be typed
    def block_copy(self, box):
        t = box._textbox
        t.configure(selectbackground=CARD, selectforeground=INK, inactiveselectbackground=CARD, exportselection=False)
        for ev in ("<Control-c>", "<Control-C>", "<<Copy>>", "<Control-Insert>", "<Button-3>", "<B1-Motion>",
                   "<Double-Button-1>", "<Triple-Button-1>", "<Control-a>", "<Control-A>", "<Shift-Left>",
                   "<Shift-Right>", "<Shift-Up>", "<Shift-Down>"):
            t.bind(ev, lambda e: "break")

    def guard_clipboard(self):
        """If a passage to be typed has been copied (from anywhere), empty the clipboard."""
        if not self.passages:
            return
        try:
            clip = self.app.clipboard_get()
        except Exception:
            return
        c = " ".join(clip.split())
        if len(c) < 40:
            return
        if any(p[i:i + 40] in c for p in self.passages for i in range(0, max(1, len(p) - 40), 40)):
            try:
                self.app.clipboard_clear()
                self.app.clipboard_append("")
            except Exception:
                pass
            self.monitor.log("test", "Copied passage removed from the clipboard (typing must be done by hand)")
            if not self.clip_warned:
                self.clip_warned = True
                Popup(self.app, "Copying not allowed", "The passage must be typed by hand.\n"
                      "Copy-paste of the given text is blocked.", danger=True, auto_close=6000)

    def hindi(self):
        path = self.folder / "Instructions_Hindi.html"
        try:
            build_instructions(self.tasks, self.title_, path)
            webbrowser.open(path.as_uri())
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"Could not open the Hindi instructions.\n{e}")

    # ---------- task display
    def pick(self, item):
        for it in self.items:
            it.select(it is item)
        self.cur = self.items.index(item)
        t = self.tasks[self.cur]
        self.task_menu.set(f"{self.cur + 1}. {APP_SHORT[t.app]} – {t.title}")
        self.show_task()

    def current(self):
        return self.tasks[self.cur]

    def move(self, d):
        i = self.cur + d
        if 0 <= i < len(self.tasks):
            self.pick(self.items[i])

    def show_task(self):
        t = self.current()
        self.h_app.configure(text=f"  {APP_LABEL[t.app]}  ", fg_color=APP_COL[t.app])
        self.h_title.configure(text=t.title)
        self.h_marks.configure(text=f"  {t.marks:g} marks  ")
        self.h_file.configure(text=f"Save as:  {t.file}")
        self.h_start.configure(text=("Starter file in the folder: " + ", ".join(t.starters)) if t.starters else
                               "Pictures for inserting are in the 'Pictures' folder.")
        render_instructions(self.text, t)
        self.b_starter.configure(state="normal" if t.starters else "disabled")
        if self.mini is not None:
            self.mini.show_task()

    def open_starter(self):
        for name in self.current().starters:
            open_path(self.folder / name)

    def hint(self):
        t = self.current()
        Popup(self.app, f"Model answer – {t.title}", lines=t.solution)

    def check_one(self):
        res = grade_task(self.current(), self.folder, self.known)
        ResultPopup(self.app, res)

    # ---------- timer & log
    def add_log(self, ev):
        self.log.configure(state="normal")
        self.log.insert("end", f"{ev['time']}  {APP_NAME.get(ev['app'], ev['app']):<10}  {ev['msg']}\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def tick(self):
        if self.done:
            return
        while True:
            try:
                self.add_log(self.events_q.get_nowait())
            except queue.Empty:
                break
        left = max(0, self.allowed - (time.monotonic() - self.started))
        if int(left) % 2 == 0:
            self.guard_clipboard()
        colour = (BAD if left <= min(300, self.allowed * 0.2) else
                  SAFFRON_DARK if left <= min(600, self.allowed * 0.4) else NAVY)
        self.timer.configure(text=fmt_secs(left), fg_color=colour)
        if self.mini is not None:
            self.mini.set_timer(fmt_secs(left), colour)
        for k, c in self.chips.items():
            c.configure(text=f"  {APP_SHORT[k]} {fmt_secs(self.monitor.app_seconds[k])}  ")
        for mins in (10, 5, 1):
            if left <= mins * 60 and mins not in self.warned and self.allowed > mins * 60 + 30:
                self.warned.add(mins)
                self.app.bell()
                self.monitor.log("test", f"{mins} minute(s) left")
                Popup(self.app, "Time warning", f"Only {mins} minute{'s' if mins > 1 else ''} left.\nSave your files now!",
                      danger=True, auto_close=8000)
        if left <= 0:
            self.monitor.log("test", "Time over - submitted automatically")
            self.finish(auto=True)
            return
        self.after(1000, self.tick)

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

    def finish(self, auto=False, quiet=False):
        if self.done:
            return
        self.done = True
        self.monitor.stop()
        if self.mini is not None:
            self.close_mini()
        used = min(time.monotonic() - self.started, self.allowed)
        wait = card(self, fg_color=NAVY)
        label(wait, "Checking your files…", 18, "bold", "white").pack(padx=40, pady=24)
        wait.place(relx=0.5, rely=0.5, anchor="center")
        self.update_idletasks()
        results = grade_tasks(self.tasks, self.folder)
        rec = make_record(self.mode, self.title_, self.folder, results, used, self.allowed,
                          self.monitor.app_seconds, self.monitor.events, spec=self.spec, attempt=self.attempt)
        try:
            save_attempt(rec)
        except Exception:
            pass
        report = self.folder / "Result_Report.html"
        try:
            build_report(rec, report)
        except Exception:
            report = None
        if quiet:
            return
        if auto:
            messagebox.showinfo(APP_TITLE, "Time is over. Your test has been submitted and checked.")
        self.app.attributes("-topmost", False)
        self.app.switch(ResultView, rec, report)


# =================================================================== popups

class Popup(ctk.CTkToplevel):
    def __init__(self, app, title, text=None, lines=None, danger=False, auto_close=None):
        super().__init__(app, fg_color=BG)
        self.title(title)
        self.attributes("-topmost", True)
        self.resizable(False, False)
        c = card(self)
        c.pack(fill="both", expand=True, padx=14, pady=14)
        label(c, title, 17, "bold", BAD if danger else NAVY).pack(anchor="w", padx=20, pady=(16, 6))
        if text:
            label(c, text, 15, "bold" if danger else "normal", justify="left").pack(anchor="w", padx=20, pady=6)
        if lines:
            t = textbox(c, height=min(60 + 34 * len(lines), 420), size=13)
            for i, s in enumerate(lines, 1):
                t.insert("end", f"{i}.  {s}\n\n")
            t.configure(state="disabled", width=560)
            t.pack(fill="both", expand=True, padx=12)
        btn(c, "OK", self.destroy, "accent" if danger else "primary", width=110).pack(pady=14)
        if auto_close:
            self.after(auto_close, lambda: self.winfo_exists() and self.destroy())


def render_instructions(tx, t):
    tx.configure(state="normal")
    tx.delete("1.0", "end")
    n = 0
    for line in t.instructions:
        if line.startswith("   "):
            tx.insert("end", line.strip() + "\n", "quote")
        else:
            n += 1
            tx.insert("end", f"{n}.  ", "num")
            tx.insert("end", line + "\n")
        tx.insert("end", "\n" if not line.startswith("   ") else "")
    tx.insert("end", "Save inside the test folder with the exact file name.", "hint")
    tx.configure(state="disabled")


class MiniBar(ctk.CTkToplevel):
    """Small always-on-top question box, so Office can stay full screen."""
    ALPHAS = (1.0, 0.85, 0.7)

    def __init__(self, view):
        super().__init__(view.app, fg_color=BG)
        self.view = view
        self.collapsed = False
        self.alpha_i = 0
        self.title("Questions")
        self.attributes("-topmost", True)
        if sys.platform == "win32":
            try:
                self.attributes("-toolwindow", True)          # thin title bar, no taskbar button
            except tk.TclError:
                pass
        scale = max(ctk.ScalingTracker.get_window_scaling(self), 0.1)
        sw = int(self.winfo_screenwidth() / scale)
        self.geometry(f"400x360+{max(0, sw - 420)}+40")
        self.minsize(300, 44)
        self.protocol("WM_DELETE_WINDOW", self.back_to_full)

        head = ctk.CTkFrame(self, fg_color=NAVY, corner_radius=0)
        head.pack(fill="x")
        self.timer = ctk.CTkLabel(head, text="", font=ctk.CTkFont(family=MONO, size=18, weight="bold"),
                                  text_color="white", fg_color=NAVY_DARK, corner_radius=8, width=78, height=30)
        self.timer.pack(side="left", padx=6, pady=6)
        self.task_lbl = ctk.CTkLabel(head, text="", font=F(12, "bold"), text_color="white", anchor="w")
        self.task_lbl.pack(side="left", fill="x", expand=True, padx=4)
        small = dict(width=30, height=28, corner_radius=8, font=F(13, "bold"), fg_color=NAVY_DARK,
                     hover_color="#06214F", text_color="white")
        ctk.CTkButton(head, text="⤢", command=self.back_to_full, **small).pack(side="right", padx=(2, 6))
        self.fold_b = ctk.CTkButton(head, text="▴", command=self.toggle, **small)
        self.fold_b.pack(side="right", padx=2)
        ctk.CTkButton(head, text="◐", command=self.see_through, **small).pack(side="right", padx=2)

        self.body = ctk.CTkFrame(self, fg_color=BG, corner_radius=0)
        self.body.pack(fill="both", expand=True)
        self.file_lbl = label(self.body, "", 12, "bold", NAVY, anchor="w")
        self.file_lbl.pack(fill="x", padx=10, pady=(6, 0))
        foot = ctk.CTkFrame(self.body, fg_color="transparent")
        foot.pack(side="bottom", fill="x", padx=8, pady=6)
        b = dict(height=30, font=F(12, "bold"))
        btn(foot, "◀", lambda: view.move(-1), "ghost", width=36, **b).pack(side="left")
        btn(foot, "▶", lambda: view.move(1), "ghost", width=36, **b).pack(side="left", padx=4)
        if view.allow_check:
            btn(foot, "Check", view.check_one, width=70, **b).pack(side="left", padx=4)
        btn(foot, "Submit", lambda: (view.close_mini(), view.submit()), "accent", width=76, **b).pack(side="right")
        self.text = textbox(self.body, size=12)
        self.text.pack(fill="both", expand=True, padx=6)
        self.text.tag_config("quote", foreground="#334", lmargin1=18, lmargin2=18, background="#EEF2FA")
        self.text.tag_config("num", foreground=NAVY)
        self.text.tag_config("hint", foreground="#7A5B00")
        view.block_copy(self.text)
        self.bind("<Left>", lambda e: view.move(-1))
        self.bind("<Right>", lambda e: view.move(1))
        self.show_task()

    def show_task(self):
        t = self.view.current()
        self.task_lbl.configure(text=f"{self.view.cur + 1}/{len(self.view.tasks)}  {APP_SHORT[t.app]} · {t.title}")
        self.file_lbl.configure(text=f"Save as: {t.file}" + (f"   (starter: {', '.join(t.starters)})" if t.starters else ""))
        render_instructions(self.text, t)

    def set_timer(self, text, colour):
        self.timer.configure(text=text, fg_color=NAVY_DARK if colour == NAVY else colour)
        if self.collapsed:
            self.title(f"{text} – Questions")

    def toggle(self):
        self.collapsed = not self.collapsed
        if self.collapsed:
            self._size = self.geometry().split("+")[0]
            self.body.pack_forget()
            w = self._size.split("x")[0]
            self.geometry(f"{w}x44")
            self.fold_b.configure(text="▾")
        else:
            self.body.pack(fill="both", expand=True)
            self.geometry(getattr(self, "_size", "400x360"))
            self.fold_b.configure(text="▴")

    def see_through(self):
        self.alpha_i = (self.alpha_i + 1) % len(self.ALPHAS)
        self.attributes("-alpha", self.ALPHAS[self.alpha_i])

    def back_to_full(self):
        self.view.close_mini()


def make_tree(parent):
    frame = ctk.CTkFrame(parent, fg_color=CARD, corner_radius=0)
    tree = ttk.Treeview(frame, columns=("marks",), show="tree headings", style="Nice.Treeview")
    tree.heading("#0", text="  Task / check", anchor="w")
    tree.heading("marks", text="Marks")
    tree.column("#0", width=240, minwidth=160, stretch=True)
    tree.column("marks", width=90, anchor="center", stretch=False)
    sb = ctk.CTkScrollbar(frame, command=tree.yview, button_color="#C9D3E6", button_hover_color="#AEB9CE")
    tree.configure(yscrollcommand=sb.set)
    tree.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    tree.tag_configure("task", font=(FAMILY, 11, "bold"), background="#F6F8FC")
    tree.tag_configure("ok", foreground=OK)
    tree.tag_configure("bad", foreground=BAD)
    tree.tag_configure("part", foreground=SAFFRON_DARK)
    tree.tag_configure("note", foreground="#7A5B00")
    return frame, tree


def fill_result_tree(tree, results):
    meta = {}
    for t in results:
        pid = tree.insert("", "end", text=f"  {APP_SHORT[t['app']]}  ·  {t['title']}", values=(f"{t['got']:g} / {t['marks']:g}",),
                          open=True, tags=("task",))
        meta[pid] = ("task", t)
        for c in t["checks"]:
            if c["desc"] == "Note":
                iid = tree.insert(pid, "end", text=f"   ⓘ  {c['found']}", values=("",), tags=("note",))
            else:
                kind = "ok" if c["ok"] else "part" if c.get("got") else "bad"
                icon = {"ok": "   ✔  ", "part": "   ◐  ", "bad": "   ✘  "}[kind]
                iid = tree.insert(pid, "end", text=icon + c["desc"], values=(f"{c['got']:g} / {c['marks']:g}",), tags=(kind,))
            meta[iid] = ("check", c)
    return meta


def make_detail(parent):
    t = textbox(parent, size=13)
    t.tag_config("h", foreground=INK)
    t.tag_config("b", foreground=NAVY)
    t.tag_config("ok", foreground=OK)
    t.tag_config("bad", foreground=BAD)
    t.tag_config("part", foreground=SAFFRON_DARK)
    t.tag_config("how", background="#FFF6E0", lmargin1=8, lmargin2=8)
    t.tag_config("muted", foreground=MUTED)
    t.configure(state="disabled")
    return t


def show_detail(text, kind, data):
    text.configure(state="normal")
    text.delete("1.0", "end")
    if kind == "task":
        text.insert("end", f"{data['title']}\n", "h")
        text.insert("end", f"{data['got']:g} / {data['marks']:g} marks  ·  file {data['file']}\n\n", "muted")
        text.insert("end", "MODEL ANSWER (step by step)\n\n", "b")
        for i, s in enumerate(data.get("solution") or [], 1):
            text.insert("end", f"{i}.  {s}\n\n")
    else:
        c = data
        text.insert("end", c["desc"] + "\n", "h")
        state = "✔ Correct" if c["ok"] else "◐ Part marks" if c.get("got") else "✘ Not correct"
        text.insert("end", state + f"   ({c['got']:g} / {c['marks']:g} marks)\n\n",
                    "ok" if c["ok"] else "part" if c.get("got") else "bad")
        text.insert("end", "WHAT WE FOUND\n", "b")
        text.insert("end", c["found"] + "\n\n")
        if c.get("expected"):
            text.insert("end", "EXPECTED\n", "b")
            text.insert("end", c["expected"] + "\n\n")
        if c.get("how"):
            text.insert("end", "HOW TO DO IT\n", "b")
            text.insert("end", c["how"] + "\n", "how")
        if c.get("topic"):
            text.insert("end", f"\nSyllabus topic: {c['topic']}\n", "muted")
    text.configure(state="disabled")


class ResultPopup(ctk.CTkToplevel):
    """Instant check of one task (practice and drills)."""
    def __init__(self, app, res):
        super().__init__(app, fg_color=BG)
        self.title(f"Check: {res.title}")
        self.geometry("900x600")
        self.attributes("-topmost", True)
        full = res.got >= res.marks
        head = card(self, fg_color=OK_SOFT if full else CARD)
        head.pack(fill="x", padx=14, pady=(14, 8))
        label(head, res.title, 17, "bold").pack(side="left", padx=18, pady=12)
        label(head, f"{res.got:g} / {res.marks:g}", 20, "bold", OK if full else NAVY).pack(side="right", padx=18)
        body = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, padx=14)
        body.grid_columnconfigure(0, weight=3)
        body.grid_columnconfigure(1, weight=2)
        body.grid_rowconfigure(0, weight=1)
        c1 = card(body)
        c1.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        fr, tree = make_tree(c1)
        fr.pack(fill="both", expand=True, padx=8, pady=8)
        c2 = card(body)
        c2.grid(row=0, column=1, sticky="nsew")
        detail = make_detail(c2)
        detail.pack(fill="both", expand=True, padx=8, pady=8)
        meta = fill_result_tree(tree, [res.to_dict()])
        tree.bind("<<TreeviewSelect>>", lambda e: tree.selection() and tree.selection()[0] in meta and
                  show_detail(detail, *meta[tree.selection()[0]]))
        first = next((i for i, (k, d) in meta.items() if k == "check" and not d["ok"]), None)
        if first:
            tree.selection_set(first)
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(side="bottom", fill="x", padx=14, pady=12, before=body)
        label(bar, "Fix the mistakes, save, and press 'Check task' again.", 12, color=MUTED).pack(side="left")
        btn(bar, "Close", self.destroy, width=100).pack(side="right")


# =================================================================== results

def verdict(pct):
    if pct >= 85:
        return "Excellent – exam ready!", OK
    if pct >= 70:
        return "Good – polish the weak points", NAVY
    if pct >= 50:
        return "Fair – keep practising", SAFFRON_DARK
    return "Needs more practice", BAD


class ResultView(ctk.CTkFrame):
    def __init__(self, parent, app: App, rec, report):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.rec = rec
        pct = rec["score"] / rec["out_of"] * 100 if rec["out_of"] else 0
        vtext, vcol = verdict(pct)

        hero = card(self)
        hero.pack(fill="x")
        left = self.h_left = ctk.CTkFrame(hero, fg_color="transparent")
        row = ctk.CTkFrame(left, fg_color="transparent")
        row.pack(anchor="w")
        label(row, rec["title"], 15, "bold", MUTED).pack(side="left")
        if rec.get("attempt"):
            pill(row, f"Attempt {rec['attempt']}", NAVY_SOFT, NAVY, size=11).pack(side="left", padx=8)
        self.big = label(left, f"{rec['score']:g} / {rec['out_of']:g}", 40, "bold", NAVY)
        self.big.pack(anchor="w")
        label(left, f"{pct:.0f}%  ·  {vtext}", 15, "bold", vcol).pack(anchor="w")
        label(left, f"Time used {fmt_secs(rec['seconds_used'])} of {fmt_secs(rec['seconds_allowed'])}", 12, color=MUTED).pack(anchor="w", pady=(4, 0))

        right = self.h_right = ctk.CTkFrame(hero, fg_color="transparent")
        self.bars = []
        for k in ("word", "excel", "ppt"):
            if k not in rec["sections"]:
                continue
            g, m = rec["sections"][k]
            r = ctk.CTkFrame(right, fg_color="transparent")
            r.pack(fill="x", pady=4)
            label(r, APP_SHORT[k], 13, "bold", APP_COL[k], width=95, anchor="w").pack(side="left")
            pb = ctk.CTkProgressBar(r, width=240, height=12, corner_radius=6, progress_color=APP_COL[k], fg_color="#E6EAF2")
            self.bars.append(pb)
            pb.set(g / m if m else 0)
            pb.pack(side="left", padx=8)
            label(r, f"{g:g} / {m:g}", 13, "bold", width=70).pack(side="left")
            t = rec.get("app_seconds", {}).get(k)
            label(r, f"⏱ {fmt_secs(t or 0)}", 12, color=MUTED).pack(side="left", padx=6)

        self.tip = label(self, "Click any line to see what was found, what was expected and how to do it. "
                               "Click a task name for its model answer.", 12, color=MUTED, wraplength=900, justify="left")
        self.tip.pack(anchor="w", pady=(10, 6))

        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.pack(side="bottom", fill="x", pady=(12, 0))
        body = self.body_ = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True)
        c1 = self.c1 = card(body)
        fr, self.tree = make_tree(c1)
        fr.pack(fill="both", expand=True, padx=8, pady=8)
        c2 = self.c2 = card(body)
        self.detail = make_detail(c2)
        self.detail.pack(fill="both", expand=True, padx=8, pady=8)
        self.meta = fill_result_tree(self.tree, rec["results"])
        self.tree.bind("<<TreeviewSelect>>", self.on_select)

        self.bar_a = ctk.CTkFrame(bar, fg_color="transparent")
        self.bar_b = ctk.CTkFrame(bar, fg_color="transparent")
        self.bar_c = ctk.CTkFrame(bar, fg_color="transparent")
        btn(self.bar_a, "← Home", app.show_home, "ghost", width=90).pack(side="left")
        btn(self.bar_a, "Report wrong marking", self.report_problem, "ghost", width=170).pack(side="left", padx=6)
        if rec.get("folder"):
            btn(self.bar_c, "Open folder", lambda: open_path(rec["folder"]), "ghost", width=104).pack(side="left")
        if report:
            btn(self.bar_c, "Printable report", lambda: webbrowser.open(Path(report).as_uri()), "ghost",
                width=140).pack(side="left", padx=6)
        if rec.get("spec"):
            btn(self.bar_b, "Retake (new numbers)", lambda: app.start(rec["spec"]), "accent", width=170).pack(side="right")
            if rec.get("folder") and Path(rec["folder"]).exists() and rec["score"] < rec["out_of"]:
                btn(self.bar_b, "Fix & recheck", lambda: app.start(rec["spec"], recheck=rec), width=130).pack(side="right", padx=6)

        self.compact = None
        self.layout(False)
        self.bind("<Configure>", lambda e: responsive(self, 1000))

        first_bad = next((i for i, (k, d) in self.meta.items() if k == "check" and not d["ok"] and d["desc"] != "Note"), None)
        if first_bad:
            self.tree.selection_set(first_bad)
            self.tree.see(first_bad)

    def on_select(self, _):
        sel = self.tree.selection()
        if sel and sel[0] in self.meta:
            show_detail(self.detail, *self.meta[sel[0]])

    def layout(self, compact):
        self.compact = compact
        for w in (self.h_left, self.h_right, self.bar_a, self.bar_b, self.bar_c):
            w.pack_forget()
        self.c1.grid_forget()
        self.c2.grid_forget()
        b = self.body_
        if compact:
            self.h_left.pack(side="top", anchor="w", padx=18, pady=(12, 4))
            self.h_right.pack(side="top", fill="x", padx=18, pady=(0, 12))
            self.big.configure(font=F(30, "bold"))
            for pb in self.bars:
                pb.configure(width=140)
            b.grid_columnconfigure(0, weight=1)
            b.grid_columnconfigure(1, weight=0)
            b.grid_rowconfigure(0, weight=3)
            b.grid_rowconfigure(1, weight=2)
            self.c1.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
            self.c2.grid(row=1, column=0, sticky="nsew")
            self.tip.configure(wraplength=380)
            self.bar_b.pack(side="top", fill="x")
            self.bar_a.pack(side="top", fill="x", pady=(8, 0))
            self.bar_c.pack(side="top", fill="x", pady=(8, 0))
        else:
            self.h_left.pack(side="left", padx=24, pady=16)
            self.h_right.pack(side="right", padx=24, pady=16, fill="y")
            self.big.configure(font=F(40, "bold"))
            for pb in self.bars:
                pb.configure(width=240)
            b.grid_columnconfigure(0, weight=3)
            b.grid_columnconfigure(1, weight=2)
            b.grid_rowconfigure(0, weight=1)
            b.grid_rowconfigure(1, weight=0)
            self.c1.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
            self.c2.grid(row=0, column=1, sticky="nsew")
            self.tip.configure(wraplength=900)
            self.bar_a.pack(side="left")
            self.bar_c.pack(side="left", padx=6)
            self.bar_b.pack(side="right")

    def report_problem(self):
        sel = self.tree.selection()
        hint = ""
        if sel and sel[0] in self.meta and self.meta[sel[0]][0] == "check":
            hint = self.meta[sel[0]][1]["desc"]
        note = ctk.CTkInputDialog(title="Report wrong marking",
                                  text="Which point was marked wrongly, and what did you do?\n"
                                       + (f"(Selected: {hint})" if hint else "(Tip: select the line first)")).get_input()
        if note is None:
            return
        zp = make_problem_report(self.rec, f"Selected check: {hint}\n\n{note}")
        messagebox.showinfo(APP_TITLE, f"A report file was made:\n\n{zp}\n\nPlease send this zip file (by WhatsApp or "
                                       "e-mail) to the person who gave you the app, so the marking can be fixed.")
        open_path(zp.parent)


# =================================================================== quiz

class QuizView(ctk.CTkFrame):
    def __init__(self, parent, app: App, questions, quiz_app=None):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        self.qs = questions
        self.quiz_app = quiz_app
        self.i = 0
        self.score = 0
        self.answers = []
        self.picked = {}

        self.box = card(self)
        self.box.place(relx=0.5, rely=0.47, anchor="center", relwidth=0.62, relheight=0.94)
        self.compact = None
        self.bind("<Configure>", lambda e: responsive(self))
        top = ctk.CTkFrame(self.box, fg_color="transparent")
        top.pack(fill="x", padx=28, pady=(22, 4))
        label(top, "Shortcut Key Quiz", 20, "bold").pack(side="left")
        self.scorelbl = pill(top, "", NAVY_SOFT, NAVY)
        self.scorelbl.pack(side="right")
        self.progress = ctk.CTkProgressBar(self.box, height=8, corner_radius=4, progress_color=SAFFRON, fg_color="#E6EAF2")
        self.progress.pack(fill="x", padx=28, pady=(6, 4))
        self.count = label(self.box, "", 12, color=MUTED)
        self.count.pack(anchor="w", padx=28)
        self.tag = pill(self.box, "", NAVY, size=11)
        self.tag.pack(anchor="w", padx=28, pady=(16, 4))
        self.q = label(self.box, "", 22, "bold", wraplength=620, justify="left")
        self.q.pack(anchor="w", padx=28, pady=(0, 16))
        self.opts = []
        for _ in range(4):
            b = ctk.CTkButton(self.box, text="", height=44, corner_radius=12, font=F(15, "bold"), anchor="w",
                              fg_color=CARD, hover_color=NAVY_SOFT, text_color=INK, border_width=1, border_color="#C9D3E6")
            b.pack(fill="x", padx=28, pady=5)
            self.opts.append(b)
        self.feedback = label(self.box, "", 15, "bold")
        self.feedback.pack(anchor="w", padx=28, pady=10)
        bar = ctk.CTkFrame(self.box, fg_color="transparent")
        bar.pack(side="bottom", fill="x", padx=28, pady=16, before=top)     # reserve room for the buttons first
        btn(bar, "← Home", app.show_home, "ghost", width=110).pack(side="left")
        self.next_b = btn(bar, "Next →", self.next, width=110)
        self.next_b.pack(side="right")
        self.prev_b = btn(bar, "← Back", self.prev, "ghost", width=90)
        self.prev_b.pack(side="right", padx=8)
        label(bar, "Keys: A–D answer · ← → move · Enter next", 11, color=MUTED).pack(side="left", padx=12)
        # keyboard: A-D or 1-4 to answer, Enter for the next question
        for keys, i in ((("a", "A", "1"), 0), (("b", "B", "2"), 1), (("c", "C", "3"), 2), (("d", "D", "4"), 3)):
            for k in keys:
                app.bind(f"<KeyPress-{k}>", lambda e, i=i: self.key_pick(i))
        app.bind("<Return>", lambda e: self.opts and self.next())
        app.bind("<Right>", lambda e: self.opts and self.next())
        app.bind("<Left>", lambda e: self.prev())
        self.show()

    def layout(self, compact):
        self.compact = compact
        self.box.place_configure(relwidth=0.98 if compact else 0.62)

    def show(self):
        q = self.qs[self.i]
        self.progress.set(self.i / len(self.qs))
        self.count.configure(text=f"Question {self.i + 1} of {len(self.qs)}")
        self.scorelbl.configure(text=f"  Score {self.score}  ")
        self.tag.configure(text=f"  {'All programs' if q['app'] == 'All' else q['app']}  ", fg_color=APP_COL.get({"Word": "word", "Excel": "excel",
                                                                           "PowerPoint": "ppt"}.get(q["app"], ""), NAVY))
        self.q.configure(text=q["q"].split("] ", 1)[-1])
        for b, opt, letter in zip(self.opts, q["options"], "ABCD"):
            b.configure(text=f"   {letter}.   {opt}", state="normal", fg_color=CARD, border_color="#C9D3E6",
                        text_color=INK, command=lambda o=opt: self.pick(o))
        self.feedback.configure(text="")
        self.next_b.configure(state="disabled")
        self.prev_b.configure(state="normal" if self.i > 0 else "disabled")
        if self.i in self.picked:                     # going back: show the earlier answer again
            self.mark(q, self.picked[self.i])

    def prev(self):
        if self.i > 0 and self.opts:
            self.i -= 1
            self.show()

    def key_pick(self, i):
        if self.winfo_exists() and self.i < len(self.qs) and self.opts and self.opts[0].winfo_exists() \
                and self.opts[i].cget("state") == "normal":
            self.pick(self.qs[self.i]["options"][i])

    def destroy(self):
        for k in ("a", "A", "1", "b", "B", "2", "c", "C", "3", "d", "D", "4"):
            self.app.unbind(f"<KeyPress-{k}>")
        for k in ("<Return>", "<Right>", "<Left>"):
            self.app.unbind(k)
        super().destroy()

    def pick(self, opt):
        q = self.qs[self.i]
        if self.i in self.picked:
            return
        self.picked[self.i] = opt
        ok = opt == q["answer"]
        self.score += ok
        self.answers.append((q, opt, ok))
        self.mark(q, opt)

    def mark(self, q, opt):
        ok = opt == q["answer"]
        for b, o in zip(self.opts, q["options"]):
            b.configure(state="disabled", text_color_disabled=INK)
            if o == q["answer"]:
                b.configure(fg_color=OK_SOFT, border_color=OK)
            elif o == opt:
                b.configure(fg_color=BAD_SOFT, border_color=BAD)
        self.feedback.configure(text="✔ Correct!" if ok else f"✘ Not correct. The answer is {q['answer']}.",
                                text_color=OK if ok else BAD)
        self.scorelbl.configure(text=f"  Score {self.score}  ")
        self.next_b.configure(state="normal")

    def next(self):
        if self.i not in self.picked:
            return
        self.i += 1
        if self.i < len(self.qs):
            self.show()
            return
        for w in self.box.winfo_children():
            w.destroy()
        pct = self.score / len(self.qs) * 100
        vtext, vcol = verdict(pct)
        self.opts = []
        bar = ctk.CTkFrame(self.box, fg_color="transparent")
        bar.pack(side="bottom", fill="x", padx=28, pady=(0, 20))
        btn(bar, "← Home", self.app.show_home, "ghost", width=110).pack(side="left")
        btn(bar, "New quiz", lambda: self.app.switch(QuizView, make_quiz(20, self.quiz_app), self.quiz_app), "accent",
            width=130).pack(side="right")
        label(self.box, "Quiz finished", 15, "bold", MUTED).pack(anchor="w", padx=28, pady=(22, 0))
        label(self.box, f"{self.score} / {len(self.qs)}", 40, "bold", NAVY).pack(anchor="w", padx=28)
        label(self.box, vtext, 15, "bold", vcol).pack(anchor="w", padx=28)
        t = textbox(self.box, size=13)
        t.tag_config("bad", foreground=BAD)
        t.tag_config("ok", foreground=OK)
        for q, opt, ok in self.answers:
            t.insert("end", ("✔ " if ok else "✘ ") + q["q"] + "  →  " + q["answer"] +
                     ("" if ok else f"   (you chose {opt})") + "\n", "ok" if ok else "bad")
        t.configure(state="disabled")
        t.pack(fill="both", expand=True, padx=20, pady=12)
        try:
            save_attempt({"when": time.strftime("%Y-%m-%dT%H:%M:%S"), "mode": "Shortcut quiz", "title": "Shortcut Key Quiz",
                          "folder": "", "score": self.score, "out_of": len(self.qs), "sections": {}, "seconds_used": 0,
                          "seconds_allowed": 0, "app_seconds": {}, "results": [], "events": []})
        except Exception:
            pass


# =================================================================== history

class HistoryView(ctk.CTkFrame):
    def __init__(self, parent, app: App):
        super().__init__(parent, fg_color="transparent")
        self.app = app
        top = ctk.CTkFrame(self, fg_color="transparent")
        top.pack(fill="x")
        label(top, "My Progress", 22, "bold").pack(side="left")
        pill(top, current_student(), NAVY_SOFT, NAVY).pack(side="left", padx=10)
        btn(top, "← Home", app.show_home, "ghost", width=90).pack(side="right")
        btn(top, "Export to Excel", self.export, "ghost", width=140).pack(side="right", padx=8)

        body = self.body_ = ctk.CTkFrame(self, fg_color="transparent")
        body.pack(fill="both", expand=True, pady=(10, 0))

        c1 = self.c1 = card(body)
        label(c1, "Score trend", 15, "bold").pack(anchor="w", padx=16, pady=(14, 0))
        self.chart = tk.Canvas(c1, height=130, bg=CARD, highlightthickness=0)
        self.chart.pack(fill="x", padx=16, pady=(4, 6))
        self.chart.bind("<Configure>", lambda e: self.draw_chart())
        label(c1, "Past attempts", 15, "bold").pack(anchor="w", padx=16, pady=(4, 6))
        bar = ctk.CTkFrame(c1, fg_color="transparent")
        bar.pack(side="bottom", fill="x", padx=12, pady=12)
        btn(bar, "Open result", self.open, "ghost", width=120).pack(side="left")
        btn(bar, "Retake", self.retake, "accent", width=100).pack(side="left", padx=8)
        btn(bar, "Delete attempt", self.delete, "danger", width=140).pack(side="right")
        frame = ctk.CTkFrame(c1, fg_color=CARD)
        frame.pack(fill="both", expand=True, padx=8)
        cols = ("when", "test", "att", "score", "word", "excel", "ppt")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", style="Nice.Treeview")
        for c, text, w, anchor in (("when", "Date", 150, "w"), ("test", "Test", 230, "w"), ("att", "Try", 50, "center"),
                                   ("score", "Score", 80, "center"), ("word", "Word", 70, "center"),
                                   ("excel", "Excel", 70, "center"), ("ppt", "PPT", 70, "center")):
            self.tree.heading(c, text=text, anchor=anchor)
            self.tree.column(c, width=w, anchor=anchor, stretch=c == "test")
        sb = ctk.CTkScrollbar(frame, command=self.tree.yview, button_color="#C9D3E6")
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.tree.bind("<Double-1>", lambda e: self.open())
        self.load()

        c2 = self.c2 = card(body)
        label(c2, "Weak areas", 15, "bold").pack(anchor="w", padx=16, pady=(14, 0))
        label(c2, "Topics you get wrong most often", 12, color=MUTED).pack(anchor="w", padx=16, pady=(0, 8))
        wl = ctk.CTkScrollableFrame(c2, fg_color="transparent")
        wl.pack(fill="both", expand=True, padx=6, pady=(0, 10))
        weak = weak_topics(load_history(), limit=15)
        if not weak:
            label(wl, "Take a test first – your weak areas will show here.", 13, color=MUTED).pack(anchor="w", padx=10)
        for app_, topic, ok, n, rate in weak:
            r = ctk.CTkFrame(wl, fg_color="transparent")
            r.pack(fill="x", pady=5, padx=6)
            hrow = ctk.CTkFrame(r, fg_color="transparent")
            hrow.pack(fill="x")
            name = topic[0].upper() + topic[1:]
            label(hrow, name if len(name) <= 26 else name[:25] + "…", 13, "bold").pack(side="left")
            label(hrow, f"{ok}/{n}", 12, "bold", MUTED).pack(side="right")
            srow = ctk.CTkFrame(r, fg_color="transparent")
            srow.pack(fill="x")
            label(srow, APP_SHORT.get(app_, app_), 11, color=APP_COL.get(app_, MUTED)).pack(side="left")
            drills = practice_drills(app_, topic)
            if drills:
                ctk.CTkButton(srow, text="Practise →", width=86, height=22, corner_radius=8, font=F(11, "bold"),
                              fg_color=NAVY_SOFT, hover_color="#D5DDEC", text_color=NAVY,
                              command=lambda d=drills: app.start({"kind": "drills", "ids": [x.id for x in d]})
                              ).pack(side="right")
            pb = ctk.CTkProgressBar(r, height=8, corner_radius=4, fg_color="#E6EAF2",
                                    progress_color=BAD if rate < 0.5 else SAFFRON if rate < 0.8 else OK)
            pb.set(rate)
            pb.pack(fill="x", pady=(3, 0))
        self.compact = None
        self.layout(False)
        self.bind("<Configure>", lambda e: responsive(self))

    def layout(self, compact):
        self.compact = compact
        b = self.body_
        self.c1.grid_forget()
        self.c2.grid_forget()
        if compact:
            b.grid_columnconfigure(0, weight=1)
            b.grid_columnconfigure(1, weight=0)
            b.grid_rowconfigure(0, weight=3)
            b.grid_rowconfigure(1, weight=2)
            self.c1.grid(row=0, column=0, sticky="nsew", pady=(0, 8))
            self.c2.grid(row=1, column=0, sticky="nsew")
            self.tree["displaycolumns"] = ("when", "test", "score")
        else:
            b.grid_columnconfigure(0, weight=3)
            b.grid_columnconfigure(1, weight=2)
            b.grid_rowconfigure(0, weight=1)
            b.grid_rowconfigure(1, weight=0)
            self.c1.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
            self.c2.grid(row=0, column=1, sticky="nsew")
            self.tree["displaycolumns"] = "#all"

    def draw_chart(self):
        c = self.chart
        c.delete("all")
        pts = [r["score"] / r["out_of"] * 100 for r in load_history() if r.get("out_of") and r.get("mode") != "Shortcut quiz"][-20:]
        w, h = max(c.winfo_width(), 200), 130
        left, right, top, bottom = 42, 12, 16, 22
        for v in (0, 50, 100):
            y = top + (h - top - bottom) * (1 - v / 100)
            c.create_line(left, y, w - right, y, fill=LINE, dash=(2, 3) if v == 50 else None)
            c.create_text(left - 6, y, text=f"{v}%", anchor="e", fill=MUTED, font=(FAMILY, 9))
        if not pts:
            c.create_text(w / 2, h / 2, text="Your scores will appear here after your first test", fill=MUTED,
                          font=(FAMILY, 11))
            return
        step = (w - left - right) / max(1, len(pts) - 1) if len(pts) > 1 else 0
        xy = [(left + (i * step if len(pts) > 1 else (w - left - right) / 2), top + (h - top - bottom) * (1 - p / 100))
              for i, p in enumerate(pts)]
        if len(xy) > 1:
            c.create_line(*[v for p in xy for v in p], fill=NAVY, width=3)
        for i, ((x, y), p) in enumerate(zip(xy, pts)):
            c.create_oval(x - 5, y - 5, x + 5, y + 5, fill=SAFFRON, outline=CARD, width=2)
            if i == len(pts) - 1 or len(pts) <= 8:
                c.create_text(x, y - 13, text=f"{p:.0f}", fill=INK, font=(FAMILY, 9, "bold"))
        c.create_text(left, h - 6, text="older", anchor="w", fill=MUTED, font=(FAMILY, 9))
        c.create_text(w - right, h - 6, text="latest", anchor="e", fill=MUTED, font=(FAMILY, 9))

    def export(self):
        path = filedialog.asksaveasfilename(title="Save results as", defaultextension=".xlsx",
                                            initialdir=str(default_root()), initialfile="ISRO_SkillTest_Results.xlsx",
                                            filetypes=[("Excel workbook", "*.xlsx")])
        if not path:
            return
        try:
            export_results(path)
        except PermissionError:
            messagebox.showerror(APP_TITLE, "Could not save – is that file open in Excel? Close it and try again.")
            return
        open_path(path)

    def load(self):
        self.tree.delete(*self.tree.get_children())
        self.rows = {}
        for rec in reversed(load_history()):
            s = rec.get("sections", {})
            fmt = lambda k: f"{s[k][0]:g}/{s[k][1]:g}" if k in s else "–"
            w = rec["when"]
            when = f"{w[8:10]}-{w[5:7]}-{w[0:4]} {w[11:16]}"
            name = rec["title"] + (" (recheck)" if rec.get("mode") == "Fix & recheck" else "")
            iid = self.tree.insert("", "end", values=(when, name,
                                                      rec.get("attempt", "–"), f"{rec['score']:g}/{rec['out_of']:g}",
                                                      fmt("word"), fmt("excel"), fmt("ppt")))
            self.rows[iid] = rec

    def selected(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo(APP_TITLE, "Select an attempt in the list first.")
            return None
        return self.rows[sel[0]]

    def open(self):
        rec = self.selected()
        if rec and rec.get("results"):
            self.app.switch(ResultView, rec, Path(rec["folder"]) / "Result_Report.html"
                            if rec.get("folder") and (Path(rec["folder"]) / "Result_Report.html").exists() else None)

    def retake(self):
        rec = self.selected()
        if not rec:
            return
        if rec.get("mode") == "Shortcut quiz":
            self.app.switch(QuizView, make_quiz(20))
        else:
            self.app.start(rec.get("spec") or {})

    def delete(self):
        rec = self.selected()
        if not rec:
            return
        has_folder = rec.get("folder") and Path(rec["folder"]).exists()
        msg = f"Delete '{rec['title']}' (attempt {rec.get('attempt', '')}) from your history?"
        if has_folder:
            msg += f"\n\nIts folder and files will also be deleted:\n{rec['folder']}"
        if messagebox.askyesno(APP_TITLE, msg, icon="warning"):
            delete_attempt(rec)
            self.app.switch(HistoryView)


def practice_drills(app, topic):
    """Drills that practise a weak topic (at most 3)."""
    found = [d for d in DRILLS if d.app == app and any(c.topic == topic for c in d.checks)]
    if not found:
        words = {w for w in topic.lower().replace("&", " ").split() if len(w) > 3}
        found = [d for d in DRILLS if d.app == app and words & set(d.topic.lower().replace("&", " ").split())]
    return found[:3]


def main():
    app = App()
    app.mainloop()
