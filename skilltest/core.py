"""Core data model: tasks, checks, file lookup and marking."""
from __future__ import annotations

import os
import re
import traceback
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional

APP_LABEL = {"word": "MS Word", "excel": "MS Excel", "ppt": "MS PowerPoint"}
EXT_KIND = {".docx": "word", ".xlsx": "excel", ".pptx": "ppt"}
OLD_EXT = {".docx": ".doc", ".xlsx": ".xls", ".pptx": ".ppt"}


class Missing(Exception):
    """Raised by a check when the file it needs is not there."""


@dataclass
class Check:
    desc: str                                  # what is being checked
    marks: float
    fn: Callable[["Ctx"], tuple]               # returns (ok, found_text)
    expected: str = ""                         # what the correct result looks like
    how: str = ""                              # steps to do it in Office
    topic: str = ""                            # syllabus topic, used for weak-area tracking


@dataclass
class Task:
    id: str
    app: str                                   # word / excel / ppt
    title: str
    file: str                                  # main file the student must save
    instructions: list
    checks: list
    solution: list = field(default_factory=list)   # model answer, step by step
    starters: dict = field(default_factory=dict)   # filename -> fn(path) that creates it
    minutes: int = 10                          # suggested time (used for drills/practice)

    @property
    def marks(self) -> float:
        return round(sum(c.marks for c in self.checks), 2)


@dataclass
class Paper:
    id: str
    title: str
    tasks: list
    minutes: int = 75

    def section(self, app):
        return [t for t in self.tasks if t.app == app]


# ---------------------------------------------------------------- results

@dataclass
class CheckResult:
    desc: str
    marks: float
    got: float
    ok: bool
    found: str
    expected: str
    how: str
    topic: str

    def to_dict(self):
        return self.__dict__.copy()


@dataclass
class TaskResult:
    id: str
    app: str
    title: str
    file: str
    marks: float
    got: float
    checks: list
    solution: list

    def to_dict(self):
        d = self.__dict__.copy()
        d["checks"] = [c.to_dict() for c in self.checks]
        return d


# ---------------------------------------------------------------- file lookup

def _is_temp(p: Path) -> bool:
    return p.name.startswith("~$") or p.name.startswith(".~") or p.suffix.lower() == ".tmp"


def list_files(folder: Path):
    if not folder.exists():
        return []
    return [p for p in folder.rglob("*") if p.is_file() and not _is_temp(p)]


class Ctx:
    """Gives checks access to the student's files for one task."""

    def __init__(self, folder: Path, task: Task, known_names=()):
        self.folder = Path(folder)
        self.task = task
        self.known = {n.lower() for n in known_names}
        self._cache = {}
        self.notes = []

    def exact(self, name: str) -> Optional[Path]:
        want = name.lower()
        for p in list_files(self.folder):
            if p.name.lower() == want:
                return p
        return None

    def path(self, name: Optional[str] = None) -> Path:
        name = name or self.task.file
        p = self.exact(name)
        if p:
            return p
        stem, ext = os.path.splitext(name)
        files = list_files(self.folder)
        old = OLD_EXT.get(ext.lower())
        if old:
            for f in files:
                if f.stem.lower() == stem.lower() and f.suffix.lower() == old:
                    raise Missing(f"Found {f.name} in the old 97-2003 format. Save it again as {ext} "
                                  f"(File > Save As > choose the {ext} type).")
        # wrong name: accept a single unclaimed file of the same type
        cands = [f for f in files if f.suffix.lower() == ext.lower() and f.name.lower() not in self.known]
        if len(cands) == 1:
            note = f"{name} not found; checked {cands[0].name} instead."
            if note not in self.notes:
                self.notes.append(note)
            return cands[0]
        raise Missing(f"File {name} was not found in the test folder.")

    def _load(self, kind, name):
        p = self.path(name)
        key = (kind, str(p))
        if key not in self._cache:
            if kind == "word":
                from .word import WordDoc
                self._cache[key] = WordDoc(p)
            elif kind == "excel":
                from .excel import Book
                self._cache[key] = Book(p)
            else:
                from .ppt import Deck
                self._cache[key] = Deck(p)
        return self._cache[key]

    def word(self, name=None):
        return self._load("word", name)

    def excel(self, name=None):
        return self._load("excel", name)

    def ppt(self, name=None):
        return self._load("ppt", name)


def grade_task(task: Task, folder: Path, known_names=()) -> TaskResult:
    ctx = Ctx(folder, task, known_names)
    out = []
    for c in task.checks:
        try:
            ok, found = c.fn(ctx)
        except Missing as e:
            ok, found = False, str(e)
        except Exception as e:  # a broken or unusual file should never crash the app
            ok, found = False, f"Could not read this part of the file ({type(e).__name__}: {e})"
            if os.environ.get("SKILLTEST_DEBUG"):
                traceback.print_exc()
        out.append(CheckResult(c.desc, c.marks, c.marks if ok else 0.0, bool(ok), str(found),
                               c.expected, c.how, c.topic))
    if ctx.notes:
        out.insert(0, CheckResult("Note", 0, 0, False, " ".join(ctx.notes), "", "", ""))
    got = round(sum(r.got for r in out), 2)
    return TaskResult(task.id, task.app, task.title, task.file, task.marks, got, out, task.solution)


FILE_RE = re.compile(r"[\w\-]+\.(?:docx|dotx|xlsx|xltx|pptx|potx|ppsx|doc|xls|ppt|pdf|rtf|txt|csv|html?)\b", re.I)


def all_file_names(tasks):
    """Every file name a paper mentions, so a wrongly named file is never mixed up with another task's file."""
    names = set()
    for t in tasks:
        names.add(t.file)
        names.update(t.starters.keys())
        for line in t.instructions:
            names.update(FILE_RE.findall(line))
    return names


def grade_tasks(tasks, folder):
    known = all_file_names(tasks)
    return [grade_task(t, folder, known) for t in tasks]


# ---------------------------------------------------------------- small text helpers

_QUOTES = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"',
                         "–": "-", "—": "-", " ": " "})


def norm(s: str) -> str:
    return re.sub(r"\s+", " ", (s or "").translate(_QUOTES)).strip().lower()


def similarity(a: str, b: str) -> float:
    from difflib import SequenceMatcher
    a, b = norm(a), norm(b)
    if not a and not b:
        return 1.0
    return SequenceMatcher(None, a, b, autojunk=False).ratio()


def short(s, n=80):
    s = re.sub(r"\s+", " ", str(s or "")).strip()
    return s if len(s) <= n else s[: n - 1] + "…"
