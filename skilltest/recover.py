"""Find old results: rebuild missing My Progress entries from the test folders in Documents\\ISRO_SkillTest.

Every submitted test left a folder with the student's Office files and Result_Report.html. Newer versions also
keep result.json there. A missing entry is taken from result.json when it exists; otherwise the saved Office
files are marked again (exact for first attempts; retakes of older versions used other numbers, so their score
is approximate and they are labelled "re-marked").
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path

from .content import PAPERS
from .content.drills import DRILLS
from .core import grade_tasks
from .session import DEFAULT_STUDENT, default_root, history_path, load_all_history, make_record, write_json

SECTION = {"word": ("Word", "MS Word"), "excel": ("Excel", "MS Excel"), "ppt": ("PowerPoint", "MS PowerPoint")}


def _key(text):
    return re.sub(r"[^a-z0-9]", "", text.lower())


def _known_tests():
    """normalised folder label -> (mode, title, spec, tasks)"""
    out = {}
    for p in PAPERS:
        out[_key(f"{p.title} - Full Test")] = ("Full mock", f"{p.title} - Full Test",
                                               {"kind": "mock", "paper": p.id}, p.tasks)
        for app, names in SECTION.items():
            for n in names:
                out[_key(f"{p.title} - {n}")] = ("Section practice", f"{p.title} - {SECTION[app][0]}",
                                                 {"kind": "section", "paper": p.id, "app": app}, p.section(app))
    for d in DRILLS:
        out[_key(f"Drill {d.id}")] = ("Topic drills", f"Drill {d.id}", {"kind": "drills", "ids": [d.id]}, [d])
    return out


def _guess_drills(folder: Path):
    """Which drills a multi-drill folder held: their starter files or answer files are in it."""
    names = {p.name.lower() for p in folder.iterdir() if p.is_file()}
    found = []
    for d in DRILLS:
        if d.starters and all(n.lower() in names for n in d.starters):
            found.append(d)
        elif not d.starters and d.file.lower() in names:
            found.append(d)
    return found


def _when(folder: Path):
    m = re.search(r"\((\d{2})-(\d{2})-(\d{4}) (\d{2})\.(\d{2})\)", folder.name)
    if m:
        d, mo, y, h, mi = m.groups()
        return f"{y}-{mo}-{d}T{h}:{mi}:00"
    m = re.match(r"(\d{4})-(\d{2})-(\d{2})_(\d{2})(\d{2})_", folder.name)
    if m:
        y, mo, d, h, mi = m.groups()
        return f"{y}-{mo}-{d}T{h}:{mi}:00"
    return datetime.fromtimestamp(folder.stat().st_mtime).isoformat(timespec="seconds")


def _attempt_folders(root: Path):
    """(folder, label, attempt number, student) for every submitted test folder under root."""
    if not root.exists():
        return
    for report in root.rglob("Result_Report.html"):
        f = report.parent
        m = re.match(r"Attempt (\d+)", f.name)
        if m:                                           # .../<student>/<test label>/Attempt N (date)
            label, attempt = f.parent.name, int(m.group(1))
            student_dir = f.parent.parent
        else:                                           # first versions: <date>_<time>_<label>
            mm = re.match(r"\d{4}-\d{2}-\d{2}_\d{4}_(.+?)(?:_\d+)?$", f.name)
            if not mm:
                continue
            label, attempt, student_dir = mm.group(1), 1, f.parent
        student = DEFAULT_STUDENT if student_dir == root else student_dir.name
        yield f, label, attempt, student


def find_old_results(progress=None):
    """Returns (added_records, skipped_count). Adds the missing entries to the history file."""
    root = default_root()
    history = load_all_history(strict=True)
    have = {str(Path(r.get("folder") or "")).lower() for r in history if r.get("folder")}
    known = _known_tests()
    added, skipped = [], 0
    folders = list(_attempt_folders(root))
    for i, (folder, label, attempt, student) in enumerate(folders):
        if progress:
            progress(i, len(folders))
        if str(folder).lower() in have:
            continue
        rec = None
        copy = folder / "result.json"
        if copy.exists():
            try:
                rec = json.loads(copy.read_text(encoding="utf-8"))
            except Exception:
                rec = None
        if rec is None:
            k = _key(label)
            if k in known:
                mode, title, spec, tasks = known[k]
            elif k.startswith("drills") or k == "topicdrills":
                tasks = _guess_drills(folder)
                if not tasks:
                    skipped += 1
                    continue
                mode, title = "Topic drills", "Drill " + tasks[0].id if len(tasks) == 1 else f"Drills ({len(tasks)})"
                spec = {"kind": "drills", "ids": [d.id for d in tasks]}
            else:
                skipped += 1
                continue
            try:
                results = grade_tasks(tasks, folder)
            except Exception:
                skipped += 1
                continue
            rec = make_record(mode, title, folder, results, 0, 0, {}, [], spec={**spec, "seed": 0}, attempt=attempt)
            rec["when"] = _when(folder)
            rec["recovered"] = True
            if attempt > 1:
                rec["remarked"] = True          # retakes used other numbers: the score is only approximate
        rec["folder"] = str(folder)
        rec.setdefault("student", student)
        added.append(rec)
    if added:
        history = load_all_history(strict=True) + added
        history.sort(key=lambda r: r.get("when", ""))
        write_json(history_path(), history)
    return added, skipped
