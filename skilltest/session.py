"""Test sessions: work folders, starter files, saved history."""
from __future__ import annotations

import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path

from .content.starters import write_pictures


def data_dir() -> Path:
    base = os.environ.get("APPDATA") if sys.platform == "win32" else None
    d = Path(base) / "ISRO_SkillTest" if base else Path.home() / ".isro_skilltest"
    d.mkdir(parents=True, exist_ok=True)
    return d


def settings_path():
    return data_dir() / "settings.json"


def load_settings():
    try:
        return json.loads(settings_path().read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_settings(s):
    settings_path().write_text(json.dumps(s, indent=2), encoding="utf-8")


def default_root() -> Path:
    s = load_settings()
    if s.get("root"):
        return Path(s["root"])
    docs = Path.home() / "Documents"
    return (docs if docs.exists() else Path.home()) / "ISRO_SkillTest"


def safe_name(label: str) -> str:
    return "".join(ch if ch.isalnum() or ch in " -_" else "-" for ch in label).strip()


# ---------------------------------------------------------------- students

DEFAULT_STUDENT = "Me"


def students():
    return load_settings().get("students") or [DEFAULT_STUDENT]


def current_student():
    s = load_settings()
    cur = s.get("student") or DEFAULT_STUDENT
    return cur if cur in students() else students()[0]


def set_student(name):
    s = load_settings()
    names = s.get("students") or [DEFAULT_STUDENT]
    if name not in names:
        names.append(name)
    s["students"], s["student"] = names, name
    save_settings(s)


def remove_student(name):
    s = load_settings()
    names = [n for n in (s.get("students") or [DEFAULT_STUDENT]) if n != name] or [DEFAULT_STUDENT]
    s["students"] = names
    if s.get("student") == name:
        s["student"] = names[0]
    save_settings(s)


def student_root(student=None) -> Path:
    student = student or current_student()
    return default_root() if student == DEFAULT_STUDENT else default_root() / safe_name(student)


def next_attempt_number(label: str) -> int:
    base = student_root() / safe_name(label)
    nums = []
    if base.exists():
        for p in base.iterdir():
            if p.is_dir() and p.name.startswith("Attempt "):
                try:
                    nums.append(int(p.name.split()[1]))
                except (IndexError, ValueError):
                    pass
    return max(nums, default=0) + 1


def new_attempt_folder(label: str):
    """Each test gets its own folder; every retake is a new 'Attempt N' inside it.
    Returns (folder, attempt_number)."""
    base = student_root() / safe_name(label)
    base.mkdir(parents=True, exist_ok=True)
    n = next_attempt_number(label)
    stamp = datetime.now().strftime("%d-%m-%Y %H.%M")
    return base / f"Attempt {n} ({stamp})", n


def prepare_folder(folder: Path, tasks):
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    write_pictures(folder / "Pictures")
    for t in tasks:
        for name, make in t.starters.items():
            p = folder / name
            if not p.exists():
                make(p)
    # remember starter times so we can tell if a starter was saved again
    return folder


# ---------------------------------------------------------------- history

def history_path():
    return data_dir() / "history.json"


def load_all_history():
    try:
        return json.loads(history_path().read_text(encoding="utf-8"))
    except Exception:
        return []


def load_history(student=None):
    """Attempts of one student (the current one by default)."""
    student = student or current_student()
    return [r for r in load_all_history() if (r.get("student") or DEFAULT_STUDENT) == student]


def save_attempt(record: dict):
    record.setdefault("student", current_student())
    h = load_all_history()
    h.append(record)
    history_path().write_text(json.dumps(h, indent=1, ensure_ascii=False), encoding="utf-8")


def make_record(mode, title, folder, results, seconds_used, seconds_allowed, app_seconds, events, spec=None, attempt=1):
    by_app = {}
    for r in results:
        a = by_app.setdefault(r.app, [0.0, 0.0])
        a[0] += r.got
        a[1] += r.marks
    return {
        "when": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "title": title,
        "spec": spec or {},
        "attempt": attempt,
        "folder": str(folder),
        "score": round(sum(r.got for r in results), 2),
        "out_of": round(sum(r.marks for r in results), 2),
        "sections": {k: [round(v[0], 2), round(v[1], 2)] for k, v in by_app.items()},
        "seconds_used": int(seconds_used),
        "seconds_allowed": int(seconds_allowed),
        "app_seconds": {k: int(v) for k, v in app_seconds.items()},
        "results": [r.to_dict() for r in results],
        "events": events[-500:],
    }


def delete_attempt(record: dict, delete_files=True):
    """Remove one attempt from history (matched by time and title) and, if asked, its folder."""
    import shutil
    h = [r for r in load_all_history() if not (r.get("when") == record.get("when") and r.get("title") == record.get("title"))]
    history_path().write_text(json.dumps(h, indent=1, ensure_ascii=False), encoding="utf-8")
    folder = record.get("folder")
    if delete_files and folder and Path(folder).exists() and default_root() in Path(folder).parents:
        shutil.rmtree(folder, ignore_errors=True)
        parent = Path(folder).parent
        if parent != default_root() and parent.exists() and not any(parent.iterdir()):
            parent.rmdir()


def stats(history):
    tests = [r for r in history if r.get("mode") not in ("Shortcut quiz", "Typing test", "Typing compare")]
    if not tests:
        return None
    pct = [r["score"] / r["out_of"] * 100 for r in tests if r.get("out_of")]
    return {"count": len(tests), "best": max(pct) if pct else 0, "last": pct[-1] if pct else 0,
            "avg": sum(pct) / len(pct) if pct else 0}


def weak_topics(history, limit=12):
    stats = {}
    for rec in history:
        for t in rec.get("results", []):
            for c in t["checks"]:
                if not c.get("topic"):
                    continue
                key = (t["app"], c["topic"])
                s = stats.setdefault(key, [0, 0])
                s[1] += 1
                s[0] += 1 if c["ok"] else 0
    rows = [(app, topic, ok, n, ok / n) for (app, topic), (ok, n) in stats.items() if n]
    rows.sort(key=lambda r: (r[4], -r[3]))
    return rows[:limit]


def export_results(path, history=None):
    """All students' attempts in one Excel file for the teacher."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    history = load_all_history() if history is None else history
    wb = Workbook()
    ws = wb.active
    ws.title = "Attempts"
    head = ["Student", "Date", "Time", "Test", "Mode", "Attempt", "Score", "Out of", "Percent", "Word", "Excel",
            "PowerPoint", "Time used (min)", "Min in Word", "Min in Excel", "Min in PowerPoint"]
    ws.append(head)
    for r in history:
        s = r.get("sections", {})
        a = r.get("app_seconds", {})
        when = r.get("when", "")
        ws.append([r.get("student") or DEFAULT_STUDENT, when[:10], when[11:16], r.get("title"), r.get("mode"),
                   r.get("attempt"), r.get("score"), r.get("out_of"),
                   round(r["score"] / r["out_of"], 4) if r.get("out_of") else None,
                   s.get("word", [None])[0], s.get("excel", [None])[0], s.get("ppt", [None])[0],
                   round(r.get("seconds_used", 0) / 60, 1),
                   round(a.get("word", 0) / 60, 1), round(a.get("excel", 0) / 60, 1), round(a.get("ppt", 0) / 60, 1)])
    for c in ws["I"][1:]:
        c.number_format = "0%"
    ws2 = wb.create_sheet("Weak topics")
    ws2.append(["Student", "Program", "Syllabus topic", "Correct", "Times checked", "Success"])
    names = sorted({r.get("student") or DEFAULT_STUDENT for r in history})
    for n in names:
        for app, topic, ok, cnt, rate in weak_topics([r for r in history if (r.get("student") or DEFAULT_STUDENT) == n], 50):
            ws2.append([n, {"word": "Word", "excel": "Excel", "ppt": "PowerPoint"}.get(app, app), topic, ok, cnt, rate])
    for c in ws2["F"][1:]:
        c.number_format = "0%"
    for sheet in (ws, ws2):
        for c in sheet[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="0B3D91")
            c.alignment = Alignment(horizontal="center", wrap_text=True)
        sheet.freeze_panes = "A2"
        sheet.auto_filter.ref = sheet.dimensions
        for i in range(1, sheet.max_column + 1):
            width = max(len(str(c.value or "")) for c in sheet[get_column_letter(i)])
            sheet.column_dimensions[get_column_letter(i)].width = min(max(width + 2, 9), 45)
    wb.save(str(path))
    return path


def make_problem_report(record: dict, note: str) -> Path:
    """Zip the attempt's files, its result and the student's note, to send to the developer."""
    import zipfile
    out_dir = default_root() / "Marking reports"
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    zp = out_dir / f"Marking_report_{safe_name(record.get('title', 'test'))}_{stamp}.zip"
    with zipfile.ZipFile(zp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("note.txt", note or "")
        rec = {k: v for k, v in record.items() if k != "events"}
        z.writestr("result.json", json.dumps(rec, indent=1, ensure_ascii=False))
        folder = Path(record.get("folder") or "")
        if folder.exists():
            for p in folder.rglob("*"):
                if p.is_file() and "Pictures" not in p.parts and not p.name.startswith("~$") and p.stat().st_size < 20e6:
                    z.write(p, "files/" + str(p.relative_to(folder)))
    return zp


def fmt_secs(s):
    s = int(s)
    return f"{s // 60}:{s % 60:02d}"


def now():
    return time.monotonic()
