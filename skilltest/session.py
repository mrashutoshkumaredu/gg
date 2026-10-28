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


# ---------------------------------------------------------------- safe saving
# A save goes to a temporary file that replaces the real one only when it is complete, and the previous good
# version is kept as .bak. So a power cut while saving can never wipe the results.

DATA_WARNINGS = []          # shown once by the window, e.g. "history was damaged, restored from backup"


def write_json(path: Path, data, indent=1):
    path = Path(path)
    tmp = path.with_name(path.name + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=indent, ensure_ascii=False)
        f.flush()
        os.fsync(f.fileno())
    if path.exists():
        try:
            json.loads(path.read_text(encoding="utf-8"))            # only a good file becomes the backup
            os.replace(path, path.with_name(path.name + ".bak"))
        except Exception:
            pass
    os.replace(tmp, path)


class DataBusy(OSError):
    """The file exists but could not be opened (for example locked by an antivirus scan)."""


def _read_text(path: Path):
    """Read a file, retrying for about 2 seconds if Windows or an antivirus has it locked."""
    last = None
    for _ in range(10):
        try:
            return path.read_text(encoding="utf-8")
        except FileNotFoundError:
            raise
        except OSError as e:
            last = e
            time.sleep(0.2)
    raise DataBusy(f"{path.name} is in use by another program ({last})")


def read_json(path: Path, default, what="data"):
    """Read saved data. A locked file raises DataBusy (it is never mistaken for a damaged one);
    only a file whose content is broken is set aside and replaced from the backup."""
    path = Path(path)
    bak = path.with_name(path.name + ".bak")
    if not path.exists() and not bak.exists():
        return default
    try:
        return json.loads(_read_text(path))
    except (DataBusy, FileNotFoundError) as e:
        if isinstance(e, DataBusy):
            raise
    except (ValueError, UnicodeDecodeError):
        pass
    # the main file is damaged: keep it aside (never overwrite it) and fall back to the backup
    if path.exists():
        stamp = datetime.now().strftime("%Y%m%d-%H%M%S")
        try:
            os.replace(path, path.with_name(f"{path.name}.damaged-{stamp}"))
        except OSError:
            pass
    try:
        data = json.loads(bak.read_text(encoding="utf-8"))
        DATA_WARNINGS.append(f"The saved {what} file was damaged. It has been restored from the backup "
                             f"(the most recent change may be missing).")
        write_json(path, data)
        return data
    except Exception:
        DATA_WARNINGS.append(f"The saved {what} file was damaged and no backup could be read. The damaged file "
                             f"was kept in {path.parent} so nothing is overwritten.")
        return default


def settings_path():
    return data_dir() / "settings.json"


def load_settings():
    try:
        s = read_json(settings_path(), {}, "settings")
    except DataBusy:
        return {}
    return s if isinstance(s, dict) else {}


def save_settings(s):
    write_json(settings_path(), s, indent=2)


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


def load_all_history(strict=False):
    """strict=True is used before saving: if the file cannot be read, saving stops instead of overwriting it."""
    try:
        h = read_json(history_path(), [], "results history")
    except DataBusy:
        if strict:
            raise
        DATA_WARNINGS.append("The results file is in use by another program, so My Progress could not be shown "
                             "fully this time. Nothing has been changed; try again in a moment.")
        return []
    return h if isinstance(h, list) else []


def load_history(student=None):
    """Attempts of one student (the current one by default)."""
    student = student or current_student()
    return [r for r in load_all_history() if (r.get("student") or DEFAULT_STUDENT) == student]


def save_attempt(record: dict):
    record.setdefault("student", current_student())
    h = load_all_history(strict=True)
    h.append(record)
    write_json(history_path(), h)


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
    h = [r for r in load_all_history(strict=True)
         if not (r.get("when") == record.get("when") and r.get("title") == record.get("title"))]
    write_json(history_path(), h)
    folder = record.get("folder")
    if delete_files and folder and Path(folder).exists() and default_root() in Path(folder).parents:
        shutil.rmtree(folder, ignore_errors=True)
        parent = Path(folder).parent
        if parent != default_root() and parent.exists() and not any(parent.iterdir()):
            parent.rmdir()


def stats(history):
    tests = [r for r in history if r.get("mode") not in ("Shortcut quiz", "Typing test", "Typing compare")
             and r.get("completed") is not False]
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
