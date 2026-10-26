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


def new_attempt_folder(label: str) -> Path:
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    safe = "".join(ch if ch.isalnum() or ch in "-_" else "_" for ch in label)
    f = default_root() / f"{stamp}_{safe}"
    n = 2
    while f.exists():
        f = default_root() / f"{stamp}_{safe}_{n}"
        n += 1
    return f


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


def load_history():
    try:
        return json.loads(history_path().read_text(encoding="utf-8"))
    except Exception:
        return []


def save_attempt(record: dict):
    h = load_history()
    h.append(record)
    history_path().write_text(json.dumps(h, indent=1, ensure_ascii=False), encoding="utf-8")


def make_record(mode, title, folder, results, seconds_used, seconds_allowed, app_seconds, events):
    by_app = {}
    for r in results:
        a = by_app.setdefault(r.app, [0.0, 0.0])
        a[0] += r.got
        a[1] += r.marks
    return {
        "when": datetime.now().isoformat(timespec="seconds"),
        "mode": mode,
        "title": title,
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


def fmt_secs(s):
    s = int(s)
    return f"{s // 60}:{s % 60:02d}"


def now():
    return time.monotonic()
