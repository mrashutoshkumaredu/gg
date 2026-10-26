"""Every check of every paper and drill must run without a code error, and starter files alone must score ~0."""
import shutil
import sys
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from skilltest.content import PAPERS
from skilltest.content.drills import DRILLS
from skilltest.core import grade_tasks
from skilltest.session import prepare_folder


def run():
    problems = 0
    tmp = Path(tempfile.mkdtemp())
    try:
        groups = [(p.id, p.tasks) for p in PAPERS] + [("drills", DRILLS)]
        ids = set()
        for name, tasks in groups:
            for t in tasks:
                assert t.id not in ids, f"duplicate id {t.id}"
                ids.add(t.id)
                assert t.solution, f"{t.id} has no model answer"
            f = tmp / name
            prepare_folder(f, tasks)
            res = grade_tasks(tasks, f)
            got = sum(r.got for r in res)
            print(f"{name}: starter-only score {got:g}")
            for r in res:
                for c in r.checks:
                    if c.found.startswith("Could not read"):
                        problems += 1
                        print("   ERROR:", r.id, c.desc, "|", c.found)
                    elif c.ok:
                        print("   passes on starter only:", r.id, c.desc, "|", c.found)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return problems


if __name__ == "__main__":
    sys.exit(1 if run() else 0)
