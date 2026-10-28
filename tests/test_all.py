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


def test_variants():
    """Retake papers (new numbers) keep 100 marks, and their marking uses their own numbers."""
    from skilltest.content import build_paper
    from skilltest.content import mock1 as M
    p = build_paper("mock1", 7)
    e2 = next(t for t in p.tasks if t.id == "M1-E2")
    import sys as _s
    v = _s.modules["skilltest.content.mock1_v7"]
    assert v.REGION_TOTAL != M.REGION_TOTAL, "variant should have different numbers"
    for pid in ("mock1", "mock2", "mock3", "mock4", "mock5"):
        for seed in (3, 11):
            q = build_paper(pid, seed)
            assert sum(t.marks for t in q.tasks) == 100
    from skilltest.content.hindi import HI
    for t in [t for p in PAPERS for t in p.tasks] + DRILLS:
        n = sum(1 for line in t.instructions if not line.startswith("   "))
        assert len(HI.get(t.id, [])) == n, f"Hindi lines for {t.id}: {len(HI.get(t.id, []))} instead of {n}"
    print("variants and Hindi OK")
    from skilltest.typing import compare, compare_partial
    o = "The Indian Space Research Organisation was set up in 1969 with the aim of using space technology, for national development."
    t = "The indian Space Reserch Organisation was setup in 1969 with aim of of using space technology for national development"
    r = compare(o, t, seconds=60).summary()
    want = {"case": 1, "wrong": 1, "spacing": 1, "missing": 1, "extra": 1, "punct": 2}
    assert all(r[k] == v for k, v in want.items()), r
    assert r["mistakes"] == 5 and compare(o, o).summary()["mistakes"] == 0
    assert compare(o, t, count_case=False, count_punct=False).summary()["half_mistakes"] == 1      # only spacing left
    assert compare_partial(o, "The Indian Space Research Organisation was set up in", 30).summary()["mistakes"] == 0
    print("typing comparison OK")
    # untimed typing: words never reached count as missing
    from skilltest.typing import PASSAGES, untyped_tail
    txt = PASSAGES[0][1]
    r = compare(txt, " ".join(txt.split()[:40]), 600)
    assert untyped_tail(r) == len(txt.split()) - 40 and r.summary()["accuracy"] < 50
    # safe saving: a damaged file is recovered from the backup and never overwritten
    import tempfile, glob
    from skilltest import session as S
    d = Path(tempfile.mkdtemp())
    f = d / "h.json"
    S.write_json(f, [1])
    S.write_json(f, [1, 2])
    f.write_text("[1, 2, 3")                          # simulated power cut
    assert S.read_json(f, []) == [1] and glob.glob(str(f) + ".damaged-*")
    S.write_json(f, [1, 9])
    assert S.read_json(f, []) == [1, 9]
    # versions: old installs (v1.0.x) must see the new releases as newer
    from skilltest import updater as U
    assert U.parse("v1.7.12") > U.parse("v1.0.9") and U.parse("v1.7.12") > (1, 7, 11)
    # update helper: PATH keeps everything except the app's temporary folder, tools are called by full path
    import os
    old_meipass, old_path = getattr(sys, "_MEIPASS", None), os.environ.get("PATH", "")
    sys._MEIPASS = "/tmp/_MEI12345"
    os.environ["PATH"] = os.pathsep.join(["/tmp/_MEI12345", "/usr/bin", "/bin"])
    try:
        env = U.clean_env()
        assert env["PATH"] == os.pathsep.join(["/usr/bin", "/bin"]), env["PATH"]
    finally:
        os.environ["PATH"] = old_path
        if old_meipass is None:
            del sys._MEIPASS
        else:
            sys._MEIPASS = old_meipass
    bat = U.helper_script(r"C:\Users\A\Desktop\ISRO_SkillTest (1).exe", r"C:\Users\A\Desktop\ISRO_SkillTest_update.exe",
                          [111, 222], r"C:\Users\A\Desktop\ISRO_SkillTest_update.json")
    assert "IMAGENAME eq ISRO_SkillTest (1).exe\"" in bat
    assert "PID eq 111" in bat and "PID eq 222" in bat and "\\System32\\ping.exe" in bat
    assert "\\System32\\tasklist.exe" in bat and "powershell.exe" in bat
    assert bat.index(":closed") < bat.index("move /y") < bat.index("start \"\" %OLD%")
    assert "\nping" not in bat and "\ntasklist" not in bat          # never relies on PATH alone
    # a locked file is never treated as damaged, and saving stops instead of overwriting it
    real = S._read_text
    S._read_text = lambda p: (_ for _ in ()).throw(S.DataBusy("locked"))
    try:
        try:
            S.read_json(f, [])
            raise AssertionError("a locked file must raise DataBusy")
        except S.DataBusy:
            pass
        assert not [x for x in glob.glob(str(f) + ".damaged-*") if x.endswith(".damaged-new")]
    finally:
        S._read_text = real
    assert S.read_json(f, []) == [1, 9]
    print("safe saving, untimed typing and versions OK")


if __name__ == "__main__":
    bad = run()
    test_variants()
    sys.exit(1 if bad else 0)
