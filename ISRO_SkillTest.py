"""Entry point used by PyInstaller to build the Windows .exe.

ISRO_SkillTest.exe --selftest out.txt   checks that every paper can be prepared and marked (used by the build).
"""
import sys


def selftest(out_path):
    import tempfile
    import traceback
    from pathlib import Path
    lines = []
    try:
        from pptx import Presentation
        from skilltest.content import PAPERS
        from skilltest.content.drills import DRILLS
        from skilltest.core import grade_tasks
        from skilltest.session import prepare_folder
        tmp = Path(tempfile.mkdtemp())
        Presentation().save(str(tmp / "blank.pptx"))
        for name, tasks in [(p.id, p.tasks) for p in PAPERS] + [("drills", DRILLS)]:
            f = tmp / name
            prepare_folder(f, tasks)
            res = grade_tasks(tasks, f)
            errors = [c for r in res for c in r.checks if c.found.startswith("Could not read")]
            lines.append(f"{name}: tasks={len(tasks)} score={sum(r.got for r in res):g} errors={len(errors)}")
            if errors:
                raise RuntimeError(errors[0].found)
        lines.append("SELFTEST OK")
    except Exception:
        lines.append("SELFTEST FAILED")
        lines.append(traceback.format_exc())
    Path(out_path).write_text("\n".join(lines), encoding="utf-8")
    return 0 if lines[-1] == "SELFTEST OK" else 1


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "--selftest":
        sys.exit(selftest(sys.argv[2]))
    from skilltest.gui import main
    main()
