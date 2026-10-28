"""HTML result report (opens in the browser, can be printed)."""
from __future__ import annotations

import html
from pathlib import Path

from .core import APP_LABEL
from .session import fmt_secs

CSS = """
:root{--bg:#f6f7fb;--card:#fff;--ink:#1d2433;--muted:#5b6475;--ok:#1a7f37;--bad:#c62828;--line:#e3e6ee;--accent:#0b4f9c}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--ink);font:15px/1.5 Segoe UI,Arial,sans-serif}
.wrap{max-width:1050px;margin:0 auto;padding:24px 16px}
h1{margin:0 0 4px;font-size:24px}h2{margin:28px 0 10px;font-size:19px}h3{margin:0;font-size:16px}
.muted{color:var(--muted)}
.cards{display:flex;flex-wrap:wrap;gap:12px;margin:16px 0}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:14px 18px;min-width:150px}
.big{font-size:28px;font-weight:700}
.task{background:var(--card);border:1px solid var(--line);border-radius:10px;margin:14px 0;overflow:hidden}
.task header{display:flex;justify-content:space-between;gap:10px;padding:12px 16px;background:#eef2fa;border-bottom:1px solid var(--line)}
table{width:100%;border-collapse:collapse}td,th{padding:8px 10px;border-bottom:1px solid var(--line);vertical-align:top;text-align:left}
th{font-size:13px;color:var(--muted);font-weight:600}
.ok{color:var(--ok);font-weight:700}.bad{color:var(--bad);font-weight:700}.part{color:#b26500;font-weight:700}
.how{background:#fff8e1;border-left:3px solid #f0b400;padding:4px 8px;margin-top:4px;font-size:14px}
.sol{padding:10px 16px 14px}.sol ol{margin:6px 0 0 18px;padding:0}
.ev{font:13px Consolas,monospace}.ev td{padding:3px 8px}
@media print{.task{break-inside:avoid}}
"""


def esc(s):
    return html.escape(str(s))


def build(record: dict, path: Path):
    r = record
    secs = r.get("sections", {})
    cards = [f'<div class="card"><div class="muted">Total</div><div class="big">{r["score"]:g} / {r["out_of"]:g}</div></div>']
    for app in ("word", "excel", "ppt"):
        if app in secs:
            g, m = secs[app]
            t = r.get("app_seconds", {}).get(app)
            extra = f'<div class="muted">time in program: {fmt_secs(t)}</div>' if t else ""
            cards.append(f'<div class="card"><div class="muted">{APP_LABEL[app]}</div><div class="big">{g:g} / {m:g}</div>{extra}</div>')
    cards.append(f'<div class="card"><div class="muted">Time used</div><div class="big">{fmt_secs(r["seconds_used"])}</div>'
                 f'<div class="muted">of {fmt_secs(r["seconds_allowed"])}</div></div>')

    parts = [f"<!doctype html><html><head><meta charset='utf-8'><title>Result – {esc(r['title'])}</title>"
             f"<meta name='viewport' content='width=device-width,initial-scale=1'><style>{CSS}</style></head><body><div class='wrap'>",
             f"<h1>{esc(r['title'])} – Result</h1>",
             f"<div class='muted'>{esc(r['when'].replace('T', ' '))} · {esc(r['mode'])} · Folder: {esc(r['folder'])}</div>",
             "<div class='cards'>" + "".join(cards) + "</div>"]

    for t in r["results"]:
        parts.append(f"<div class='task'><header><h3>{esc(APP_LABEL[t['app']])} – {esc(t['title'])} "
                     f"<span class='muted'>({esc(t['file'])})</span></h3><b>{t['got']:g} / {t['marks']:g}</b></header>")
        parts.append("<table><tr><th style='width:34%'>What was checked</th><th>Result</th><th style='width:8%'>Marks</th></tr>")
        for c in t["checks"]:
            if c["desc"] == "Note":
                parts.append(f"<tr><td colspan='3' class='muted'>Note: {esc(c['found'])}</td></tr>")
                continue
            if c["ok"]:
                mark = "<span class='ok'>✔ Correct</span>"
            elif c.get("got"):
                mark = "<span class='part'>◐ Part marks</span>"
            else:
                mark = "<span class='bad'>✘ Not correct</span>"
            body = f"{mark}<div>Found: {esc(c['found'])}</div>"
            if not c["ok"]:
                if c["expected"]:
                    body += f"<div>Expected: {esc(c['expected'])}</div>"
                if c["how"]:
                    body += f"<div class='how'><b>How to do it:</b> {esc(c['how'])}</div>"
            parts.append(f"<tr><td>{esc(c['desc'])}</td><td>{body}</td><td>{c['got']:g}/{c['marks']:g}</td></tr>")
        parts.append("</table>")
        if t.get("solution"):
            parts.append("<div class='sol'><b>Model answer (step by step):</b><ol>" +
                         "".join(f"<li>{esc(s)}</li>" for s in t["solution"]) + "</ol></div>")
        parts.append("</div>")

    ev = r.get("events") or []
    if ev:
        parts.append("<h2>Activity timeline</h2><div class='task'><table class='ev'>")
        for e in ev:
            parts.append(f"<tr><td>{esc(e['time'])}</td><td>+{fmt_secs(e['elapsed'])}</td><td>{esc(e['app'])}</td>"
                         f"<td>{esc(e['msg'])}</td></tr>")
        parts.append("</table></div>")
    parts.append("</div></body></html>")
    Path(path).write_text("\n".join(parts), encoding="utf-8")
    return path


def build_instructions(tasks, title, path):
    """Bilingual (English + Hindi) question paper, opened in the browser where Hindi shows correctly."""
    from .content.hindi import UI, hindi_lines
    parts = [f"<!doctype html><html lang='hi'><head><meta charset='utf-8'><title>{esc(title)} – Instructions</title>"
             "<meta name='viewport' content='width=device-width,initial-scale=1'>"
             "<style>" + CSS + "\n.hi{color:#0b4f9c;font-family:'Nirmala UI','Mangal','Noto Sans Devanagari',sans-serif;"
             "font-size:16px;margin:2px 0 10px}.en{margin:0}.q{font:14px Consolas,monospace;background:#f6f8fc;"
             "border-left:3px solid #c9d3e6;padding:4px 10px;margin:2px 0 2px 18px;white-space:pre-wrap;"
             "user-select:none;-webkit-user-select:none}"
             "ol{padding-left:22px}li{margin-bottom:6px}</style></head><body><div class='wrap'>",
             f"<h1>{esc(title)}</h1><div class='hi'>{esc(UI['title'])}</div>",
             f"<p class='muted'>{esc(UI['note'])}</p>"]
    for t in tasks:
        parts.append(f"<div class='task'><header><h3>{esc(APP_LABEL[t.app])} – {esc(t.title)}</h3>"
                     f"<b>{t.marks:g} marks / {esc(UI['marks'])}</b></header><div class='sol'>")
        parts.append(f"<p><b>Save as / {esc(UI['save_as'])}:</b> {esc(t.file)}</p>")
        if t.starters:
            parts.append(f"<p><b>Starter file / {esc(UI['starter'])}:</b> {esc(', '.join(t.starters))}</p>")
        parts.append("<ol>")
        open_li = False
        for en, hi, quote in hindi_lines(t):
            if quote:
                parts.append(f"<div class='q'>{esc(en)}</div>")
                continue
            if open_li:
                parts.append("</li>")
            parts.append(f"<li><p class='en'>{esc(en)}</p>" + (f"<p class='hi'>{esc(hi)}</p>" if hi else ""))
            open_li = True
        if open_li:
            parts.append("</li>")
        parts.append("</ol></div></div>")
    # the passages are for typing practice: block copying them
    parts.append("<script>document.addEventListener('copy',function(e){var s=String(window.getSelection());"
                 "if(s.length>20){e.preventDefault();alert('Copying is turned off. Please type the text yourself.');}});"
                 "document.addEventListener('contextmenu',function(e){if(e.target.closest('.q'))e.preventDefault();});</script>")
    parts.append("</div></body></html>")
    Path(path).write_text("\n".join(parts), encoding="utf-8")
    return path


def build_typing(rec, path):
    """Printable typing / dictation report with every mistake marked."""
    from .typing import KIND_LABEL
    s = rec["typing"]
    css = CSS + (".bad{color:#c62828;text-decoration:line-through}.fix{color:#1a7f37}.miss{color:#b26500;"
                 "text-decoration:underline}.extra{color:#7b3fb5;text-decoration:line-through}.half{color:#0b6e8a;"
                 "text-decoration:underline}.txt{font-size:17px;line-height:1.9;background:#fff;padding:16px;"
                 "border:1px solid var(--line);border-radius:10px}")
    out = [f"<!doctype html><html><head><meta charset='utf-8'><title>{esc(rec['title'])}</title>"
           f"<meta name='viewport' content='width=device-width,initial-scale=1'><style>{css}</style></head><body>"
           f"<div class='wrap'><h1>{esc(rec['title'])}</h1><div class='muted'>{esc(rec['when'].replace('T', ' '))}</div>",
           "<div class='cards'>"]
    if s.get("net_wpm") is not None:
        out.append(f"<div class='card'><div class='muted'>Net speed</div><div class='big'>{s['net_wpm']:g} WPM</div>"
                   f"<div class='muted'>gross {s['gross_wpm']:g} WPM · {fmt_secs(s.get('seconds', 0))}</div></div>")
    out.append(f"<div class='card'><div class='muted'>Accuracy</div><div class='big'>{s['accuracy']:g}%</div>"
               f"<div class='muted'>errors {s['error_pct']:g}%</div></div>")
    out.append(f"<div class='card'><div class='muted'>Mistakes</div><div class='big'>{s['mistakes']:g}</div>"
               f"<div class='muted'>{s['full_mistakes']} full · {s['half_mistakes']} half</div></div></div>")
    out.append("<h2>Your typing, checked</h2><div class='txt'>")
    for k, o, t in rec["items"]:
        if k == "ok":
            out.append(esc(t) + " ")
        elif k == "wrong":
            out.append(f"<span class='bad'>{esc(t)}</span> <span class='fix'>{esc(o)}</span> ")
        elif k == "missing":
            out.append(f"<span class='miss'>[{esc(o)}]</span> ")
        elif k == "extra":
            out.append(f"<span class='extra'>{esc(t)}</span> ")
        else:
            out.append(f"<span class='half'>{esc(t)}</span> <span class='fix'>{esc(o)}</span> ")
    out.append("</div><p class='muted'>Struck red = wrong (correct word in green) · [orange] = missing · "
               "struck purple = extra · underlined teal = half mistake (capitals, punctuation, spacing)</p>")
    out.append("<h2>List of mistakes</h2><div class='task'><table><tr><th>Mistake</th><th>Should be</th>"
               "<th>You typed</th></tr>")
    for k, o, t in rec["items"]:
        if k != "ok":
            out.append(f"<tr><td>{esc(KIND_LABEL[k])}{' (½)' if k in ('case', 'punct', 'spacing') else ''}</td>"
                       f"<td>{esc(o) or '—'}</td><td>{esc(t) or '—'}</td></tr>")
    out.append("</table></div><h2>Original passage</h2>"
               f"<div class='txt'>{esc(rec.get('original', ''))}</div></div></body></html>")
    Path(path).write_text("".join(out), encoding="utf-8")
    return path
