"""Reading a .pptx and the MS PowerPoint check builders."""
from __future__ import annotations

import re
import zipfile

from lxml import etree
from pptx import Presentation
from pptx.util import Emu

from .core import Check, norm, short

T_CREATE = "Creating a presentation"
T_FMT = "Formatting a presentation"
T_GFX = "Adding graphics"
T_OPEN = "Opening new presentation / templates / layouts"


class Deck:
    def __init__(self, path):
        self.path = path
        self.prs = Presentation(str(path))
        self.slides = list(self.prs.slides)
        self.zip = zipfile.ZipFile(str(path))

    def slide(self, n):
        """n is 1-based."""
        if n > len(self.slides):
            raise IndexError(f"There is no slide {n} (the file has {len(self.slides)} slide(s))")
        return self.slides[n - 1]

    def xml(self, part):
        return etree.tostring(part._element).decode() if hasattr(part, "_element") else ""

    def slide_xml(self, n):
        return etree.tostring(self.slide(n)._element).decode()

    def slide_text(self, n):
        s = self.slide(n)
        out = []
        for sh in s.shapes:
            out.append(" ".join(sh._element.itertext()))
        return " ".join(out)

    def theme_name(self):
        for n in self.zip.namelist():
            if re.match(r"ppt/theme/theme\d+\.xml", n):
                x = self.zip.read(n).decode("utf8", "ignore")
                m = re.search(r'<a:theme [^>]*name="([^"]*)"', x)
                # the first theme used by slide master 1
                return m.group(1) if m else ""
        return ""


def _slides(deck, which):
    if which == "all":
        return list(range(1, len(deck.slides) + 1))
    if isinstance(which, int):
        return [which]
    return list(which)


def _desc_slides(which):
    if which == "all":
        return "all slides"
    if isinstance(which, int):
        return f"slide {which}"
    return "slides " + ", ".join(map(str, which))


def slide_count(n, marks=1.0, file=None, at_least=False):
    def fn(ctx):
        d = ctx.ppt(file)
        c = len(d.slides)
        return (c >= n if at_least else c == n), f"{c} slide(s)"
    return Check(f"Presentation has {'at least ' if at_least else ''}{n} slides", marks, fn, expected=f"{n} slides",
                 how="Home > New Slide (Ctrl+M) to add slides.", topic=T_CREATE)


LAYOUT_ALIASES = {
    "title slide": ["title slide"],
    "title and content": ["title and content", "title and text", "title, content"],
    "two content": ["two content", "two contents", "title and 2 content", "two objects"],
    "comparison": ["comparison"],
    "title only": ["title only"],
    "blank": ["blank"],
    "section header": ["section header"],
    "picture with caption": ["picture with caption", "content with caption"],
}


def layout(n, name, marks=1.0, file=None):
    def fn(ctx):
        s = ctx.ppt(file).slide(n)
        got = s.slide_layout.name
        ok = any(a in got.lower() for a in LAYOUT_ALIASES.get(name.lower(), [name.lower()]))
        return ok, f"Slide {n} layout: {got}"
    return Check(f"Slide {n} uses the '{name}' layout", marks, fn, expected=name,
                 how=f"Select slide {n} > Home > Layout > {name}. (For a new slide: Home > New Slide drop-down > {name}.)",
                 topic=T_OPEN)


def title(n, text, marks=1.0, file=None):
    def fn(ctx):
        s = ctx.ppt(file).slide(n)
        t = s.shapes.title.text_frame.text if s.shapes.title is not None and s.shapes.title.has_text_frame else ""
        if norm(text) in norm(t):
            return True, f"Title: '{short(t, 50)}'"
        if norm(text) in norm(ctx.ppt(file).slide_text(n)):
            return True, "Text is on the slide (not in the title placeholder)"
        return False, f"Title: '{short(t, 50) or 'empty'}'"
    return Check(f"Slide {n} title '{short(text, 40)}'", marks, fn, expected=text,
                 how="Click 'Click to add title' and type.", topic=T_CREATE)


def text_on(n, texts, marks=1.0, file=None, what=None):
    texts = [texts] if isinstance(texts, str) else texts

    def fn(ctx):
        t = norm(ctx.ppt(file).slide_text(n))
        miss = [x for x in texts if norm(x) not in t]
        return (not miss), ("All text present" if not miss else "Missing: " + ", ".join(short(m, 25) for m in miss[:4]))
    return Check(what or f"Slide {n} has the given text", marks, fn, expected="; ".join(short(x, 30) for x in texts[:4]),
                 how="Click 'Click to add text' and type each point; press Enter for a new bullet.", topic="adding text")


def subtitle(n, text, marks=1.0, file=None):
    return text_on(n, [text], marks, file, what=f"Slide {n} subtitle '{short(text, 35)}'")


def design_applied(marks=1.0, file=None):
    def fn(ctx):
        d = ctx.ppt(file)
        name = d.theme_name()
        ok = name and name.lower() not in ("office theme", "office", "")
        return bool(ok), f"Design/theme: {name or 'default Office Theme'}"
    return Check("A design template (theme) applied", marks, fn, expected="Any theme other than the blank Office Theme",
                 how="Design tab > Themes > click a theme (e.g. Facet, Ion, Retrospect).", topic="presentation templates")


def _effective_bg(slide):
    for part in (slide, slide.slide_layout, slide.slide_layout.slide_master):
        bg = part._element.find(".//{http://schemas.openxmlformats.org/presentationml/2006/main}bg")
        if bg is not None:
            x = etree.tostring(bg).decode()
            if "<p:bgPr" in x:
                return part, x
    return None, ""


BG_KIND = {"gradient": "a:gradFill", "picture": "a:blipFill", "solid": "a:solidFill", "pattern": "a:pattFill",
           "texture": "a:blipFill"}


def background(which="all", kind="any", marks=1.0, file=None):
    def fn(ctx):
        d = ctx.ppt(file)
        bad = []
        for n in _slides(d, which):
            part, x = _effective_bg(d.slide(n))
            if part is None:
                bad.append(f"slide {n}: default")
            elif kind != "any" and BG_KIND[kind] not in x:
                bad.append(f"slide {n}: not {kind}")
            elif kind == "any" and part is d.slide(n).slide_layout.slide_master and "<p:bgPr" not in x:
                bad.append(f"slide {n}: default")
        return (not bad), ("Background OK" if not bad else "; ".join(bad[:4]))
    word = {"gradient": "Gradient fill", "picture": "Picture fill", "solid": "Solid fill", "any": "a"}[kind]
    return Check(f"{word} background on {_desc_slides(which)}", marks, fn, expected=f"{word} background",
                 how="Design > Format Background (or right-click the slide > Format Background) > Fill: "
                     f"{word if kind != 'any' else 'Solid/Gradient'} > choose colours" +
                     (" > Apply to All." if which == "all" else "."), topic="Slide background")


def shape_gradient(n, marks=1.0, file=None):
    def fn(ctx):
        x = ctx.ppt(file).slide_xml(n)
        sp = re.findall(r"<p:spPr>.*?</p:spPr>", x, re.S)
        ok = any("<a:gradFill" in s for s in sp)
        return ok, "Gradient fill on a shape" if ok else "No shape has a gradient fill"
    return Check(f"Shape on slide {n} has a gradient fill", marks, fn, expected="Gradient fill",
                 how="Select the shape > Shape Format (Drawing Tools Format) > Shape Fill > Gradient > pick one.",
                 topic="gradient fills")


def shape_present(n, marks=1.0, file=None, what=None):
    def fn(ctx):
        s = ctx.ppt(file).slide(n)
        c = [sh for sh in s.shapes if sh.shape_type is not None and sh.shape_type in (1, 5, 9) and not sh.is_placeholder]
        return bool(c), f"{len(c)} shape(s)/text box(es)"
    return Check(what or f"Shape drawn on slide {n}", marks, fn, expected="A shape",
                 how="Insert > Shapes > pick > drag on the slide.", topic="arranging objects")


def textbox(n, text, marks=1.0, file=None):
    def fn(ctx):
        x = ctx.ppt(file).slide_xml(n)
        boxes = re.findall(r'<p:sp>(?:(?!</p:sp>).)*?txBox="1".*?</p:sp>', x, re.S)
        t = " ".join(re.sub(r"<[^>]+>", " ", b) for b in boxes)
        ok = norm(text) in norm(t)
        return ok, "Text box found" if ok else ("Text box without that text" if boxes else "No text box")
    return Check(f"Text box on slide {n} with '{short(text, 30)}'", marks, fn, expected=text,
                 how="Insert > Text Box > click on the slide > type.", topic="adding text")


def picture(n, marks=1.0, file=None):
    def fn(ctx):
        x = ctx.ppt(file).slide_xml(n)
        c = x.count("<p:pic>") + x.count("<p:pic ")
        if not c and "<a:blip " in x:
            c = 1
        return c > 0, f"{c} picture(s) on slide {n}"
    return Check(f"Picture inserted on slide {n}", marks, fn, expected="A picture",
                 how="Insert > Pictures > pick an image from the test folder > Insert.", topic=T_GFX)


def table(n, rows=None, cols=None, marks=1.0, file=None):
    def fn(ctx):
        s = ctx.ppt(file).slide(n)
        for sh in s.shapes:
            if sh.has_table:
                r, c = len(sh.table.rows), len(sh.table.columns)
                ok = (rows is None or r == rows) and (cols is None or c == cols)
                return ok, f"Table {r} x {c}"
        return False, "No table"
    size = f" of {rows} x {cols}" if rows else ""
    return Check(f"Table{size} on slide {n}", marks, fn, expected=f"Table{size}",
                 how=f"Insert > Table > drag {cols} x {rows} boxes (or Insert Table > type columns/rows).", topic=T_GFX)


def table_text(n, texts, marks=1.0, file=None):
    def fn(ctx):
        s = ctx.ppt(file).slide(n)
        for sh in s.shapes:
            if sh.has_table:
                t = norm(" | ".join(c.text for r in sh.table.rows for c in r.cells))
                miss = [x for x in texts if norm(x) not in t]
                return (not miss), ("All entries present" if not miss else "Missing: " + ", ".join(miss[:4]))
        return False, "No table"
    return Check(f"Table data on slide {n} entered", marks, fn, expected=", ".join(texts[:5]),
                 how="Click in each cell and type; Tab moves to the next cell.", topic=T_GFX)


def chart(n, kind=None, marks=1.0, file=None):
    def fn(ctx):
        s = ctx.ppt(file).slide(n)
        for sh in s.shapes:
            if getattr(sh, "has_chart", False) and sh.has_chart:
                x = etree.tostring(sh.chart._chartSpace).decode()
                if kind and kind not in x.lower():
                    return False, "Chart is of another type"
                return True, "Chart found"
        return False, "No chart"
    return Check(f"Chart on slide {n}" + (f" ({kind})" if kind else ""), marks, fn, expected="A chart",
                 how="Insert > Chart > pick the type > OK > type the data in the small sheet > close it.", topic=T_GFX)


def smartart(n, marks=1.0, file=None):
    def fn(ctx):
        x = ctx.ppt(file).slide_xml(n)
        ok = "drawingml/2006/diagram" in x
        return ok, "SmartArt found" if ok else "No SmartArt"
    return Check(f"SmartArt on slide {n}", marks, fn, expected="A SmartArt graphic",
                 how="Insert > SmartArt > pick (e.g. Process > Basic Process) > OK > type in the text pane.", topic=T_GFX)


def transition(which="all", marks=1.0, file=None):
    def fn(ctx):
        d = ctx.ppt(file)
        bad = [n for n in _slides(d, which) if "transition" not in d.slide_xml(n)]
        return (not bad), ("Transition on every slide asked" if not bad else "No transition on slide(s) " + ", ".join(map(str, bad)))
    return Check(f"Transition on {_desc_slides(which)}", marks, fn, expected="A transition effect",
                 how="Transitions tab > pick an effect (e.g. Fade) > click Apply To All.", topic="Slide transition")


def animation(n, marks=1.0, file=None):
    def fn(ctx):
        x = ctx.ppt(file).slide_xml(n)
        ok = "<p:timing" in x and ("presetClass" in x or "<p:anim" in x)
        return ok, "Animation found" if ok else "No animation"
    return Check(f"Animation on slide {n}", marks, fn, expected="An animation effect",
                 how="Select the object > Animations tab > pick an effect (e.g. Fly In).", topic="Animation")


def _placeholder_types(slide):
    out = []
    for sh in slide.shapes:
        if sh.is_placeholder:
            try:
                out.append((str(sh.placeholder_format.type), sh))
            except Exception:
                pass
    return out


def footer(text, which="all", marks=1.0, file=None):
    def fn(ctx):
        d = ctx.ppt(file)
        bad = []
        for n in _slides(d, which):
            ph = [sh for t, sh in _placeholder_types(d.slide(n)) if "FOOTER" in t]
            if not ph or norm(text) not in norm(" ".join(sh.text_frame.text for sh in ph if sh.has_text_frame)):
                bad.append(n)
        return (not bad), ("Footer on all slides asked" if not bad else "No footer on slide(s) " + ", ".join(map(str, bad[:6])))
    return Check(f"Footer '{short(text, 30)}' on {_desc_slides(which)}", marks, fn, expected=text,
                 how="Insert > Header & Footer > tick Footer > type the text > Apply to All.", topic="adding header & footer")


def slide_numbers(which="all", marks=1.0, file=None):
    def fn(ctx):
        d = ctx.ppt(file)
        bad = [n for n in _slides(d, which) if not any("SLIDE_NUMBER" in t for t, _ in _placeholder_types(d.slide(n)))]
        return (not bad), ("Slide numbers shown" if not bad else "No slide number on slide(s) " + ", ".join(map(str, bad[:6])))
    return Check(f"Slide numbers on {_desc_slides(which)}", marks, fn, expected="Slide number",
                 how="Insert > Header & Footer > tick Slide number > Apply to All.", topic="adding header & footer")


def date_shown(which="all", marks=1.0, file=None):
    def fn(ctx):
        d = ctx.ppt(file)
        bad = [n for n in _slides(d, which) if not any("DATE" in t for t, _ in _placeholder_types(d.slide(n)))]
        return (not bad), ("Date shown" if not bad else "No date on slide(s) " + ", ".join(map(str, bad[:6])))
    return Check(f"Date on {_desc_slides(which)}", marks, fn, expected="Date and time",
                 how="Insert > Header & Footer > tick Date and time > Apply to All.", topic="adding header & footer")


def _runs_with(slide, text):
    for sh in slide.shapes:
        if sh.has_text_frame:
            for p in sh.text_frame.paragraphs:
                if norm(text) in norm("".join(r.text for r in p.runs)):
                    return sh, p, [r for r in p.runs if r.text.strip()]
    return None, None, []


def text_bold(n, text, marks=1.0, file=None):
    def fn(ctx):
        sh, p, runs = _runs_with(ctx.ppt(file).slide(n), text)
        if not runs:
            return False, "Text not found"
        ok = all(r.font.bold for r in runs)
        return ok, "Bold" if ok else "Not bold"
    return Check(f"'{short(text, 30)}' on slide {n} is Bold", marks, fn, expected="Bold",
                 how="Select the text > Ctrl+B.", topic=T_FMT)


def text_size(n, text, pt, marks=1.0, file=None):
    def fn(ctx):
        sh, p, runs = _runs_with(ctx.ppt(file).slide(n), text)
        if not runs:
            return False, "Text not found"
        sizes = {r.font.size.pt if r.font.size else None for r in runs}
        ok = sizes == {pt}
        return ok, "Size: " + ", ".join("default" if s is None else f"{s:g}" for s in sizes)
    return Check(f"'{short(text, 30)}' on slide {n} is {pt} pt", marks, fn, expected=f"{pt} pt",
                 how=f"Select the text > Home > Font Size > {pt}.", topic=T_FMT)


def text_color(n, text, marks=1.0, file=None):
    def fn(ctx):
        sh, p, runs = _runs_with(ctx.ppt(file).slide(n), text)
        if not runs:
            return False, "Text not found"
        def has(r):
            rpr = r._r.find("{http://schemas.openxmlformats.org/drawingml/2006/main}rPr")
            return rpr is not None and rpr.find("{http://schemas.openxmlformats.org/drawingml/2006/main}solidFill") is not None
        ok = all(has(r) for r in runs)
        return ok, "Colour changed" if ok else "Default colour"
    return Check(f"'{short(text, 30)}' on slide {n} colour changed", marks, fn, expected="A new font colour",
                 how="Select the text > Home > Font Color > pick a colour.", topic="colour")


def notes(n, text, marks=1.0, file=None):
    def fn(ctx):
        s = ctx.ppt(file).slide(n)
        t = s.notes_slide.notes_text_frame.text if s.has_notes_slide else ""
        ok = norm(text) in norm(t)
        return ok, f"Notes: '{short(t, 50) or 'none'}'"
    return Check(f"Speaker notes on slide {n}", marks, fn, expected=text,
                 how="View > Notes (or click 'Click to add notes' under the slide) > type.", topic=T_CREATE)


def grouped(n, marks=1.0, file=None):
    def fn(ctx):
        x = ctx.ppt(file).slide_xml(n)
        ok = "<p:grpSp>" in x
        return ok, "Grouped objects found" if ok else "No group"
    return Check(f"Objects grouped on slide {n}", marks, fn, expected="A group of shapes",
                 how="Select the shapes with Shift+click > Shape Format > Group > Group (Ctrl+G).", topic="arranging objects")


def aligned(n, edge="top", marks=1.0, file=None):
    def fn(ctx):
        s = ctx.ppt(file).slide(n)
        def flat(shs):
            for sh in shs:
                if sh.shape_type == 6:          # group: look at the shapes inside it
                    yield from flat(sh.shapes)
                else:
                    yield sh
        shapes = [sh for sh in flat(s.shapes) if sh.shape_type in (1, 13) and not sh.is_placeholder]  # shapes, pictures
        if len(shapes) < 2:
            return False, f"Only {len(shapes)} non-placeholder object(s)"
        if edge == "top":
            vals = [sh.top for sh in shapes]
        elif edge == "middle":
            vals = [sh.top + sh.height // 2 for sh in shapes]
        else:
            vals = [sh.left for sh in shapes]
        ok = max(vals) - min(vals) <= Emu(12700 * 3)
        return ok, f"{len(shapes)} objects; {'aligned' if ok else 'not lined up'}"
    return Check(f"Objects on slide {n} aligned ({edge})", marks, fn, expected=f"Same {edge}",
                 how=f"Select all the shapes > Shape Format > Align > Align {edge.capitalize()}.", topic="arranging objects")


def hyperlink(n, marks=1.0, file=None):
    def fn(ctx):
        x = ctx.ppt(file).slide_xml(n)
        ok = "hlinkClick" in x
        return ok, "Hyperlink found" if ok else "No hyperlink"
    return Check(f"Hyperlink on slide {n}", marks, fn, expected="A link",
                 how="Select the text > Ctrl+K > type the address > OK.", topic=T_CREATE)


def slide_size(ratio="16:9", marks=1.0, file=None):
    def fn(ctx):
        p = ctx.ppt(file).prs
        r = p.slide_width / p.slide_height
        got = "16:9" if abs(r - 16 / 9) < 0.02 else "4:3" if abs(r - 4 / 3) < 0.02 else f"{r:.2f}"
        return got == ratio, f"Slide size {got}"
    return Check(f"Slide size {ratio}", marks, fn, expected=ratio,
                 how=f"Design > Slide Size > {'Widescreen (16:9)' if ratio == '16:9' else 'Standard (4:3)'}.",
                 topic="Slide layout")


def master_text(text, marks=1.0, file=None):
    def fn(ctx):
        d = ctx.ppt(file)
        x = " ".join(etree.tostring(m._element).decode() for m in d.prs.slide_masters)
        x += " ".join(etree.tostring(l._element).decode() for m in d.prs.slide_masters for l in m.slide_layouts)
        ok = norm(text) in norm(re.sub(r"<[^>]+>", " ", x))
        return ok, "Found on the Slide Master" if ok else "Not on the Slide Master"
    return Check(f"'{text}' added on the Slide Master", marks, fn, expected="Appears on every slide via the master",
                 how="View > Slide Master > click the top (largest) master slide > Insert > Text Box > type > Close Master View.",
                 topic="Slide layout")
