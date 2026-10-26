"""Helpers for building 'model answer' files in tests."""
import shutil
import subprocess
import tempfile
from pathlib import Path

from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH


def recalc_xlsx(path):
    """Open and re-save an xlsx with LibreOffice so formula results are stored (like Excel does)."""
    path = Path(path)
    out = Path(tempfile.mkdtemp())
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(out),
                    str(path)], check=True, capture_output=True, timeout=120)
    shutil.move(str(out / path.name), str(path))


def fake_pdf(path):
    Path(path).write_bytes(b"%PDF-1.4\n%test\n")


def add_field(paragraph, instr, result=""):
    fld = parse_xml(f'<w:fldSimple {nsdecls("w")} w:instr="{instr}"><w:r><w:t>{result}</w:t></w:r></w:fldSimple>')
    paragraph._p.append(fld)


def shade(cell, fill="D9E2F3"):
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.append(parse_xml(f'<w:shd {nsdecls("w")} w:val="clear" w:color="auto" w:fill="{fill}"/>'))


def para(doc, text, bold=False, italic=False, underline=False, size=None, font=None, align=None, spacing=None, style=None):
    p = doc.add_paragraph(style=style)
    r = p.add_run(text)
    r.bold, r.italic, r.underline = bold or None, italic or None, underline or None
    if size:
        r.font.size = Pt(size)
    if font:
        r.font.name = font
    if align:
        p.alignment = getattr(WD_ALIGN_PARAGRAPH, align)
    if spacing:
        p.paragraph_format.line_spacing = spacing
    return p


def set_margins(section, inch):
    for a in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(section, a, Inches(inch))
