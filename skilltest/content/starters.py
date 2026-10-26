"""Helpers that create the starter files given to the student at the start of a test."""
from __future__ import annotations

import struct
import zlib
from pathlib import Path

from docx import Document
from docx.shared import Pt
from openpyxl import Workbook
from openpyxl.styles import Font


def docx_with(paragraphs, title=None):
    """Return a starter maker for a Word file holding the given plain paragraphs."""
    def make(path):
        d = Document()
        st = d.styles["Normal"]
        st.font.name = "Calibri"
        st.font.size = Pt(11)
        if title:
            d.add_paragraph(title)
        for p in paragraphs:
            d.add_paragraph(p)
        d.save(str(path))
    return make


def docx_table(headers, rows, before=None):
    def make(path):
        d = Document()
        if before:
            d.add_paragraph(before)
        t = d.add_table(rows=1, cols=len(headers))
        for i, h in enumerate(headers):
            t.rows[0].cells[i].text = str(h)
        for r in rows:
            cells = t.add_row().cells
            for i, v in enumerate(r):
                cells[i].text = str(v)
        d.save(str(path))
    return make


def xlsx(sheets, bold_header=True):
    """sheets = {name: (headers, rows)}"""
    def make(path):
        wb = Workbook()
        wb.remove(wb.active)
        for name, (headers, rows) in sheets.items():
            ws = wb.create_sheet(name)
            ws.append(list(headers))
            for r in rows:
                ws.append(list(r))
            if bold_header:
                for c in ws[1]:
                    c.font = Font(bold=False)
        wb.save(str(path))
    return make


def text_file(text):
    def make(path):
        Path(path).write_text(text, encoding="utf-8")
    return make


# ---------------------------------------------------------------- sample pictures

def _png(w, h, pixel):
    raw = b"".join(b"\x00" + b"".join(bytes(pixel(x, y)) for x in range(w)) for y in range(h))

    def chunk(t, data):
        c = struct.pack(">I", len(data)) + t + data
        return c + struct.pack(">I", zlib.crc32(t + data) & 0xFFFFFFFF)
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", w, h, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 9)) + chunk(b"IEND", b""))


def _rocket(x, y):
    w, h = 160, 240
    cx = w // 2
    if 40 < y < 200 and abs(x - cx) < 18:
        return (235, 235, 240)
    if 15 < y <= 40 and abs(x - cx) < (y - 15) * 18 // 25:
        return (220, 60, 40)
    if 160 < y < 205 and 18 <= abs(x - cx) < 18 + (y - 160) // 2:
        return (220, 60, 40)
    if 200 <= y < 235 and abs(x - cx) < 14 - (y - 200) // 3:
        return (255, 170, 30)
    return (20, 30, 80) if (x * 7 + y * 13) % 97 else (255, 255, 255)


def _satellite(x, y):
    if 90 < x < 150 and 70 < y < 130:
        return (230, 190, 60)
    if 70 < y < 130 and (10 < x < 85 or 155 < x < 230) and (x // 12) % 2 == 0:
        return (40, 80, 200)
    if 70 < y < 130 and (10 < x < 85 or 155 < x < 230):
        return (60, 110, 230)
    return (10, 15, 40)


def _logo(x, y):
    dx, dy = x - 100, y - 100
    r = (dx * dx + dy * dy) ** 0.5
    if r < 60:
        return (255, 120, 0)
    if r < 75:
        return (0, 70, 160)
    return (255, 255, 255)


def write_pictures(folder: Path):
    folder.mkdir(parents=True, exist_ok=True)
    pics = {"Rocket.png": (160, 240, _rocket), "Satellite.png": (240, 200, _satellite), "Logo.png": (200, 200, _logo)}
    for name, (w, h, f) in pics.items():
        p = folder / name
        if not p.exists():
            p.write_bytes(_png(w, h, f))
