"""Small white line icons, drawn once with Pillow, so every screen uses the same style."""
from __future__ import annotations

from functools import lru_cache

from PIL import Image, ImageDraw

S = 96          # drawn large, then scaled down smoothly
W = "white"


def _mock(d):                       # question paper with a tick
    d.rounded_rectangle((24, 12, 72, 84), radius=8, outline=W, width=7)
    d.line((35, 50, 45, 61, 62, 38), fill=W, width=8, joint="curve")


def _section(d):                    # three sections, the first one chosen
    d.rounded_rectangle((18, 18, 78, 34), radius=7, fill=W)
    d.rounded_rectangle((18, 41, 78, 57), radius=7, outline=W, width=5)
    d.rounded_rectangle((18, 64, 78, 80), radius=7, outline=W, width=5)


def _drill(d):                      # target
    d.ellipse((14, 14, 82, 82), outline=W, width=7)
    d.ellipse((30, 30, 66, 66), outline=W, width=7)
    d.ellipse((42, 42, 54, 54), fill=W)


def _quiz(d):                       # lightning: quick questions
    d.polygon([(56, 8), (24, 54), (46, 54), (38, 88), (72, 40), (50, 40)], fill=W)


def _typing(d):                     # keyboard
    d.rounded_rectangle((10, 26, 86, 72), radius=9, outline=W, width=6)
    for row, y in enumerate((38, 50)):
        for i in range(5 - row):
            x = 22 + i * 13 + row * 6
            d.rounded_rectangle((x, y - 3, x + 7, y + 3), radius=2, fill=W)
    d.rounded_rectangle((30, 58, 66, 63), radius=2, fill=W)


def _compare(d):                    # two pages side by side
    d.rounded_rectangle((12, 18, 46, 78), radius=6, outline=W, width=6)
    d.rounded_rectangle((50, 18, 84, 78), radius=6, outline=W, width=6)
    for y in (34, 46, 58):
        d.line((20, y, 38, y), fill=W, width=5)
        d.line((58, y, 76, y), fill=W, width=5)


def _settings(d):                   # gear
    import math
    cx = cy = 48
    for i in range(8):
        a = i * math.pi / 4
        x, y = cx + 34 * math.cos(a), cy + 34 * math.sin(a)
        d.ellipse((x - 9, y - 9, x + 9, y + 9), fill=W)
    d.ellipse((cx - 28, cy - 28, cx + 28, cy + 28), fill=W)
    d.ellipse((cx - 12, cy - 12, cx + 12, cy + 12), fill=(0, 0, 0, 0))


DRAW = {"mock": _mock, "section": _section, "drill": _drill, "quiz": _quiz, "typing": _typing,
        "compare": _compare, "settings": _settings}


GLYPH = {"mock": "✓", "section": "≡", "drill": "◎", "quiz": "ϟ", "typing": "⌨", "compare": "⇆", "settings": "⚙"}


@lru_cache(maxsize=None)
def image(name: str, size: int = 22):
    """A customtkinter image of the icon, white on transparent; None if pictures cannot be shown."""
    try:
        return _image(name, size)
    except Exception:
        return None


def _image(name, size):
    import customtkinter as ctk
    img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    DRAW[name](ImageDraw.Draw(img))
    px = max(16, size * 2)              # sharp on high-DPI screens
    img = img.resize((px, px), Image.LANCZOS)
    return ctk.CTkImage(light_image=img, dark_image=img, size=(size, size))
