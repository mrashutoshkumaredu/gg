"""Typing test: passages, word-by-word comparison and speed / accuracy figures.

Mistakes are counted the way government typing tests usually do:
- full mistake: a missing word, an extra word, a wrong or misspelt word
- half mistake: only the capital letters differ, only the punctuation differs, or two words are joined / one word is split
"""
from __future__ import annotations

import re
from dataclasses import dataclass, field
from difflib import SequenceMatcher

_QUOTES = str.maketrans({"‘": "'", "’": "'", "“": '"', "”": '"', "–": "-", "—": "-",
                         " ": " "})
_PUNCT = re.compile(r"[^\w]+", re.UNICODE)


def words(text):
    return (text or "").translate(_QUOTES).split()


def _bare(w):
    """word without punctuation, lower case"""
    return _PUNCT.sub("", w).lower()


@dataclass
class Item:
    kind: str          # ok, wrong, missing, extra, case, punct, spacing
    original: str = ""
    typed: str = ""

    @property
    def weight(self):
        return {"ok": 0, "wrong": 1, "missing": 1, "extra": 1}.get(self.kind, 0.5)


@dataclass
class Result:
    items: list = field(default_factory=list)
    original_words: int = 0
    typed_words: int = 0
    typed_chars: int = 0
    seconds: float = 0
    count_case: bool = True
    count_punct: bool = True

    def count(self, kind):
        return sum(1 for i in self.items if i.kind == kind)

    @property
    def full(self):
        return sum(1 for i in self.items if i.kind in ("wrong", "missing", "extra"))

    @property
    def half(self):
        return sum(1 for i in self.items if i.kind in ("case", "punct", "spacing"))

    @property
    def mistakes(self):
        return self.full + 0.5 * self.half

    @property
    def error_pct(self):
        return 100 * self.mistakes / self.original_words if self.original_words else 0.0

    @property
    def accuracy(self):
        return max(0.0, 100 - self.error_pct)

    @property
    def minutes(self):
        return self.seconds / 60 if self.seconds else 0

    @property
    def gross_wpm(self):
        """standard words (5 keystrokes) per minute"""
        return (self.typed_chars / 5) / self.minutes if self.minutes else 0.0

    @property
    def net_wpm(self):
        return max(0.0, self.gross_wpm - self.mistakes / self.minutes) if self.minutes else 0.0

    def summary(self):
        out = {"original_words": self.original_words, "typed_words": self.typed_words,
               "full_mistakes": self.full, "half_mistakes": self.half, "mistakes": self.mistakes,
               "error_pct": round(self.error_pct, 2), "accuracy": round(self.accuracy, 2),
               "wrong": self.count("wrong"), "missing": self.count("missing"), "extra": self.count("extra"),
               "case": self.count("case"), "punct": self.count("punct"), "spacing": self.count("spacing")}
        if self.seconds:
            out.update(seconds=int(self.seconds), gross_wpm=round(self.gross_wpm, 1), net_wpm=round(self.net_wpm, 1))
        return out


def _pair(o, t, count_case, count_punct):
    if o == t:
        return Item("ok", o, t)
    if _bare(o) == _bare(t) and _bare(o):
        if o.lower() == t.lower():                       # only capital letters differ
            return Item("case" if count_case else "ok", o, t)
        if _PUNCT.sub("", o) == _PUNCT.sub("", t):          # only punctuation differs
            return Item("punct" if count_punct else "ok", o, t)
        # both capitals and punctuation differ
        if not count_case and not count_punct:
            return Item("ok", o, t)
        return Item("punct" if count_punct else "case", o, t)
    return Item("wrong", o, t)


def _block(ow, tw, count_case, count_punct):
    """Compare a stretch of original words with a stretch of typed words that did not match exactly."""
    out = []
    i = j = 0
    while i < len(ow) or j < len(tw):
        if i < len(ow) and j < len(tw):
            o, t = ow[i], tw[j]
            # joined words: "in the" typed as "inthe"
            if i + 1 < len(ow) and _bare(o + ow[i + 1]) == _bare(t) and _bare(t):
                out.append(Item("spacing", f"{o} {ow[i + 1]}", t))
                i, j = i + 2, j + 1
                continue
            # split word: "together" typed as "to gether"
            if j + 1 < len(tw) and _bare(t + tw[j + 1]) == _bare(o) and _bare(o):
                out.append(Item("spacing", o, f"{t} {tw[j + 1]}"))
                i, j = i + 1, j + 2
                continue
            # when one side is longer, decide whether this typed word belongs to a later original word
            if len(ow) - i != len(tw) - j:
                look_o = [_bare(x) for x in ow[i + 1:i + 4]]
                look_t = [_bare(x) for x in tw[j + 1:j + 4]]
                if len(ow) - i > len(tw) - j and _bare(t) in look_o and _bare(o) != _bare(t):
                    out.append(Item("missing", o, ""))
                    i += 1
                    continue
                if len(tw) - j > len(ow) - i and _bare(o) in look_t and _bare(o) != _bare(t):
                    out.append(Item("extra", "", t))
                    j += 1
                    continue
            out.append(_pair(o, t, count_case, count_punct))
            i, j = i + 1, j + 1
        elif i < len(ow):
            out.append(Item("missing", ow[i], ""))
            i += 1
        else:
            out.append(Item("extra", "", tw[j]))
            j += 1
    return out


def compare(original, typed, seconds=0, count_case=True, count_punct=True):
    ow, tw = words(original), words(typed)
    # match on simplified words so that capital / punctuation slips stay lined up
    sm = SequenceMatcher(None, [_bare(w) or w for w in ow], [_bare(w) or w for w in tw], autojunk=False)
    items = []
    for op, a1, a2, b1, b2 in sm.get_opcodes():
        if op == "equal":
            for o, t in zip(ow[a1:a2], tw[b1:b2]):
                items.append(_pair(o, t, count_case, count_punct))
        else:
            items.extend(_block(ow[a1:a2], tw[b1:b2], count_case, count_punct))
    return Result(items=items, original_words=len(ow), typed_words=len(tw),
                  typed_chars=len(" ".join(tw)), seconds=seconds, count_case=count_case, count_punct=count_punct)


def compare_partial(original, typed, seconds=0, count_case=True, count_punct=True):
    """For a timed test that ended early: only compare with the part of the passage that was reached."""
    ow, tw = words(original), words(typed)
    if not tw:
        return compare("", "", seconds)
    # find how far into the passage the typist got (best match for the last few typed words)
    best_end, best = len(ow), -1.0
    n = len(tw)
    for end in range(max(1, n - 15), min(len(ow), n + 15) + 1):
        tail_o = " ".join(_bare(w) for w in ow[max(0, end - 6):end])
        tail_t = " ".join(_bare(w) for w in tw[-6:])
        s = SequenceMatcher(None, tail_o, tail_t).ratio()
        if s > best:
            best, best_end = s, end
    return compare(" ".join(ow[:best_end]), typed, seconds, count_case, count_punct)


def untyped_tail(result):
    """How many words at the end of the passage were never reached."""
    n = 0
    for it in reversed(result.items):
        if it.kind != "missing":
            break
        n += 1
    return n


KIND_LABEL = {"wrong": "Wrong / misspelt word", "missing": "Missing word", "extra": "Extra word",
              "case": "Capital letter", "punct": "Punctuation", "spacing": "Spacing (joined/split)"}


# ---------------------------------------------------------------- passages

PASSAGES = [
    ("Space programme", "The Indian Space Research Organisation was set up in 1969 with the aim of using space "
     "technology for national development. Over the years it has built satellites for communication, weather "
     "forecasting, education and disaster management. The launch vehicles PSLV and GSLV have made the country "
     "self-reliant in sending satellites to space. Missions such as Chandrayaan and Mangalyaan have brought "
     "international recognition at a very low cost. Today, the organisation is preparing for the Gaganyaan mission, "
     "which will carry Indian astronauts into space for the first time."),
    ("Office discipline", "Every employee is expected to reach the office on time and to mark attendance in the "
     "biometric system. Leave must be applied for in advance through the online portal, except in emergencies. "
     "Files should be disposed of within the time limits laid down in the office manual. Confidential documents "
     "must be kept under lock and key and must not be shared with unauthorised persons. Courteous behaviour with "
     "colleagues and visitors is an essential part of good office conduct."),
    ("Digital India", "The Digital India programme aims to transform the country into a digitally empowered "
     "society and knowledge economy. It focuses on providing high speed internet in rural areas, delivering "
     "government services electronically and improving digital literacy. Common Service Centres in villages help "
     "citizens to apply for certificates, pay bills and access banking services. The use of electronic files in "
     "government offices has reduced delays and made the working of departments more transparent."),
    ("Environment", "Protection of the environment is the shared responsibility of the government and the citizens. "
     "Rapid growth of cities has increased the pollution of air, water and soil. Planting trees, saving water and "
     "reducing the use of plastic are simple steps that every family can take. Offices can contribute by switching "
     "off lights and computers when they are not in use and by printing only when it is necessary. Small changes "
     "in daily habits, when adopted by millions of people, can make a big difference."),
    ("Right to Information", "The Right to Information Act was passed in 2005 to promote transparency and "
     "accountability in the working of every public authority. Under this Act, any citizen may request information "
     "from a public authority, which is required to reply within thirty days. Each office appoints a Public "
     "Information Officer to receive such requests. Information that affects the security of the nation or the "
     "privacy of an individual may be withheld. The Act has helped citizens to question delays and to know how "
     "public money is spent."),
    ("Satellites and weather", "Weather satellites take pictures of clouds and measure the temperature of land and "
     "sea from high above the earth. This information helps the meteorological department to forecast rain, storms "
     "and cyclones several days in advance. Timely warnings have saved thousands of lives in coastal districts. "
     "Farmers also use these forecasts to decide when to sow and when to harvest their crops. Fishermen are advised "
     "not to go into the sea when a cyclone is expected."),
    ("Hindi in offices", "Hindi is the official language of the Union and its progressive use in government work is "
     "encouraged. Letters received in Hindi must be replied to in Hindi. Name boards, forms and rubber stamps are "
     "prepared in both Hindi and English. Officers and staff are trained in Hindi typing and noting through regular "
     "workshops. Hindi Fortnight is celebrated every year in September, during which competitions are held and "
     "prizes are given to encourage employees."),
    ("Computer safety", "A strong password should have at least eight characters and should mix capital letters, "
     "small letters, numbers and symbols. Passwords must never be written on paper or shared with anyone, even "
     "with colleagues. Employees should not open e-mail attachments from unknown senders, as they may contain "
     "viruses. Important files must be backed up regularly to a safe location. Any unusual behaviour of a computer "
     "should be reported at once to the IT section so that the problem can be checked."),
]


def passage_text(title):
    return next((t for n, t in PASSAGES if n == title), "")
