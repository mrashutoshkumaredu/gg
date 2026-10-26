"""Mock Test 4."""
from .. import common as C
from .. import excel as X
from .. import ppt as P
from .. import word as W
from ..core import Paper, Task
from . import variants as V
from .starters import docx_with, text_file, xlsx

# ------------------------------------------------------------------ WORD

A_DATE = "Date: 26 October 2026"
A_SUBJ = "Subject: Request for Earned Leave"
A_P1 = ("I wish to state that I have to attend the marriage of my younger sister at my native place. I therefore "
        "request you to kindly sanction earned leave for ten days from 3 November to 12 November.")
A_P2 = ("During my absence my work will be looked after by Shri Rajesh, Assistant of the same section. I will be "
        "available on my mobile phone for any urgent official work.")
A_ENC = ["Leave application form", "Copy of invitation card", "Charge handover note"]

w1 = Task(
    id="M4-W1", app="word", title="Leave Application", file="Application.docx", minutes=12,
    instructions=[
        "Start a new blank document and type this letter:",
        f"    {A_DATE}",
        "    To, The Head of Division, ISRO, Bengaluru",
        f"    {A_SUBJ}",
        "    Sir,",
        f"    {A_P1}",
        f"    {A_P2}",
        "    Enclosures:",
        *[f"    {e}" for e in A_ENC],
        "    Yours faithfully, (Name)",
        "Right align the date line.",
        "Make the Subject line Bold and Underlined.",
        "Set the two main paragraphs to font Times New Roman, size 12, Justify, line spacing 2.0 (Double).",
        "Give the three enclosures Numbering (1, 2, 3).",
        "Custom margins: Top 1\", Bottom 1\", Left 1.25\", Right 1\".",
        "Save as Application.docx and also save a copy in Word 97-2003 format as Application.doc.",
    ],
    checks=[
        C.file_saved("Application.docx", 1),
        W.typed(A_P1, 2, what="First paragraph typed"), W.typed(A_P2, 1.5, what="Second paragraph typed"),
        W.alignment(A_DATE, "Right", 1),
        W.bold(A_SUBJ, 1), W.underline(A_SUBJ, 1),
        W.font_name(A_P1, "Times New Roman", 0.75), W.font_name(A_P2, "Times New Roman", 0.75),
        W.font_size(A_P1, 12, 0.5), W.font_size(A_P2, 12, 0.5),
        W.alignment(A_P1, "Justify", 0.5), W.alignment(A_P2, "Justify", 0.5),
        W.line_spacing(A_P1, 2.0, 0.75), W.line_spacing(A_P2, 2.0, 0.75),
        W.listed(A_ENC, "number", 1.5),
        W.margins(1, 1, 1.25, 1, marks=2.5),
        C.file_format("Application.doc", "doc97", 1.5),
    ],
    solution=[
        "Type the whole letter, Enter after each line.",
        "Click the date > Ctrl+R. Select the subject > Ctrl+B, Ctrl+U.",
        "Select both paragraphs > Font Times New Roman, 12 > Ctrl+J > Line spacing 2.0 (Ctrl+2).",
        "Select the three enclosure lines > Numbering.",
        "Layout > Margins > Custom Margins > Top 1, Bottom 1, Left 1.25, Right 1 > OK.",
        "F12 > Application > Save. F12 > Save as type Word 97-2003 Document > Save > Continue.",
    ],
)

TT_HEAD = ["Day", "9:30-10:30", "10:30-11:30", "11:30-12:30", "12:30-1:30", "1:30-2:30"]
TT = [["Monday", "Word", "Excel", "Typing", "Lunch", "PowerPoint"],
      ["Tuesday", "Excel", "Word", "PowerPoint", "", "Typing"],
      ["Wednesday", "Typing", "PowerPoint", "Word", "", "Excel"],
      ["Thursday", "PowerPoint", "Typing", "Excel", "", "Word"]]

w2 = Task(
    id="M4-W2", app="word", title="Training Timetable", file="Timetable.docx", minutes=13,
    instructions=[
        "Start a new document. Type the heading 'Training Timetable' and apply the style Heading 2.",
        "Insert a table of 5 rows × 6 columns and fill it:",
        "    " + " | ".join(TT_HEAD),
        *["    " + " | ".join(c or "(blank)" for c in r) for r in TT],
        "Merge the four cells of the 12:30-1:30 column below the header (rows 2 to 5) into one cell containing Lunch.",
        "Make the header row Bold and apply shading to it; keep All Borders.",
        "Add one more row at the bottom with Friday | Revision | Revision | Test | Lunch | Test "
        "(it may join the merged Lunch cell – that's fine).",
        "Set the page to Landscape and paper size A4.",
        "Header text: ISRO Training Cell. Insert page numbers at the bottom.",
        "Save as Timetable.docx and as Timetable.pdf.",
    ],
    checks=[
        C.file_saved("Timetable.docx", 1),
        W.style("Training Timetable", "Heading 2", 1),
        W.table_size(6, 6, 2, at_least=True),
        W.table_has(["Monday", "Thursday", "Friday", "Revision", "Lunch", "PowerPoint", "1:30-2:30"], 2),
        W.table_merged(2),
        W.bold("9:30-10:30", 1), W.table_shading(1), W.table_borders(1),
        W.orientation("landscape", 1), W.paper_size("A4", 0.5),
        W.header_text("ISRO Training Cell", 1), W.page_numbers(1),
        C.file_format("Timetable.pdf", "pdf", 1.5),
        W.typed("Training Timetable", 1, what="Heading typed"),
    ],
    solution=[
        "Type the heading > Styles > Heading 2 > Enter.",
        "Insert > Table > Insert Table > 6 columns, 5 rows. Type the data with Tab.",
        "Select the 4 cells under 12:30-1:30 > Layout > Merge Cells > type Lunch.",
        "Click in the last row > Layout > Insert Below > type Friday's data.",
        "Select header row > Ctrl+B > Shading. Borders > All Borders.",
        "Layout > Orientation > Landscape; Size > A4.",
        "Insert > Header > type ISRO Training Cell. Insert > Page Number > Bottom of Page.",
        "F12 > Timetable > Save. F12 > PDF > Save.",
    ],
)

# ------------------------------------------------------------------ EXCEL

_CONS0 = [("Faisal", 5520, 5784), ("Anjali", 4512, 4630), ("Hari", 2200, 2460), ("Chitra", 1210, 1295),
          ("Esther", 9050, 9122), ("Balaji", 7801, 8015), ("Geeta", 6400, 6512), ("Dinesh", 3388, 3611)]
CONS = []
for (n0, p0, c0), n in zip(_CONS0, V.names([r[0] for r in _CONS0])):
    p = V.num(p0, 0.3)
    CONS.append((n, p, p + V.num(c0 - p0, 0.6, low=20)))
if [r[0] for r in CONS] == sorted(r[0] for r in CONS):
    CONS.reverse()                                   # the sorting task needs unsorted names


def _units(r):
    return r["C"] - r["B"]


def _charge(r):
    u = _units(r)
    return u * 3 if u <= 100 else (300 + (u - 100) * 5 if u <= 200 else 800 + (u - 200) * 7)


e1 = Task(
    id="M4-E1", app="excel", title="Electricity Bill", file="Bill.xlsx", minutes=13,
    starters={"Bill.xlsx": xlsx({"Bills": (["Consumer", "Previous", "Current"], CONS)})},
    instructions=[
        "Open Bill.xlsx (sheet Bills, rows 2-9).",
        "D1: Units. D2:D9 = Current − Previous.",
        "E1: Charge. Use a nested IF: up to 100 units ₹3 per unit; 101-200 units ₹300 + ₹5 per unit above 100; "
        "above 200 units ₹800 + ₹7 per unit above 200.",
        "   e.g. =IF(D2<=100,D2*3,IF(D2<=200,300+(D2-100)*5,800+(D2-200)*7))",
        "F1: Total. F2:F9 = ROUND(Charge × 1.05, 0) (5% tax).",
        "Format E2:F9 as Currency.",
        "A11: Total Collection, F11: SUM of F2:F9. A12: Average Units, D12: AVERAGE of D2:D9. "
        "A13: Maximum Units, D13: MAX of D2:D9.",
        "Add Data Validation on C2:C9: Whole number greater than or equal to 0.",
        "Sort rows 2-9 by Consumer name A to Z.",
        "Apply All Borders to A1:F9 and Bold to A1:F1.",
        "Save (Ctrl+S).",
    ],
    checks=[
        X.rowwise("Bills", "D", 2, 9, r"-", _units, 2, what="Units = Current − Previous", how="D2: =C2-B2, fill down."),
        X.rowwise("Bills", "E", 2, 9, "IF", _charge, 4, what="Charge with nested IF",
                  how="E2: =IF(D2<=100,D2*3,IF(D2<=200,300+(D2-100)*5,800+(D2-200)*7)), fill down."),
        X.rowwise("Bills", "F", 2, 9, "ROUND", lambda r: round(_charge(r) * 1.05), 2.5, what="Total with ROUND",
                  how="F2: =ROUND(E2*1.05,0), fill down."),
        X.number_format("Bills", "E2:F9", "currency", 1),
        X.formula("Bills", "F11", "SUM", None, 1.5),
        X.formula("Bills", "D12", "AVERAGE", sum(c - p for _, p, c in CONS) / len(CONS), 1),
        X.formula("Bills", "D13", "MAX", max(c - p for _, p, c in CONS), 1),
        X.validation("Bills", "C2", "whole", 1.5),
        X.sorted_by("Bills", "A", 2, 9, keep=("A", "B", {n: p for n, p, _ in CONS}), marks=1.5),
        X.borders("Bills", "A1:F9", 1), X.bold("Bills", "A1:F1", 1),
    ],
    solution=[
        "D1 Units, D2: =C2-B2. E1 Charge, E2: the nested IF. F1 Total, F2: =ROUND(E2*1.05,0). "
        "Select D2:F2 > double-click the fill handle.",
        "Select E2:F9 > Currency format.",
        "F11: =SUM(F2:F9). D12: =AVERAGE(D2:D9). D13: =MAX(D2:D9).",
        "Select C2:C9 > Data > Data Validation > Whole number > greater than or equal to > 0 > OK.",
        "Select A1:F9 > Data > Sort > Consumer > A to Z.",
        "A1:F9 > All Borders; A1:F1 > Ctrl+B. Ctrl+S.",
    ],
)

BOOKS = [("Let Us C", "Comp Sci", 2019, 450), ("Higher Engineering Mathemetics", "Maths", 2017, 980),
         ("Wings of Fire", "General", 1999, 350), ("Data Structures", "Comp Sci", 2020, 620),
         ("Linear Algebra", "Maths", 2016, 540), ("Ignited Minds", "General", 2002, 300),
         ("Operating Systems", "Comp Sci", 2018, 710)]

e2 = Task(
    id="M4-E2", app="excel", title="Library Catalogue", file="Library.xlsx", minutes=12,
    starters={"Library.xlsx": xlsx({"Books": (["Title", "Subject", "Year", "Price"], BOOKS)})},
    instructions=[
        "Open Library.xlsx.",
        "Insert a new column before column A and fill it with S.No 1 to 7 (heading S.No in A1).",
        "Use Find & Replace to replace 'Comp Sci' with 'Computer Science'.",
        "Run Spell Check (F7) and correct 'Mathemetics'.",
        "Insert a new row 1 for the title: type CENTRAL LIBRARY in A1 and Merge & Centre A1:E1.",
        "In the Price heading cell (E2) change the text to Price (₹) using Insert > Symbol.",
        "Turn on Wrap Text for B2:B9 and set column B width to 25.",
        "Insert the picture Logo.png on the sheet.",
        "In G2 type Website and in G3 make a hyperlink to www.isro.gov.in.",
        "Delete the row of the book 'Ignited Minds'.",
        "Save (Ctrl+S) and also save a copy as Library.csv.",
    ],
    checks=[
        X.series("Books", "A3:A7", 1, 1, 2),
        X.text_in("Books", "C3:C9", ["Computer Science"], 1, what="'Computer Science' replaced in"),
        X.Check("No 'Comp Sci' left", 1, lambda ctx: _no_text(ctx, "Comp Sci"), expected="All replaced",
                how="Ctrl+H > Find what: Comp Sci > Replace with: Computer Science > Replace All.",
                topic="find, search & replace"),
        X.text_in("Books", "B3:B9", ["Higher Engineering Mathematics"], 1.5, what="Spelling corrected (Mathematics)",
                  topic="spell checks"),
        X.merged("Books", "A1:E1", 1.5), X.value("Books", "A1", "CENTRAL LIBRARY", 0.5),
        X.symbol("Books", "E2", "₹", 1.5),
        X.wrap("Books", "B3:B8", 1), X.col_width("Books", "B", 25, 1),
        X.picture(1.5),
        X.hyperlink("Books", "G3", 1.5),
        X.Check("Row of 'Ignited Minds' deleted", 1.5, lambda ctx: _no_text(ctx, "Ignited Minds"),
                expected="That book's row removed", how="Right-click its row number > Delete.", topic=X.T_DATA),
        C.file_format("Library.csv", "csv", 1.5),
    ],
    solution=[
        "Right-click column A header > Insert. A1: S.No, A2: 1, A3: 2 > select both > fill down to 7.",
        "Ctrl+H > Comp Sci → Computer Science > Replace All. F7 > Change Mathemetics to Mathematics.",
        "Right-click row 1 > Insert. A1: CENTRAL LIBRARY > select A1:E1 > Merge & Center.",
        "Click E2 > edit to 'Price ' then Insert > Symbol > ₹ > Insert > type ')'.",
        "Select B3:B9 > Wrap Text. Right-click B > Column Width 25.",
        "Insert > Pictures > Pictures\\Logo.png.",
        "G2 Website; G3: Ctrl+K > www.isro.gov.in.",
        "Right-click the 'Ignited Minds' row number > Delete.",
        "Ctrl+S. F12 > Save as type CSV (Comma delimited) > Save > OK/Yes.",
    ],
)


def _no_text(ctx, text):
    b = ctx.excel()
    for ws in b.wbv.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and text.lower() in c.value.lower():
                    return False, f"Still found in {ws.title}!{c.coordinate}"
    return True, "Not found any more"


# ------------------------------------------------------------------ POWERPOINT

p1 = Task(
    id="M4-P1", app="ppt", title="Seminar Template", file="Seminar.pptx", minutes=11,
    instructions=[
        "New presentation; apply any Design theme.",
        "Slide 1 (Title Slide): 'Seminar on Cyber Security' / 'IT Division'.",
        "Slide 2 (Title and Content): title 'Agenda'; bullets: Passwords, Phishing, Safe Browsing, Backups.",
        "Slide 3 (Title Only): title 'Reporting Structure'; insert a SmartArt Hierarchy (Organisation Chart).",
        "Slide 4 (Title Only): title 'Process'; draw two rectangles and an arrow between them.",
        "Make the title on slide 2 Bold and size 36.",
        "Insert Date and Slide Number on all slides.",
        "Save as Seminar.pptx and also as a PowerPoint Template Seminar.potx in the test folder.",
    ],
    checks=[
        C.file_saved("Seminar.pptx", 1), P.design_applied(1.5), P.slide_count(4, 1),
        P.title(1, "Seminar on Cyber Security", 0.5), P.subtitle(1, "IT Division", 0.5),
        P.title(2, "Agenda", 0.5), P.text_on(2, ["Passwords", "Phishing", "Safe Browsing", "Backups"], 1),
        P.text_bold(2, "Agenda", 0.5), P.text_size(2, "Agenda", 36, 0.5),
        P.title(3, "Reporting Structure", 0.5), P.smartart(3, 2.5),
        P.layout(4, "Title Only", 0.5), P.shape_present(4, 1.5, what="Shapes and arrow drawn on slide 4"),
        P.date_shown("all", 0.5), P.slide_numbers("all", 0.5),
        C.file_format("Seminar.potx", "potx", 2),
    ],
    solution=[
        "Design > pick a theme. Type slide 1.",
        "New Slide > Title and Content > Agenda + 4 bullets. Select 'Agenda' > Ctrl+B > 36.",
        "New Slide > Title Only > Insert > SmartArt > Hierarchy > Organization Chart.",
        "New Slide > Title Only > Insert > Shapes > Rectangle (twice) and Block Arrow / Line Arrow.",
        "Insert > Header & Footer > Date and time, Slide number > Apply to All.",
        "F12 > Seminar > Save. F12 > Save as type PowerPoint Template (*.potx) > change folder back to the test folder > Save.",
    ],
)

p2 = Task(
    id="M4-P2", app="ppt", title="Results Presentation", file="Results.pptx", minutes=11,
    instructions=[
        "New blank presentation with 3 slides; slide size Widescreen (16:9).",
        "Slide 1 (Title Slide): 'Recruitment Results', subtitle 'Skill Test 2026'. Change the title font colour.",
        "Slide 2 (Title and Content): title 'Result Summary'; insert a table of 3 rows × 3 columns:",
        "    Post | Selected | Waiting",
        "    Assistant | 45 | 10",
        "    UDC | 30 | 8",
        "Slide 3 (Title and Content): title 'Selected by Post'; insert a Line chart (any data).",
        "Apply a Gradient background to all slides.",
        "Apply a transition to slides 2 and 3 only.",
        "Footer text 'Recruitment Cell' on all slides.",
        "Save as Results.pptx.",
    ],
    checks=[
        C.file_saved("Results.pptx", 1), P.slide_size("16:9", 0.5), P.slide_count(3, 1),
        P.title(1, "Recruitment Results", 0.5), P.subtitle(1, "Skill Test 2026", 0.5),
        P.text_color(1, "Recruitment Results", 1),
        P.title(2, "Result Summary", 0.5), P.table(2, 3, 3, 1),
        P.table_text(2, ["Post", "Selected", "Assistant", "45", "UDC", "30"], 1.5),
        P.title(3, "Selected by Post", 0.5), P.chart(3, "line", 2),
        P.background("all", "gradient", 2),
        P.transition([2, 3], 1.5), P.footer("Recruitment Cell", "all", 1.5),
    ],
    solution=[
        "Blank > Design > Slide Size > Widescreen. Type slide 1; select the title > Font Color.",
        "New Slide > Title and Content > table icon > 3×3 > type.",
        "New Slide > chart icon > Line > OK.",
        "Design > Format Background > Gradient fill > Apply to All.",
        "Select slides 2 and 3 in the left pane (Ctrl+click) > Transitions > Fade (do not Apply To All).",
        "Insert > Header & Footer > Footer: Recruitment Cell > Apply to All. F12 > Results > Save.",
    ],
)

PAPER = Paper(id="mock4", title="Mock Test 4", tasks=[w1, w2, e1, e2, p1, p2])
