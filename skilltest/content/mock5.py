"""Mock Test 5."""
from .. import common as C
from .. import excel as X
from .. import ppt as P
from .. import word as W
from ..core import Paper, Task
from .starters import text_file, xlsx

# ------------------------------------------------------------------ WORD

R_INTRO = ("The Administration Division handled more than four thousand files during the year. Most of the files "
           "were cleared within seven working days with the help of the new e-office system.")
R_ACTS = ["Conducted Hindi workshops", "Organised a blood donation camp", "Completed annual stock verification"]
R_ANNEX = "Annexure: The list of officers trained during the year is kept in the training register of the division."

w1 = Task(
    id="M5-W1", app="word", title="Annual Report", file="Report.docx", minutes=13,
    starters={"Annexure.txt": text_file(R_ANNEX)},
    instructions=[
        "Start a new blank document and type:",
        "    Annual Activity Report             → style Title",
        "    Introduction                       → style Heading 1",
        f"    {R_INTRO}",
        "    Activities                         → style Heading 1",
        *[f"    {a}" for a in R_ACTS],
        "    Annexure                           → style Heading 1",
        "Give the three activity lines bullets.",
        "Introduction paragraph: font size 12, line spacing 1.15.",
        "Insert a Page Break so that 'Annexure' starts on a new page.",
        "Below 'Annexure' insert the contents of Annexure.txt using Insert > Object > Text from File.",
        "Header: Annual Report 2026. Page numbers at the bottom.",
        "Margins: Moderate.",
        "Save as Report.docx and also save as a Web Page: Report.htm.",
    ],
    checks=[
        C.file_saved("Report.docx", 1),
        W.style("Annual Activity Report", "Title", 1),
        W.style("Introduction", "Heading 1", 1), W.style("Activities", "Heading 1", 0.5), W.style("Annexure", "Heading 1", 0.5),
        W.typed(R_INTRO, 2, what="Introduction paragraph typed"),
        W.listed(R_ACTS, "bullet", 1.5),
        W.font_size(R_INTRO, 12, 1), W.line_spacing(R_INTRO, 1.15, 1),
        W.page_break(1.5),
        W.contains("list of officers trained during the year", 2, what="Text inserted from Annexure.txt",
                   topic="Importing & Exporting documents",
                   how="Insert > Object (drop-down) > Text from File > choose Annexure.txt (set file type to All Files) > Insert."),
        W.header_text("Annual Report 2026", 1), W.page_numbers(1),
        W.margins(preset="Moderate", marks=1),
        C.file_format("Report.htm", "html", 2),
    ],
    solution=[
        "Type all lines. Apply styles from Home > Styles: Title, Heading 1 (×3).",
        "Select the activity lines > Bullets. Select the intro paragraph > size 12 > Line spacing 1.15.",
        "Click just before 'Annexure' > Ctrl+Enter.",
        "Click at the end of 'Annexure' > Enter > Insert > Object ▼ > Text from File > All Files > Annexure.txt > Insert.",
        "Insert > Header > Annual Report 2026. Insert > Page Number > Bottom of Page.",
        "Layout > Margins > Moderate. F12 > Report > Save. F12 > Save as type Web Page (*.htm) > Save.",
    ],
)

Q_ITEMS = [("Desktop Computer", 5, 52000), ("Laser Printer", 2, 18500), ("UPS 1KVA", 5, 4600), ("Scanner", 1, 21500)]
Q_TOTAL = sum(q * r for _, q, r in Q_ITEMS)
Q_NOTE = "Prices are inclusive of GST and valid for thirty days."

w2 = Task(
    id="M5-W2", app="word", title="Quotation", file="Quotation.docx", minutes=12,
    instructions=[
        "Start a new document. Type QUOTATION and make it Bold, size 20, Centre.",
        "Insert a table 5 rows × 4 columns: Item | Qty | Rate (₹) | Amount (₹). Use Insert > Symbol for ₹.",
        *[f"    {n} | {q} | {r} | {q * r}" for n, q, r in Q_ITEMS],
        "Sort the table by Amount, largest first (Number type, Descending, header row).",
        "Add a last row: Total in the first cell and =SUM(ABOVE) formula in the Amount cell.",
        "Header row Bold with shading; All Borders.",
        f"Below the table type: {Q_NOTE} and make it Italic.",
        "Save as Quotation.docx. Save another copy as Quotation_Locked.docx with the password quote123.",
    ],
    checks=[
        C.file_saved("Quotation.docx", 1),
        W.bold("QUOTATION", 0.5), W.font_size("QUOTATION", 20, 0.5), W.alignment("QUOTATION", "Center", 0.5),
        W.table_size(6, 4, 1.5),
        W.table_has([n for n, _, _ in Q_ITEMS] + ["260000", "37000"], 2),
        W.symbol("₹", 1),
        W.table_sorted(3, desc=True, numeric=True, marks=2),
        W.table_formula("SUM", Q_TOTAL, 2.5),
        W.bold("Item", 0.5), W.table_shading(1), W.table_borders(0.5),
        W.italic(Q_NOTE, 1),
        C.password("Quotation_Locked.docx", "quote123", 2.5),
    ],
    solution=[
        "Type QUOTATION > Ctrl+B > 20 > Ctrl+E. Enter.",
        "Insert > Table 4×5 > type data (₹ from Insert > Symbol).",
        "Table Layout > Sort > Sort by Amount (₹), Type Number, Descending, Header row > OK.",
        "Layout > Insert Below > Total; last cell > Formula > =SUM(ABOVE) > OK.",
        "Header row Ctrl+B + Shading; Borders > All Borders.",
        "Below the table type the note > Ctrl+I. Ctrl+S as Quotation.",
        "F12 > Quotation_Locked > Tools > General Options > Password to open: quote123 > OK > re-enter > Save.",
    ],
)

# ------------------------------------------------------------------ EXCEL

STUD = [(101, "Asha", 88, 79), (102, "Bharat", 45, 52), (103, "Chandni", 67, 71), (104, "Dev", 33, 29),
        (105, "Ekta", 92, 95), (106, "Feroz", 58, 49), (107, "Gauri", 74, 81), (108, "Hitesh", 39, 44),
        (109, "Isha", 61, 66), (110, "Jatin", 80, 72)]


def _pct(r):
    return (r["C"] + r["D"]) / 200


def _grade(r):
    p = _pct(r)
    return "A" if p >= 0.75 else "B" if p >= 0.6 else "C" if p >= 0.4 else "F"


A_COUNT = sum(1 for _, _, a, b in STUD if (a + b) / 200 >= 0.75)

e1 = Task(
    id="M5-E1", app="excel", title="Exam Analysis", file="Exam.xlsx", minutes=13,
    starters={"Exam.xlsx": xlsx({"Sheet1": (["Roll", "Name", "Paper1", "Paper2"], STUD)})},
    instructions=[
        "Open Exam.xlsx. Rename Sheet1 to Result.",
        "E1: Total, E2:E11 = Paper1 + Paper2 using SUM.",
        "F1: Percent, F2:F11 = Total ÷ 200, formatted as Percentage.",
        "G1: Grade, G2:G11 nested IF: A if Percent ≥ 75%, B if ≥ 60%, C if ≥ 40%, otherwise F.",
        "   e.g. =IF(F2>=0.75,\"A\",IF(F2>=0.6,\"B\",IF(F2>=0.4,\"C\",\"F\")))",
        "I1: Number of A grades, J1: COUNTIF on G2:G11 for \"A\".",
        "I2: Class Average, J2: AVERAGE of E2:E11.",
        "Conditional formatting: make Grade F cells red (Highlight Cells Rules > Equal To > F).",
        "Freeze the top row.",
        "Insert a Line chart of Paper1 and Paper2 (with names) titled Paper-wise Marks.",
        "Turn on AutoFilter and show only Grade A rows.",
        "Save (Ctrl+S).",
    ],
    checks=[
        X.sheet_exists("Result", 1),
        X.rowwise("Result|Sheet1", "E", 2, 11, "SUM", lambda r: r["C"] + r["D"], 2),
        X.rowwise("Result|Sheet1", "F", 2, 11, "", _pct, 2, what="Percent = Total ÷ 200", how="F2: =E2/200, fill down."),
        X.number_format("Result|Sheet1", "F2:F11", "percent", 1),
        X.rowwise("Result|Sheet1", "G", 2, 11, "IF", _grade, 3.5, what="Grade with nested IF"),
        X.formula("Result|Sheet1", "J1", "COUNTIF", A_COUNT, 1.5, how='J1: =COUNTIF(G2:G11,"A")'),
        X.formula("Result|Sheet1", "J2", "AVERAGE", sum(a + b for _, _, a, b in STUD) / 10, 1.5),
        X.cond_format("Result|Sheet1", 1.5),
        X.freeze("Result|Sheet1", "A2", 1),
        X.chart("line", "Paper-wise Marks", 2),
        X.autofilter("Result|Sheet1", 1),
    ],
    solution=[
        "Double-click Sheet1 tab > Result.",
        "E2: =SUM(C2:D2)  F2: =E2/200  G2: nested IF. Select E2:G2 > double-click the fill handle. F2:F11 > %.",
        "J1: =COUNTIF(G2:G11,\"A\")  J2: =AVERAGE(E2:E11).",
        "G2:G11 > Conditional Formatting > Highlight Cells Rules > Equal To > F > Light Red Fill.",
        "View > Freeze Panes > Freeze Top Row.",
        "Select B1:D11 > Insert > Line > Line with Markers > title Paper-wise Marks.",
        "Data > Filter > Grade arrow > only A. Ctrl+S.",
    ],
)

WAGES = [("Murugan", 26, 650), ("Savitri", 0, 600), ("Prakash", 24, 650), ("Rekha", 22, 600), ("Tarun", 0, 700),
         ("Uma", 25, 650)]

e2 = Task(
    id="M5-E2", app="excel", title="Payroll Sheets", file="Payroll.xlsx", minutes=12,
    starters={"Payroll.xlsx": xlsx({"Jan": (["Name", "Days Present", "Daily Wage"], WAGES),
                                    "Notes": (["Rough"], [("delete me",)])})},
    instructions=[
        "Open Payroll.xlsx.",
        "Delete the sheet Notes.",
        "On Jan: D1 Wage, D2:D7 = Days Present × Daily Wage.",
        "E1: Per Day. E2:E7 = Wage ÷ Days Present. Rows with 0 days give #DIV/0! – "
        "wrap the formula in IFERROR so that they show 0 instead.",
        "Hide column C.",
        "Make a copy of the Jan sheet (right-click > Move or Copy > Create a copy) and rename the copy Feb.",
        "On Jan: print header Payroll January, print titles row 1, Landscape.",
        "Protect the Jan sheet.",
        "Save (Ctrl+S). Also save a copy as an Excel Template Payroll.xltx in the test folder.",
    ],
    checks=[
        X.sheet_absent("Notes", 1),
        X.rowwise("Jan", "D", 2, 7, r"\*|PRODUCT\(", lambda r: r["B"] * r["C"], 2.5, what="Wage = Days × Daily Wage",
                  how="D2: =B2*C2, fill down."),
        X.rowwise("Jan", "E", 2, 7, "IFERROR", lambda r: (r["D"] / r["B"]) if r["B"] else 0, 3,
                  what="IFERROR used, no #DIV/0!", how="E2: =IFERROR(D2/B2,0), fill down.", topic="Error checking"),
        X.no_errors("Jan", "E2:E7", 1),
        X.hidden_col("Jan", "C", 1.5),
        X.sheet_exists("Feb", 2),
        X.sheet_header("Jan", "Payroll January", 1), X.print_titles("Jan", "1:1", 1), X.landscape("Jan", 0.5),
        X.protected("Jan", 1.5),
        C.file_format("Payroll.xltx", "xltx", 2),
    ],
    solution=[
        "Right-click Notes > Delete.",
        "D2: =B2*C2. E2: =IFERROR(D2/B2,0). Select D2:E2 > double-click the fill handle.",
        "Right-click column C header > Hide.",
        "Right-click Jan > Move or Copy > tick Create a copy > OK > rename 'Jan (2)' to Feb.",
        "On Jan: Page Layout > Print Titles > Rows to repeat $1:$1; Header/Footer > Custom Header > Payroll January; "
        "Orientation Landscape.",
        "Review > Protect Sheet > OK. Ctrl+S.",
        "F12 > Save as type Excel Template (*.xltx) > go back to the test folder > Payroll > Save.",
    ],
)

# ------------------------------------------------------------------ POWERPOINT

p1 = Task(
    id="M5-P1", app="ppt", title="Gaganyaan", file="Gaganyaan.pptx", minutes=11,
    instructions=[
        "New presentation; apply any Design theme.",
        "Slide 1 (Title Slide): 'Gaganyaan Mission' / 'India's Human Spaceflight Programme'.",
        "Slide 2 (Section Header): title 'Objectives'.",
        "Slide 3 (Title and Content): title 'Key Points'; bullets: Crew of three, Low Earth Orbit, Safe return.",
        "Slide 4 (Two Content): title 'Launch Vehicle'; left: text 'Human-rated LVM3'; right: picture Rocket.png.",
        "Slide 5 (Title Only): title 'Thank You'; add a Text Box with 'Jai Hind'.",
        "Animate the bullets on slide 3. Transitions on all slides.",
        "Speaker note on slide 4: Show the rocket model.",
        "Save as Gaganyaan.pptx.",
    ],
    checks=[
        C.file_saved("Gaganyaan.pptx", 1), P.design_applied(1), P.slide_count(5, 1),
        P.title(1, "Gaganyaan Mission", 0.5), P.subtitle(1, "Human Spaceflight Programme", 0.5),
        P.layout(2, "Section Header", 1), P.title(2, "Objectives", 0.5),
        P.text_on(3, ["Crew of three", "Low Earth Orbit", "Safe return"], 1), P.animation(3, 1.5),
        P.layout(4, "Two Content", 0.5), P.picture(4, 1), P.text_on(4, ["Human-rated LVM3"], 0.5),
        P.textbox(5, "Jai Hind", 1.5), P.transition("all", 1.5), P.notes(4, "Show the rocket model", 1.5),
        P.title(5, "Thank You", 0.5),
    ],
    solution=[
        "Design > theme. Slide 1 text.",
        "New Slide > Section Header > Objectives. New Slide > Title and Content > bullets > Animations > Fly In.",
        "New Slide > Two Content > text left, picture icon right > Rocket.png. Notes: Show the rocket model.",
        "New Slide > Title Only > Thank You > Insert > Text Box > Jai Hind.",
        "Transitions > Apply To All. F12 > Gaganyaan > Save.",
    ],
)

p2 = Task(
    id="M5-P2", app="ppt", title="Office Layout", file="OfficeLayout.pptx", minutes=11,
    instructions=[
        "New presentation with 2 slides.",
        "Slide 1 (Title Slide): 'New Office Layout' / 'Block A, Ground Floor'.",
        "Apply a Picture background (Logo.png) to all slides.",
        "Slide 2 (Title Only): title 'Seating Plan'. Draw 4 rectangles (cabins). Align them Left, then Group them.",
        "Give one rectangle a Gradient fill.",
        "Add a Text Box 'Reception' on slide 2 and change its font colour.",
        "Insert Date, Slide Number and footer 'Estate Office' on all slides.",
        "Save as OfficeLayout.pptx and OfficeLayout.pdf.",
    ],
    checks=[
        C.file_saved("OfficeLayout.pptx", 1), P.slide_count(2, 1),
        P.title(1, "New Office Layout", 0.5), P.subtitle(1, "Block A, Ground Floor", 0.5),
        P.background("all", "picture", 1.5),
        P.title(2, "Seating Plan", 0.5), P.aligned(2, "left", 1.5), P.grouped(2, 1.5), P.shape_gradient(2, 1.5),
        P.textbox(2, "Reception", 1), P.text_color(2, "Reception", 1),
        P.date_shown("all", 0.5), P.slide_numbers("all", 0.5), P.footer("Estate Office", "all", 1),
        C.file_format("OfficeLayout.pdf", "pdf", 1.5),
    ],
    solution=[
        "Type slide 1. Design > Format Background > Picture fill > Insert > Logo.png > Apply to All.",
        "New Slide > Title Only > Seating Plan. Draw 4 rectangles. Select all > Align Left > Group.",
        "Click one rectangle inside the group > Shape Fill > Gradient.",
        "Insert > Text Box > Reception > Font Color.",
        "Insert > Header & Footer > Date, Slide number, Footer 'Estate Office' > Apply to All.",
        "F12 > OfficeLayout > Save. F12 > PDF > Save.",
    ],
)

PAPER = Paper(id="mock5", title="Mock Test 5", tasks=[w1, w2, e1, e2, p1, p2])
