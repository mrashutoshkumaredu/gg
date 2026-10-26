"""Mock Test 3."""
from .. import common as C
from .. import excel as X
from .. import ppt as P
from .. import word as W
from ..core import Paper, Task
from .starters import docx_with, xlsx

# ------------------------------------------------------------------ WORD

NEWS_P1 = ("The Indian Space Research Organisation successfully launched an earth observation satellite from "
           "Sriharikota this month. The satellite will help farmers and disaster management agencies with timely data.")
NEWS_P2 = ("The canteen committee has decided to introduce millet based food items from next week. Staff members are "
           "invited to give their suggestions in the box kept near the reception.")
NEWS_NOTE = "Note: Suggestions must reach the committee before 30 November."
NEWS_LAST = "ISRO 2026"

w1 = Task(
    id="M3-W1", app="word", title="Newsletter", file="Newsletter.docx", minutes=13,
    instructions=[
        "Start a new blank document.",
        "Insert the heading ISRO NEWS as WordArt (Insert > WordArt).",
        "Type these paragraphs:",
        f"    {NEWS_P1}",
        f"    {NEWS_P2}",
        f"    {NEWS_NOTE}",
        f"    © {NEWS_LAST}     (insert © with Insert > Symbol)",
        "Put the first two paragraphs in Two columns. Insert a Column Break before the second paragraph so it starts in column 2.",
        "Apply a Drop Cap to the first paragraph.",
        "Give the Note paragraph a box border and a shading colour (Borders and Shading).",
        "Set Spacing After of the Note paragraph to 12 pt.",
        "Make the last line (© ISRO 2026) Italic.",
        "Insert the picture Satellite.png and draw any one shape (e.g. a star).",
        "Give the page a light Page Colour. Add the footer text Monthly Newsletter.",
        "Save as Newsletter.docx.",
    ],
    checks=[
        C.file_saved("Newsletter.docx", 1),
        W.wordart("ISRO NEWS", 1.5),
        W.typed(NEWS_P1, 2, what="First paragraph typed"), W.typed(NEWS_P2, 1, what="Second paragraph typed"),
        W.columns(2, 1.5), W.column_break(1), W.drop_cap(1.5),
        W.para_border_or_shading(NEWS_NOTE, 1.5), W.spacing_after(NEWS_NOTE, 12, 1),
        W.symbol("©", 1), W.italic(NEWS_LAST, 1),
        W.picture(1, 1), W.shape(1),
        W.page_color(1), W.footer_text("Monthly Newsletter", 1),
    ],
    solution=[
        "Insert > WordArt > pick a style > type ISRO NEWS. Click below it and type the paragraphs.",
        "Select the first two paragraphs > Layout > Columns > Two.",
        "Click at the start of paragraph 2 > Layout > Breaks > Column.",
        "Click in paragraph 1 > Insert > Drop Cap > Dropped.",
        "Select the Note paragraph > Borders > Borders and Shading > Box; Shading tab > Fill colour > Apply to Paragraph > OK. "
        "Layout > Spacing After: 12 pt.",
        "For ©: Insert > Symbol > © (or Alt+Ctrl+C). Select the last line > Ctrl+I.",
        "Insert > Pictures > Pictures\\Satellite.png. Insert > Shapes > Star > drag on the page.",
        "Design > Page Color > a light colour. Insert > Footer > Edit > type Monthly Newsletter.",
        "F12 > Newsletter > Save.",
    ],
)

R1 = "Casual leave of eight days is allowed in a calendar year."
R2 = "Earned leave can be accumulated up to three hundred days. This rule is temporary."
R3 = "Medical leave requires a certificate from an authorised doctor."
NEW = "Leave must be applied through the online portal."

w2 = Task(
    id="M3-W2", app="word", title="Review a Policy", file="Policy_Reviewed.docx", minutes=12,
    starters={"Policy.docx": docx_with([R1, R2, R3], title="Leave Policy")},
    instructions=[
        "Open Policy.docx.",
        "Apply the style Title to the first line 'Leave Policy'.",
        "Give the three rules Numbering (1, 2, 3).",
        "Turn ON Track Changes (Review > Track Changes). Then:",
        "    a) delete the sentence 'This rule is temporary.'",
        f"    b) add a new rule at the end: {NEW}",
        "    c) add a comment on the word 'Casual' with the text Check with HR.",
        "Save as Policy_Reviewed.docx (keep the tracked changes – do not accept them).",
        "Save another copy as Policy_Secure.docx protected with the password isro@123.",
        "Save another copy as Plain Text: Policy.txt.",
    ],
    checks=[
        C.file_saved("Policy_Reviewed.docx", 1),
        W.style("Leave Policy", "Title", 1),
        W.listed([R1, "Earned leave can be accumulated", R3], "number", 2),
        W.track_changes(2),
        W.not_contains("This rule is temporary", 2, what="Sentence 'This rule is temporary.' deleted",
                       how="With Track Changes on, select the sentence and press Delete."),
        W.contains(NEW, 2, what="New rule added"),
        W.comment(2, text="Check with HR"),
        C.password("Policy_Secure.docx", "isro@123", 3),
        C.file_format("Policy.txt", "txt", 2),
    ],
    solution=[
        "Open Policy.docx. Click in 'Leave Policy' > Styles > Title.",
        "Select the three rules > Home > Numbering.",
        "Review > Track Changes (Ctrl+Shift+E) so it is highlighted.",
        "Select 'This rule is temporary.' > Delete. Click at the end of rule 3 > Enter > type the new rule.",
        "Select 'Casual' > Review > New Comment > Check with HR.",
        "F12 > Policy_Reviewed > Save.",
        "F12 > Policy_Secure > Save. Then File > Info > Protect Document > Encrypt with Password > isro@123 twice > Ctrl+S. "
        "(Or in the Save As box: Tools > General Options > Password to open.)",
        "F12 > Save as type: Plain Text > Policy > Save > OK.",
    ],
)

# ------------------------------------------------------------------ EXCEL

NAMES = ["Anil", "Bindu", "Charan", "Deepika", "Elango", "Fathima", "Girish", "Hema", "Irfan", "Janaki"]
ATT = []
for i, n in enumerate(NAMES):
    row = [n] + [("A" if (d * (i + 3)) % (5 + i % 4) == 0 else "P") for d in range(1, 21)]
    ATT.append(row)
PRESENT = {r[0]: r[1:].count("P") for r in ATT}

e1 = Task(
    id="M3-E1", app="excel", title="Attendance Register", file="Attendance.xlsx", minutes=13,
    starters={"Attendance.xlsx": xlsx({"March": (["Name"], ATT)})},
    instructions=[
        "Open Attendance.xlsx (sheet March). Columns B to U hold P/A for 20 days; row 1 has no day numbers yet.",
        "Rename the sheet to March 2026.",
        "Fill B1:U1 with the numbers 1 to 20 using AutoFill/Series.",
        "Set the width of columns B to U to 4.",
        "V1: Present. In V2:V11 use COUNTIF to count the P's in each row.",
        "W1: Percent. In W2:W11 calculate Present ÷ 20 and format as Percentage.",
        "Make row 1 (A1:W1) Bold and centre B1:W1.",
        "Conditional formatting on W2:W11: highlight values Less Than 75%.",
        "In Y1 enter today's date using the TODAY function.",
        "Freeze panes at B2 (both the names column and the day row stay visible).",
        "Make a Bar chart of the names and their Percent with the title Attendance %, and show data labels.",
        "Save (Ctrl+S).",
    ],
    checks=[
        X.sheet_exists("March 2026", 0.5),
        X.series("March 2026|March", "B1:U1", 1, 1, 2),
        X.col_width("March 2026|March", "B", 4, 1),
        X.rowwise("March 2026|March", "V", 2, 11, "COUNTIF", lambda r: PRESENT[r["A"]], 3,
                  what="Present days counted with COUNTIF", how='In V2 type =COUNTIF(B2:U2,"P") and fill down.'),
        X.rowwise("March 2026|March", "W", 2, 11, "", lambda r: PRESENT[r["A"]] / 20, 2,
                  what="Percent = Present ÷ 20", how="In W2 type =V2/20 and fill down."),
        X.number_format("March 2026|March", "W2:W11", "percent", 1),
        X.bold("March 2026|March", "A1:W1", 1), X.align("March 2026|March", "B1:W1", "center", 1),
        X.cond_format("March 2026|March", 1.5),
        X.formula("March 2026|March", "Y1", "TODAY", None, 1),
        X.freeze("March 2026|March", "B2", 1),
        X.chart("bar", "Attendance %", 2), X.chart_data_labels(1),
    ],
    solution=[
        "Double-click the March tab > March 2026.",
        "B1: 1, C1: 2 > select both > drag the fill handle to U1.",
        "Select column headers B to U > right-click > Column Width > 4.",
        "V1 Present, V2: =COUNTIF(B2:U2,\"P\"). W1 Percent, W2: =V2/20. Select V2:W2 > double-click fill handle.",
        "Select W2:W11 > % button. Select A1:W1 > Ctrl+B; B1:W1 > Center.",
        "W2:W11 > Conditional Formatting > Highlight Cells Rules > Less Than > 75% > OK.",
        "Y1: =TODAY(). Click B2 > View > Freeze Panes > Freeze Panes.",
        "Select A1:A11 and (Ctrl) W1:W11 > Insert > Bar > Clustered Bar > title Attendance % > + > Data Labels.",
        "Ctrl+S.",
    ],
)

HEADS = ["Travel", "Stationery", "Electricity", "Telephone"]
QDATA = {"Q1": [42000, 8500, 61000, 7200], "Q2": [38500, 9100, 58000, 6900], "Q3": [45500, 7800, 66000, 7400]}
QTOTAL = {h: sum(QDATA[q][i] for q in QDATA) for i, h in enumerate(HEADS)}
BILLS = [("Library", "Books", 18500), ("Accounts", "Calculator", 2400), ("Admin", "Chairs", 32000),
         ("Library", "Racks", 26000), ("Purchase", "Files", 3500), ("Accounts", "Ledger", 1800),
         ("Admin", "Water Cooler", 21000), ("Purchase", "Printer", 14500), ("Library", "Journals", 9500),
         ("Accounts", "Software", 12000), ("Admin", "Curtains", 7600), ("Purchase", "Scanner", 9900)]
DEPT_TOTAL = {}
for d, _, a in BILLS:
    DEPT_TOTAL[d] = DEPT_TOTAL.get(d, 0) + a

e2 = Task(
    id="M3-E2", app="excel", title="Quarterly Expenses", file="Expenses.xlsx", minutes=12,
    starters={"Expenses.xlsx": xlsx({**{q: (["Head", "Amount"], list(zip(HEADS, v))) for q, v in QDATA.items()},
                                     "Bills": (["Department", "Item", "Amount"], BILLS)})},
    instructions=[
        "Open Expenses.xlsx. Sheets Q1, Q2 and Q3 have expenses by head; sheet Bills has bills by department.",
        "Insert a new sheet named Total. Click A1 and use Data > Consolidate (Function: Sum) to add up "
        "Q1!A1:B5, Q2!A1:B5 and Q3!A1:B5, with 'Top row' and 'Left column' ticked.",
        "Give the Total sheet tab a colour and make a Column chart of the totals (any title).",
        "On Bills: sort by Department (A to Z), then use Data > Subtotal: at each change in Department, "
        "use Sum, add subtotal to Amount.",
        "Format the Amount column on Bills as Currency.",
        "Bills print setup: Landscape, footer text Expense Report, and insert a page break above row 8.",
        "Save (Ctrl+S).",
    ],
    checks=[
        X.sheet_exists("Total", 1),
        X.labelled_values("Total", "A", "B", QTOTAL, 4, what="Consolidated totals on Total sheet",
                          how="Total sheet A1 > Data > Consolidate > Sum > add each Q sheet range A1:B5 > tick Top row "
                              "and Left column > OK.", topic="consolidation and subtotal"),
        X.tab_color("Total", 1),
        X.chart("column", None, 2, what="Column chart of the totals"),
        X.subtotal(3, sheet="Bills", min_groups=4),
        X.labelled_values("Bills", "A", "C", {f"{d} Total": v for d, v in DEPT_TOTAL.items()}, 3,
                          what="Correct subtotal for each Department",
                          how="Sort by Department first, then Data > Subtotal.", topic="consolidation and subtotal"),
        X.number_format("Bills", "C2:C13", "currency", 1),
        X.landscape("Bills", 0.5), X.sheet_header("Bills", "Expense Report", 1, footer=True), X.page_break("Bills", 0.5),
    ],
    solution=[
        "Click + > rename the new sheet Total. Click A1 > Data > Consolidate > Function Sum.",
        "Reference: click Q1 tab, select A1:B5 > Add. Repeat for Q2 and Q3. Tick Top row and Left column > OK.",
        "Right-click Total tab > Tab Color. Select A2:B5 > Insert > Column > Clustered Column.",
        "Bills: click in the data > Data > Sort A to Z on Department. Data > Subtotal > At each change in: Department, "
        "Use function: Sum, Add subtotal to: Amount > OK.",
        "Select the Amount column cells > Currency.",
        "Page Layout > Orientation Landscape. Insert > Header & Footer > footer: Expense Report. "
        "Click A8 > Page Layout > Breaks > Insert Page Break.",
        "Ctrl+S.",
    ],
)

# ------------------------------------------------------------------ POWERPOINT

p1 = Task(
    id="M3-P1", app="ppt", title="Annual Review", file="Annual.pptx", minutes=11,
    instructions=[
        "New presentation; apply any Design theme.",
        "Slide 1 (Title Slide): 'Annual Review 2026' / 'Space Applications Centre'.",
        "Slide 2 (Two Content): title 'Highlights'; left box bullets: 12 launches, 5 new centres; "
        "right box: insert a Pie chart (any data). Animate the chart.",
        "Slide 3 (Title and Content): title 'Staff Strength'; insert a table of 4 rows × 2 columns:",
        "    Cadre | Number",
        "    Assistant | 120",
        "    UDC | 85",
        "    JPA | 60",
        "Slide 4 (Picture with Caption): insert Satellite.png in the picture box; title 'Our Satellites'.",
        "Footer 'Annual Review' and slide numbers on all slides. Transition on all slides.",
        "Save as Annual.pptx.",
    ],
    checks=[
        C.file_saved("Annual.pptx", 1), P.design_applied(1), P.slide_count(4, 1),
        P.title(1, "Annual Review 2026", 0.5), P.subtitle(1, "Space Applications Centre", 0.5),
        P.layout(2, "Two Content", 0.5), P.text_on(2, ["12 launches", "5 new centres"], 0.5), P.chart(2, "pie", 2),
        P.animation(2, 1),
        P.table(3, 4, 2, 1), P.table_text(3, ["Cadre", "Assistant", "120", "UDC", "85", "JPA", "60"], 1.5),
        P.layout(4, "Picture with Caption", 0.5), P.picture(4, 1), P.title(4, "Our Satellites", 0.5),
        P.footer("Annual Review", "all", 1), P.slide_numbers("all", 1), P.transition("all", 0.5),
    ],
    solution=[
        "Blank presentation > Design > theme. Type title & subtitle.",
        "New Slide > Two Content > type bullets left; click the chart icon in the right box > Pie > OK. "
        "Select the chart > Animations > Wipe.",
        "New Slide > Title and Content > table icon > 2 columns, 4 rows > type data.",
        "New Slide > Picture with Caption > click the picture icon > Satellite.png > type the title.",
        "Insert > Header & Footer > Slide number, Footer: Annual Review > Apply to All. Transitions > Apply To All.",
        "F12 > Annual > Save.",
    ],
)

p2 = Task(
    id="M3-P2", app="ppt", title="Quiz Slide Design", file="Quiz.pptx", minutes=10,
    instructions=[
        "New presentation with 2 slides.",
        "Slide 1 (Title Slide): 'General Awareness Quiz', subtitle 'Round 1'.",
        "Apply a Solid fill background colour to all slides.",
        "Slide 2 (Blank): draw three shapes (any). Give one of them a gradient fill. "
        "Align all three to the Middle, then Group them.",
        "On slide 2 add a Text Box with the text Next Round and make it a hyperlink to www.isro.gov.in.",
        "Add the speaker note 'Read questions slowly' on slide 2.",
        "Save as Quiz.pptx and also as Quiz.pdf.",
    ],
    checks=[
        C.file_saved("Quiz.pptx", 1), P.slide_count(2, 1),
        P.title(1, "General Awareness Quiz", 0.5), P.subtitle(1, "Round 1", 0.5),
        P.background("all", "solid", 2),
        P.layout(2, "Blank", 0.5), P.shape_gradient(2, 1.5), P.aligned(2, "middle", 1.5), P.grouped(2, 1.5),
        P.textbox(2, "Next Round", 1), P.hyperlink(2, 1.5),
        P.notes(2, "Read questions slowly", 1),
        C.file_format("Quiz.pdf", "pdf", 1.5),
    ],
    solution=[
        "Type the title slide text. Design > Format Background > Solid fill > colour > Apply to All.",
        "New Slide > Blank > Insert > Shapes three times. Select one > Shape Fill > Gradient.",
        "Select all three (Shift+click) > Shape Format > Align > Align Middle > Group.",
        "Insert > Text Box > Next Round > select text > Ctrl+K > www.isro.gov.in.",
        "Notes: Read questions slowly. F12 > Quiz > Save; F12 > PDF > Save.",
    ],
)

PAPER = Paper(id="mock3", title="Mock Test 3", tasks=[w1, w2, e1, e2, p1, p2])
