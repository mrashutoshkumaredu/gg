"""Mock Test 1."""
from .. import common as C
from .. import excel as X
from .. import ppt as P
from .. import word as W
from ..core import Paper, Task
from .starters import xlsx

# ------------------------------------------------------------------ WORD

N_HEAD = "OFFICE NOTICE"
N_P1 = ("All employees of the Administration Division are informed that the annual skill development workshop "
        "will be held from 10 November to 14 November in the main conference hall. Attendance is compulsory for "
        "all Assistants and Upper Division Clerks.")
N_P2 = ("Participants must bring their identity cards and report fifteen minutes before the session begins. "
        "Lunch will be provided by the office canteen.")
N_ITEMS = ["Bring a notebook and pen", "Switch off mobile phones during sessions", "Sign the attendance register daily"]
N_SIGN = "Head of Administration"

w1 = Task(
    id="M1-W1", app="word", title="Office Notice", file="Notice.docx", minutes=13,
    instructions=[
        "Open MS Word and start a new blank document.",
        "Type the text below exactly (heading, two paragraphs, three item lines and the last line):",
        f"    {N_HEAD}",
        f"    {N_P1}",
        f"    {N_P2}",
        *[f"    {i}" for i in N_ITEMS],
        f"    {N_SIGN}",
        f"Format the heading '{N_HEAD}': Bold, Underline, font size 16, Centre aligned.",
        "Both paragraphs: Justify alignment and line spacing 1.5.",
        "Give the three item lines bullets.",
        f"'{N_SIGN}': Right aligned and Italic.",
        "Page setup: Margins = Narrow, Paper size = A4.",
        "Footer: type 'Administration Division, ISRO' and also insert page numbers at the bottom of the page.",
        "Save as Notice.docx in the test folder. Then also save a copy as Notice.pdf (PDF format).",
    ],
    checks=[
        C.file_saved("Notice.docx", 1),
        W.typed(N_P1, 1.5, what="First paragraph typed correctly"),
        W.typed(N_P2, 1, what="Second paragraph typed correctly"),
        W.bold(N_HEAD, 1), W.underline(N_HEAD, 1), W.font_size(N_HEAD, 16, 1), W.alignment(N_HEAD, "Center", 1),
        W.alignment(N_P1, "Justify", 0.75), W.alignment(N_P2, "Justify", 0.75),
        W.line_spacing(N_P1, 1.5, 0.75), W.line_spacing(N_P2, 1.5, 0.75),
        W.listed(N_ITEMS, "bullet", 1.5),
        W.alignment(N_SIGN, "Right", 0.5), W.italic(N_SIGN, 0.5),
        W.margins(preset="Narrow", marks=1), W.paper_size("A4", 0.5),
        W.footer_text("Administration Division, ISRO", 1), W.page_numbers(1),
        C.file_format("Notice.pdf", "pdf", 1.5),
    ],
    solution=[
        "Start Word > Blank document. Type all the lines, pressing Enter after each line.",
        "Select the heading > Ctrl+B, Ctrl+U > Font Size box: 16 > Ctrl+E (centre).",
        "Select both paragraphs > Ctrl+J (justify) > Home > Line and Paragraph Spacing > 1.5.",
        "Select the three item lines > Home > Bullets.",
        f"Click in '{N_SIGN}' > Ctrl+R (right) > select it > Ctrl+I.",
        "Layout (Page Layout) > Margins > Narrow. Layout > Size > A4.",
        "Insert > Footer > Edit Footer > type 'Administration Division, ISRO' > Enter > "
        "Page Number > Current Position > Plain Number > Close Header and Footer.",
        "F12 > choose the test folder > File name: Notice > type Word Document > Save.",
        "F12 again > Save as type: PDF > File name Notice > Save.",
    ],
)

STAFF = [("1", "Ravi Kumar", "Assistant", "35400"), ("2", "Anita Sharma", "UDC", "25500"),
         ("3", "Mohan Das", "JPA", "21700"), ("4", "Deepa Nair", "Stenographer", "25500"),
         ("5", "Suresh Babu", "Assistant", "35400")]
S_TITLE = "STAFF DETAILS - ACCOUNTS SECTION"

w2 = Task(
    id="M1-W2", app="word", title="Staff Table", file="Staff.docx", minutes=12,
    instructions=[
        "Start a new blank Word document.",
        f"Type the title '{S_TITLE}' and format it: font Arial, size 14, Bold.",
        "Below the title insert a table with 4 columns and 6 rows. Fill it:",
        "    S.No | Name | Designation | Basic Pay (₹)",
        *[f"    {' | '.join(r)}" for r in STAFF],
        "Insert the ₹ symbol with Insert > Symbol (or type it if your keyboard has it).",
        "Sort the table by Name in ascending order (the header row must stay on top).",
        "Add one more row at the bottom. Merge its first three cells and type 'Total'. In the last cell use "
        "Table Tools Layout > Formula to calculate the total Basic Pay (=SUM(ABOVE)).",
        "Make the header row Bold and give it a shading colour. Keep borders on the table.",
        "Change the page orientation to Landscape.",
        "Add the header text 'Confidential'.",
        "Save as Staff.docx.",
    ],
    checks=[
        C.file_saved("Staff.docx", 1),
        W.font_name(S_TITLE, "Arial", 1), W.font_size(S_TITLE, 14, 1), W.bold(S_TITLE, 1),
        W.table_size(7, 4, 1),
        W.table_has([r[1] for r in STAFF] + ["Stenographer", "21700", "Basic Pay"], 1.5),
        W.symbol("₹", 1),
        W.table_sorted(1, marks=2),
        W.table_merged(1),
        W.table_formula("SUM", 143500, 2),
        W.bold("Designation", 1),
        W.table_shading(1), W.table_borders(0.5),
        W.orientation("landscape", 1),
        W.header_text("Confidential", 1),
    ],
    solution=[
        f"Type the title > select it > Font: Arial, Size: 14 > Ctrl+B. Press Enter at the end of the title.",
        "Insert > Table > Insert Table > Columns 4, Rows 6 > OK. Type the data, using Tab to move between cells.",
        "For ₹: Insert > Symbol > More Symbols > Subset: Currency Symbols > ₹ > Insert.",
        "Click in the table > Table Tools Layout > Sort > Sort by: Name, Type: Text, Ascending; My list has: Header row > OK.",
        "Click in the last row > Layout > Insert Below. Select its first 3 cells > Merge Cells > type Total.",
        "Click the last cell > Layout > Formula > =SUM(ABOVE) > OK (it shows 143500).",
        "Select the header row > Ctrl+B > Table Design > Shading > pick a colour. Borders > All Borders.",
        "Layout > Orientation > Landscape.",
        "Insert > Header > Blank > type Confidential > Close Header and Footer.",
        "F12 > test folder > File name Staff > Save.",
    ],
)

# ------------------------------------------------------------------ EXCEL

MARKS = [("Aarav", 78, 85, 69), ("Bhavna", 45, 38, 52), ("Chetan", 92, 88, 95), ("Divya", 35, 30, 41),
         ("Eshan", 67, 72, 58), ("Farah", 55, 61, 49), ("Gopal", 28, 35, 33), ("Harini", 81, 79, 90),
         ("Imran", 60, 58, 66), ("Jyoti", 73, 69, 77)]
TOTALS = {n: a + b + c for n, a, b, c in MARKS}


def _tot(r):
    return r["B"] + r["C"] + r["D"]


e1 = Task(
    id="M1-E1", app="excel", title="Class Result", file="Result.xlsx", minutes=13,
    starters={"Marks.xlsx": xlsx({"Sheet1": (["Name", "Maths", "Science", "English"], MARKS)})},
    instructions=[
        "Open Marks.xlsx from the test folder (it has 10 students' marks).",
        "Insert a new row above row 1. In A1 type CLASS RESULT 2026. Merge & Centre A1:G1, make it Bold, size 14.",
        "Now the headings are in row 2 and the data in rows 3 to 12.",
        "Type the headings E2: Total, F2: Average, G2: Result.",
        "E3:E12 – total of the three subjects using SUM.",
        "F3:F12 – average of the three subjects using AVERAGE, shown with 2 decimal places.",
        "G3:G12 – use IF: \"Pass\" if Average is 40 or more, otherwise \"Fail\".",
        "Make A2:G2 Bold with a fill colour. Apply All Borders to A2:G12.",
        "Sort the student rows (A3:G12) by Total, largest to smallest.",
        "Freeze panes so that rows 1 and 2 stay visible while scrolling (select A3 > Freeze Panes).",
        "Make a Clustered Column chart of Name and Total. Chart title: Total Marks.",
        "Save the workbook as Result.xlsx (File > Save As, new name).",
    ],
    checks=[
        C.file_saved("Result.xlsx", 1),
        X.value(None, "A1", "CLASS RESULT 2026", 1, what="New row inserted with title in A1"),
        X.merged(None, "A1:G1", 1), X.bold(None, "A1", 0.5), X.font_size(None, "A1", 14, 0.5),
        X.text_in(None, "E2:G2", ["Total", "Average", "Result"], 1, what="Headings Total / Average / Result typed"),
        X.rowwise(None, "E", 3, 12, "SUM", _tot, 2, what="Total (SUM) in E3:E12"),
        X.rowwise(None, "F", 3, 12, "AVERAGE", lambda r: _tot(r) / 3, 1.5, what="Average in F3:F12"),
        X.number_format(None, "F3:F12", "decimal2", 1),
        X.rowwise(None, "G", 3, 12, "IF", lambda r: "Pass" if _tot(r) / 3 >= 40 else "Fail", 2,
                  what="Result with IF in G3:G12",
                  how='In G3 type =IF(F3>=40,"Pass","Fail") and fill down.'),
        X.bold(None, "A2:G2", 0.5), X.fill(None, "A2:G2", 1),
        X.borders(None, "A2:G12", 1),
        X.sorted_by(None, "E", 3, 12, desc=True, numeric=True, keep=("A", "E", TOTALS), marks=2),
        X.freeze(None, "A3", 1),
        X.chart("column", "Total Marks", 1),
    ],
    solution=[
        "Open Marks.xlsx. Right-click row number 1 > Insert.",
        "In A1 type CLASS RESULT 2026. Select A1:G1 > Home > Merge & Center > Ctrl+B > size 14.",
        "Type Total, Average, Result in E2, F2, G2.",
        "E3: =SUM(B3:D3)   F3: =AVERAGE(B3:D3)   G3: =IF(F3>=40,\"Pass\",\"Fail\").",
        "Select E3:G3 and double-click the fill handle to copy down to row 12.",
        "Select F3:F12 > Home > Increase/Decrease Decimal until 2 places (or Ctrl+1 > Number > 2 decimals).",
        "Select A2:G2 > Ctrl+B > Fill Color. Select A2:G12 > Borders > All Borders.",
        "Select A2:G12 > Data > Sort > Sort by Total, Largest to Smallest (My data has headers ticked) > OK.",
        "Click A3 > View > Freeze Panes > Freeze Panes.",
        "Select A2:A12, hold Ctrl and select E2:E12 > Insert > Column > Clustered Column. "
        "Click the chart title > type Total Marks.",
        "F12 > test folder > File name Result > Save.",
    ],
)

REGIONS = ["North", "South", "East", "West"]
PRODUCTS = ["Printer", "Scanner", "Laptop", "Monitor", "UPS"]
RATES = {"Printer": 12500, "Scanner": 8400, "Laptop": 52000, "Monitor": 9800, "UPS": 4600}
SALES = []
for i in range(20):
    reg = REGIONS[i % 4]
    prod = PRODUCTS[(i * 3) % 5]
    units = [3, 5, 2, 8, 4, 6, 1, 7, 2, 5][i % 10] + (i // 10)
    SALES.append((reg, ["Jan", "Feb", "Mar", "Apr"][i // 5], prod, units, RATES[prod]))
REGION_TOTAL = {r: sum(u * rt for rg, m, p, u, rt in SALES if rg == r) for r in REGIONS}

e2 = Task(
    id="M1-E2", app="excel", title="Sales Report", file="Sales.xlsx", minutes=12,
    starters={"Sales.xlsx": xlsx({"Data": (["Region", "Month", "Product", "Units", "Rate"], SALES)})},
    instructions=[
        "Open Sales.xlsx from the test folder (sheet 'Data').",
        "In F1 type Amount. In F2:F21 calculate Amount = Units × Rate with a formula.",
        "Format F2:F21 as Currency (₹).",
        "Set the width of column C to 20.",
        "Apply conditional formatting on F2:F21: highlight values Greater Than 50000.",
        "In G1 type Status. On G2:G21 add Data Validation: a List with the items Paid,Pending.",
        "Insert a new sheet named Summary. In A1 type Region and in B1 Total Sales. "
        "A2:A5: North, South, East, West. In B2:B5 use SUMIF to add up Amount for each region from the Data sheet.",
        "On the Summary sheet make a Pie chart of the regions with the title Region-wise Sales.",
        "Give the Summary sheet tab a colour.",
        "Data sheet print setup: Landscape, row 1 repeats at top of every page (Print Titles), "
        "and a header with the text Sales Report.",
        "Save the file (Ctrl+S, same name Sales.xlsx).",
    ],
    checks=[
        X.rowwise("Data", "F", 2, 21, r"\*|PRODUCT\(", lambda r: r["D"] * r["E"], 3,
                  what="Amount = Units × Rate in F2:F21", how="In F2 type =D2*E2 and fill down to F21."),
        X.number_format("Data", "F2:F21", "currency", 1),
        X.col_width("Data", "C", 20, 1),
        X.cond_format("Data", 1.5),
        X.validation("Data", "G2", "list", 1.5),
        X.sheet_exists("Summary", 1),
        X.text_in("Summary", "A2:A5", REGIONS, 1, what="Region names on Summary"),
        X.rowwise("Summary", "B", 2, 5, "SUMIF", lambda r: REGION_TOTAL[r["A"]], 3,
                  what="SUMIF totals on Summary B2:B5",
                  how="In B2 type =SUMIF(Data!A:A,A2,Data!F:F) and fill down to B5."),
        X.chart("pie", "Region-wise Sales", 2),
        X.tab_color("Summary", 0.5),
        X.landscape("Data", 0.5), X.print_titles("Data", "1:1", 0.5), X.sheet_header("Data", "Sales Report", 0.5),
    ],
    solution=[
        "F1: Amount. F2: =D2*E2 > Enter > double-click the fill handle.",
        "Select F2:F21 > Home > Number Format > Currency (or Ctrl+1 > Currency > symbol ₹).",
        "Right-click column C header > Column Width > 20.",
        "Select F2:F21 > Conditional Formatting > Highlight Cells Rules > Greater Than > 50000 > OK.",
        "G1: Status. Select G2:G21 > Data > Data Validation > Allow: List > Source: Paid,Pending > OK.",
        "Click + next to the sheet tabs > double-click the new tab > Summary.",
        "Type Region / Total Sales and the 4 region names. B2: =SUMIF(Data!A:A,A2,Data!F:F) > fill down.",
        "Select A1:B5 > Insert > Pie > 2-D Pie > click title > Region-wise Sales.",
        "Right-click the Summary tab > Tab Color > pick a colour.",
        "Go to Data sheet > Page Layout > Orientation > Landscape > Print Titles > Rows to repeat at top: $1:$1 > "
        "Header/Footer tab > Custom Header > Center: Sales Report > OK.",
        "Ctrl+S.",
    ],
)

# ------------------------------------------------------------------ POWERPOINT

p1 = Task(
    id="M1-P1", app="ppt", title="ISRO Presentation", file="ISRO.pptx", minutes=12,
    instructions=[
        "Start PowerPoint with a blank presentation and apply any Design theme.",
        "Slide 1 (Title Slide): title 'Indian Space Research Organisation', subtitle 'Major Missions'.",
        "Slide 2 (Title and Content): title 'Recent Missions', bullets: Chandrayaan-3, Aditya-L1, Gaganyaan.",
        "Slide 3 (Two Content): title 'Launch Vehicles'; left box text 'PSLV and GSLV'; "
        "right box: insert the picture Rocket.png from the Pictures folder in the test folder.",
        "Slide 4 (Title Only): title 'Mission Table'; insert a table of 3 rows × 3 columns:",
        "    Mission | Year | Purpose",
        "    Chandrayaan-3 | 2023 | Moon",
        "    Aditya-L1 | 2023 | Sun",
        "Slide 5 (Title Only): title 'Thank You'.",
        "Apply any transition to all slides.",
        "Insert slide numbers and the footer 'ISRO Skill Test' on all slides except the title slide.",
        "Save as ISRO.pptx.",
    ],
    checks=[
        C.file_saved("ISRO.pptx", 1),
        P.slide_count(5, 1), P.design_applied(1.5),
        P.layout(1, "Title Slide", 0.5), P.title(1, "Indian Space Research Organisation", 0.5), P.subtitle(1, "Major Missions", 0.5),
        P.layout(2, "Title and Content", 0.5), P.title(2, "Recent Missions", 0.5),
        P.text_on(2, ["Chandrayaan-3", "Aditya-L1", "Gaganyaan"], 1, what="Slide 2 bullet points"),
        P.layout(3, "Two Content", 0.5), P.text_on(3, ["PSLV and GSLV"], 0.5), P.picture(3, 1),
        P.layout(4, "Title Only", 0.5), P.table(4, 3, 3, 1), P.table_text(4, ["Mission", "Chandrayaan-3", "Aditya-L1", "Sun", "Moon"], 1),
        P.title(5, "Thank You", 0.5),
        P.transition("all", 1),
        P.slide_numbers([2, 3, 4, 5], 1), P.footer("ISRO Skill Test", [2, 3, 4, 5], 1),
    ],
    solution=[
        "PowerPoint > Blank Presentation > Design tab > click a theme.",
        "Slide 1: type the title and subtitle.",
        "Home > New Slide > Title and Content > type the title and 3 bullets.",
        "New Slide > Two Content > type the title and left text; in the right box click the Pictures icon > Rocket.png.",
        "New Slide > Title Only > title 'Mission Table' > Insert > Table > 3×3 > type the data.",
        "New Slide > Title Only > 'Thank You'.",
        "Transitions tab > pick Fade > Apply To All.",
        "Insert > Header & Footer > tick Slide number, tick Footer and type ISRO Skill Test, "
        "tick Don't show on title slide > Apply to All.",
        "F12 > test folder > ISRO > Save.",
    ],
)

p2 = Task(
    id="M1-P2", app="ppt", title="Training Slides", file="Training.pptx", minutes=10,
    instructions=[
        "Start a new blank presentation with 3 slides.",
        "Slide 1 (Title Slide): title 'Office Automation Training', subtitle 'Skill Test Preparation'. "
        "Add the speaker note: Welcome all participants.",
        "Apply a Gradient fill background to ALL slides (Format Background > Gradient fill > Apply to All).",
        "Slide 2 (Title Only): title 'Training Process'. Insert a SmartArt 'Basic Process' with the text Word, Excel, PowerPoint.",
        "Slide 3 (Blank): draw a rectangle and an oval. Give the rectangle a gradient fill. "
        "Align both shapes to the Top, then Group them. Add an animation to the group.",
        "On slide 3 also add a Text Box with the text Thank You.",
        "Save as Training.pptx and also save a copy as Training.pdf.",
    ],
    checks=[
        C.file_saved("Training.pptx", 1),
        P.slide_count(3, 1),
        P.title(1, "Office Automation Training", 0.5), P.subtitle(1, "Skill Test Preparation", 0.5),
        P.notes(1, "Welcome all participants", 1),
        P.background("all", "gradient", 2),
        P.title(2, "Training Process", 0.5), P.smartart(2, 1.5),
        P.layout(3, "Blank", 0.5), P.shape_gradient(3, 1.5), P.aligned(3, "top", 0.5), P.grouped(3, 1),
        P.animation(3, 1), P.textbox(3, "Thank You", 1),
        C.file_format("Training.pdf", "pdf", 1.5),
    ],
    solution=[
        "New blank presentation. Type the title and subtitle on slide 1. Click under the slide (Notes) and type the note.",
        "Design > Format Background > Gradient fill > choose a preset > Apply to All.",
        "Home > New Slide > Title Only > type the title > Insert > SmartArt > Process > Basic Process > type Word, Excel, PowerPoint.",
        "New Slide > Blank. Insert > Shapes > Rectangle, then Oval. Select the rectangle > Shape Fill > Gradient > pick one.",
        "Shift+click both shapes > Shape Format > Align > Align Top > Group > Group.",
        "With the group selected > Animations > Fly In.",
        "Insert > Text Box > click on the slide > type Thank You.",
        "F12 > Training > Save. F12 again > Save as type PDF > Save.",
    ],
)

PAPER = Paper(id="mock1", title="Mock Test 1", tasks=[w1, w2, e1, e2, p1, p2])
