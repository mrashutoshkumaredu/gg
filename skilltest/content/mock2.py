"""Mock Test 2."""
from .. import common as C
from .. import excel as X
from .. import ppt as P
from .. import word as W
from ..core import Paper, Task
from .starters import docx_with, xlsx

# ------------------------------------------------------------------ WORD

CIRC_HEAD = "Circular on Office Timings"
CIRC_P1 = ("It is hereby informed that the office timings of the Bangalore campus will change from next month. "
           "All staff must recieve the new identity cards from the security section before the change.")
CIRC_P2 = ("The goverment has also advised that the work enviroment should be kept clean. Staff working at the "
           "Bangalore and Hassan centres should acheive the target of a paper-free office by March. "
           "Visitors to Bangalore office must use the seperate gate.")
CIRC_P3 = "For details visit the ISRO website. This circular is issued with the approval of the competent authority Bangalore."
SPELL = {"recieve": "receive", "goverment": "government", "enviroment": "environment", "acheive": "achieve",
         "seperate": "separate"}

w1 = Task(
    id="M2-W1", app="word", title="Correct a Circular", file="Circular.docx", minutes=12,
    starters={"Circular_Draft.docx": docx_with([CIRC_P1, CIRC_P2, CIRC_P3], title=CIRC_HEAD)},
    instructions=[
        "Open Circular_Draft.docx from the test folder.",
        "Use Spelling & Grammar (F7) to correct all the spelling mistakes.",
        "Use Find and Replace (Ctrl+H) to replace every 'Bangalore' with 'Bengaluru'.",
        f"Apply the style Heading 1 to the first line, then change it to UPPERCASE using Change Case.",
        "Set the three paragraphs below the heading in font Times New Roman, and give them a first line indent of 0.5 inch.",
        "Highlight the words 'paper-free office' in yellow.",
        "Make the words 'ISRO website' a hyperlink to www.isro.gov.in.",
        "Put a page border around the page.",
        "Insert the picture Logo.png (from the Pictures folder) in the header.",
        "Save as Circular.docx (Save As, new name) and also save a copy as Circular.rtf (Rich Text Format).",
    ],
    checks=[
        C.file_saved("Circular.docx", 1),
        W.spell_fixed(SPELL, 3),
        W.not_contains("Bangalore", 1, what="No 'Bangalore' left", topic="Find and Replace",
                       how="Ctrl+H > Find what: Bangalore > Replace with: Bengaluru > Replace All."),
        W.count_of("Bengaluru", 4, 1, how="Ctrl+H > Find what: Bangalore > Replace with: Bengaluru > Replace All."),
        W.style(CIRC_HEAD, "Heading 1", 1),
        W.uppercase(CIRC_HEAD, 1),
        W.font_name("office timings of the", "Times New Roman", 1),
        W.font_name("should be kept clean", "Times New Roman", 1),
        W.indent("office timings of the", first=0.5, marks=1),
        W.indent("should be kept clean", first=0.5, marks=1),
        W.highlighted("paper-free office", 1, sub="paper-free office"),
        W.hyperlink(1),
        W.page_border(1),
        W.picture(1, 1.5, where="header"),
        C.file_format("Circular.rtf", "rtf", 1.5),
    ],
    solution=[
        "Open the draft. Press F7 and click Change for: recieve, goverment, enviroment, acheive, seperate.",
        "Ctrl+H > Find: Bangalore > Replace: Bengaluru > Replace All (4 replacements).",
        "Click in the first line > Home > Styles > Heading 1. Select it > Change Case (Aa) > UPPERCASE.",
        "Select the three paragraphs > Font: Times New Roman. Paragraph dialog > Special: First line, By: 0.5\" > OK.",
        "Select 'paper-free office' > Text Highlight Color > Yellow.",
        "Select 'ISRO website' > Ctrl+K > Address: www.isro.gov.in > OK.",
        "Design > Page Borders > Box > OK.",
        "Insert > Header > Edit Header > Insert > Pictures > Pictures\\Logo.png > Close Header.",
        "F12 > Circular > Word Document > Save. F12 > Save as type: Rich Text Format > Save.",
    ],
)

RECIP = [("Shri Arun Mehta", "Section Officer", "Thiruvananthapuram"),
         ("Smt. Kavya Rao", "Assistant", "Sriharikota"),
         ("Shri Vikram Singh", "Accounts Officer", "Ahmedabad")]
LETTER_BODY = ("You are requested to attend the review meeting on Hindi implementation to be held at ISRO Headquarters "
               "on 5 December at 11 AM.")

w2 = Task(
    id="M2-W2", app="word", title="Mail Merge Letters", file="MergeLetter.docx", minutes=13,
    starters={"Recipients.xlsx": xlsx({"Sheet1": (["Name", "Designation", "City"], RECIP)})},
    instructions=[
        "Start a new Word document. Type the letter below:",
        "    To,",
        "    <Name>",
        "    <Designation>, <City>",
        "    Subject: Review meeting",
        f"    {LETTER_BODY}",
        "    Yours faithfully,",
        "Use Mailings > Start Mail Merge > Letters > Select Recipients > Use an Existing List > Recipients.xlsx.",
        "In place of <Name>, <Designation> and <City> insert the merge fields Name, Designation and City.",
        "Make the line 'Subject: Review meeting' Bold.",
        "Save this main letter as MergeLetter.docx.",
        "Finish & Merge > Edit Individual Documents > All. Save the merged letters as AllLetters.docx.",
        "Finally, save the main letter also as a Word Template named LetterTemplate.dotx in the test folder.",
    ],
    checks=[
        C.file_saved("MergeLetter.docx", 1),
        W.typed(LETTER_BODY, 2, what="Letter body typed correctly"),
        W.contains("Yours faithfully", 1),
        W.bold("Subject: Review meeting", 1),
        W.mergefields(["Name", "Designation", "City"], 4),
        C.file_saved("AllLetters.docx", 1, topic="Mail Merge", what="merged letters"),
        W.merged_output([r[0] for r in RECIP] + [r[2] for r in RECIP], 4, file="AllLetters.docx"),
        C.file_format("LetterTemplate.dotx", "dotx", 3, topic="Templates"),
    ],
    solution=[
        "Blank document > type the letter, leaving the three placeholder lines empty.",
        "Mailings > Start Mail Merge > Letters. Select Recipients > Use an Existing List > Recipients.xlsx > Sheet1$ > OK.",
        "Put the cursor on the Name line > Insert Merge Field > Name. Same for Designation and City.",
        "Select the subject line > Ctrl+B. Ctrl+S > MergeLetter.",
        "Finish & Merge > Edit Individual Documents > All > OK. In the new document Ctrl+S > AllLetters.",
        "Go back to MergeLetter > F12 > Save as type: Word Template (*.dotx) > change the folder back to the test "
        "folder (Word jumps to the Templates folder!) > LetterTemplate > Save.",
    ],
)

# ------------------------------------------------------------------ EXCEL

EMP = [("E101", "Ramesh", 35400), ("E102", "Sunita", 25500), ("E103", "Ajay", 44900), ("E104", "Lakshmi", 29200),
       ("E105", "Naveen", 21700), ("E106", "Pooja", 56100), ("E107", "Kiran", 25500), ("E108", "Farhan", 35400)]
DA, HRA = 0.46, 0.24


def _gross(r):
    return round(r["C"] * (1 + DA + HRA))


e1 = Task(
    id="M2-E1", app="excel", title="Salary Sheet", file="Salary.xlsx", minutes=13,
    starters={"Salary.xlsx": xlsx({"Pay": (["Emp ID", "Name", "Basic", "DA", "HRA", "Gross"], EMP)})},
    instructions=[
        "Open Salary.xlsx (sheet 'Pay', rows 2 to 9).",
        "In H1 type DA Rate and in I1 type 0.46. In H2 type HRA Rate and in I2 type 0.24. Format I1:I2 as Percentage.",
        "D2:D9: DA = Basic × DA rate. Use an absolute reference to the rate, e.g. =C2*$I$1.",
        "E2:E9: HRA = Basic × HRA rate using $I$2.",
        "F2:F9: Gross = ROUND(Basic + DA + HRA, 0).",
        "Format C2:F9 in Comma Style (thousand separator).",
        "In A11 type Highest Gross and in F11 use MAX. In A12 type Lowest Gross and in F12 use MIN.",
        "In A13 type Basic 25500 count and in F13 use COUNTIF to count how many have Basic equal to 25500.",
        "In H5 type Find ID, in I5 type E104. In I6 use VLOOKUP to show the Name of the ID in I5.",
        "Make row 1 Bold with Wrap Text on, and set its row height to 30.",
        "Save (Ctrl+S).",
    ],
    checks=[
        X.number_format("Pay", "I1:I2", "percent", 1),
        X.rowwise("Pay", "D", 2, 9, r"\$I\$1", lambda r: r["C"] * DA, 2.5, what="DA uses absolute reference $I$1",
                  how="In D2 type =C2*$I$1 (press F4 after clicking I1 to add the $ signs) and fill down."),
        X.rowwise("Pay", "E", 2, 9, r"\$I\$2", lambda r: r["C"] * HRA, 2, what="HRA uses absolute reference $I$2",
                  how="In E2 type =C2*$I$2 and fill down."),
        X.rowwise("Pay", "F", 2, 9, "ROUND", _gross, 2.5, what="Gross with ROUND",
                  how="In F2 type =ROUND(C2+D2+E2,0) and fill down."),
        X.number_format("Pay", "C2:F9", "comma", 1),
        X.formula("Pay", "F11", "MAX", round(56100 * 1.7), 1.5),
        X.formula("Pay", "F12", "MIN", round(21700 * 1.7), 1.5),
        X.formula("Pay", "F13", "COUNTIF", 2, 2, how='In F13 type =COUNTIF(C2:C9,25500).'),
        X.formula("Pay", "I6", "VLOOKUP", "Lakshmi", 2, how="In I6 type =VLOOKUP(I5,A2:F9,2,FALSE)."),
        X.bold("Pay", "A1:F1", 0.5), X.wrap("Pay", "A1:F1", 0.5), X.row_height("Pay", 1, 30, 1),
    ],
    solution=[
        "Type the rates in H1:I2. Select I1:I2 > Home > % (Percent Style).",
        "D2: =C2*$I$1   E2: =C2*$I$2   F2: =ROUND(C2+D2+E2,0). Select D2:F2 > double-click the fill handle.",
        "Select C2:F9 > Home > Comma Style (,).",
        "F11: =MAX(F2:F9)   F12: =MIN(F2:F9)   F13: =COUNTIF(C2:C9,25500).",
        "I6: =VLOOKUP(I5,A2:F9,2,FALSE) – it shows Lakshmi.",
        "Select A1:F1 > Ctrl+B > Wrap Text. Right-click row 1 > Row Height > 30.",
        "Ctrl+S.",
    ],
)

ITEMS = [("Pen", "Stationery", 250, 10), ("Stapler", "Stationery", 40, 85), ("Toner", "Printer", 12, 2400),
         ("A4 Paper", "Stationery", 120, 260), ("Mouse", "Computer", 18, 450), ("Keyboard", "Computer", 15, 650),
         ("Drum Unit", "Printer", 4, 3800), ("File Cover", "Stationery", 300, 12), ("Pen Drive", "Computer", 25, 550),
         ("Ink Cartridge", "Printer", 9, 1200)]

e2 = Task(
    id="M2-E2", app="excel", title="Inventory Register", file="Inventory.xlsx", minutes=12,
    starters={"Inventory.xlsx": xlsx({"Stock": (["Item Name and Description", "Category", "Quantity", "Unit Price"], ITEMS),
                                      "Old": (["Remarks"], [("not needed",)])})},
    instructions=[
        "Open Inventory.xlsx.",
        "Delete the sheet named Old.",
        "Rename the sheet Stock to Store Items.",
        "AutoFit column A so the full heading is visible (double-click its right border).",
        "In E1 type Value. In E2:E11 calculate Quantity × Unit Price.",
        "Hide row 11 (the last item).",
        "Insert a comment on A1 with the text Updated by store keeper.",
        "Turn on AutoFilter and filter Category to show only Stationery.",
        "Print setup: set the print area to A1:E11 and fit the sheet on one page.",
        "Protect the sheet (Review > Protect Sheet, no password needed).",
        "Save (Ctrl+S). Then also save a copy as Inventory.xls (Excel 97-2003 Workbook).",
    ],
    checks=[
        X.sheet_absent("Old", 1),
        X.sheet_exists("Store Items", 1),
        X.autofit("Store Items|Stock", "A", 20, 1),
        X.rowwise("Store Items|Stock", "E", 2, 11, r"\*|PRODUCT\(", lambda r: r["C"] * r["D"], 3,
                  what="Value = Quantity × Unit Price", how="In E2 type =C2*D2 and fill down to E11."),
        X.hidden_row("Store Items|Stock", 11, 1.5),
        X.comment("Store Items|Stock", "A1", 1.5),
        X.autofilter("Store Items|Stock", 2.5),
        X.print_area("Store Items|Stock", "A1:E11", 1),
        X.fit_one_page("Store Items|Stock", 1),
        X.protected("Store Items|Stock", 1.5),
        C.file_format("Inventory.xls", "xls97", 2),
    ],
    solution=[
        "Right-click the Old tab > Delete > Delete.",
        "Double-click the Stock tab > type Store Items > Enter.",
        "Double-click the line between column headers A and B.",
        "E1: Value. E2: =C2*D2 > double-click the fill handle.",
        "Right-click row number 11 > Hide.",
        "Right-click A1 > Insert Comment (New Note) > type Updated by store keeper.",
        "Click in the data > Data > Filter > Category arrow > tick only Stationery > OK.",
        "Select A1:E11 > Page Layout > Print Area > Set Print Area. Page Setup > Fit to 1 page wide by 1 tall.",
        "Review > Protect Sheet > OK.",
        "Ctrl+S. Then F12 > Save as type: Excel 97-2003 Workbook > Inventory > Save > Continue.",
    ],
)

# ------------------------------------------------------------------ POWERPOINT

p1 = Task(
    id="M2-P1", app="ppt", title="Budget Presentation", file="Budget.pptx", minutes=12,
    instructions=[
        "Start a new blank presentation. Change the slide size to Standard (4:3).",
        "Slide 1 (Title Slide): title 'Annual Budget 2026-27', subtitle 'Finance Division'.",
        "Make the title Bold, size 40 and change its font colour.",
        "Slide 2 (Comparison): title 'Plan vs Actual'. Type 'Planned' in the left heading box and 'Actual' in the right heading box.",
        "Slide 3 (Title and Content): title 'Expenditure Chart'. Insert a Clustered Column chart (any data).",
        "Slide 4 (Title and Content): title 'More Information'; text 'Visit ISRO website'. "
        "Make 'Visit ISRO website' a hyperlink to www.isro.gov.in.",
        "Add the speaker note 'Explain the chart' on slide 3.",
        "Add a transition to all slides.",
        "Save as Budget.pptx.",
    ],
    checks=[
        C.file_saved("Budget.pptx", 1),
        P.slide_size("4:3", 1.5),
        P.slide_count(4, 1),
        P.title(1, "Annual Budget 2026-27", 0.5), P.subtitle(1, "Finance Division", 0.5),
        P.text_bold(1, "Annual Budget 2026-27", 1), P.text_size(1, "Annual Budget 2026-27", 40, 1),
        P.text_color(1, "Annual Budget 2026-27", 1),
        P.layout(2, "Comparison", 1), P.text_on(2, ["Planned", "Actual"], 1),
        P.title(3, "Expenditure Chart", 0.5), P.chart(3, None, 2),
        P.hyperlink(4, 1.5),
        P.notes(3, "Explain the chart", 1),
        P.transition("all", 0.5),
    ],
    solution=[
        "Blank presentation > Design > Slide Size > Standard (4:3) > Ensure Fit.",
        "Type the title and subtitle. Select the title > Ctrl+B > size 40 > Font Color.",
        "New Slide > Comparison > type the title, Planned and Actual.",
        "New Slide > Title and Content > title > click the chart icon > Clustered Column > OK > close the data window.",
        "Notes pane on slide 3: Explain the chart.",
        "New Slide > title 'More Information', text 'Visit ISRO website' > select it > Ctrl+K > www.isro.gov.in > OK.",
        "Transitions > pick one > Apply To All. F12 > Budget > Save.",
    ],
)

p2 = Task(
    id="M2-P2", app="ppt", title="Safety Slides", file="Safety.pptx", minutes=11,
    instructions=[
        "Start a new blank presentation with 3 slides.",
        "Open View > Slide Master. On the top master slide insert a Text Box with 'ISRO Internal'. Close Master View.",
        "Slide 1 (Title Slide): title 'Fire Safety Drill', subtitle 'Dos and Don'ts'.",
        "Give slide 1 a Picture background using Satellite.png from the Pictures folder (this slide only).",
        "Slide 2 (Section Header): title 'Before the Drill'.",
        "Slide 3 (Title and Content): title 'Assembly Points'; bullets: Gate 1, Gate 3, Parking Area. "
        "Add an animation to the bullet text box.",
        "Add the date and slide number on all slides.",
        "Save as Safety.pptx and also save it as a PowerPoint Show named Safety.ppsx.",
    ],
    checks=[
        C.file_saved("Safety.pptx", 1),
        P.master_text("ISRO Internal", 3),
        P.slide_count(3, 1),
        P.title(1, "Fire Safety Drill", 0.5),
        P.background(1, "picture", 2.5),
        P.layout(2, "Section Header", 1), P.title(2, "Before the Drill", 0.5),
        P.text_on(3, ["Gate 1", "Gate 3", "Parking Area"], 1), P.animation(3, 1.5),
        P.date_shown("all", 1), P.slide_numbers("all", 1),
        C.file_format("Safety.ppsx", "ppsx", 1),
    ],
    solution=[
        "View > Slide Master > click the top big slide > Insert > Text Box > type ISRO Internal > Close Master View.",
        "Slide 1: type the title and subtitle. Right-click slide > Format Background > Picture or texture fill > "
        "Insert (File) > Pictures\\Satellite.png. Do NOT click Apply to All.",
        "New Slide > Section Header > title. New Slide > Title and Content > title and 3 bullets.",
        "Select the bullet box > Animations > Appear (or Fly In).",
        "Insert > Header & Footer > tick Date and time and Slide number > Apply to All.",
        "F12 > Safety > Save. F12 > Save as type: PowerPoint Show (*.ppsx) > Save.",
    ],
)

PAPER = Paper(id="mock2", title="Mock Test 2", tasks=[w1, w2, e1, e2, p1, p2])
