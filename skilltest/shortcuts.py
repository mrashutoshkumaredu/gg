"""Shortcut key quiz questions: (app, action, correct keys, wrong options)."""
import random

Q = [
    # common
    ("All", "Save the current file", "Ctrl+S", ["Ctrl+V", "Alt+S", "Shift+S"]),
    ("All", "Open the Save As box", "F12", ["F2", "Ctrl+Shift+S", "F10"]),
    ("All", "Open a file", "Ctrl+O", ["Ctrl+P", "Alt+O", "Ctrl+N"]),
    ("All", "Create a new file", "Ctrl+N", ["Ctrl+M", "Ctrl+O", "Alt+N"]),
    ("All", "Print", "Ctrl+P", ["Ctrl+Shift+P", "F9", "Alt+P"]),
    ("All", "Undo the last action", "Ctrl+Z", ["Ctrl+Y", "Ctrl+U", "Alt+Z"]),
    ("All", "Redo the last undone action", "Ctrl+Y", ["Ctrl+Z", "Ctrl+R", "Shift+Z"]),
    ("All", "Cut", "Ctrl+X", ["Ctrl+C", "Ctrl+K", "Ctrl+T"]),
    ("All", "Copy", "Ctrl+C", ["Ctrl+X", "Ctrl+V", "Ctrl+P"]),
    ("All", "Paste", "Ctrl+V", ["Ctrl+P", "Ctrl+B", "Alt+V"]),
    ("All", "Select all", "Ctrl+A", ["Ctrl+S", "Alt+A", "Shift+A"]),
    ("All", "Find", "Ctrl+F", ["Ctrl+H", "F3", "Alt+F"]),
    ("All", "Find and Replace", "Ctrl+H", ["Ctrl+R", "Ctrl+F", "Alt+H"]),
    ("All", "Bold", "Ctrl+B", ["Ctrl+D", "Alt+B", "Ctrl+Shift+D"]),
    ("All", "Italic", "Ctrl+I", ["Ctrl+T", "Alt+I", "Shift+I"]),
    ("All", "Underline", "Ctrl+U", ["Ctrl+L", "Alt+U", "Ctrl+Shift+U"]),
    ("All", "Insert a hyperlink", "Ctrl+K", ["Ctrl+L", "Ctrl+H", "Alt+K"]),
    ("All", "Spelling and grammar check", "F7", ["F5", "F8", "Shift+F7"]),
    ("All", "Close the current file", "Ctrl+W", ["Ctrl+Q", "Alt+W", "Ctrl+E"]),
    ("All", "Open Help", "F1", ["F2", "Ctrl+H", "F11"]),
    ("All", "Close the program window", "Alt+F4", ["Ctrl+F4", "Alt+F1", "Shift+F4"]),
    # word
    ("Word", "Centre align a paragraph", "Ctrl+E", ["Ctrl+C", "Ctrl+M", "Ctrl+Shift+C"]),
    ("Word", "Justify a paragraph", "Ctrl+J", ["Ctrl+Shift+J", "Ctrl+F", "Alt+J"]),
    ("Word", "Right align a paragraph", "Ctrl+R", ["Ctrl+Shift+R", "Alt+R", "Ctrl+T"]),
    ("Word", "Left align a paragraph", "Ctrl+L", ["Ctrl+Shift+L", "Alt+L", "Ctrl+K"]),
    ("Word", "Line spacing 1.5", "Ctrl+5", ["Ctrl+1", "Ctrl+2", "Alt+5"]),
    ("Word", "Double line spacing", "Ctrl+2", ["Ctrl+D", "Ctrl+5", "Alt+2"]),
    ("Word", "Single line spacing", "Ctrl+1", ["Ctrl+0", "Ctrl+2", "Alt+1"]),
    ("Word", "Insert a page break", "Ctrl+Enter", ["Shift+Enter", "Alt+Enter", "Ctrl+Shift+Enter"]),
    ("Word", "Insert a line break (new line, same paragraph)", "Shift+Enter", ["Ctrl+Enter", "Alt+Enter", "Tab"]),
    ("Word", "Insert a column break", "Ctrl+Shift+Enter", ["Ctrl+Enter", "Shift+Enter", "Alt+Shift+Enter"]),
    ("Word", "Change case (cycle UPPER / lower / Title)", "Shift+F3", ["Ctrl+F3", "Alt+F3", "F3"]),
    ("Word", "Increase font size", "Ctrl+Shift+>", ["Ctrl+]", "Ctrl+Shift+<", "Alt+>"]),
    ("Word", "Decrease font size", "Ctrl+Shift+<", ["Ctrl+[", "Ctrl+Shift+>", "Alt+<"]),
    ("Word", "Open the Font dialog box", "Ctrl+D", ["Ctrl+F", "Ctrl+Shift+F", "Alt+D"]),
    ("Word", "Turn Track Changes on/off", "Ctrl+Shift+E", ["Ctrl+T", "Alt+T", "Ctrl+Shift+T"]),
    ("Word", "Insert © symbol", "Alt+Ctrl+C", ["Ctrl+Shift+C", "Alt+C", "Ctrl+C"]),
    ("Word", "Go to a page (Go To)", "Ctrl+G", ["Ctrl+P", "F6", "Alt+G"]),
    ("Word", "Subscript (e.g. H2O)", "Ctrl+=", ["Ctrl+Shift+=", "Ctrl+-", "Alt+="]),
    ("Word", "Superscript (e.g. x²)", "Ctrl+Shift+=", ["Ctrl+=", "Ctrl+Shift+-", "Alt+Shift+="]),
    ("Word", "Indent a paragraph from the left", "Ctrl+M", ["Ctrl+Shift+M", "Tab+M", "Alt+M"]),
    ("Word", "Insert a footnote", "Alt+Ctrl+F", ["Ctrl+F", "Alt+F", "Ctrl+Shift+F"]),
    # excel
    ("Excel", "Edit the active cell", "F2", ["F4", "Ctrl+E", "Enter"]),
    ("Excel", "Make a reference absolute ($A$1) while typing a formula", "F4", ["F2", "Ctrl+4", "Shift+4"]),
    ("Excel", "AutoSum", "Alt+=", ["Ctrl+=", "Alt+S", "Ctrl+Shift+S"]),
    ("Excel", "Insert a new worksheet", "Shift+F11", ["Ctrl+F11", "Alt+F11", "Ctrl+N"]),
    ("Excel", "Create a chart on a new chart sheet", "F11", ["Alt+F1", "Ctrl+F11", "Shift+F1"]),
    ("Excel", "Insert a chart on the same sheet", "Alt+F1", ["F11", "Ctrl+F1", "Alt+C"]),
    ("Excel", "Open Format Cells dialog", "Ctrl+1", ["Ctrl+F", "Alt+1", "Ctrl+Shift+1"]),
    ("Excel", "Turn AutoFilter on/off", "Ctrl+Shift+L", ["Ctrl+L", "Alt+F", "Ctrl+Shift+F"]),
    ("Excel", "Insert today's date as a value", "Ctrl+;", ["Ctrl+:", "Ctrl+D", "Alt+;"]),
    ("Excel", "Insert the current time", "Ctrl+Shift+;", ["Ctrl+;", "Ctrl+T", "Alt+Shift+;"]),
    ("Excel", "Fill down (copy the cell above)", "Ctrl+D", ["Ctrl+R", "Ctrl+F", "Alt+D"]),
    ("Excel", "Fill right (copy the cell to the left)", "Ctrl+R", ["Ctrl+D", "Ctrl+L", "Alt+R"]),
    ("Excel", "Select the whole column", "Ctrl+Space", ["Shift+Space", "Alt+Space", "Ctrl+C"]),
    ("Excel", "Select the whole row", "Shift+Space", ["Ctrl+Space", "Alt+Space", "Ctrl+R"]),
    ("Excel", "Insert a new line inside a cell", "Alt+Enter", ["Ctrl+Enter", "Shift+Enter", "Tab"]),
    ("Excel", "Go to cell A1", "Ctrl+Home", ["Home", "Ctrl+Up", "Alt+Home"]),
    ("Excel", "Move to the next sheet", "Ctrl+Page Down", ["Ctrl+Tab", "Alt+Page Down", "Shift+Page Down"]),
    ("Excel", "Apply Currency format", "Ctrl+Shift+$", ["Ctrl+$", "Ctrl+Shift+%", "Alt+$"]),
    ("Excel", "Apply Percentage format", "Ctrl+Shift+%", ["Ctrl+%", "Ctrl+Shift+$", "Alt+%"]),
    ("Excel", "Recalculate all formulas", "F9", ["F5", "F2", "Ctrl+F9"]),
    ("Excel", "Insert a comment / note", "Shift+F2", ["Ctrl+F2", "Alt+F2", "F2"]),
    ("Excel", "Hide the selected rows", "Ctrl+9", ["Ctrl+0", "Ctrl+H", "Alt+9"]),
    ("Excel", "Hide the selected columns", "Ctrl+0", ["Ctrl+9", "Ctrl+H", "Alt+0"]),
    ("Excel", "Insert cells/rows/columns", "Ctrl+Shift++", ["Ctrl+-", "Alt+I", "Ctrl+I"]),
    ("Excel", "Delete cells/rows/columns", "Ctrl+-", ["Ctrl+Shift++", "Delete", "Alt+-"]),
    ("Excel", "Show formulas instead of results", "Ctrl+`", ["Ctrl+F", "Ctrl+Shift+`", "Alt+`"]),
    # powerpoint
    ("PowerPoint", "Insert a new slide", "Ctrl+M", ["Ctrl+N", "Ctrl+Shift+M", "Alt+M"]),
    ("PowerPoint", "Start the slide show from the beginning", "F5", ["Shift+F5", "F6", "Ctrl+F5"]),
    ("PowerPoint", "Start the slide show from the current slide", "Shift+F5", ["F5", "Ctrl+F5", "Alt+F5"]),
    ("PowerPoint", "End the slide show", "Esc", ["Ctrl+E", "Alt+F4", "End"]),
    ("PowerPoint", "Duplicate the selected slide/object", "Ctrl+D", ["Ctrl+C", "Ctrl+Shift+D", "Alt+D"]),
    ("PowerPoint", "Group the selected objects", "Ctrl+G", ["Ctrl+Shift+G", "Alt+G", "Ctrl+U"]),
    ("PowerPoint", "Ungroup objects", "Ctrl+Shift+G", ["Ctrl+G", "Alt+Shift+G", "Ctrl+U"]),
    ("PowerPoint", "Black screen during a slide show", "B", ["W", "K", "Ctrl+B"]),
    ("PowerPoint", "White screen during a slide show", "W", ["B", "Ctrl+W", "S"]),
    ("PowerPoint", "Go to slide number 5 during a show", "5 then Enter", ["Ctrl+5", "F5", "Alt+5"]),
    ("PowerPoint", "Centre align text in a box", "Ctrl+E", ["Ctrl+C", "Ctrl+Shift+E", "Alt+E"]),
]


def make_quiz(n=20, app=None, seed=None):
    rnd = random.Random(seed)
    pool = [q for q in Q if app in (None, "All") or q[0] in (app, "All")]
    picked = rnd.sample(pool, min(n, len(pool)))
    out = []
    for app_, action, right, wrong in picked:
        opts = [right] + list(wrong)
        rnd.shuffle(opts)
        out.append({"app": app_, "q": f"[{app_}] {action}", "options": opts, "answer": right})
    return out
