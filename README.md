# ISRO Skill Test Practice

A Windows practice app for the computer skill test for Assistant / UDC / JPA / Stenographer posts
(MS Word 35 marks, MS Excel 35 marks, MS PowerPoint 30 marks, 75 minutes).

You do the tasks in the real MS Word, Excel and PowerPoint. The app tracks what you do, then opens your saved
files and checks each point. For every mistake it shows what it found, what was expected and how to do it,
along with a step-by-step model answer for each task.

## Download

Go to **Releases** (right side of the GitHub page), open the latest release and download `ISRO_SkillTest.exe`.
Double-click it. Nothing needs to be installed. MS Office (2010 or newer) must be on the computer.

If Windows shows "Windows protected your PC", click **More info > Run anyway**. This happens because the
file is not signed.

## What is inside

- **5 full mock tests**: each one has 2 Word, 2 Excel and 2 PowerPoint tasks, 100 marks in total, with a 75-minute timer.
- **Section practice**: only the Word, Excel or PowerPoint part of a mock, with a shorter timer. You can check each task as you go.
- **45 topic drills**: short tasks on one syllabus topic each (Mail Merge, Track Changes, VLOOKUP, Subtotal,
  Data Validation, Slide Master, gradient fills and so on). You can check them instantly and see the model answer.
- **Shortcut key quiz**: 20 random questions from about 80 Word, Excel and PowerPoint shortcuts.
- **My Progress**: past scores and your weak topics.

## How a test works

1. Choose a test. A new test folder is made in `Documents\ISRO_SkillTest` and opened for you. Any starter
   files and a `Pictures` folder are put in it.
2. Read each task in the app and do it in Office. Save each file in the test folder with the exact name asked.
3. While the test runs, the app records:
   - which Office program is open and how long you spend in each one
   - files created, saved or deleted in the test folder
   - documents opened or saved, Track Changes switched on or off, slides added, sheets added, charts inserted
     and formulas typed in Excel
4. The timer warns you at 10, 5 and 1 minute left, and submits the test by itself when time runs out.
5. After you submit, the app checks your files and shows the result. A printable report
   (`Result_Report.html`) with the full activity timeline is saved in the test folder.

## Run from source (optional)

```
pip install -r requirements.txt
python ISRO_SkillTest.py
```

Tests: `python tests/test_all.py`, `python tests/test_mock1.py` and `python tests/test_excel_mocks.py`
(the last two need LibreOffice to recalculate formulas).

## Project layout

- `skilltest/word.py`, `excel.py`, `ppt.py`: read Office files and hold the check builders
- `skilltest/content/mock1.py` … `mock5.py`, `drills.py`: question papers (instructions, checks, model answers)
- `skilltest/monitor.py`: live activity tracking
- `skilltest/gui.py`: the app window
- `.github/workflows/build.yml`: builds the `.exe` on every push to `main` and publishes it as a release
