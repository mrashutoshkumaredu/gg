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

## Updates

The app checks for a new version when it opens. If there is one, the Home screen shows
**New version … is available** with **What's new** and **Update now**. Update now downloads the new .exe,
closes the app, replaces the file and opens the new version. Scores and settings are kept.
You can also press **Check for updates** at the bottom of the Home screen.

If the check fails, the message says why in plain words (for example a security certificate problem, a proxy or
antivirus block, or a time-out) and offers to open the download page. The app tries three ways of reaching GitHub
before giving up.

The version shown in the window title (for example v1.7.12) is the same as the release name.

Each release also publishes a SHA-256 fingerprint of the .exe. Update now checks the downloaded file against it and
installs only if they match; otherwise the download is deleted and nothing is changed.

**For the developer:** every push to `main` is built and tested, but a release (which students receive through
Update now) is made only when the commit message contains `[release]`, or when the workflow is run by hand from
the Actions tab ("Run workflow", with optional notes). The version is `__version__` in `skilltest/__init__.py`
plus the build number. Before anything is built, a test stage on GitHub checks that the model answers still get
full marks (Mock 1 and the Excel tasks of Mocks 2-5, including retake numbers); if any fails, no release is made.

## What is inside

- **5 full mock tests**: each one has 2 Word, 2 Excel and 2 PowerPoint tasks, 100 marks in total, with a 75-minute timer.
- **Section practice**: only the Word, Excel or PowerPoint part of a mock, with a shorter timer. You can check each task as you go.
- **45 topic drills**: short tasks on one syllabus topic each (Mail Merge, Track Changes, VLOOKUP, Subtotal,
  Data Validation, Slide Master, gradient fills and so on). You can check them instantly and see the model answer.
- **Shortcut key quiz**: 20 random questions from about 80 Word, Excel and PowerPoint shortcuts.
- **Typing test**: 8 office-style passages (or your own), 2/5/10/15-minute or untimed, Backspace on/off.
  In a timed test only the part you reached is checked; in an untimed test the whole passage is checked,
  so words not typed count as missing.
  Gives gross and net speed (WPM), accuracy, error %, and full and half mistakes, with every mistake marked.
  Paste is blocked in the typing box.
- **Compare two texts** (dictation check): paste or load (.txt/.docx) the paragraph that was read out and
  what you typed. Every wrong, missing or extra word, capital letter, punctuation and spacing slip is marked.
  Capitals and punctuation can be switched off. Results can be saved as a printable report.
- **My Progress**: past scores and your weak topics. **Find old results** looks through the test folders and puts
  back any result missing from the list (older retakes are re-marked, so their score is approximate).
  A test closed without submitting is shown as "not completed" and left out of the averages.

## How a test works

1. Choose a test. A folder is made for it in `Documents\ISRO_SkillTest`, for example
   `Mock Test 1 - Full Test\Attempt 1 (26-10-2026 18.15)`, and opened for you. Any starter files and a
   `Pictures` folder are put in it.
2. Read each task in the app and do it in Office. Save each file in the test folder with the exact name asked.
3. While the test runs, the app records:
   - which Office program is open and how long you spend in each one
   - files created, saved or deleted in the test folder
   - documents opened or saved, Track Changes switched on or off, slides added, sheets added, charts inserted
     and formulas typed in Excel
4. The timer warns you at 10, 5 and 1 minute left, and submits the test by itself when time runs out.
5. After you submit, the app checks your files and shows the result. A printable report
   (`Result_Report.html`) with the full activity timeline is saved in the test folder.

## Any screen size

The top bar (Home, My progress, Typing, Updates, Help) is always visible. Every screen can be scrolled with the
mouse wheel or touchpad when the window is too small to show everything, so all options can be reached on any
laptop. The window can be made as small as half the screen, and every screen rearranges itself when it is narrow.
During a test, press **Split screen**: open Office windows move to the left and the app moves to the right, on top.
Windows key + Left/Right arrow (Windows Snap) also works.

**Mini bar**: keeps Office full screen. A small always-on-top box shows the timer, the current task's steps,
◀ ▶ to change task, Check and Submit. Drag it anywhere, resize it, press ◐ to make it see-through,
▴ to fold it down to just the timer, and ⤢ to go back to the full window.

## Keyboard

- Quiz: **A–D** (or 1–4) to answer, **→** or **Enter** for the next question, **←** to look back.
- Test screen and mini bar: **←** and **→** move between tasks.

## Retaking a test

Every retake makes a new `Attempt` folder inside the same test folder (Attempt 1, Attempt 2 …), so your old
work and score are kept for comparison. Attempt 1 uses the original paper; every retake uses **new names and
numbers**, so answers cannot simply be remembered. Use **Retake (new numbers)** on the result screen, or open
**My progress & retakes** to retake, open or delete any past attempt. Deleting an attempt also removes its folder.

**Fix & recheck** reopens the same files so you can correct the mistakes and have them marked again.

## Fair marking

- Part marks: for example, if 3 of 4 bullets are right, or 8 of 10 formulas, you get that share of the marks.
- Typing: each mistake costs 10% of that passage's marks, and the result lists every wrong, missing or extra word.
- Copying the passages to be typed is blocked, both in the app and on the clipboard.
- **Strict exam mode** (switch on the Home screen) hides "Check this task" and model answers until you submit.
- **Report wrong marking** on the result screen makes a zip of the files and result to send to the developer.

## Learning and teachers

- **Practise →** next to each weak area starts the matching drills.
- **Score trend** chart on the My Progress screen.
- **Instructions in Hindi** opens the question paper in English and Hindi in the browser.
- **Students**: add several students on the Home screen; each has separate folders, scores and weak areas.
- **Export all results to Excel** (My Progress screen) gives a sheet of every attempt by every student, plus weak topics.

## Your saved results

Results and settings are kept in `%APPDATA%\ISRO_SkillTest` on each computer. Every save is written to a
temporary file first and the previous good copy is kept as a `.bak`, so a power cut cannot wipe them. If a file is
ever damaged, the app restores it from the backup, keeps the damaged copy aside, and tells you. If the file is only
in use by another program for a moment, nothing is changed. Each test folder also keeps its own `result.json`.

Live activity only lists files in the test folder (and new unsaved documents), not other files you have open.

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
