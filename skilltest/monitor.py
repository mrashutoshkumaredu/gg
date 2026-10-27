"""Live activity tracking while a test runs.

What is tracked:
- files created / saved / deleted in the test folder (works everywhere)
- which Office program is in front, and for how long (Windows)
- Word/Excel/PowerPoint started or closed (Windows)
- documents opened, saved, Track Changes switched, slides added, formulas typed (Windows + pywin32)
"""
from __future__ import annotations

import os
import sys
import threading
import time
from datetime import datetime
from pathlib import Path

IS_WIN = sys.platform == "win32"
EXE_APP = {"winword.exe": "word", "excel.exe": "excel", "powerpnt.exe": "ppt"}
APP_NAME = {"word": "Word", "excel": "Excel", "ppt": "PowerPoint", "folder": "Folder", "test": "Test app",
            "other": "Other"}


# ---------------------------------------------------------------- Windows helpers (ctypes only)

if IS_WIN:
    import ctypes
    from ctypes import wintypes

    _user32 = ctypes.windll.user32
    _kernel32 = ctypes.windll.kernel32

    class PROCESSENTRY32W(ctypes.Structure):
        _fields_ = [("dwSize", wintypes.DWORD), ("cntUsage", wintypes.DWORD), ("th32ProcessID", wintypes.DWORD),
                    ("th32DefaultHeapID", ctypes.c_void_p), ("th32ModuleID", wintypes.DWORD),
                    ("cntThreads", wintypes.DWORD), ("th32ParentProcessID", wintypes.DWORD),
                    ("pcPriClassBase", ctypes.c_long), ("dwFlags", wintypes.DWORD), ("szExeFile", ctypes.c_wchar * 260)]

    def running_exes():
        snap = _kernel32.CreateToolhelp32Snapshot(0x2, 0)
        out = set()
        if snap in (0, -1):
            return out
        e = PROCESSENTRY32W()
        e.dwSize = ctypes.sizeof(PROCESSENTRY32W)
        ok = _kernel32.Process32FirstW(snap, ctypes.byref(e))
        while ok:
            out.add(e.szExeFile.lower())
            ok = _kernel32.Process32NextW(snap, ctypes.byref(e))
        _kernel32.CloseHandle(snap)
        return out

    def foreground():
        """Return (exe name, window title) of the window in front."""
        hwnd = _user32.GetForegroundWindow()
        if not hwnd:
            return "", ""
        pid = wintypes.DWORD()
        _user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        length = _user32.GetWindowTextLengthW(hwnd)
        buf = ctypes.create_unicode_buffer(length + 1)
        _user32.GetWindowTextW(hwnd, buf, length + 1)
        h = _kernel32.OpenProcess(0x1000, False, pid.value)
        exe = ""
        if h:
            size = wintypes.DWORD(1024)
            pbuf = ctypes.create_unicode_buffer(1024)
            if _kernel32.QueryFullProcessImageNameW(h, 0, pbuf, ctypes.byref(size)):
                exe = os.path.basename(pbuf.value).lower()
            _kernel32.CloseHandle(h)
        return exe, buf.value
else:
    def running_exes():
        return set()

    def foreground():
        return "", ""


def _col_letter(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class Monitor:
    def __init__(self, folder: Path, on_event=None):
        self.folder = Path(folder)
        self.on_event = on_event
        self.events = []
        self.app_seconds = {"word": 0.0, "excel": 0.0, "ppt": 0.0}
        self.first_seen = {}
        self.unsaved = []                  # documents with unsaved changes, as seen at the last check
        self.com_ok = False
        self._stop = threading.Event()
        self._t0 = time.monotonic()
        self._files = {}
        self._procs = set()
        self._fg = None
        self._docs = {}
        self._formulas = {}
        self._slides = {}
        self._sheets = {}
        self._charts = {}
        self._lock = threading.Lock()
        self._thread = threading.Thread(target=self._run, daemon=True)

    # ---------- public
    def start(self):
        self._files = self._scan_files()
        self._procs = running_exes() if IS_WIN else set()
        self.log("test", "Test started. Test folder: " + str(self.folder))
        self._thread.start()

    def stop(self):
        self._stop.set()
        self.log("test", "Test submitted")

    def log(self, app, msg):
        ev = {"time": datetime.now().strftime("%H:%M:%S"), "elapsed": int(time.monotonic() - self._t0),
              "app": app, "msg": msg}
        with self._lock:
            self.events.append(ev)
        if self.on_event:
            try:
                self.on_event(ev)
            except Exception:
                pass

    # ---------- loop
    def _run(self):
        com = None
        if IS_WIN:
            try:
                import pythoncom
                import win32com.client
                pythoncom.CoInitialize()
                com = win32com.client
                self.com_ok = True
            except Exception:
                com = None
        tick = 0
        last = time.monotonic()
        while not self._stop.is_set():
            now = time.monotonic()
            dt = now - last
            last = now
            try:
                self._check_foreground(dt)
            except Exception:
                pass
            if tick % 2 == 0:
                try:
                    self._check_files()
                except Exception:
                    pass
                try:
                    self._check_processes()
                except Exception:
                    pass
            if com is not None and tick % 3 == 0:
                for fn in (self._com_word, self._com_excel, self._com_ppt):
                    try:
                        fn(com)
                    except Exception:
                        pass       # Office busy (e.g. editing a cell); try again next time
            tick += 1
            self._stop.wait(1.0)

    # ---------- checks
    def _check_foreground(self, dt):
        exe, title = foreground()
        app = EXE_APP.get(exe)
        if app:
            self.app_seconds[app] += min(dt, 5)
            if app not in self.first_seen:
                self.first_seen[app] = int(time.monotonic() - self._t0)
        if app != self._fg and exe:
            if app:
                self.log(app, f"Switched to {APP_NAME[app]}" + (f" – {title}" if title else ""))
            self._fg = app

    def _scan_files(self):
        out = {}
        if not self.folder.exists():
            return out
        for p in self.folder.rglob("*"):
            if p.is_file() and not p.name.startswith("~$") and "Pictures" not in p.parts:
                try:
                    st = p.stat()
                    out[str(p.relative_to(self.folder))] = (st.st_mtime, st.st_size)
                except OSError:
                    pass
        return out

    def _check_files(self):
        now = self._scan_files()
        for name, (mt, sz) in now.items():
            if name not in self._files:
                self.log("folder", f"File created: {name}")
            elif self._files[name][0] != mt:
                self.log("folder", f"File saved: {name}")
        for name in self._files:
            if name not in now:
                self.log("folder", f"File deleted/renamed: {name}")
        self._files = now

    def _check_processes(self):
        if not IS_WIN:
            return
        now = running_exes()
        for exe, app in EXE_APP.items():
            if exe in now and exe not in self._procs:
                self.log(app, f"{APP_NAME[app]} opened")
            if exe not in now and exe in self._procs:
                self.log(app, f"{APP_NAME[app]} closed")
        self._procs = now

    def _get(self, com, progid):
        try:
            return com.GetActiveObject(progid)
        except Exception:
            return None

    def _track_docs(self, app, items):
        """items: {fullname: (saved, extra_dict)}"""
        prev = {k: v for k, v in self._docs.items() if k[0] == app}
        for name, (saved, extra) in items.items():
            key = (app, name)
            short = os.path.basename(name) or name
            if key not in self._docs:
                self.log(app, f"Opened/created: {short}")
            else:
                old_saved, old_extra = self._docs[key]
                if saved and not old_saved:
                    self.log(app, f"Saved: {short}")
                for k, v in extra.items():
                    if old_extra.get(k) != v:
                        self.log(app, f"{k} {'ON' if v else 'OFF'} in {short}" if isinstance(v, bool) else f"{k}: {v} ({short})")
            self._docs[key] = (saved, extra)
        for key in prev:
            if key[1] not in items:
                self.log(app, f"Closed: {os.path.basename(key[1]) or key[1]}")
                self._docs.pop(key, None)
        unsaved = [os.path.basename(n) or n for n, (s, _) in items.items() if not s]
        self.unsaved = [u for u in self.unsaved if not u.startswith(app + ":")] + [f"{app}:{u}" for u in unsaved]

    def _com_word(self, com):
        w = self._get(com, "Word.Application")
        if w is None:
            return
        items = {}
        for i in range(1, w.Documents.Count + 1):
            d = w.Documents.Item(i)
            items[d.FullName] = (bool(d.Saved), {"Track Changes": bool(d.TrackRevisions)})
        self._track_docs("word", items)

    def _com_ppt(self, com):
        p = self._get(com, "PowerPoint.Application")
        if p is None:
            return
        items = {}
        for i in range(1, p.Presentations.Count + 1):
            pr = p.Presentations.Item(i)
            items[pr.FullName] = (bool(pr.Saved), {"Slides": int(pr.Slides.Count)})
        self._track_docs("ppt", items)

    def _com_excel(self, com):
        x = self._get(com, "Excel.Application")
        if x is None:
            return
        items = {}
        for i in range(1, x.Workbooks.Count + 1):
            wb = x.Workbooks.Item(i)
            name = wb.FullName
            items[name] = (bool(wb.Saved), {})
            in_folder = not os.path.dirname(name) or str(self.folder).lower() in name.lower()
            if not in_folder:
                continue
            sheets = [wb.Worksheets.Item(j) for j in range(1, wb.Worksheets.Count + 1)]
            names = [s.Name for s in sheets]
            old = self._sheets.get(name)
            if old is not None and old != names:
                added = [n for n in names if n not in old]
                gone = [n for n in old if n not in names]
                if added:
                    self.log("excel", f"Sheet added/renamed: {', '.join(added)} ({os.path.basename(name)})")
                if gone:
                    self.log("excel", f"Sheet removed/renamed: {', '.join(gone)}")
            self._sheets[name] = names
            new_logs = []
            for s in sheets:
                try:
                    nch = int(s.ChartObjects().Count)
                except Exception:
                    nch = 0
                ck = (name, s.Name)
                prev = self._charts.get(ck)
                if prev is not None and nch > prev:
                    self.log("excel", f"Chart inserted on {s.Name}")
                self._charts[ck] = nch
                ur = s.UsedRange
                if int(ur.Count) > 5000:
                    continue
                vals = ur.Formula
                r0, c0 = int(ur.Row), int(ur.Column)
                if not isinstance(vals, tuple):
                    vals = ((vals,),)
                for ri, row in enumerate(vals):
                    for ci, v in enumerate(row):
                        if isinstance(v, str) and v.startswith("="):
                            key = (name, s.Name, r0 + ri, c0 + ci)
                            if self._formulas.get(key) != v:
                                new_logs.append(f"{s.Name}!{_col_letter(c0 + ci)}{r0 + ri}: {v}")
                                self._formulas[key] = v
            if new_logs:
                shown = new_logs[:4]
                more = f" (+{len(new_logs) - 4} more)" if len(new_logs) > 4 else ""
                self.log("excel", "Formula entered – " + "; ".join(shown) + more)
        self._track_docs("excel", items)


# ---------------------------------------------------------------- split screen

def arrange_split(tk_root, app_share=0.42):
    """Put the open Office windows on the left and the app on the right of the screen.
    Returns how many Office windows were moved (always 0 outside Windows)."""
    if not IS_WIN:
        sw, sh = tk_root.winfo_screenwidth(), tk_root.winfo_screenheight()
        w = int(sw * app_share)
        tk_root.wm_geometry(f"{w}x{sh - 80}+{sw - w}+0")
        return 0
    user32 = ctypes.windll.user32

    class RECT(ctypes.Structure):
        _fields_ = [("left", ctypes.c_long), ("top", ctypes.c_long), ("right", ctypes.c_long), ("bottom", ctypes.c_long)]

    work = RECT()
    user32.SystemParametersInfoW(48, 0, ctypes.byref(work), 0)          # SPI_GETWORKAREA (without the taskbar)
    width, height = work.right - work.left, work.bottom - work.top
    split = work.left + int(width * (1 - app_share))

    office = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, wintypes.HWND, wintypes.LPARAM)
    def each(hwnd, _):
        if user32.IsWindowVisible(hwnd) and user32.GetWindowTextLengthW(hwnd) > 0 and not user32.GetWindow(hwnd, 4):
            pid = wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            h = _kernel32.OpenProcess(0x1000, False, pid.value)
            if h:
                size = wintypes.DWORD(1024)
                buf = ctypes.create_unicode_buffer(1024)
                if _kernel32.QueryFullProcessImageNameW(h, 0, buf, ctypes.byref(size)):
                    if os.path.basename(buf.value).lower() in EXE_APP:
                        office.append(hwnd)
                _kernel32.CloseHandle(h)
        return True

    user32.EnumWindows(each, 0)
    flags = 0x0004 | 0x0040                                              # SWP_NOZORDER | SWP_SHOWWINDOW
    for hwnd in office:
        user32.ShowWindow(hwnd, 9)                                       # SW_RESTORE (un-maximise / un-minimise)
        user32.SetWindowPos(hwnd, 0, work.left, work.top, split - work.left, height, flags)
    try:
        me = int(tk_root.wm_frame(), 16)
        tk_root.update_idletasks()
        user32.ShowWindow(me, 9)
        user32.SetWindowPos(me, 0, split, work.top, work.right - split, height, flags)
    except Exception:
        pass
    return len(office)
