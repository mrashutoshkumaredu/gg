"""Check GitHub for a newer release and replace the running .exe with it."""
from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import urllib.request
from pathlib import Path

REPO = "mrashutoshkumaredu/gg"
API = f"https://api.github.com/repos/{REPO}/releases/latest"
PAGE = f"https://github.com/{REPO}/releases/latest"

try:
    from ._build import BUILD          # written by the GitHub build (the run number)
except ImportError:
    BUILD = 0


def frozen():
    return getattr(sys, "frozen", False)


def build_of(tag):
    """'v1.0.12' -> 12"""
    try:
        return int(str(tag).lstrip("v").split(".")[-1])
    except ValueError:
        return 0


def latest(timeout=8):
    """Return dict(tag, build, notes, url, size) of the newest release, or None if offline."""
    req = urllib.request.Request(API, headers={"Accept": "application/vnd.github+json", "User-Agent": "ISRO-SkillTest"})
    with urllib.request.urlopen(req, timeout=timeout) as r:
        d = json.load(r)
    asset = next((a for a in d.get("assets", []) if a["name"].lower().endswith(".exe")), None)
    if not asset:
        return None
    return {"tag": d.get("tag_name", ""), "build": build_of(d.get("tag_name")), "notes": d.get("body") or "",
            "url": asset["browser_download_url"], "size": asset.get("size", 0), "page": d.get("html_url", PAGE)}


def update_available(info):
    return bool(info) and info["build"] > BUILD and BUILD > 0


def download(info, progress=None):
    """Download the new exe next to the running one. progress(done_bytes, total_bytes)."""
    target = Path(sys.executable).with_name("ISRO_SkillTest_update.exe")
    req = urllib.request.Request(info["url"], headers={"User-Agent": "ISRO-SkillTest"})
    with urllib.request.urlopen(req, timeout=30) as r, open(target, "wb") as f:
        total = int(r.headers.get("Content-Length") or info.get("size") or 0)
        done = 0
        while True:
            chunk = r.read(256 * 1024)
            if not chunk:
                break
            f.write(chunk)
            done += len(chunk)
            if progress:
                progress(done, total)
    if total and target.stat().st_size != total:
        target.unlink(missing_ok=True)
        raise IOError("The download was incomplete. Please try again.")
    return target


def install_and_restart(new_exe):
    """A running .exe cannot replace itself, so a tiny batch file waits for the app to close,
    swaps the files and starts the new version. The caller must quit the app right after this."""
    old = Path(sys.executable)
    bat = Path(tempfile.gettempdir()) / "isro_skilltest_update.bat"
    bat.write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        f"set OLD=\"{old}\"\r\n"
        f"set NEW=\"{new_exe}\"\r\n"
        "set N=0\r\n"
        ":wait\r\n"
        "ping -n 2 127.0.0.1 >nul\r\n"                # wait about 1 second (works without a console)
        "move /y %NEW% %OLD% >nul 2>&1\r\n"
        "if not errorlevel 1 goto done\r\n"
        "set /a N+=1\r\n"
        "if %N% lss 30 goto wait\r\n"
        ":done\r\n"
        "start \"\" %OLD%\r\n"
        "del \"%~f0\"\r\n", encoding="utf-8")
    flags = 0x00000008 | 0x00000200 | 0x08000000        # DETACHED, NEW_PROCESS_GROUP, NO_WINDOW
    subprocess.Popen(["cmd", "/c", str(bat)], creationflags=flags, close_fds=True)


def cleanup_leftovers():
    if frozen():
        left = Path(sys.executable).with_name("ISRO_SkillTest_update.exe")
        try:
            if left.exists():
                left.unlink()
        except OSError:
            pass
