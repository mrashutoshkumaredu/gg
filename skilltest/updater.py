"""Check GitHub for a newer release and replace the running .exe with it.

Reaching GitHub is tried three ways, because many PCs block one of them:
1. Python's normal connection, using Windows' list of trusted certificates
2. the same, with the certificate list that is bundled inside the app (certifi)
3. Windows PowerShell, which uses Windows' own downloader, proxy settings and certificate updates
"""
from __future__ import annotations

import hashlib
import json
import os
import socket
import ssl
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
from pathlib import Path

from . import __version__

REPO = "mrashutoshkumaredu/gg"
API = f"https://api.github.com/repos/{REPO}/releases/latest"
PAGE = f"https://github.com/{REPO}/releases/latest"
HEADERS = {"Accept": "application/vnd.github+json", "User-Agent": "ISRO-SkillTest"}

try:
    from ._build import BUILD          # written by the GitHub build (the run number)
except ImportError:
    BUILD = 0


def frozen():
    return getattr(sys, "frozen", False)


def version_string():
    """The one version shown everywhere, e.g. 1.7.12 (same as the release name)."""
    return f"{__version__}.{BUILD}" if BUILD else f"{__version__} (dev)"


def parse(tag):
    """'v1.7.12' -> (1, 7, 12)"""
    out = []
    for part in str(tag).lstrip("vV").split("."):
        try:
            out.append(int(part))
        except ValueError:
            out.append(0)
    return tuple((out + [0, 0, 0])[:3])


def current():
    return parse(f"{__version__}.{BUILD}")


class UpdateError(Exception):
    """Carries a plain-words reason for the user and the technical detail."""
    def __init__(self, reason, detail=""):
        super().__init__(reason)
        self.reason, self.detail = reason, detail


def explain(e):
    """Turn a connection error into words a student understands."""
    text = str(e)
    reason = getattr(e, "reason", None)
    if isinstance(e, urllib.error.HTTPError):
        if e.code == 403 and "rate limit" in text.lower():
            return "GitHub is busy (too many checks from this network). Please try again in an hour."
        if e.code in (403, 407):
            return "The connection was blocked by a proxy, firewall or antivirus."
        if e.code == 404:
            return "No released version was found on GitHub."
        return f"GitHub replied with an error ({e.code}). Please try again later."
    if isinstance(e, ssl.SSLError) or isinstance(reason, ssl.SSLError) or "CERTIFICATE" in text.upper():
        return ("Security certificate problem: this PC (or its antivirus) does not trust GitHub's certificate. "
                "Updating Windows, or opening github.com once in Edge, usually fixes it.")
    if isinstance(e, (socket.timeout, TimeoutError)) or isinstance(reason, (socket.timeout, TimeoutError)) \
            or "timed out" in text.lower():
        return "The connection timed out. The internet may be slow; please try again."
    if isinstance(reason, socket.gaierror) or "getaddrinfo" in text or "Name or service" in text:
        return "No internet connection (github.com could not be found)."
    if "refused" in text.lower() or "proxy" in text.lower():
        return "The connection was refused, possibly by a proxy or firewall."
    return "Could not connect to GitHub."


def _contexts():
    yield "system", None                       # Windows certificate store
    try:
        import certifi
        yield "bundled", ssl.create_default_context(cafile=certifi.where())
    except Exception:
        pass


def _powershell(args, timeout):
    if sys.platform != "win32":
        raise UpdateError("PowerShell is only available on Windows")
    cmd = ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command",
           "[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; "
           "$ProgressPreference = 'SilentlyContinue'; " + args]
    r = subprocess.run(cmd, capture_output=True, timeout=timeout, creationflags=0x08000000)
    if r.returncode != 0:
        raise UpdateError("PowerShell download failed", r.stderr.decode(errors="ignore")[-300:])
    return r.stdout


def _get_json(url, timeout):
    errors = []
    for name, ctx in _contexts():
        try:
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                return json.load(r)
        except urllib.error.HTTPError as e:
            raise UpdateError(explain(e), f"{name}: {e}")          # GitHub answered: retrying will not help
        except Exception as e:
            errors.append((name, e))
    try:
        out = _powershell(f"(Invoke-WebRequest -UseBasicParsing -Headers @{{'User-Agent'='ISRO-SkillTest'}} "
                          f"-Uri '{url}').Content", timeout + 10)
        return json.loads(out.decode("utf-8-sig"))
    except Exception as e:
        errors.append(("powershell", e))
    first = errors[0][1] if errors else Exception("unknown")
    raise UpdateError(explain(first), "; ".join(f"{n}: {type(e).__name__}: {e}" for n, e in errors)[:600])


def _get_text(url, timeout):
    """Small text file (the fingerprint), using the same three ways as _get_json."""
    errors = []
    for name, ctx in _contexts():
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "ISRO-SkillTest"})
            with urllib.request.urlopen(req, timeout=timeout, context=ctx) as r:
                return r.read(4096).decode("utf-8", "ignore")
        except urllib.error.HTTPError as e:
            raise UpdateError(explain(e), f"{name}: {e}")
        except Exception as e:
            errors.append((name, e))
    try:
        out = _powershell(f"(Invoke-WebRequest -UseBasicParsing -Uri '{url}').Content", timeout + 10)
        return out.decode("utf-8-sig", "ignore")
    except Exception as e:
        errors.append(("powershell", e))
    raise UpdateError(explain(errors[0][1]), "; ".join(f"{n}: {e}" for n, e in errors)[:600])


def sha256_of(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1024 * 1024), b""):
            h.update(chunk)
    return h.hexdigest()


def latest(timeout=20):
    """Return dict(tag, notes, url, size, page) of the newest release. Raises UpdateError."""
    d = _get_json(API, timeout)
    asset = next((a for a in d.get("assets", []) if a["name"].lower().endswith(".exe")), None)
    if not asset:
        return None
    sha = next((a for a in d.get("assets", []) if a["name"].lower().endswith(".sha256")), None)
    return {"tag": d.get("tag_name", ""), "notes": d.get("body") or "", "url": asset["browser_download_url"],
            "sha_url": sha["browser_download_url"] if sha else None,
            "size": asset.get("size", 0), "page": d.get("html_url", PAGE)}


def update_available(info):
    return bool(info) and BUILD > 0 and parse(info["tag"]) > current()


def download(info, progress=None):
    """Download the new exe next to the running one and check its fingerprint. progress(done_bytes, total_bytes)."""
    if not info.get("sha_url"):
        raise UpdateError("This release has no fingerprint (checksum), so it cannot be checked and will not be "
                          "installed automatically. Please download it from the download page.")
    text = _get_text(info["sha_url"], 20).strip()
    expected = text.split()[0].lower() if text else ""
    if len(expected) != 64:
        raise UpdateError("The release fingerprint could not be read, so the update was not installed.")
    target = Path(sys.executable).with_name("ISRO_SkillTest_update.exe")
    total = int(info.get("size") or 0)
    errors = []
    for name, ctx in _contexts():
        try:
            req = urllib.request.Request(info["url"], headers={"User-Agent": "ISRO-SkillTest"})
            with urllib.request.urlopen(req, timeout=60, context=ctx) as r, open(target, "wb") as f:
                total = int(r.headers.get("Content-Length") or total or 0)
                done = 0
                while True:
                    chunk = r.read(256 * 1024)
                    if not chunk:
                        break
                    f.write(chunk)
                    done += len(chunk)
                    if progress:
                        progress(done, total)
            break
        except PermissionError as e:
            raise UpdateError("The app's folder is write-protected (for example Program Files). Move "
                              "ISRO_SkillTest.exe to Desktop or Downloads and try again.", str(e))
        except Exception as e:
            errors.append((name, e))
    else:
        try:
            if progress:
                progress(0, total)
            _powershell(f"Invoke-WebRequest -UseBasicParsing -Uri '{info['url']}' -OutFile '{target}'", 600)
        except Exception as e:
            errors.append(("powershell", e))
            first = errors[0][1]
            raise UpdateError(explain(first), "; ".join(f"{n}: {e}" for n, e in errors)[:600])
    if not target.exists() or (total and target.stat().st_size != total):
        target.unlink(missing_ok=True)
        raise UpdateError("The download was incomplete. Please try again.")
    got = sha256_of(target)
    if got != expected:
        target.unlink(missing_ok=True)
        raise UpdateError("The downloaded file is not genuine or is damaged (its fingerprint does not match). "
                          "Nothing was changed.", f"expected {expected[:16]}…, got {got[:16]}…")
    if progress:
        progress(total or 1, total or 1)
    return target


def clean_env():
    """Environment for starting the new version as a separate program.

    A one-file .exe unpacks itself into a temporary _MEI folder and deletes it when it closes. Settings that point
    into that folder must not be passed on, otherwise the new version looks for its files in the old (deleted)
    folder and fails with "Failed to load Python DLL"."""
    env = dict(os.environ)
    temp_dir = getattr(sys, "_MEIPASS", None)
    for k in list(env):
        v = env[k]
        if k.startswith(("_PYI", "_MEI", "PYINSTALLER")) or (temp_dir and temp_dir.lower() in str(v).lower()):
            del env[k]
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"           # PyInstaller: start fresh, not as a child of this app
    return env


def install_and_restart(new_exe):
    """A running .exe cannot replace itself, so a tiny batch file waits for the app to close,
    swaps the files and starts the new version. The caller must quit the app right after this."""
    old = Path(sys.executable)
    bat = Path(tempfile.gettempdir()) / "isro_skilltest_update.bat"
    note = ("Update installed. If the app did not open by itself, please open "
            f"{old.name} again.").replace("'", "")
    fail = ("The update could not be installed because the file is in use. Please download it from "
            f"{PAGE}").replace("'", "")
    popup = "powershell -NoProfile -WindowStyle Hidden -Command \"Add-Type -AssemblyName PresentationFramework;" \
            "[System.Windows.MessageBox]::Show('{}', 'ISRO Skill Test Practice')\""
    bat.write_text(
        "@echo off\r\n"
        "chcp 65001 >nul\r\n"
        f"set OLD=\"{old}\"\r\n"
        f"set NEW=\"{new_exe}\"\r\n"
        "set N=0\r\n"
        "ping -n 4 127.0.0.1 >nul\r\n"                # give the old app ~3 s to close and remove its temp folder
        ":wait\r\n"
        "move /y %NEW% %OLD% >nul 2>&1\r\n"
        "if not errorlevel 1 goto moved\r\n"
        "set /a N+=1\r\n"
        "ping -n 2 127.0.0.1 >nul\r\n"
        "if %N% lss 30 goto wait\r\n"
        + popup.format(fail) + "\r\n"
        "goto end\r\n"
        ":moved\r\n"
        "ping -n 2 127.0.0.1 >nul\r\n"
        "start \"\" %OLD%\r\n"
        "ping -n 16 127.0.0.1 >nul\r\n"               # after ~15 s, check that the new version is running
        f"tasklist /FI \"IMAGENAME eq {old.name}\" | find /I \"{old.name}\" >nul || " + popup.format(note) + "\r\n"
        ":end\r\n"
        "del \"%~f0\"\r\n", encoding="utf-8")
    flags = 0x00000008 | 0x00000200 | 0x08000000        # DETACHED, NEW_PROCESS_GROUP, NO_WINDOW
    subprocess.Popen(["cmd", "/c", str(bat)], creationflags=flags, close_fds=True, env=clean_env())


def cleanup_leftovers():
    if frozen():
        left = Path(sys.executable).with_name("ISRO_SkillTest_update.exe")
        try:
            if left.exists():
                left.unlink()
        except OSError:
            pass
