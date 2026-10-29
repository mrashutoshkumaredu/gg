"""Check GitHub for a newer release and replace the running .exe with it.

Reaching GitHub is tried three ways, because many PCs block one of them:
1. Python's normal connection, using Windows' list of trusted certificates
2. the same, with the certificate list that is bundled inside the app (certifi)
3. Windows PowerShell, which uses Windows' own downloader, proxy settings and certificate updates
"""
from __future__ import annotations

import hashlib
import json
import ntpath
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
HEADERS = {"Accept": "application/vnd.github+json", "User-Agent": "ComputerSkillTest"}

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
        out = _powershell(f"(Invoke-WebRequest -UseBasicParsing -Headers @{{'User-Agent'='ComputerSkillTest'}} "
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
            req = urllib.request.Request(url, headers={"User-Agent": "ComputerSkillTest"})
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
    target = Path(sys.executable).with_name(PENDING)
    total = int(info.get("size") or 0)
    errors = []
    for name, ctx in _contexts():
        try:
            req = urllib.request.Request(info["url"], headers={"User-Agent": "ComputerSkillTest"})
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
                              f"{Path(sys.executable).name} to Desktop or Downloads and try again.", str(e))
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
    remember_pending(target, info.get("tag", ""), got)
    if progress:
        progress(total or 1, total or 1)
    return target


def clean_env():
    """Environment for starting the new version as a separate program.

    A one-file .exe unpacks itself into a temporary _MEI folder and deletes it when it closes. Settings that point
    into that folder must not be passed on, otherwise the new version looks for its files in the old (deleted)
    folder and fails with "Failed to load Python DLL". Only those entries are removed: PATH itself must stay,
    or Windows cannot find its own tools (ping, tasklist, PowerShell) that the update helper uses."""
    env = dict(os.environ)
    temp_dir = (getattr(sys, "_MEIPASS", None) or "").rstrip("\\/").lower()
    for k in list(env):
        if k.upper().startswith(("_PYI", "_MEI", "PYINSTALLER")):
            del env[k]
        elif temp_dir and temp_dir in env[k].lower():
            keep = [p for p in env[k].split(os.pathsep) if p and temp_dir not in p.lower()]
            if keep:
                env[k] = os.pathsep.join(keep)
            else:
                del env[k]
    env["PYINSTALLER_RESET_ENVIRONMENT"] = "1"           # PyInstaller: start fresh, not as a child of this app
    return env


PENDING = "ComputerSkillTest_update.exe"     # downloaded and checked, waiting to replace the app
PENDING_NOTE = "ComputerSkillTest_update.json"
OLD_PENDING = ("ISRO_SkillTest_update.exe", "ISRO_SkillTest_update.json")     # left by versions before the new name


def _tools():
    """Full paths of the Windows tools the helper uses, so it works even if PATH is missing or damaged."""
    root = os.environ.get("SYSTEMROOT") or os.environ.get("SystemRoot") or os.environ.get("WINDIR") or r"C:\Windows"
    s32 = ntpath.join(root, "System32")
    ps = ntpath.join(s32, "WindowsPowerShell", "v1.0")
    return {"root": root, "s32": s32, "ps_dir": ps, "ps": ntpath.join(ps, "powershell.exe"),
            "ping": ntpath.join(s32, "ping.exe"), "tasklist": ntpath.join(s32, "tasklist.exe"),
            "find": ntpath.join(s32, "find.exe"), "chcp": ntpath.join(s32, "chcp.com"),
            "cmd": ntpath.join(s32, "cmd.exe")}


def helper_script(old, new_exe, pids, note_file=None):
    """Text of the batch file that waits for the app to close, swaps the files and starts the new version."""
    t = _tools()
    name = ntpath.basename(str(old))
    note = ("Update installed. If the app did not open by itself, please open "
            f"{name} again.").replace("'", "")
    fail = ("The update could not be installed because the file is in use. Close the app and open it again, "
            f"or download the new version from {PAGE}").replace("'", "")
    popup = f"\"{t['ps']}\" -NoProfile -WindowStyle Hidden -Command \"Add-Type -AssemblyName PresentationFramework;" \
            "[System.Windows.MessageBox]::Show('{}', 'Computer Skill Test')\""
    sleep = f"\"{t['ping']}\" -n {{}} 127.0.0.1 >nul"
    lines = ["@echo off",
             f"set \"PATH={t['s32']};{t['root']};{t['ps_dir']};%PATH%\"",
             f"\"{t['chcp']}\" 65001 >nul",
             f"set OLD=\"{old}\"",
             f"set NEW=\"{new_exe}\"",
             "set N=0",
             ":closing"]                                  # wait (up to ~60 s) until the old app has really closed
    lines.append("set BUSY=0")
    for pid in pids:
        lines.append(f"\"{t['tasklist']}\" /NH /FI \"PID eq {int(pid)}\" /FI \"IMAGENAME eq {name}\" 2>nul "
                     f"| \"{t['find']}\" /I \"{name}\" >nul && set BUSY=1")
    lines += ["if %BUSY%==0 goto closed",
              "set /a N+=1",
              sleep.format(2),
              "if %N% lss 60 goto closing",
              ":closed",
              sleep.format(3),                            # let it finish removing its temporary folder
              "set N=0",
              ":wait",
              "move /y %NEW% %OLD% >nul 2>&1",
              "if not errorlevel 1 goto moved",
              "set /a N+=1",
              sleep.format(2),
              "if %N% lss 30 goto wait",
              popup.format(fail),
              "goto end",
              ":moved"]
    if note_file:
        lines.append(f"del \"{note_file}\" >nul 2>&1")
    lines += [sleep.format(2),
              "start \"\" %OLD%",
              sleep.format(16),                           # after ~15 s, check that the new version is running
              f"\"{t['tasklist']}\" /NH /FI \"IMAGENAME eq {name}\" 2>nul | \"{t['find']}\" /I \"{name}\" >nul "
              "|| " + popup.format(note),
              ":end",
              "del \"%~f0\""]
    return "\r\n".join(lines) + "\r\n"


def install_and_restart(new_exe):
    """A running .exe cannot replace itself, so a tiny batch file waits for the app to close,
    swaps the files and starts the new version. The caller must quit the app right after this."""
    old = Path(sys.executable)
    pids = [os.getpid()]
    try:
        pids.append(os.getppid())          # one-file .exe: the small starter process that holds the file too
    except (AttributeError, OSError):
        pass
    bat = Path(tempfile.gettempdir()) / "computerskilltest_update.bat"
    bat.write_text(helper_script(old, new_exe, pids, old.with_name(PENDING_NOTE)), encoding="utf-8")
    flags = 0x00000008 | 0x00000200 | 0x08000000        # DETACHED, NEW_PROCESS_GROUP, NO_WINDOW
    subprocess.Popen([_tools()["cmd"], "/c", str(bat)], creationflags=flags, close_fds=True, env=clean_env())


def remember_pending(target, tag, digest):
    """Note that a checked new version is waiting, so it can still be installed if the swap is interrupted."""
    try:
        Path(target).with_name(PENDING_NOTE).write_text(
            json.dumps({"tag": tag, "sha256": digest, "tries": 0}), encoding="utf-8")
    except OSError:
        pass


def finish_pending():
    """Run when the app starts. If an earlier update was downloaded and checked but not swapped in, finish it.

    Returns True when the helper was started; the app must then quit at once (it reopens as the new version).
    A leftover that is not newer, does not match its fingerprint, or has already failed twice is deleted."""
    if not frozen():
        return False
    exe = Path(sys.executable)
    new, note = exe.with_name(PENDING), exe.with_name(PENDING_NOTE)
    try:
        if new.exists() and note.exists():
            m = json.loads(note.read_text(encoding="utf-8"))
            tries = int(m.get("tries", 0))
            if tries < 2 and parse(m.get("tag", "")) > current() and sha256_of(new) == m.get("sha256"):
                m["tries"] = tries + 1
                note.write_text(json.dumps(m), encoding="utf-8")
                install_and_restart(new)
                return True
    except Exception:
        pass
    for f in (new, note) + tuple(exe.with_name(n) for n in OLD_PENDING):
        try:
            f.unlink(missing_ok=True)
        except OSError:
            pass
    return False
