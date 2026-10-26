"""Checks that look at files themselves (saved, format, password) rather than their content."""
from __future__ import annotations

import zipfile

from .core import Check, Missing

CFB = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"   # header of old Office files and of password-protected files

FORMAT_INFO = {
    # kind: (label, how)
    "pdf": ("PDF (.pdf)", "File > Save As (or Export) > choose 'PDF (*.pdf)' as the type > Save."),
    "rtf": ("Rich Text Format (.rtf)", "File > Save As > Save as type: 'Rich Text Format (*.rtf)' > Save."),
    "txt": ("Plain Text (.txt)", "File > Save As > Save as type: 'Plain Text (*.txt)' > Save > OK on the encoding box."),
    "doc97": ("Word 97-2003 Document (.doc)", "File > Save As > Save as type: 'Word 97-2003 Document (*.doc)' > Save."),
    "xls97": ("Excel 97-2003 Workbook (.xls)", "File > Save As > Save as type: 'Excel 97-2003 Workbook (*.xls)' > Save."),
    "ppt97": ("PowerPoint 97-2003 (.ppt)", "File > Save As > Save as type: 'PowerPoint 97-2003 Presentation (*.ppt)' > Save."),
    "dotx": ("Word Template (.dotx)", "File > Save As > Save as type: 'Word Template (*.dotx)'. Check the folder is still the test folder, then Save."),
    "xltx": ("Excel Template (.xltx)", "File > Save As > Save as type: 'Excel Template (*.xltx)'. Check the folder is still the test folder, then Save."),
    "potx": ("PowerPoint Template (.potx)", "File > Save As > Save as type: 'PowerPoint Template (*.potx)'. Check the folder, then Save."),
    "ppsx": ("PowerPoint Show (.ppsx)", "File > Save As > Save as type: 'PowerPoint Show (*.ppsx)' > Save."),
    "csv": ("CSV (.csv)", "File > Save As > Save as type: 'CSV (Comma delimited) (*.csv)' > Save > Yes/OK to keep the format."),
    "html": ("Web Page (.htm/.html)", "File > Save As > Save as type: 'Web Page (*.htm; *.html)' > Save."),
}


def _head(p, n=8):
    with open(p, "rb") as f:
        return f.read(n)


def _zip_content_types(p):
    with zipfile.ZipFile(p) as z:
        return z.read("[Content_Types].xml").decode("utf8", "ignore")


def detect_format(p) -> str:
    """Return a short description of what the file really is."""
    h = _head(p, 8)
    if h.startswith(b"%PDF"):
        return "pdf"
    if h.startswith(b"{\\rtf"):
        return "rtf"
    if h == CFB:
        try:
            import olefile
            ole = olefile.OleFileIO(str(p))
            names = ["/".join(e) for e in ole.listdir()]
            ole.close()
            if "EncryptedPackage" in names:
                return "encrypted"
            if "WordDocument" in names:
                return "doc97"
            if "Workbook" in names or "Book" in names:
                return "xls97"
            if "PowerPoint Document" in names:
                return "ppt97"
        except Exception:
            pass
        return "cfb"
    if h.startswith(b"PK"):
        try:
            ct = _zip_content_types(p)
        except Exception:
            return "zip"
        for key, kind in (("template.main", None), ("slideshow.main", "ppsx")):
            if key in ct:
                if kind:
                    return kind
                if "wordprocessingml" in ct:
                    return "dotx"
                if "spreadsheetml" in ct:
                    return "xltx"
                if "presentationml" in ct:
                    return "potx"
        if "wordprocessingml.document" in ct:
            return "docx"
        if "spreadsheetml.sheet" in ct:
            return "xlsx"
        if "presentationml.presentation" in ct:
            return "pptx"
        return "zip"
    low = h.lower()
    if low.startswith(b"<html") or low.startswith(b"<!doc") or low.startswith(b"mime-") or low.startswith(b"<?xml"):
        return "html"
    return "text"


def file_saved(name, marks=1.0, topic="Saving files", what="file"):
    def fn(ctx):
        p = ctx.exact(name)
        if p:
            rel = p.relative_to(ctx.folder)
            return True, f"Found {rel}"
        try:
            other = ctx.path(name)
            return False, f"{name} not found. A file called {other.name} was found instead (wrong name)."
        except Missing as e:
            return False, str(e)
    return Check(f"{what.capitalize()} saved as {name}", marks, fn,
                 expected=f"A file named exactly {name} in the test folder",
                 how=f"Press Ctrl+S (or F12 for Save As). Choose the test folder, type the name {name.rsplit('.', 1)[0]} "
                     f"and keep the correct file type, then click Save.",
                 topic=topic)


def file_format(name, kind, marks=1.0, topic="Converting files to different formats"):
    label, how = FORMAT_INFO.get(kind, (kind, ""))

    def fn(ctx):
        p = ctx.exact(name)
        if not p:
            raise Missing(f"File {name} was not found in the test folder.")
        real = detect_format(p)
        if kind == "html" and real in ("html", "text"):
            return True, f"{p.name} is a web page file"
        if kind in ("txt", "csv") and real == "text":
            return True, f"{p.name} is a plain text file"
        if real == kind:
            return True, f"{p.name} is a {label} file"
        return False, f"{p.name} exists but it is really a '{real}' file, not {label}. It was probably renamed instead of using Save As."
    return Check(f"Saved a copy as {label}: {name}", marks, fn, expected=f"{name} saved in {label} format",
                 how=how, topic=topic)


def password(name, pw, marks=2.0, topic="Security"):
    def fn(ctx):
        p = ctx.exact(name)
        if not p:
            raise Missing(f"File {name} was not found in the test folder.")
        if detect_format(p) != "encrypted":
            return False, f"{p.name} opens without a password (it is not protected)."
        try:
            import io
            import msoffcrypto
            with open(p, "rb") as f:
                of = msoffcrypto.OfficeFile(f)
                of.load_key(password=pw)
                of.decrypt(io.BytesIO())
            return True, f"{p.name} is protected and the password '{pw}' opens it"
        except ImportError:
            return True, f"{p.name} is password protected"
        except Exception:
            return False, f"{p.name} is protected, but '{pw}' does not open it (a different password was used)."
    return Check(f"{name} protected with password '{pw}'", marks, fn,
                 expected=f"Opening {name} asks for a password; '{pw}' opens it",
                 how="File > Info > Protect Document/Workbook/Presentation > Encrypt with Password > type the password "
                     "> OK > type it again > OK > Save. (Office 2007: Office button > Prepare > Encrypt Document.)",
                 topic=topic)


def decrypt_to_temp(path, pw):
    """Open a password-protected file and return a temp path to a normal copy, so its content can be checked."""
    import io
    import tempfile
    import msoffcrypto
    out = io.BytesIO()
    with open(path, "rb") as f:
        of = msoffcrypto.OfficeFile(f)
        of.load_key(password=pw)
        of.decrypt(out)
    suffix = "." + str(path).rsplit(".", 1)[-1]
    fd, tmp = tempfile.mkstemp(suffix=suffix)
    with open(fd, "wb") as f:
        f.write(out.getvalue())
    return tmp
