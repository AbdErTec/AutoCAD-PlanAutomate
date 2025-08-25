
import time
import pythoncom
import win32com.client
from pathlib import Path
import sys

def resource_path(*parts) -> str:
    """Return an absolute path to bundled resources (works dev + PyInstaller)."""
    base = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parents[2]))
    return str(base.joinpath(*parts))

def wait_for_autocad_ready(app, timeout_s=15):
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        try:
            _ = app.ActiveDocument  # will raise until ready
            return True
        except pythoncom.com_error:
            time.sleep(0.25)
    return False

def ensure_app_and_doc(template_path=None):
    """
    Returns (app, doc, created_new_doc: bool)
    - attaches to existing AutoCAD if present, else launches it
    - uses existing ActiveDocument if present, else creates a new one
    """
    app = None
    doc = None

    # 1) try to attach to a running instance
    try:
        app = win32com.client.GetObject(None, "AutoCAD.Application")
    except Exception:
        # 2) or create one
        app = win32com.client.Dispatch("AutoCAD.Application")

    # show it (standalone UX)
    try:
        app.Visible = True
    except Exception:
        pass

    # wait until the app is ready
    wait_for_autocad_ready(app, timeout_s=20)

    # try use active document first
    try:
        doc = app.ActiveDocument
    except Exception:
        doc = None

    created_new = False
    if doc is None:
        # create a fresh doc
        if template_path:
            doc = app.Documents.Add(template_path)
        else:
            doc = app.Documents.Add()
        created_new = True

    return app, doc, created_new

def reacquire_doc(app, title=None, fullname=None):
    """
    Find an open document by FullName (preferred) or Name.
    Use this in another thread instead of passing COM objects across threads.
    """
    docs = app.Documents
    # COM collections are 1-based
    for i in range(1, docs.Count + 1):
        d = docs.Item(i)
        try:
            if fullname and getattr(d, "FullName", "") == fullname:
                return d
        except Exception:
            pass
        if title and d.Name == title:
            return d
    # fallback
    try:
        return app.ActiveDocument
    except Exception:
        return None
