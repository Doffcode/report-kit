"""Fill TOC / List of Figures / List of Tables fields with real page numbers.

Pandoc leaves TOC fields empty; Google Docs and LibreOffice never refresh
them. This opens the DOCX in headless LibreOffice, updates every index
twice (the second pass picks up page shifts caused by the first), and saves.
Also exports a PDF of the DOCX so the layout inspector can check Word output.

Runs under the SYSTEM python3 (it needs LibreOffice's `uno` module):
    python3 lo_refresh.py report.docx [docx-preview.pdf]
"""
import os
import subprocess
import sys
import tempfile
import time

import uno
from com.sun.star.beans import PropertyValue


def prop(name, value):
    p = PropertyValue()
    p.Name, p.Value = name, value
    return p


def insert_caption_indexes(doc):
    """Replace [[KIT-LIST:<style>]] markers (written by docx_fix.py) with an
    index built from every paragraph in that caption style (LOF / LOT)."""
    para_styles = doc.getStyleFamilies().getByName("ParagraphStyles")
    while True:
        sd = doc.createSearchDescriptor()
        sd.SearchString = "[[KIT-LIST:"
        found = doc.findFirst(sd)
        if found is None:
            return
        cursor = found.getText().createTextCursorByRange(found)
        cursor.gotoEndOfParagraph(True)
        style = cursor.getString()[len("[[KIT-LIST:"):-2]
        cursor.setString("")
        name = next((n for n in para_styles.getElementNames()
                     if n == style or para_styles.getByName(n).DisplayName == style), style)
        idx = doc.createInstance("com.sun.star.text.ContentIndex")
        idx.CreateFromOutline = False
        idx.CreateFromMarks = False
        idx.CreateFromLevelParagraphStyles = True
        idx.Title = ""
        uno.invoke(idx.LevelParagraphStyles, "replaceByIndex", (0, uno.Any("[]string", (name,))))
        cursor.getText().insertTextContent(cursor, idx, False)


def main(docx_path, pdf_path=None):
    docx_path = os.path.abspath(docx_path)
    profile = tempfile.mkdtemp(prefix="lo-profile-")
    pipe = f"reportkit{os.getpid()}"
    office = subprocess.Popen([
        "soffice", "--headless", "--invisible", "--nologo", "--norestore",
        f"-env:UserInstallation=file://{profile}",
        f"--accept=pipe,name={pipe};urp;StarOffice.ComponentContext",
    ], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    try:
        local = uno.getComponentContext()
        resolver = local.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local)
        for _ in range(100):
            try:
                ctx = resolver.resolve(f"uno:pipe,name={pipe};urp;StarOffice.ComponentContext")
                break
            except Exception:
                time.sleep(0.2)
        else:
            raise SystemExit("LibreOffice did not start")
        desktop = ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        url = uno.systemPathToFileUrl(docx_path)
        doc = desktop.loadComponentFromURL(url, "_blank", 0, (prop("Hidden", True),))
        insert_caption_indexes(doc)
        indexes = doc.getDocumentIndexes()
        for _ in range(2):
            for i in range(indexes.getCount()):
                indexes.getByIndex(i).update()
            doc.getTextFields().refresh()
        print(f"updated {indexes.getCount()} index(es)")
        doc.storeToURL(url, (prop("FilterName", "MS Word 2007 XML"),))
        if pdf_path:
            doc.storeToURL(uno.systemPathToFileUrl(os.path.abspath(pdf_path)),
                           (prop("FilterName", "writer_pdf_Export"),))
        doc.close(True)
    finally:
        office.terminate()
        office.wait(timeout=20)


if __name__ == "__main__":
    main(sys.argv[1], sys.argv[2] if len(sys.argv) > 2 else None)
