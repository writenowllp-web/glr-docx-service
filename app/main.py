"""
GLR DOCX Service — Option B.

A small, self-owned service that n8n calls over HTTP. Does the one thing that is
genuinely hard: filling a Word template WITHOUT destroying it.

Endpoints
---------
POST /scan     : template.docx        -> every token + where it lives + surrounding text
POST /fill     : template.docx + JSON -> filled .docx  (headers/footers/styles intact)
POST /convert  : template.docx        -> [TOKEN] converted to {TOKEN} (for docxtemplater)
GET  /health

Why this exists
---------------
- HTML -> docx renderers DESTROY headers, footers, and page numbering.
- Regex-on-XML replacement CORRUPTS .docx files.
- Word's own Find/Replace MISSES tokens split across runs.
This service does the correct thing: edits the XML tree, stitching runs first.

Run:
    uvicorn app.main:app --host 0.0.0.0 --port 8000
"""

from __future__ import annotations

import base64
import json

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import JSONResponse, Response

from .docx_fill import DEFAULT_PATTERN, fill_docx, scan_docx

app = FastAPI(title="GLR DOCX Service", version="1.0")


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/scan")
async def scan(file: UploadFile = File(...), pattern: str = Form(DEFAULT_PATTERN)):
    """
    List every token in the template.

    Feed the response to your AI writer. The `context` field is important: it shows the
    sentence each token sits in, so the model knows e.g. that the template already prints
    a '$' before XM8_COV_RCV_1 and should return '2,614.70' not '$2,614.70'.
    """
    data = await file.read()
    try:
        return scan_docx(data, pattern)
    except Exception as e:
        raise HTTPException(400, f"scan failed: {type(e).__name__}: {e}")


@app.post("/fill")
async def fill(
    file: UploadFile = File(...),
    values: str = Form(...),                 # JSON object: {"XM8_FILE_NO": "..."}
    pattern: str = Form(DEFAULT_PATTERN),
    leave_missing: bool = Form(True),
    as_base64: bool = Form(False),           # True -> JSON w/ base64 (easier in n8n)
):
    """
    Fill the template and return the .docx.

    leave_missing=True  (default): an unfilled token stays visible as [XM8_FOO] so a human
                                   catches it. Safer than silently blanking it.
    leave_missing=False:           unfilled tokens become empty string. Warning: this can
                                   leave dangling punctuation, e.g. "received on ."
    """
    data = await file.read()
    try:
        vals = json.loads(values)
        if not isinstance(vals, dict):
            raise ValueError("values must be a JSON object")
    except Exception as e:
        raise HTTPException(400, f"bad values JSON: {e}")

    try:
        res = fill_docx(data, vals, pattern, leave_missing)
    except Exception as e:
        raise HTTPException(400, f"fill failed: {type(e).__name__}: {e}")

    if as_base64:
        return JSONResponse({
            "docx_base64": base64.b64encode(res.data).decode(),
            "filled": res.filled,
            "missing": res.missing,          # <-- QA on this. Non-empty = incomplete report.
            "paragraphs_changed": res.paragraphs_changed,
        })

    return Response(
        content=res.data,
        media_type=("application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"),
        headers={
            "Content-Disposition": 'attachment; filename="filled.docx"',
            "X-Filled-Count": str(len(res.filled)),
            "X-Missing": ",".join(res.missing) or "none",
        },
    )


@app.post("/convert")
async def convert(file: UploadFile = File(...)):
    """
    Convert [TOKEN] -> {TOKEN} for docxtemplater (Option A).

    Do NOT do this with Word's Find/Replace: Word splits some tokens across internal runs
    (e.g. '[XM8_INSURED_H_' + 'STREET]') and find/replace misses them, so a raw token ships
    in the report. This stitches runs first, and also converts headers/footers.
    """
    import io
    import re
    import zipfile

    from lxml import etree

    from .docx_fill import FILLABLE_PART, W, XML_SPACE

    data = await file.read()
    token_re = re.compile(r"\[([A-Z][A-Z0-9_]*)\]")
    found: list[str] = []

    src = zipfile.ZipFile(io.BytesIO(data))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            blob = src.read(item.filename)
            if FILLABLE_PART.match(item.filename):
                root = etree.fromstring(blob)
                for p in root.iter(f"{W}p"):
                    ts = list(p.iter(f"{W}t"))
                    if not ts:
                        continue
                    joined = "".join(t.text or "" for t in ts)
                    hits = token_re.findall(joined)
                    if not hits:
                        continue
                    found.extend(hits)
                    ts[0].text = token_re.sub(r"{\1}", joined)
                    ts[0].set(XML_SPACE, "preserve")
                    for t in ts[1:]:
                        t.text = ""
                blob = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            out.writestr(item, blob)
    src.close()

    return Response(
        content=buf.getvalue(),
        media_type=("application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"),
        headers={
            "Content-Disposition": 'attachment; filename="converted.docx"',
            "X-Tokens-Converted": str(len(found)),
            "X-Unique-Tokens": str(len(set(found))),
        },
    )


@app.post("/extract_text")
async def extract_text(file: UploadFile = File(...)):
    """
    Return the template's full paragraph text, numbered, with style names.

    This is what the TOKENIZER (Claude) reads. It needs the whole document, not just
    existing tokens, because most carrier templates carry their fillable spots as PROSE:

        "Dwelling is a (one story, raised ranch, two story, etc.) wood framed house"
        "Roof has a (25 year 3-tab, 30 year laminate, etc.) roof with a */12 pitch"
        "(Put N/A if this is the original inspection and delete this paragraph.)"
        "Inspection of the right elevation revealed no visible storm related damage."

    None of those have tokens. All of them are fillable. A regex cannot tell the last one
    (deliverable default prose -> keep it) from the others (needs replacing). Claude can.
    """
    import io, zipfile
    from lxml import etree
    from .docx_fill import W

    data = await file.read()
    src = zipfile.ZipFile(io.BytesIO(data))
    out = []
    for partname, kind in [("word/document.xml", "body"),
                           ("word/header1.xml", "header"),
                           ("word/header2.xml", "header"),
                           ("word/footer1.xml", "footer")]:
        try:
            root = etree.fromstring(src.read(partname))
        except KeyError:
            continue
        for i, p in enumerate(root.iter(f"{W}p")):
            text = "".join(t.text or "" for t in p.iter(f"{W}t"))
            if not text.strip():
                continue
            style = ""
            pstyle = p.find(f"{W}pPr/{W}pStyle")
            if pstyle is not None:
                style = pstyle.get(f"{W}val", "")
            out.append({"i": len(out), "part": kind, "style": style, "text": text})
    src.close()
    return {"paragraphs": out, "count": len(out)}


@app.post("/tokenize")
async def tokenize(
    file: UploadFile = File(...),
    edits: str = Form(...),   # [{"i":11,"new_text":"Dwelling is a [DWELLING_STORIES] ..."}]
):
    """
    Write Claude's tokenized text back into the .docx, IN PLACE.

    `edits` is a list of {"i": <paragraph index from /extract_text>, "new_text": "..."}.
    Only listed paragraphs change. Everything else -- headers, footers, styles, numbering,
    page setup -- is untouched, because we edit the XML tree rather than rebuilding.

    The result is a template with real tokens where there used to be prose, ready for /fill.
    """
    import io, json as _json, zipfile
    from lxml import etree
    from .docx_fill import W, XML_SPACE

    data = await file.read()
    try:
        edit_list = _json.loads(edits)
        by_index = {int(e["i"]): e["new_text"] for e in edit_list}
    except Exception as e:
        raise HTTPException(400, f"bad edits JSON: {e}")

    src = zipfile.ZipFile(io.BytesIO(data))
    buf = io.BytesIO()
    applied = 0
    counter = 0

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            blob = src.read(item.filename)
            if item.filename in ("word/document.xml", "word/header1.xml",
                                 "word/header2.xml", "word/footer1.xml"):
                root = etree.fromstring(blob)
                for p in root.iter(f"{W}p"):
                    ts = list(p.iter(f"{W}t"))
                    text = "".join(t.text or "" for t in ts)
                    if not text.strip():
                        continue
                    idx = counter
                    counter += 1
                    if idx not in by_index or not ts:
                        continue
                    ts[0].text = by_index[idx]
                    ts[0].set(XML_SPACE, "preserve")
                    for t in ts[1:]:
                        t.text = ""
                    applied += 1
                blob = etree.tostring(root, xml_declaration=True,
                                      encoding="UTF-8", standalone=True)
            out.writestr(item, blob)
    src.close()

    return Response(
        content=buf.getvalue(),
        media_type=("application/vnd.openxmlformats-officedocument"
                    ".wordprocessingml.document"),
        headers={
            "Content-Disposition": 'attachment; filename="tokenized.docx"',
            "X-Edits-Applied": str(applied),
            "X-Edits-Requested": str(len(by_index)),
        },
    )
