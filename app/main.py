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
