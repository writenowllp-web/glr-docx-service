"""
DOCX fill engine.

Fills a .docx IN PLACE by editing the XML tree. Never rebuilds the document, so
headers, footers, page numbering, styles, numbering.xml, images, and section
properties all survive untouched.

Two failure modes this handles that naive implementations do not:

1. TOKENS SPLIT ACROSS RUNS.
   Word frequently breaks a token across multiple <w:r> elements:
       <w:t>[XM8_INSURED_H_</w:t> ... <w:t>STREET]</w:t>
   A plain string replace MISSES these and the raw token ships in the report.
   We stitch all <w:t> text within a paragraph before matching.

2. HEADERS AND FOOTERS.
   These live in separate XML parts (word/header1.xml etc). Any HTML-based
   renderer destroys them. We fill them too.

We edit via lxml (a real XML parser), NOT regex on raw XML. Regexing OOXML
reliably corrupts the file -- verified the hard way.
"""

from __future__ import annotations

import io
import re
import zipfile
from typing import Iterable

from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Which parts of the zip get filled. document + all headers + all footers.
FILLABLE_PART = re.compile(r"word/(document|header\d*|footer\d*)\.xml$")

# Default token style: [XM8_FILE_NO]. Configurable per request.
DEFAULT_PATTERN = r"\[([A-Z][A-Z0-9_]*)\]"


class FillResult:
    def __init__(self, data: bytes, filled: list[str], missing: list[str],
                 paragraphs_changed: int):
        self.data = data
        self.filled = filled
        self.missing = missing
        self.paragraphs_changed = paragraphs_changed


def _fill_part(xml_bytes: bytes, values: dict[str, str], token_re: re.Pattern,
               leave_missing: bool) -> tuple[bytes, list[str], list[str], int]:
    root = etree.fromstring(xml_bytes)
    filled: list[str] = []
    missing: list[str] = []
    changed = 0

    for p in root.iter(f"{W}p"):
        ts = list(p.iter(f"{W}t"))
        if not ts:
            continue

        joined = "".join(t.text or "" for t in ts)
        if not token_re.search(joined):
            continue

        def sub(m: re.Match) -> str:
            name = m.group(1)
            if name in values and values[name] is not None:
                filled.append(name)
                return str(values[name])
            missing.append(name)
            # leave_missing=True keeps the raw token visible so a human spots it.
            # leave_missing=False blanks it (may leave dangling punctuation).
            return m.group(0) if leave_missing else ""

        new = token_re.sub(sub, joined)
        if new == joined:
            continue

        # Collapse all text into the first run (keeps ITS formatting), blank the rest.
        ts[0].text = new
        ts[0].set(XML_SPACE, "preserve")
        for t in ts[1:]:
            t.text = ""
        changed += 1

    out = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    return out, filled, missing, changed


def fill_docx(template: bytes, values: dict[str, str],
              pattern: str = DEFAULT_PATTERN,
              leave_missing: bool = True) -> FillResult:
    """Fill a .docx template. Returns the filled bytes plus what was/wasn't filled."""
    token_re = re.compile(pattern)
    all_filled: list[str] = []
    all_missing: list[str] = []
    total_changed = 0

    src = zipfile.ZipFile(io.BytesIO(template))
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as out:
        for item in src.infolist():
            data = src.read(item.filename)
            if FILLABLE_PART.match(item.filename):
                data, f, m, n = _fill_part(data, values, token_re, leave_missing)
                all_filled.extend(f)
                all_missing.extend(m)
                total_changed += n
            out.writestr(item, data)
    src.close()

    return FillResult(
        data=buf.getvalue(),
        filled=sorted(set(all_filled)),
        missing=sorted(set(all_missing)),
        paragraphs_changed=total_changed,
    )


def scan_docx(template: bytes, pattern: str = DEFAULT_PATTERN) -> dict:
    """
    List every token in a template, and WHERE it lives (body vs header vs footer),
    with the surrounding sentence.

    The surrounding text matters: it tells the AI writer whether the template already
    supplies a '$' before a money token (so the value should be '2,614.70', not
    '$2,614.70'). This is how you avoid '$$2,614.70' in a real report.
    """
    token_re = re.compile(pattern)
    tokens: dict[str, dict] = {}

    src = zipfile.ZipFile(io.BytesIO(template))
    for item in src.infolist():
        if not FILLABLE_PART.match(item.filename):
            continue
        part = ("header" if "header" in item.filename
                else "footer" if "footer" in item.filename
                else "body")
        root = etree.fromstring(src.read(item.filename))
        for p in root.iter(f"{W}p"):
            ts = list(p.iter(f"{W}t"))
            if not ts:
                continue
            joined = "".join(t.text or "" for t in ts)
            for m in token_re.finditer(joined):
                name = m.group(1)
                rec = tokens.setdefault(name, {"name": name, "count": 0,
                                               "parts": set(), "context": []})
                rec["count"] += 1
                rec["parts"].add(part)
                if len(rec["context"]) < 3:
                    rec["context"].append(joined.strip()[:220])
    src.close()

    return {
        "tokens": [
            {"name": t["name"], "count": t["count"],
             "parts": sorted(t["parts"]), "context": t["context"]}
            for t in sorted(tokens.values(), key=lambda x: x["name"])
        ],
        "total": sum(t["count"] for t in tokens.values()),
        "unique": len(tokens),
    }
