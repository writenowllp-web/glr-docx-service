# GLR DOCX Service (Option B)

A small service you own. Does the one genuinely hard thing: fills a Word template
**without destroying it**.

Verified against your real `First Report - SAG Storm.docx`:
headers filled, page numbering intact, 21 bold runs preserved, split tokens resolved,
`$2,614.70` rendered with ONE dollar sign, 26 unfilled tokens reported back.

## Why this exists

| Approach | What happens |
|---|---|
| HTML → docx | **destroys headers, footers, page numbering** |
| regex on raw XML | **corrupts the .docx** (verified — it throws XMLSyntaxError) |
| Word Find/Replace | **misses tokens split across runs** (`[XM8_INSURED_H_` + `STREET]`) |
| **this service** | edits the XML tree, stitches runs first, fills header/footer parts too |

## Endpoints

### `POST /scan` → the token schema
```
file: template.docx
```
Returns every token, where it lives (body/header/footer), and **the sentence around it**.

Feed this to the AI writer. The surrounding text is what stops `$$2,614.70`:
```json
{"name":"XM8_COV_RCV_1","parts":["body"],
 "context":["...repairs in the RCV of $[XM8_COV_RCV_1]. Depreciation was..."]}
```
The model sees the template already prints `$`, so it returns `2,614.70`.

### `POST /fill` → the filled .docx
```
file:          template.docx
values:        {"XM8_FILE_NO":"THOMPSON-DAVE", ...}   (JSON string)
leave_missing: true    # unfilled tokens stay VISIBLE so a human catches them
as_base64:     true    # easier to handle in n8n
```
Returns the docx plus `filled[]` and **`missing[]`**.

**`missing` is your QA gate.** Non-empty = the report is incomplete. Fail it.

### `POST /convert` → `[TOKEN]` to `{TOKEN}`
For Option A (docxtemplater). Do NOT use Word's Find/Replace for this — it misses
split tokens and you ship a raw `[XM8_...]` in a carrier report.

## Deploy

**Railway / Render / Fly** — point at this repo, it has a Dockerfile. ~$5/mo.

Local:
```
pip install -r requirements.txt
uvicorn app.main:app --port 8000
```

## Calling it from n8n

HTTP Request node (you already use these everywhere):

- **Method:** POST → `https://your-service/fill`
- **Body:** Form-Data
  - `file` → binary property holding the template
  - `values` → `={{ JSON.stringify($json.context) }}`
  - `as_base64` → `true`
- Then a Code node: `Buffer.from(item.json.docx_base64,'base64')` → binary.

## The design point

Your SAG template has 38 `XM8_*` tokens. **Every one is an Xactimate data field** —
insured name, claim number, RCV, dates. These are *lookups*, not writing.

The AI writes only the narrative prose. The merge fields are a dict.
Keep the two straight and the hallucination surface shrinks dramatically.
