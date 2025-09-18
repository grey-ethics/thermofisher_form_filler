# services/extract_input.py
"""
Extract Data From Input
-----------------------
- keep_first_page_and_text(): copy page 1 to a new doc, return (saved_path, text)
- call_llm(): ask OpenAI to extract regions-in-scope + Medical Yes/No (strict JSON)
- extract_and_map(): orchestration for the /extract route

Assumptions:
- Input is a .docx with exactly 5 pages (we program defensively and just copy page 1).
- We only send page-1 text to the LLM (privacy + determinism).
- We prefill the GP row (r=16) and MIRROR the same values to the MD row (r=17).
"""

from __future__ import annotations
import os, json, threading, tempfile, re
import pythoncom
import win32com.client as com
import requests
from services.storage import relpath_from_output

# COM constants
_wdGoToPage = 1          # wdGoToPage
_wdGoToAbsolute = 1      # wdGoToAbsolute

# serialize Word access (shared with other COM code)
_WORD_LOCK = threading.Lock()


def _open_word():
    pythoncom.CoInitialize()
    app = com.DispatchEx("Word.Application")
    app.Visible = False
    app.DisplayAlerts = 0
    return app

def _quit_word(app):
    try:
        app.Quit()
    finally:
        pythoncom.CoUninitialize()

def keep_first_page_and_text(src_docx_path: str, out_dir: str) -> tuple[str, str]:
    """
    Copy Page 1 from src_docx_path into a new document and return:
    (saved_page1_docx_path, page1_text)
    """
    os.makedirs(out_dir, exist_ok=True)
    page1_path = os.path.join(out_dir, "input_first_page.docx")

    with _WORD_LOCK:
        app = _open_word()
        try:
            doc = app.Documents.Open(os.path.abspath(src_docx_path))
            doc.Activate()
            sel = app.Selection

            # Go to Page 1 start
            sel.GoTo(What=_wdGoToPage, Which=_wdGoToAbsolute, Count=1)
            start = sel.Start
            # Go to Page 2 start (end of our range)
            sel.GoTo(What=_wdGoToPage, Which=_wdGoToAbsolute, Count=2)
            end = sel.Start

            rng = doc.Range(Start=start, End=end)
            rng.Copy()

            newdoc = app.Documents.Add()
            newdoc.Range(0, 0).Paste()

            # Save the 1-page docx
            newdoc.SaveAs2(page1_path)
            page1_text = newdoc.Content.Text or ""

            newdoc.Close(SaveChanges=False)
            doc.Close(SaveChanges=False)
        finally:
            _quit_word(app)

    return page1_path, page1_text


def _clean_text(t: str) -> str:
    t = t.replace("\r", "\n")
    t = t.replace("\x07", "")
    t = re.sub(r"\n{3,}", "\n\n", t).strip()
    return t


def call_llm(page1_text: str) -> dict:
    """
    Ask OpenAI to extract:
      - regions in scope among: "N. America", "EMEA", "LATAM", "APAC"
      - "Medical" yes/no from the row "1.0 Medical: Yes/No"
    Returns:
    {
      "regions": { "N. America": true/false, "EMEA": ..., "LATAM": ..., "APAC": ... },
      "medical": true/false
    }
    """
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        raise RuntimeError("OPENAI_API_KEY not set in environment.")

    model = os.environ.get("OPENAI_MODEL", "gpt-4o-mini")

    system = (
        "You are a precise information extraction engine. "
        "Only return strict JSON. No explanation. No markdown."
    )
    user = f"""
From the page text below, extract:
1) Which regions are in scope among exactly these labels:
   - "N. America" (or North America / N America)
   - "EMEA"
   - "LATAM"
   - "APAC"
2) Read the 'Regulatory status of the product' table first row:
   It appears as "1.0 Medical: Yes" OR "1.0 Medical: No". Extract that as a boolean.

Return STRICT JSON with keys exactly:
{{
  "regions": {{
    "N. America": true|false,
    "EMEA": true|false,
    "LATAM": true|false,
    "APAC": true|false
  }},
  "medical": true|false
}}

PAGE_TEXT:
\"\"\"{_clean_text(page1_text)[:50000]}\"\"\""""
    payload = {
        "model": model,
        "temperature": 0,
        "messages": [
            {"role": "system", "content": system},
            {"role": "user", "content": user},
        ],
    }

    resp = requests.post(
        "https://api.openai.com/v1/chat/completions",
        headers={"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"},
        json=payload,
        timeout=60,
    )
    resp.raise_for_status()
    content = resp.json()["choices"][0]["message"]["content"]

    def _safe_json(s: str) -> dict:
        s = s.strip()
        if s.startswith("```"):
            s = re.sub(r"^```[a-zA-Z0-9]*\s*", "", s)
            s = re.sub(r"\s*```$", "", s)
        i, j = s.find("{"), s.rfind("}")
        if i >= 0 and j > i:
            s = s[i : j + 1]
        return json.loads(s)

    data = _safe_json(content)

    regions = data.get("regions") or {}
    for k in ["N. America", "EMEA", "LATAM", "APAC"]:
        regions[k] = bool(regions.get(k, False))
    medical = bool(data.get("medical", False))

    return {"regions": regions, "medical": medical}


def build_gp_ticks(regions: dict[str, bool], medical_yes: bool) -> dict[str, bool]:
    """
    Build ticks for the GP row only (row=16). Columns:
      c2=N. America, c3=EMEA, c4=LATAM, c5=APAC

    Current rule: region_in_scope AND (Medical == Yes).
    """
    colmap = [("N. America", 2), ("EMEA", 3), ("LATAM", 4), ("APAC", 5)]
    ticks = {}
    for label, col in colmap:
        val = bool(regions.get(label, False)) and bool(medical_yes)
        ticks[f"glyph_r16_c{col}"] = val
    return ticks


def _mirror_row_ticks(src_ticks: dict[str, bool], src_row: int = 16, dst_row: int = 17) -> dict[str, bool]:
    """
    Copy ticks from src_row to dst_row (keeps column numbers).
    Example: glyph_r16_c2 -> glyph_r17_c2
    """
    out = {}
    for k, v in src_ticks.items():
        m = re.match(rf"^glyph_r{src_row}_c(\d+)$", k)
        if m:
            out[f"glyph_r{dst_row}_c{m.group(1)}"] = bool(v)
    return out


def extract_and_map(file_storage, out_dir: str) -> dict:
    """
    Orchestration for /extract:
    - Save upload -> keep first page -> call LLM -> compute GP ticks -> MIRROR to MD -> build lines
    """
    os.makedirs(out_dir, exist_ok=True)
    # Save upload
    src_path = os.path.join(out_dir, "uploaded_standard.docx")
    file_storage.save(src_path)

    # Keep page 1 & read its text
    page1_path, page1_text = keep_first_page_and_text(src_path, out_dir)

    # LLM extraction
    info = call_llm(page1_text)
    regions = info["regions"]
    medical = info["medical"]

    gp_ticks = build_gp_ticks(regions, medical)
    md_ticks = _mirror_row_ticks(gp_ticks, src_row=16, dst_row=17)

    ticks = {**gp_ticks, **md_ticks}

    # Lines for UI (Yes/No shown is the same rule as GP/MD ticks)
    lines = [
        f"N. America = {'Yes' if regions['N. America'] and medical else 'No'}",
        f"EMEA = {'Yes' if regions['EMEA'] and medical else 'No'}",
        f"LATAM = {'Yes' if regions['LATAM'] and medical else 'No'}",
        f"APAC = {'Yes' if regions['APAC'] and medical else 'No'}",
    ]

    return {
        "medical": "Yes" if medical else "No",
        "regions": {
            "N. America": "Yes" if regions["N. America"] else "No",
            "EMEA": "Yes" if regions["EMEA"] else "No",
            "LATAM": "Yes" if regions["LATAM"] else "No",
            "APAC": "Yes" if regions["APAC"] else "No",
        },
        "ticks": ticks,  # now includes r16 AND r17
        "lines": lines,
        "first_page_docx_rel": relpath_from_output(page1_path),
    }
