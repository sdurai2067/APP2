# =============================================================================
# app.py — FSD Pipeline (Lossless JSON1 → Quality Tests)
# =============================================================================
#
# PIPELINE:
#
#  PDF/DOCX/TXT Upload
#       ↓
#  CALL 1 → /api/extract     FSD text → JSON1 (lossless, 1100+ words)
#       ↓ JSON1 stored in browser
#  CALL 2 → /api/summary     JSON1 → FSD Summary
#  CALL 3 → /api/tests       JSON1 → Test Cases (12-18, defined format)
#       ↓ (optional, user triggered)
#  CALL 4 → /api/css         JSON1 → CSS Release Gate
#  CALL 5 → /api/gap         JSON1 → Gap Analysis + Score
#  CALL 6 → /api/refine      JSON1 + Comments → JSON2 (enhanced)
#       ↓
#  EXCEL → /api/export       ZERO calls (pure Python)
#
# SYNC RULE: Each call only starts after previous call completes successfully.
#            Frontend shows progress per step. No parallel calls.
#            JSON1 is single source of truth — passed to every downstream call.
# =============================================================================

import os
import io
import re
import json
import traceback
from pathlib import Path
from datetime import datetime

import requests
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

from flask import Flask, request, jsonify, send_file, render_template

import docx
from pypdf import PdfReader
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import prompts


# =============================================================================
# ENV + CONFIG
# =============================================================================
def load_env(path=".env"):
    env_path = Path(path)
    if not env_path.exists():
        print(f"[WARN] No .env file found. Create one with GEMINI_API_KEY=AIzaSy...")
        return
    with open(env_path, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, _, val = line.partition("=")
                os.environ.setdefault(key.strip(), val.strip())

load_env()

GEMINI_API_KEY = os.environ.get("GEMINI_API_KEY", "")
GEMINI_MODEL   = os.environ.get("GEMINI_MODEL", "gemini-2.0-flash")
PORT           = int(os.environ.get("PORT", 5000))
GEMINI_URL     = (
    f"https://generativelanguage.googleapis.com/v1beta/models/"
    f"{GEMINI_MODEL}:generateContent"
)
OUTPUTS_DIR = Path("results")
OUTPUTS_DIR.mkdir(exist_ok=True)

app = Flask(__name__, template_folder=".")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


# =============================================================================
# GEMINI CALLER — single function, all calls go through here
# =============================================================================
def call_gemini(prompt: str, label: str = "") -> str:
    """
    Send prompt to Gemini. Returns text response.
    label = human-readable call name for logging (e.g. "CALL1-EXTRACT")
    verify=False = bypass Zscaler corporate SSL proxy
    """
    if not GEMINI_API_KEY:
        raise RuntimeError(
            "GEMINI_API_KEY not set. "
            "Add to .env: GEMINI_API_KEY=AIzaSy..."
        )

    if label:
        print(f"[GEMINI] Starting {label}...")

    payload = {
        "contents": [{"parts": [{"text": prompt}]}],
        "generationConfig": {
            "temperature":     0.2,   # very low = consistent extraction
            "maxOutputTokens": 8192,
            "topP":            0.9,
        }
    }

    url  = f"{GEMINI_URL}?key={GEMINI_API_KEY}"
    resp = requests.post(
        url,
        headers={"Content-Type": "application/json"},
        json=payload,
        timeout=180,
        verify=False
    )

    if label:
        print(f"[GEMINI] {label} → status {resp.status_code}")

    if resp.status_code == 429:
        raise RuntimeError(
            "QUOTA EXCEEDED (429). "
            "Wait 60 seconds OR create new key at "
            "https://aistudio.google.com/apikey"
        )
    if resp.status_code == 400:
        raise RuntimeError(
            f"Bad request. Check GEMINI_MODEL in .env. "
            f"Current: {GEMINI_MODEL}. Use: gemini-2.0-flash"
        )
    if resp.status_code == 403:
        raise RuntimeError(
            "API key invalid. Get new key at "
            "https://aistudio.google.com/apikey"
        )
    if resp.status_code != 200:
        raise RuntimeError(
            f"Gemini error {resp.status_code}: {resp.text[:300]}"
        )

    data = resp.json()
    try:
        text = data["candidates"][0]["content"]["parts"][0]["text"]
        if label:
            print(f"[GEMINI] {label} → {len(text)} chars returned")
        return text
    except (KeyError, IndexError) as e:
        raise RuntimeError(f"Unexpected Gemini response: {data}") from e


# =============================================================================
# TEXT EXTRACTOR — handles PDF, DOCX, TXT
# =============================================================================
def extract_text(file_storage) -> tuple:
    """
    Extract text from uploaded file.
    Returns (text, page_count) tuple.
    Skips first 2 pages for PDF (cover + author pages).
    For DOCX: skips first ~500 chars (cover area).
    """
    filename = file_storage.filename.lower()
    raw      = file_storage.read()

    if filename.endswith(".pdf"):
        reader     = PdfReader(io.BytesIO(raw))
        total_pages = len(reader.pages)
        # Skip pages 1-2 (index 0-1), start from page 3 (index 2)
        start_page = min(2, total_pages - 1)
        pages = []
        for i, page in enumerate(reader.pages):
            if i < start_page:
                continue   # skip cover/author pages
            text = page.extract_text()
            if text and text.strip():
                pages.append(f"[PAGE {i+1}]\n{text}")
        return "\n\n".join(pages), total_pages

    elif filename.endswith(".docx"):
        doc   = docx.Document(io.BytesIO(raw))
        paras = [p.text for p in doc.paragraphs if p.text.strip()]
        # Skip first ~3 paragraphs (title, author, date)
        content = paras[3:] if len(paras) > 5 else paras
        return "\n".join(content), len(paras)

    elif filename.endswith(".txt"):
        text = raw.decode("utf-8", errors="ignore")
        # Skip first 200 chars for txt files
        return text[200:] if len(text) > 400 else text, 1

    else:
        text = raw.decode("utf-8", errors="ignore")
        return text, 1


# =============================================================================
# JSON PARSER — handles AI markdown fences
# =============================================================================
def parse_json(text: str):
    """
    Extract valid JSON from AI response.
    AI sometimes wraps JSON in ```json ... ``` — this strips it.
    """
    text  = re.sub(r"```(?:json)?", "", text).strip().rstrip("`").strip()
    start = next((i for i, c in enumerate(text) if c in "[{"), -1)
    if start == -1:
        return None

    depth, in_str, esc = 0, False, False
    opener = text[start]
    closer = "]" if opener == "[" else "}"

    for i in range(start, len(text)):
        ch = text[i]
        if esc:        esc = False;     continue
        if ch == "\\":  esc = True;     continue
        if ch == '"':   in_str = not in_str
        if not in_str:
            if ch == opener: depth += 1
            elif ch == closer:
                depth -= 1
                if depth == 0:
                    try:    return json.loads(text[start:i+1])
                    except: break

    try:    return json.loads(text)
    except: return None


# =============================================================================
# EXCEL BUILDER — 4 sheets, ZERO API calls
# =============================================================================
def build_excel(test_cases: list, summary: dict,
                json1: dict, css: list, filename: str) -> bytes:
    """
    Build 4-sheet Excel workbook from pipeline outputs.
    Sheet 1: Test Cases
    Sheet 2: FSD Summary
    Sheet 3: JSON1 Structure (FS sections overview)
    Sheet 4: Coverage Matrix
    Pure Python — no AI call.
    """
    wb   = Workbook()
    NAVY = "0F172A"; BLUE = "2563EB"; GRAY = "374151"
    WHITE = "FFFFFF"; TEAL = "0F766E"; AMBER = "92400E"

    SCN_BG = {"POSITIVE":"D1FAE5","NEGATIVE":"FEE2E2","BOUNDARY":"FEF3C7",
               "EXCEPTION":"FCE7F3","REGRESSION":"EDE9FE"}
    SCN_FG = {"POSITIVE":"065F46","NEGATIVE":"991B1B","BOUNDARY":"92400E",
               "EXCEPTION":"9D174D","REGRESSION":"4C1D95"}

    thin  = Border(
        left=Side(style="thin",color="CBD5E1"),
        right=Side(style="thin",color="CBD5E1"),
        top=Side(style="thin",color="CBD5E1"),
        bottom=Side(style="thin",color="CBD5E1"))
    thick = Border(
        left=Side(style="medium",color=NAVY),
        right=Side(style="medium",color=NAVY),
        top=Side(style="medium",color=NAVY),
        bottom=Side(style="medium",color=NAVY))

    def hc(ws, r, c, v, bg=NAVY, fg=WHITE, sz=10):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name="Arial", size=sz, bold=True, color=fg)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thick
        return cell

    def dc(ws, r, c, v, bg=WHITE, sz=9, ha="left", bold=False):
        cell = ws.cell(row=r, column=c, value=v)
        cell.font = Font(name="Arial", size=sz, color="1F2937", bold=bold)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.alignment = Alignment(horizontal=ha, vertical="top", wrap_text=True)
        cell.border = thin
        return cell

    # ── SHEET 1: Test Cases ───────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Test Cases"
    ws1.sheet_view.showGridLines = False

    ws1.merge_cells("A1:I1"); ws1.row_dimensions[1].height = 36
    t = ws1["A1"]
    t.value = f"TEST CASE DOCUMENT  |  {filename}"
    t.font = Font(name="Arial", size=14, bold=True, color=WHITE)
    t.fill = PatternFill("solid", fgColor=NAVY)
    t.alignment = Alignment(horizontal="center", vertical="center")

    ws1.merge_cells("A2:I2"); ws1.row_dimensions[2].height = 18
    s2 = ws1["A2"]
    s2.value = (f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                f"   |   Total: {len(test_cases)} test cases"
                f"   |   Model: {GEMINI_MODEL}"
                f"   |   Pipeline: JSON1 → Tests")
    s2.font = Font(name="Arial", size=9, color=WHITE)
    s2.fill = PatternFill("solid", fgColor=BLUE)
    s2.alignment = Alignment(horizontal="center", vertical="center")

    hdrs   = ["S.No","TC No","FS Ref","Test Description","Scenario",
              "Steps","Expected Result","Pass/Fail","Priority"]
    widths = [6, 10, 8, 30, 13, 52, 48, 11, 10]
    ws1.row_dimensions[3].height = 26
    for ci, (h, w) in enumerate(zip(hdrs, widths), 1):
        hc(ws1, 3, ci, h, bg=GRAY)
        ws1.column_dimensions[get_column_letter(ci)].width = w

    for ri, tc in enumerate(test_cases):
        r   = ri + 4
        ws1.row_dimensions[r].height = 78
        alt = "F8FAFF" if ri % 2 else WHITE
        scn = str(tc.get("scenario", "POSITIVE")).upper()
        vals = [
            tc.get("sno", ri+1),
            tc.get("tc_no", f"TC-{ri+1:03d}"),
            tc.get("fs_ref", ""),
            tc.get("test_description", ""),
            scn,
            tc.get("steps", ""),
            tc.get("expected_result", ""),
            tc.get("pass_fail", ""),
            tc.get("priority", "HIGH"),
        ]
        for ci, v in enumerate(vals, 1):
            if ci == 5:  # scenario
                cell = ws1.cell(row=r, column=ci, value=v)
                cell.font = Font(name="Arial", size=9, bold=True,
                                 color=SCN_FG.get(scn,"1F2937"))
                cell.fill = PatternFill("solid", fgColor=SCN_BG.get(scn, alt))
                cell.alignment = Alignment(horizontal="center",
                                           vertical="center", wrap_text=True)
                cell.border = thin
            else:
                dc(ws1, r, ci, v, bg=alt,
                   ha="center" if ci in [1, 8, 9] else "left")

    ws1.freeze_panes = "A4"

    # ── SHEET 2: FSD Summary ──────────────────────────────────────
    ws2 = wb.create_sheet("FSD Summary")
    ws2.sheet_view.showGridLines = False
    ws2.column_dimensions["A"].width = 28
    ws2.column_dimensions["B"].width = 80

    ws2.merge_cells("A1:B1"); ws2.row_dimensions[1].height = 34
    b = ws2["A1"]
    b.value = "FSD SUMMARY — PIPELINE OUTPUT"
    b.font = Font(name="Arial", size=13, bold=True, color=WHITE)
    b.fill = PatternFill("solid", fgColor=NAVY)
    b.alignment = Alignment(horizontal="center", vertical="center")

    def write_sec(ws, row, title, items, bg="DBEAFE"):
        items = [str(i) for i in (items or [])]
        ws.merge_cells(f"A{row}:B{row}")
        ws.row_dimensions[row].height = 22
        sc = ws.cell(row=row, column=1, value=f"  {title}")
        sc.font = Font(name="Arial", size=10, bold=True, color=NAVY)
        sc.fill = PatternFill("solid", fgColor=bg)
        sc.border = Border(bottom=Side(style="medium", color=NAVY))
        for i, item in enumerate(items):
            r2 = row + 1 + i
            ws.row_dimensions[r2].height = 17
            bg2 = WHITE if i % 2 == 0 else "F9FAFB"
            lc = ws.cell(row=r2, column=1, value="  •")
            lc.font = Font(name="Arial", size=9)
            lc.fill = PatternFill("solid", fgColor=bg2)
            lc.border = thin
            vc = ws.cell(row=r2, column=2, value=f"  {item}")
            vc.font = Font(name="Arial", size=9)
            vc.alignment = Alignment(wrap_text=True, vertical="top")
            vc.fill = PatternFill("solid", fgColor=bg2)
            vc.border = thin
        return row + 1 + len(items) + 1

    s = summary or {}
    row = 2
    row = write_sec(ws2, row, "CENTER POINT", [s.get("center_point","")], bg="DBEAFE")
    row = write_sec(ws2, row, "TRANSACTION SCOPE", s.get("transaction_types",[]), bg="EDE9FE")
    row = write_sec(ws2, row, "BUSINESS FLOW", s.get("business_flow",[]), bg="D1FAE5")
    row = write_sec(ws2, row, "KEY BUSINESS RULES", s.get("key_business_rules",[]), bg="FEF3C7")
    row = write_sec(ws2, row, "CRITICAL INPUT FIELDS", s.get("critical_inputs",[]), bg="CCFBF1")
    row = write_sec(ws2, row, "CRITICAL OUTPUT FIELDS", s.get("critical_outputs",[]), bg="FEF9C3")
    row = write_sec(ws2, row, "INTEGRATION MAP", s.get("integration_map",[]), bg="FCE7F3")
    row = write_sec(ws2, row, "INGENIUM TEST DATA", s.get("test_data_needed",[]), bg="FFF7ED")
    row = write_sec(ws2, row, "TESTING IMPLICATIONS", s.get("testing_implications",[]), bg="F0FDF4")
    if css:
        css_items = [f"{c.get('css_id','')} [{c.get('priority','')}] {c.get('scenario','')} → {c.get('go_no_go','')}"
                     for c in css]
        row = write_sec(ws2, row, "CSS RELEASE GATE", css_items, bg="FEE2E2")
    row = write_sec(ws2, row, "RISKS & ASSUMPTIONS",
                    s.get("risks",[]) + s.get("assumptions",[]), bg="FEE2E2")
    ws2.freeze_panes = "A2"

    # ── SHEET 3: JSON1 FS Sections ─────────────────────────────────
    ws3 = wb.create_sheet("JSON1 FS Sections")
    ws3.sheet_view.showGridLines = False

    ws3.merge_cells("A1:E1"); ws3.row_dimensions[1].height = 30
    j1h = ws3["A1"]
    j1h.value = "JSON1 — LOSSLESS FSD EXTRACTION (FS SECTIONS)"
    j1h.font = Font(name="Arial", size=12, bold=True, color=WHITE)
    j1h.fill = PatternFill("solid", fgColor=TEAL)
    j1h.alignment = Alignment(horizontal="center", vertical="center")

    j1_hdrs = ["FS Ref","Section Title","Business Rules","Calculations","Error Conditions"]
    j1_widths = [8, 30, 60, 50, 50]
    for ci, (h, w) in enumerate(zip(j1_hdrs, j1_widths), 1):
        hc(ws3, 2, ci, h, bg=GRAY)
        ws3.column_dimensions[get_column_letter(ci)].width = w

    sections = (json1 or {}).get("sections", {})
    for ri, (fs_key, fs_val) in enumerate(sections.items(), 3):
        ws3.row_dimensions[ri].height = 80
        alt = "F8FAFF" if ri % 2 else WHITE
        brs   = "\n".join(fs_val.get("business_rules", [])[:5])
        calcs = "\n".join(fs_val.get("calculations", [])[:3])
        errs  = "\n".join([
            f"{e.get('error_code','')}: {e.get('message','')}"
            if isinstance(e, dict) else str(e)
            for e in fs_val.get("error_conditions", [])[:3]
        ])
        row_vals = [
            fs_key,
            fs_val.get("section_title",""),
            brs,
            calcs,
            errs,
        ]
        for ci, v in enumerate(row_vals, 1):
            dc(ws3, ri, ci, v, bg=alt,
               ha="center" if ci == 1 else "left",
               bold=(ci == 1))
    ws3.freeze_panes = "A3"

    # ── SHEET 4: Coverage Matrix ───────────────────────────────────
    ws4 = wb.create_sheet("Coverage Matrix")
    ws4.sheet_view.showGridLines = False

    ws4.merge_cells("A1:F1"); ws4.row_dimensions[1].height = 28
    cm = ws4["A1"]
    cm.value = "TEST COVERAGE MATRIX"
    cm.font = Font(name="Arial", size=12, bold=True, color=WHITE)
    cm.fill = PatternFill("solid", fgColor=NAVY)
    cm.alignment = Alignment(horizontal="center", vertical="center")

    ch = ["Scenario","Count","% Coverage","TC IDs","Risk","Status"]
    cw = [16, 8, 12, 52, 10, 10]
    for ci, (c_n, c_w) in enumerate(zip(ch, cw), 1):
        hc(ws4, 2, ci, c_n, bg=GRAY)
        ws4.column_dimensions[get_column_letter(ci)].width = c_w

    counts, tc_ids = {}, {}
    for tc in test_cases:
        scn = str(tc.get("scenario","POSITIVE")).upper()
        counts[scn] = counts.get(scn, 0) + 1
        tc_ids.setdefault(scn, []).append(str(tc.get("tc_no","")))

    total = len(test_cases) or 1
    risk  = {"POSITIVE":"LOW","NEGATIVE":"HIGH","BOUNDARY":"MEDIUM",
             "EXCEPTION":"HIGH","REGRESSION":"MEDIUM"}
    for ri, (scn, cnt) in enumerate(counts.items(), 3):
        ws4.row_dimensions[ri].height = 17
        rd = [scn, cnt, f"{round(cnt/total*100,1)}%",
              ", ".join(tc_ids.get(scn,[])), risk.get(scn,"MED"), "PENDING"]
        for ci, v in enumerate(rd, 1):
            c2 = ws4.cell(row=ri, column=ci, value=v)
            c2.font = Font(name="Arial", size=9, bold=(ci==1),
                           color=SCN_FG.get(scn,"1F2937") if ci==1 else "1F2937")
            c2.fill = PatternFill("solid", fgColor=SCN_BG.get(scn, WHITE))
            c2.alignment = Alignment(
                horizontal="center" if ci!=4 else "left",
                vertical="center", wrap_text=True)
            c2.border = thin

    tr = len(counts) + 3
    for ci, v in enumerate(["TOTAL", total, "100%","","",""],1):
        c3 = ws4.cell(row=tr, column=ci, value=v)
        c3.font = Font(name="Arial", size=10, bold=True, color=WHITE)
        c3.fill = PatternFill("solid", fgColor=BLUE)
        c3.alignment = Alignment(horizontal="center", vertical="center")
        c3.border = thick
    ws4.freeze_panes = "A3"

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.read()


# =============================================================================
# ROUTES
# =============================================================================

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/health")
def health():
    """Ping Gemini with minimal token use to verify connectivity."""
    if not GEMINI_API_KEY:
        return jsonify({"ok": False, "reason": "GEMINI_API_KEY not set"})
    try:
        call_gemini("OK", label="HEALTH-PING")
        return jsonify({"ok": True, "model": GEMINI_MODEL})
    except Exception as e:
        return jsonify({"ok": False, "reason": str(e)})


@app.route("/api/extract", methods=["POST"])
def extract():
    """
    CALL 1 — FSD → JSON1
    Upload FSD → extract text (skip cover pages) → Gemini → JSON1
    JSON1 is the lossless structured representation of the FSD.
    Returns JSON1 to frontend for storage and downstream use.
    """
    if "fsd" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    f = request.files["fsd"]
    text, page_count = extract_text(f)
    fname = f.filename

    if not text.strip():
        return jsonify({
            "error": "No text extracted. Try DOCX or TXT format."
        }), 400

    print(f"[EXTRACT] {fname} → {len(text)} chars from {page_count} pages")

    try:
        raw  = call_gemini(
            prompts.build(prompts.FSD_TO_JSON1, fsd_content=text[:10000]),
            label="CALL1-FSD-TO-JSON1"
        )
        json1 = parse_json(raw)

        if not json1:
            return jsonify({
                "error": "JSON1 extraction failed. FSD may be too short or unstructured.",
                "raw_response": raw[:500]
            }), 500

        # Save JSON1 to results folder for reference
        ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
        j1path = OUTPUTS_DIR / f"JSON1_{Path(fname).stem}_{ts}.json"
        j1path.write_text(json.dumps(json1, indent=2), encoding="utf-8")
        print(f"[EXTRACT] JSON1 saved → {j1path}")

        return jsonify({
            "ok":         True,
            "fsd_filename": fname,
            "fsd_preview":  text[:1000],
            "page_count":   page_count,
            "json1":        json1,
            "json1_path":   str(j1path),
            "sections_found": len(json1.get("sections", {})),
            "word_equivalent": json1.get("word_equivalent", "calculated"),
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/summary", methods=["POST"])
def summary():
    """
    CALL 2 — JSON1 → FSD Summary
    Triggered automatically after CALL 1 succeeds.
    Uses JSON1 — no FSD re-read.
    """
    data  = request.get_json() or {}
    json1 = data.get("json1", {})

    if not json1:
        return jsonify({"error": "JSON1 not provided. Run Extract first."}), 400

    try:
        raw     = call_gemini(
            prompts.build(prompts.JSON1_TO_SUMMARY,
                          json1=json.dumps(json1, indent=2)[:8000]),
            label="CALL2-JSON1-TO-SUMMARY"
        )
        summary = parse_json(raw) or {}
        return jsonify({"ok": True, "summary": summary})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/tests", methods=["POST"])
def tests():
    """
    CALL 3 — JSON1 → Test Cases
    Triggered by Generate Test Document button.
    Uses JSON1 — no FSD re-read.
    Returns test cases in defined format.
    """
    data  = request.get_json() or {}
    json1 = data.get("json1", {})

    if not json1:
        return jsonify({"error": "JSON1 not provided. Run Extract first."}), 400

    try:
        raw = call_gemini(
            prompts.build(prompts.JSON1_TO_TESTS,
                          json1=json.dumps(json1, indent=2)[:8000]),
            label="CALL3-JSON1-TO-TESTS"
        )
        tcs = parse_json(raw) or []

        # Add fs_ref if missing
        for i, tc in enumerate(tcs):
            if not tc.get("fs_ref"):
                tc["fs_ref"] = "FS1"
            tc["sno"] = i + 1

        return jsonify({
            "ok":         True,
            "test_cases": tcs,
            "tc_count":   len(tcs),
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/css", methods=["POST"])
def css():
    """
    CALL 4 — JSON1 → CSS Register (optional, user triggered)
    """
    data  = request.get_json() or {}
    json1 = data.get("json1", {})
    if not json1:
        return jsonify({"error": "JSON1 not provided."}), 400

    try:
        raw    = call_gemini(
            prompts.build(prompts.JSON1_TO_CSS,
                          json1=json.dumps(json1, indent=2)[:7000]),
            label="CALL4-JSON1-TO-CSS"
        )
        css_data = parse_json(raw) or []
        return jsonify({"ok": True, "css": css_data})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/gap", methods=["POST"])
def gap():
    """
    CALL 5 — JSON1 → Gap Analysis (optional, user triggered)
    """
    data  = request.get_json() or {}
    json1 = data.get("json1", {})
    if not json1:
        return jsonify({"error": "JSON1 not provided."}), 400

    try:
        raw  = call_gemini(
            prompts.build(prompts.JSON1_TO_GAP,
                          json1=json.dumps(json1, indent=2)[:7000]),
            label="CALL5-JSON1-TO-GAP"
        )
        gap  = parse_json(raw) or {"quality_score": 0, "gaps": []}
        return jsonify({"ok": True, **gap})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/refine", methods=["POST"])
def refine():
    """
    CALL 6 — JSON1 + Comments → JSON2 (optional, user triggered)
    User adds comments/corrections → AI produces enhanced JSON2.
    JSON2 replaces JSON1 for better quality downstream calls.
    """
    data          = request.get_json() or {}
    json1         = data.get("json1", {})
    user_comments = data.get("user_comments", "")

    if not json1:
        return jsonify({"error": "JSON1 not provided."}), 400
    if not user_comments.strip():
        return jsonify({"error": "No comments provided."}), 400

    try:
        raw   = call_gemini(
            prompts.build(
                prompts.JSON1_TO_JSON2,
                json1=json.dumps(json1, indent=2)[:6000],
                user_comments=user_comments,
            ),
            label="CALL6-JSON1-TO-JSON2"
        )
        json2 = parse_json(raw) or {}

        # Save JSON2
        ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
        j2path = OUTPUTS_DIR / f"JSON2_refined_{ts}.json"
        j2path.write_text(json.dumps(json2, indent=2), encoding="utf-8")

        return jsonify({
            "ok":      True,
            "json2":   json2,
            "changes": json2.get("json2_changes", []),
            "path":    str(j2path),
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/chat", methods=["POST"])
def chat():
    """Chat refine — uses JSON1 or JSON2 as context."""
    data      = request.get_json() or {}
    json_ctx  = data.get("json2") or data.get("json1", {})
    user_req  = data.get("user_request", "")

    if not user_req.strip():
        return jsonify({"error": "No message provided"}), 400

    try:
        raw    = call_gemini(
            prompts.build(
                prompts.CHAT_REFINE,
                json_context=json.dumps(json_ctx, indent=2)[:5000],
                user_request=user_req,
            ),
            label="CHAT-REFINE"
        )
        parsed = parse_json(raw)
        return jsonify({"ok": True, "response": raw, "parsed": parsed})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


@app.route("/api/export", methods=["POST"])
def export():
    """
    ZERO API CALLS — pure Python Excel generation.
    Builds 4-sheet workbook from pipeline outputs.
    """
    data       = request.get_json() or {}
    test_cases = data.get("test_cases", [])
    summary    = data.get("summary", {})
    json1      = data.get("json1", {})
    css_data   = data.get("css", [])
    fname      = data.get("fsd_filename", "FSD")

    if not test_cases:
        return jsonify({"error": "No test cases. Generate tests first."}), 400

    try:
        xlsx  = build_excel(test_cases, summary, json1, css_data, fname)
        ts    = datetime.now().strftime("%Y%m%d_%H%M%S")
        oname = f"TestDoc_{Path(fname).stem}_{ts}.xlsx"
        (OUTPUTS_DIR / oname).write_bytes(xlsx)

        return send_file(
            io.BytesIO(xlsx),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            as_attachment=True,
            download_name=oname,
        )

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500


# =============================================================================
# ENTRY POINT
# =============================================================================
if __name__ == "__main__":
    print("=" * 65)
    print("  FSD Pipeline — Lossless JSON1 → Quality Test Generation")
    print(f"  http://localhost:{PORT}")
    print(f"  Model  : {GEMINI_MODEL}")
    print(f"  API Key: {'SET ✓' if GEMINI_API_KEY else 'MISSING!'}")
    print(f"  Results: ./{OUTPUTS_DIR}/")
    print("  ─────────────────────────────────────────────────────────")
    print("  PIPELINE CALLS:")
    print("  CALL 1 /api/extract  → FSD → JSON1 (lossless)")
    print("  CALL 2 /api/summary  → JSON1 → Summary")
    print("  CALL 3 /api/tests    → JSON1 → Test Cases")
    print("  CALL 4 /api/css      → JSON1 → CSS (optional)")
    print("  CALL 5 /api/gap      → JSON1 → Gap Analysis (optional)")
    print("  CALL 6 /api/refine   → JSON1+Comments → JSON2 (optional)")
    print("  EXCEL  /api/export   → ZERO API calls")
    print("=" * 65)
    app.run(debug=False, host="0.0.0.0", port=PORT)
