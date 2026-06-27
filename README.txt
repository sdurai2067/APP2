================================================================
FSD2TESTDOC App2 — FSD Intelligence Pipeline
Ingenium · COBOL · Policy Admin · QA Test Generator
================================================================

PROJECT FOLDER LAYOUT
----------------------
app2/
├── app.py            Backend server (Flask + Gemini + Excel)
├── index.html        Frontend UI   (HTML + CSS + JavaScript)
├── prompts.py        All AI prompts (edit to tune AI behaviour)
├── requirements.txt  Python packages to install
├── README.txt        This file
└── results/          Auto-created — JSON1 + Excel files saved here


QUICK START (5 steps)
---------------------

STEP 1 — Get Gemini API Key (free)
  Go to: https://aistudio.google.com/apikey
  Sign in with Google account (no credit card needed)
  Click "Create API key" → "Create API key in new project"
  Copy the key — it starts with: AIzaSy...

STEP 2 — Create .env file
  In your terminal, inside the app2 folder, run:

  python -c "f=open('.env','w',encoding='utf-8');f.write('GEMINI_API_KEY=AIzaSyPasteYourKeyHere\nGEMINI_MODEL=gemini-2.0-flash\nPORT=5000\n');f.close();print('Done!')"

  Replace AIzaSyPasteYourKeyHere with your actual key.

STEP 3 — Install Python packages (one time only)
  python -m pip install -r requirements.txt

STEP 4 — Run the app
  python app.py

STEP 5 — Open browser
  http://localhost:5000
  Green dot in header = Gemini connected and working


HOW TO USE
----------

1. UPLOAD FSD
   Drag and drop your FSD onto the upload zone.
   Supports: PDF · DOCX · TXT (up to 50 MB)
   Cover pages and author pages are automatically skipped.

2. STEP 1 — Extract JSON1 (1 Gemini call)
   Click "Step 1 — Extract JSON1"
   AI reads FSD and creates JSON1 — a structured JSON with
   every section, business rule, calculation, field, error code.
   JSON1 is saved to results/JSON1_*.json automatically.
   Watch the Live Log tab to see progress in real time.

3. STEP 2 — Generate Summary + Tests (2 Gemini calls)
   Click "Step 2 — Generate Tests"
   Call 2A: JSON1 → Business summary (flow, rules, integrations)
   Call 2B: JSON1 → 12-15 Ingenium-aware test cases
   Test cases appear in the Test Cases tab.

4. EXPORT EXCEL (0 Gemini calls)
   Click "Export Excel (0 calls)"
   Downloads a 3-sheet formatted .xlsx workbook:
     Sheet 1: Test Cases (colour-coded by scenario type)
     Sheet 2: FSD Summary (metadata, flow, rules, risks)
     Sheet 3: Coverage Matrix (scenario/priority stats)
   Also saved to results/TestDoc_*.xlsx

5. STEP 3 — Gap Analysis (1 Gemini call, optional)
   Click "Step 3 — Gap Analysis"
   Reviews JSON1 for missing or ambiguous information.
   Produces a quality score (0-100) and gap register.

6. AI CHAT (1 call per message, optional)
   Go to the AI Chat tab.
   Ask questions in plain English:
     "Generate 3 more negative tests for FS1"
     "What are the boundary values for the premium field?"
     "List all error conditions from FS2"
     "What Ingenium screens are involved?"
   If AI returns test cases, click "Apply → Tests Tab".


PIPELINE ARCHITECTURE
---------------------

PDF/DOCX/TXT Upload
      |
  /api/step1  (1 Gemini call)
  FSD text → JSON1
  JSON1 = lossless structured extraction of every FS section
  Saved to: results/JSON1_*.json
      |
  JSON1 stored in browser memory
      |
  /api/step2  (2 Gemini calls)
  Call 2A: JSON1 → Summary
  Call 2B: JSON1 → Test Cases (12-15)
      |
  /api/export  (0 Gemini calls — pure Python)
  JSON1 + Tests → 3-sheet Excel download
  Saved to: results/TestDoc_*.xlsx

OPTIONAL:
  /api/step3  (1 call) — Gap Analysis
  /api/chat   (1 call per message) — AI questions


API CALL BUDGET
---------------
Action               Calls  Notes
-------------------  -----  -----------------------------------
Step 1 Extract        1     Only time FSD text goes to Gemini
Step 2 Summary        1     Uses JSON1 — no FSD re-read
Step 2 Test Cases     1     Uses JSON1 — no FSD re-read
Step 3 Gap Analysis   1     Optional, user triggered
AI Chat               1     Per message, optional
Excel Export          0     Pure Python, no Gemini needed
Health Check          1     Page load only
-------------------  -----  -----------------------------------
Core session total    3     (step1 + step2x2)
Maximum session       7     All features used


SUPPORTED FSD TYPES
-------------------
Any insurance / policy administration FSD including:
  SURR  : Surrender transactions
  NBS   : New Business Submission
  LOAN  : Policy Loan
  GSV   : Guaranteed Surrender Value
  SSV   : Special Surrender Value
  REI   : Reinstatement
  CLM   : Claims (death / maturity)
  CHG   : Policy Change transactions
  CAN   : Cancellation
  REN   : Renewal
  Any other Ingenium / COBOL / PAS transaction


CONFIGURATION (.env file)
-------------------------
GEMINI_API_KEY=AIzaSyYourKeyHere    Required. Get from aistudio.google.com
GEMINI_MODEL=gemini-2.0-flash       Model to use (free tier, recommended)
PORT=5000                           Port number (change if 5000 is in use)

Model options:
  gemini-2.0-flash       Fast, free tier, recommended
  gemini-2.0-flash-lite  Higher free quota, slightly less capable


TROUBLESHOOTING
---------------
Problem              Cause                Fix
-------------------  -------------------  ---------------------------
Red dot (offline)    Wrong API key        Check key starts with AIzaSy
429 Quota Exceeded   Free limit hit       Wait 60s OR create new key
404 Model Not Found  Wrong model name     Use gemini-2.0-flash in .env
Failed to fetch      Wrong folder         Must open http://localhost:5000
                                          NOT the index.html file directly
SSL Error            Corporate proxy      verify=False already set in code
PDF no text          Scanned image PDF    Save as DOCX first
Port in use          Another app on 5000  Change PORT=5001 in .env
API Key MISSING      .env not found       Run python app.py from app2 folder


SECURITY NOTES
--------------
* Your .env file contains the API key — NEVER share it
* Never paste the API key in chat, email, or code
* Add .env to .gitignore if using Git
* verify=False is set for corporate Zscaler SSL proxy
* All FSD text is sent to Google Gemini API


RESULTS FOLDER
--------------
Every run saves files automatically:
  results/JSON1_<filename>_<timestamp>.json   Lossless FSD extraction
  results/TestDoc_<filename>_<timestamp>.xlsx  Test document


TUNING AI BEHAVIOUR
-------------------
All AI instructions are in prompts.py.
To change how AI extracts, summarises, or generates tests:
  Open prompts.py
  Edit the relevant PROMPT_ constant
  Save the file
  Restart python app.py
  No other file needs changing.

PROMPT REFERENCE:
  PROMPT_JSON1    FSD text → JSON1 lossless extraction
  PROMPT_SUMMARY  JSON1 → business summary
  PROMPT_TESTS    JSON1 → test cases (12-15)
  PROMPT_GAP      JSON1 → gap analysis + quality score
  PROMPT_CHAT     JSON1 → conversational AI answers


================================================================
FSD2TESTDOC App2
Python 3.9+ · Flask · Gemini AI · openpyxl · pypdf · python-docx
Pipeline: JSON1 Architecture — Single Source of Truth
================================================================
