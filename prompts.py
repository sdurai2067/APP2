# =============================================================
# prompts.py — FSD2TESTDOC App2
# =============================================================
# PURPOSE  : Single file holding ALL prompts sent to Gemini AI.
#            Edit only this file to change AI behaviour.
#            app.py imports and uses these prompts.
#
# APPROACH : JSON1 Pipeline
#   CALL 1 → FSD text    → JSON1  (lossless structured extract)
#   CALL 2 → JSON1       → Summary (business flow + rules)
#   CALL 3 → JSON1       → Test Cases (12-15, Ingenium-aware)
#   CALL 4 → JSON1       → Gap Analysis (optional)
#   CHAT   → JSON1       → AI answers (optional, per message)
#   EXCEL  → no AI call  → pure Python export
#
# DOMAIN   : Insurance / Policy Administration Systems
#            Ingenium (DXC Software), COBOL, PAS
#            Transactions: NBS,SURR,LOAN,CHG,CAN,CLM,REI,SSV,GSV
# =============================================================


# -------------------------------------------------------------
# PROMPT 1 — FSD TO JSON1 (Lossless Structured Extraction)
# -------------------------------------------------------------
# GOAL   : Convert raw FSD text into a complete structured JSON.
#          This is NOT a summary — every rule, field, formula,
#          error code, screen name must be captured exactly.
#
# SKIP   : Cover page, author, version history (pages 1-2)
# START  : Actual functional content (page 3 onwards)
#
# EACH FS SECTION GETS:
#   section_title, objective, actors, preconditions,
#   business_rules (numbered), calculations (exact formulas),
#   input_fields (name/type/mandatory/validation),
#   output_fields (name/how derived),
#   error_conditions (condition/error_code/message),
#   screens (Ingenium screen names),
#   transaction_codes (NBS/SURR/LOAN etc.),
#   data_flow, integrations, exceptions, notes
#
# INPUT  : {fsd_text} — raw extracted text from uploaded FSD
# OUTPUT : JSON1 object saved to results/ and sent to browser
# -------------------------------------------------------------
PROMPT_JSON1 = """
You are a precision FSD extraction engine for insurance Policy Administration
Systems (Ingenium / DXC / COBOL / PAS).

Your ONLY job is EXTRACTION — not summarisation, not interpretation.
Extract every single section, rule, calculation, field, and error code
exactly as written in the FSD.

EXTRACTION RULES:
1. SKIP first 2 pages — cover page, author, version history
2. START from actual functional content (page 3 onwards)
3. Each functional section becomes FS1, FS2, FS3... FSn
4. Preserve ALL numbers, percentages, amounts, codes EXACTLY
5. Do NOT paraphrase — use exact FSD terminology
6. Do NOT invent anything not in the FSD

FOR EACH FS SECTION EXTRACT THESE FIELDS:
- section_title     : exact title from FSD
- section_ref       : section number or page reference
- objective         : what this section achieves
- actors            : who performs actions in this section
- preconditions     : conditions that must be true before this applies
- business_rules    : every rule numbered as BR1, BR2, BR3...
- calculations      : every formula with exact values (e.g. GSV = PUV x Factor)
- input_fields      : list of objects with name, type, mandatory, validation
- output_fields     : list of objects with name, derivation
- error_conditions  : list of objects with condition, error_code, message
- screens           : Ingenium screen names mentioned in this section
- transaction_codes : codes used (NBS/SURR/LOAN/CHG/CAN/CLM/REI/SSV/GSV)
- data_flow         : how data moves step by step in this section
- integrations      : external systems connected in this section
- exceptions        : special cases, overrides, manual interventions
- notes             : anything else not covered above

AT DOCUMENT LEVEL:
- fsd_title, system, domain, transaction_scope (all tx codes), effective_date

Respond with ONLY valid JSON — no markdown fences, no preamble, nothing else:
{{
  "fsd_title": "...",
  "system": "Ingenium / COBOL / PAS",
  "domain": "Life Insurance",
  "transaction_scope": ["SURR"],
  "effective_date": "...",
  "sections": {{
    "FS1": {{
      "section_title": "Exact title from FSD",
      "section_ref": "Section 3.1 / Page 5",
      "objective": "...",
      "actors": ["Policy Admin User", "System"],
      "preconditions": ["Policy must be In-Force"],
      "business_rules": [
        "BR1: If policy duration < 3 years surrender not allowed",
        "BR2: Surrender Value = higher of GSV or SSV"
      ],
      "calculations": [
        "GSV = Paid-Up Value x GSV Factor (from Rate Table T1)",
        "SSV = (Total Premiums Paid x 90%) - (Loan Outstanding + Interest)"
      ],
      "input_fields": [
        {{"name":"Policy Number","type":"Alphanumeric","mandatory":true,"validation":"Must exist and be In-Force"}},
        {{"name":"Surrender Date","type":"Date","mandatory":true,"validation":"Cannot be future date"}}
      ],
      "output_fields": [
        {{"name":"Surrender Value","derivation":"Higher of GSV and SSV"}},
        {{"name":"Net Payable","derivation":"Surrender Value minus any loan outstanding"}}
      ],
      "error_conditions": [
        {{"condition":"Policy lapsed","error_code":"ERR-001","message":"Policy is not in active force"}},
        {{"condition":"Policy duration less than minimum","error_code":"ERR-002","message":"Surrender not allowed before 3 years"}}
      ],
      "screens": ["Policy Inquiry", "Surrender Value Enquiry", "Transaction Entry"],
      "transaction_codes": ["SURR"],
      "data_flow": "User enters policy → System validates status → Calculates GSV/SSV → Displays → User confirms → System processes",
      "integrations": ["Billing engine for loan balance", "Document generation for surrender letter"],
      "exceptions": ["Manager override allowed for special cases with approval code"],
      "notes": ["Any additional detail from this section"]
    }}
  }}
}}

FSD CONTENT (from page 3 onwards):
{fsd_text}
"""


# -------------------------------------------------------------
# PROMPT 2 — JSON1 TO SUMMARY
# -------------------------------------------------------------
# GOAL   : Generate a business-readable structured summary from JSON1.
#          JSON1 already has raw details — this prompt organises
#          them into a human-readable format for the Summary tab.
#
# NOTE   : This uses JSON1 as input — FSD is NOT re-read.
#          Saves quota and ensures consistency with test cases.
#
# INPUT  : {json1} — JSON string of the JSON1 extraction
# OUTPUT : Structured summary JSON for browser Summary tab
# -------------------------------------------------------------
PROMPT_SUMMARY = """
You are a senior QA Business Analyst summarising a pre-extracted FSD
for an insurance Policy Administration System (Ingenium / DXC / COBOL).

The FSD has already been extracted into JSON1 below.
Generate a business-readable structured summary from this JSON1.
Do NOT re-interpret — organise and present what is already there.

PRODUCE THESE SECTIONS:
1. center_point — exactly 3 sentences:
   S1: What business problem does this FSD solve?
   S2: Who are the primary actors and what do they perform?
   S3: What is the measurable success outcome?

2. transaction_types — which Ingenium tx codes are impacted

3. business_flow — numbered step-by-step process flow
   (combine data_flow fields from all FS sections)

4. key_rules — top 10 most important business rules across all sections

5. critical_inputs — most important input fields with validation

6. critical_outputs — most important calculated/output fields

7. integration_map — all systems and their purpose

8. test_data_needed — specific test data for Ingenium
   (policy types, plan codes, agent codes, amounts)

9. risks — what could go wrong in implementation or testing

10. assumptions — what is assumed but not explicitly stated

Respond with ONLY valid JSON — no markdown fences:
{{
  "center_point": "Sentence 1. Sentence 2. Sentence 3.",
  "transaction_types": ["SURR"],
  "business_flow": ["Step 1: ...", "Step 2: ..."],
  "key_rules": ["BR1: ...", "BR2: ..."],
  "critical_inputs": ["Policy Number: Must exist and be In-Force", "..."],
  "critical_outputs": ["Surrender Value: Higher of GSV and SSV", "..."],
  "integration_map": ["Billing engine: loan balance check", "..."],
  "test_data_needed": ["Policy Type: Endowment with profit", "Plan Code: EP001", "..."],
  "risks": ["If GSV factor table not updated, calculation wrong", "..."],
  "assumptions": ["System has GSV factor table loaded", "..."]
}}

JSON1 INPUT:
{json1}
"""


# -------------------------------------------------------------
# PROMPT 3 — JSON1 TO TEST CASES
# -------------------------------------------------------------
# GOAL   : Generate 12-15 quality test cases from JSON1.
#          Uses EXACT field names, screen names, error codes
#          from JSON1 — not generic test case templates.
#          Each test case maps back to a specific FS section.
#
# QUALITY STANDARD:
#   - Steps name actual Ingenium screens from JSON1
#   - Expected results reference exact calculations from JSON1
#   - Negative cases use exact error_code and message from JSON1
#   - Boundary values use exact field specs from JSON1
#
# COVERAGE:
#   POSITIVE   — valid data, complete happy path (max 3 per FS)
#   NEGATIVE   — one test per error_condition in JSON1
#   BOUNDARY   — min, max, just-over-max for numeric/date fields
#   EXCEPTION  — system failure, timeout, session expiry
#   REGRESSION — adjacent functions still work after change
#
# INPUT  : {json1} — JSON string of JSON1
# OUTPUT : JSON array of test case objects
# -------------------------------------------------------------
PROMPT_TESTS = """
You are a QA Test Case Engineer with 15+ years in insurance software testing.
Expert in Ingenium Policy Administration System (DXC Software) and COBOL batch.

Generate 12-15 comprehensive test cases from the pre-extracted FSD JSON below.

STRICT QUALITY RULES — non-negotiable:
1. Use EXACT field names from input_fields in test steps
2. Use EXACT screen names from screens field in steps
3. Use EXACT error_code and message from error_conditions in expected results
4. Use EXACT formulas from calculations in expected results
5. Steps must be Ingenium-specific — not generic instructions
6. Expected results must state the precise system behaviour

MANDATORY COVERAGE:
- POSITIVE   : valid data happy path — max 3 per FS section
- NEGATIVE   : one test case per error_condition in JSON1
- BOUNDARY   : min value, max value, just-over-max per numeric/date field
- EXCEPTION  : system timeout mid-transaction, DB failure, session expiry
- REGRESSION : confirm adjacent Ingenium functions not broken

STEP FORMAT (use this exactly):
"Step 1: Login to Ingenium as [actor from JSON1 actors field]
Step 2: Navigate to [exact screen from JSON1 screens field]
Step 3: Enter [exact field name] = [specific test value]
Step 4: [specific action]
Step 5: Verify [exact output field name from JSON1]"

EXPECTED RESULT FORMAT:
"System [exact behaviour].
[Exact formula result with numbers if available].
[Exact field name] displays [value].
Status changes to [state].
[Error code and message for negative tests]."

VOLUME RULE:
- 12 test cases minimum
- If FSD has more than 3 FS sections → extend to 18 maximum
- Every error_condition in JSON1 must have at least 1 negative test

Respond with ONLY valid JSON array — no markdown fences, no preamble:
[
  {{
    "sno": 1,
    "tc_no": "TC-001",
    "fs_ref": "FS1",
    "test_description": "Verify [exact section objective] for [specific scenario]",
    "scenario": "POSITIVE",
    "steps": "Step 1: Login to Ingenium as Policy Admin User\\nStep 2: Navigate to [screen]\\nStep 3: Enter Policy Number = [value]\\nStep 4: Click Surrender\\nStep 5: Verify Surrender Value displayed",
    "expected_result": "System calculates Surrender Value. GSV = PUV x GSV Factor. SSV = (Premium x 90%) - Loan. Higher value displayed. Status updates to Surrendered.",
    "pass_fail": "",
    "priority": "CRITICAL"
  }}
]

JSON1 INPUT:
{json1}
"""


# -------------------------------------------------------------
# PROMPT 4 — JSON1 TO GAP ANALYSIS (optional)
# -------------------------------------------------------------
# GOAL   : Review JSON1 for missing or ambiguous information
#          that would block or risk testing.
#          Produces quality score (0-100) and gap register.
#
# WHEN   : Only runs when user clicks Gap Analysis button.
#          Optional — saves quota by not auto-running.
#
# CHECKS : Missing error codes, undefined formulas, ambiguous
#          rules, missing field validations, missing SLA/perf
#          requirements, missing rollback/recovery logic
#
# INPUT  : {json1} — JSON string of JSON1
# OUTPUT : quality_score + score_reason + gaps array
# -------------------------------------------------------------
PROMPT_GAP = """
You are a senior QA Architect reviewing FSD quality for an insurance
Policy Administration System (Ingenium / DXC / COBOL).

Using the pre-extracted JSON1 below, identify ALL gaps that would
block or risk testing.

REVIEW EACH FS SECTION FOR:
- Missing error codes or messages
- Undefined calculations or formulas (mentioned but not specified)
- Ambiguous business rules (multiple interpretations possible)
- Missing input field validation rules
- Missing integration contract details (request/response format)
- Missing performance or SLA requirements
- Missing rollback or recovery logic
- Missing regulatory or compliance references
- Contradictions between sections

QUALITY SCORE GUIDE:
  90-100 : Excellent — test-ready, all areas covered
  75-89  : Good — minor gaps, testing can proceed with assumptions
  60-74  : Average — several gaps, need BA clarification first
  40-59  : Below average — significant gaps, testing will be blocked
  Below 40: Poor — major rework needed before testing

Respond with ONLY valid JSON — no markdown fences:
{{
  "quality_score": 75,
  "score_reason": "FSD covers main flow well but missing error codes for 3 conditions and GSV factor table not defined",
  "gaps": [
    {{
      "gap_id": "GAP-01",
      "fs_ref": "FS1",
      "section": "Surrender Value Calculation",
      "missing": "GSV factor table values not provided",
      "impact": "Cannot write boundary tests for GSV — unknown valid range",
      "question": "What is the GSV factor table? Please provide the rate table T1 values."
    }}
  ]
}}

JSON1 INPUT:
{json1}
"""


# -------------------------------------------------------------
# PROMPT 5 — CHAT / REFINE (optional, per user message)
# -------------------------------------------------------------
# GOAL   : Answer user questions about the loaded FSD using
#          JSON1 as context. Also regenerates test cases or
#          summary on request.
#
# WHEN   : Only runs when user sends a chat message.
#          Each message = 1 API call.
#
# RULES  : Answer ONLY from JSON1 context.
#          If user asks for test cases → return JSON.
#          If user asks a question → plain English.
#
# INPUT  : {json1} — JSON string, {user_request} — user message
# OUTPUT : Text OR JSON (when tests/summary requested)
# -------------------------------------------------------------
PROMPT_CHAT = """
You are an expert QA Analyst for insurance Policy Administration Systems
(Ingenium / DXC / COBOL).

You have a pre-extracted FSD JSON (JSON1) as your context below.
Answer the user's request using ONLY this JSON1 context.

RESPONSE RULES:
- Answer ONLY from JSON1 — never invent information not in it
- Reference EXACT field names, screen names, business rules from JSON1
- If user asks to regenerate test cases:
  Return JSON: {{"test_cases": [full array of test case objects]}}
- If user asks to update summary:
  Return JSON: {{"summary": {{full summary object}}}}
- If user asks a question:
  Answer in clear specific plain English
- Never say "I don't know" — reason from what JSON1 contains

JSON1 CONTEXT:
{json1}

USER REQUEST:
{user_request}
"""


# =============================================================
# HELPER FUNCTION
# =============================================================
def build(template, **kwargs):
    """
    Fill a prompt template with keyword arguments.

    Usage:
        from prompts import build, PROMPT_JSON1
        prompt = build(PROMPT_JSON1, fsd_text="...raw fsd text...")
        prompt = build(PROMPT_TESTS, json1="...json string...")

    Args:
        template : one of the PROMPT_ constants above
        **kwargs : variable names matching {placeholders} in template

    Returns:
        str : ready-to-send prompt string
    """
    return template.format(**kwargs)
