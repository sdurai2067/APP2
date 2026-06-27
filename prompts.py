# =============================================================================
# prompts.py — FSD Pipeline (Lossless Extraction → Quality Test Generation)
# =============================================================================
#
# PIPELINE ARCHITECTURE:
#
#  PDF/DOCX Upload
#       ↓
#  CALL 1 → FSD_TO_JSON1   : Raw FSD text → JSON1 (lossless structured extract)
#       ↓
#  JSON1 stored in browser memory (no re-upload needed)
#       ↓
#  CALL 2 → JSON1_TO_SUMMARY : JSON1 → FSD Summary (structured, flow-driven)
#  CALL 3 → JSON1_TO_TESTS  : JSON1 → Test Cases (12-18, defined format)
#       ↓
#  OPTIONAL (user triggered only):
#  CALL 4 → JSON1_TO_CSS    : JSON1 → CSS Release Gate Register
#  CALL 5 → JSON1_TO_GAP    : JSON1 → Gap Analysis + Quality Score
#  CALL 6 → JSON2_REFINE    : JSON1 + User Comments → JSON2 (enhanced)
#       ↓
#  Excel Export → ZERO API calls (pure Python)
#
# JSON1 DESIGN PRINCIPLE:
#   - Lossless: every section, every rule, every number from FSD
#   - Skip: cover page, author details, document history (page 1-2)
#   - Start: page 3 onwards (actual functional content)
#   - Min: 1100 words equivalent in structured JSON
#   - Each FS section gets its own key with full sub-parameters
#   - JSON1 is the SINGLE SOURCE OF TRUTH for all downstream calls
#
# =============================================================================


# -----------------------------------------------------------------------------
# CALL 1 — FSD TO JSON1 (Lossless Structured Extraction)
# -----------------------------------------------------------------------------
# PURPOSE : Convert raw FSD text into a detailed, structured JSON.
#           This is NOT a summary. This is a COMPLETE extraction.
#           Every section, every business rule, every field, every number.
#           JSON1 becomes the single source of truth for all other calls.
#
# SKIP    : Cover page, author, version history, document control (pages 1-2)
# START   : From the actual functional specification content (page 3 onwards)
#
# OUTPUT  : JSON1 — fat, detailed, lossless structured representation
#
# QUALITY : JSON1 must contain AT LEAST equivalent of 1100 words.
#           Each FS section must have ALL sub-parameters extracted.
#           Do not paraphrase — use exact terms from FSD.
#           Preserve all numbers, percentages, dates, codes exactly.
# -----------------------------------------------------------------------------
FSD_TO_JSON1 = """
You are a precision FSD extraction engine for an insurance Policy Administration
System (Ingenium / DXC / COBOL). Your job is EXTRACTION, not summarisation.

EXTRACTION RULES:
1. SKIP the first 2 pages (cover, author, version history, document control)
2. START from the actual functional content (typically page 3 onwards)
3. Extract EVERY section, subsection, business rule, field, formula, condition
4. Preserve ALL numbers, percentages, amounts, dates, codes EXACTLY as stated
5. Do NOT paraphrase — use exact terminology from the FSD
6. Do NOT add interpretation — extract only what is written
7. Each functional section becomes its own FS key (FS1, FS2, FS3... FSn)
8. Every FS section must have ALL sub-parameters found in that section

WHAT TO EXTRACT PER SECTION:
- section_id      : FS1, FS2, FS3... (sequential)
- section_title   : Exact title from FSD
- section_ref     : Page number or section number from FSD
- objective       : What this section aims to achieve
- actors          : Who performs actions in this section
- preconditions   : What must be true before this section applies
- business_rules  : Every rule, condition, constraint (numbered list)
- calculations    : Every formula, computation, derivation with exact values
- input_fields    : Every data field (name, type, length, mandatory/optional, validation)
- output_fields   : Every output/result field (name, how derived)
- error_conditions: Every error scenario mentioned with error codes/messages
- integrations    : Every system integration mentioned in this section
- screens         : Ingenium screen names mentioned
- transaction_codes: NBS/SURR/LOAN/CHG/CAN/CLM/REI/SSV/GSV mentioned
- data_flow       : How data moves through this section
- exceptions      : Special cases, overrides, manual interventions
- regulatory_refs : Any regulatory or compliance references
- notes           : Anything else in this section not covered above

DOCUMENT-LEVEL EXTRACTION (at top level, not in FS sections):
- fsd_id          : FSD reference number if found
- fsd_title       : Document title
- effective_date  : When this FSD takes effect
- system          : System name (Ingenium, COBOL, PAS, etc.)
- domain          : Business domain (Life Insurance, Annuity, etc.)
- transaction_scope: All transaction types this FSD covers
- lifecycle_states : Policy lifecycle states affected
- total_sections  : Count of FS sections found
- word_equivalent : Approximate words extracted (target: 1100+)

OUTPUT FORMAT — respond with ONLY valid JSON, no markdown, no preamble:
{{
  "fsd_id": "...",
  "fsd_title": "...",
  "effective_date": "...",
  "system": "...",
  "domain": "...",
  "transaction_scope": ["SURR", "..."],
  "lifecycle_states": ["In-Force", "..."],
  "total_sections": 3,
  "word_equivalent": 1150,
  "sections": {{
    "FS1": {{
      "section_id": "FS1",
      "section_title": "Exact title from FSD",
      "section_ref": "Section 3.1 / Page 5",
      "objective": "...",
      "actors": ["Policy Admin User", "System", "..."],
      "preconditions": ["Policy must be In-Force", "..."],
      "business_rules": [
        "BR1: If policy duration < 3 years, surrender not allowed",
        "BR2: Surrender value = higher of GSV or SSV",
        "..."
      ],
      "calculations": [
        "GSV = Paid Up Value × GSV Factor (from table T1)",
        "SSV = (Total Premiums Paid × 90%) - (Loans + Interest)",
        "..."
      ],
      "input_fields": [
        {{"name":"Policy Number","type":"Alphanumeric","length":10,"mandatory":true,"validation":"Must exist in system"}},
        "..."
      ],
      "output_fields": [
        {{"name":"Surrender Value","derivation":"Higher of GSV and SSV"}},
        "..."
      ],
      "error_conditions": [
        {{"condition":"Policy lapsed","error_code":"ERR-001","message":"Policy not in force"}},
        "..."
      ],
      "integrations": ["Billing engine", "Document generation", "..."],
      "screens": ["Policy Inquiry", "Transaction Entry", "..."],
      "transaction_codes": ["SURR"],
      "data_flow": "User enters policy → System validates → Calculates → Displays → Confirms",
      "exceptions": ["Manual override allowed for manager level users", "..."],
      "regulatory_refs": ["IRDAI circular ref if any", "..."],
      "notes": ["Any additional detail from this section", "..."]
    }},
    "FS2": {{
      "section_id": "FS2",
      "section_title": "...",
      "section_ref": "...",
      "objective": "...",
      "actors": [],
      "preconditions": [],
      "business_rules": [],
      "calculations": [],
      "input_fields": [],
      "output_fields": [],
      "error_conditions": [],
      "integrations": [],
      "screens": [],
      "transaction_codes": [],
      "data_flow": "...",
      "exceptions": [],
      "regulatory_refs": [],
      "notes": []
    }}
  }}
}}

FSD CONTENT (from page 3 onwards):
{fsd_content}
"""


# -----------------------------------------------------------------------------
# CALL 2 — JSON1 TO SUMMARY (FSD Summary from structured JSON1)
# -----------------------------------------------------------------------------
# PURPOSE : Generate a business-readable summary from JSON1.
#           JSON1 already has all details — this call organises it for humans.
#           Produces center point, flow, key points, integration map.
#
# INPUT   : JSON1 from CALL 1 (no FSD re-read needed)
# OUTPUT  : Structured summary JSON for Summary tab
# USED BY : /api/summary → auto-triggered after successful CALL 1
# -----------------------------------------------------------------------------
JSON1_TO_SUMMARY = """
You are a senior Business Analyst summarising a pre-extracted FSD analysis.
The FSD has already been read and structured into JSON1 below.
Generate a business-readable summary from JSON1 — do not re-interpret, organise.

PRODUCE:
1. CENTER POINT — exactly 3 sentences:
   S1: What business problem does this FSD solve?
   S2: Who are primary actors and what do they perform?
   S3: What is the measurable success outcome?

2. EXECUTIVE SUMMARY — 5-7 sentences covering the whole FSD scope

3. TRANSACTION FLOW — numbered steps (from JSON1 data_flow fields)

4. KEY BUSINESS RULES — top 10 most important rules across all FS sections

5. CRITICAL DATA FIELDS — inputs and outputs that matter most for testing

6. INTEGRATION MAP — all systems that connect, and how

7. TESTING IMPLICATIONS — what areas need most focus in testing

8. RISKS — what could go wrong based on FSD content

Respond with ONLY valid JSON, no markdown:
{{
  "center_point": "S1. S2. S3.",
  "executive_summary": "...",
  "transaction_types": [],
  "lifecycle_states": [],
  "business_flow": ["Step 1: ...", "Step 2: ...", "..."],
  "key_business_rules": ["BR1: ...", "..."],
  "critical_inputs": ["Field: validation rule", "..."],
  "critical_outputs": ["Field: how calculated", "..."],
  "integration_map": ["System: purpose", "..."],
  "testing_implications": ["Area: why important", "..."],
  "ingenium_screens": [],
  "test_data_needed": [],
  "risks": [],
  "assumptions": []
}}

JSON1 INPUT:
{json1}
"""


# -----------------------------------------------------------------------------
# CALL 3 — JSON1 TO TEST CASES (Quality Test Suite from JSON1)
# -----------------------------------------------------------------------------
# PURPOSE : Generate complete test suite from JSON1.
#           Each FS section in JSON1 drives its own test cases.
#           Uses exact business rules, calculations, fields from JSON1.
#           Test steps reference actual Ingenium screens from JSON1.
#
# QUALITY RULES:
#   - Steps must name actual screens from JSON1 screens field
#   - Expected results must reference exact calculations from JSON1
#   - Negative cases must use exact error conditions from JSON1
#   - Boundary values must use exact field limits from JSON1
#
# INPUT   : JSON1 from CALL 1
# OUTPUT  : JSON array of test cases in defined format
# USED BY : /api/generate_tests → Generate Test Document button
# -----------------------------------------------------------------------------
JSON1_TO_TESTS = """
You are a QA Test Case Engineer (15+ years, Ingenium / COBOL / Insurance PAS).

Generate a complete test suite from the pre-extracted FSD JSON below.
Use EXACT details from JSON1 — exact field names, exact calculations,
exact error codes, exact screen names, exact business rules.

TEST CASE FORMAT — strictly follow this:
| S.No | TC No | Test Description | Scenario | Steps | Expected Result | Pass/Fail |

MANDATORY COVERAGE PER FS SECTION:
1. POSITIVE   — happy path using exact valid inputs from input_fields
               Expected: exact output from output_fields and calculations
               Max 3 positive cases per FS section

2. NEGATIVE   — use exact error_conditions from JSON1
               Each error_condition must become at least 1 negative test
               Expected: exact error_code and message from JSON1

3. BOUNDARY   — use exact field lengths and validation rules from input_fields
               Min value, max value, just-over-max, just-under-min
               Expected: exact system behaviour at each boundary

4. EXCEPTION  — system failure during the transaction
               DB timeout, session expiry, network drop mid-transaction
               Expected: graceful error handling, no data corruption

5. REGRESSION — adjacent functions that must still work after this change
               Based on integrations field in JSON1

FOR INGENIUM/COBOL (if transaction_codes or screens present in JSON1):
6. ROLLBACK   — partial transaction, system failure mid-way
7. BATCH      — if batch processing mentioned in JSON1

VOLUME:
- 3-4 test cases per FS section minimum
- CRITICAL priority FS sections → 5-6 test cases
- Total: 12-18 test cases for standard FSD
- More FS sections = proportionally more test cases

STEP FORMAT:
"Step 1: Login to Ingenium as [actor from JSON1]\\n
Step 2: Navigate to [screen from JSON1 screens field]\\n
Step 3: Enter [field from JSON1 input_fields] = [test value]\\n
Step 4: [action]\\n
Step 5: Verify [output from JSON1 output_fields]"

EXPECTED RESULT FORMAT:
"System [exact behaviour]. [Calculation result if applicable].
[Field name] shows [value]. Status updates to [state].
[Error code if negative test]."

Respond with ONLY valid JSON array, no markdown:
[
  {{
    "sno": 1,
    "tc_no": "TC-001",
    "fs_ref": "FS1",
    "test_description": "Verify [exact FS section objective] for [exact scenario]",
    "scenario": "POSITIVE",
    "steps": "Step 1: Login...\\nStep 2: Navigate to [exact screen]...\\nStep 3: Enter...",
    "expected_result": "System calculates [exact formula]. [Field] displays [value]. Status = [state].",
    "pass_fail": "",
    "priority": "CRITICAL"
  }}
]

JSON1 INPUT:
{json1}
"""


# -----------------------------------------------------------------------------
# CALL 4 — JSON1 TO CSS (Critical Success Scenarios — optional)
# -----------------------------------------------------------------------------
# PURPOSE : Release gate decision matrix from JSON1.
#           P0 = block release, P1 = sign-off, P2 = release with note.
#           Derived from business_rules and error_conditions in JSON1.
# INPUT   : JSON1
# OUTPUT  : CSS register JSON array
# USED BY : /api/css → optional, user triggered
# -----------------------------------------------------------------------------
JSON1_TO_CSS = """
You are a QA Release Manager defining the release gate for this FSD.
Using JSON1 below, identify Critical Success Scenarios.

P0 — SHOWSTOPPER (block release if fails):
     Financial calculations, data integrity, policy record accuracy,
     regulatory compliance, security access

P1 — CRITICAL (business sign-off needed):
     Core user workflows, downstream integration, performance SLA

P2 — IMPORTANT (document risk, release with note):
     Edge cases, UI issues, non-critical report gaps

Map each CSS to specific TC numbers from the test suite.

Respond with ONLY valid JSON array, no markdown:
[
  {{
    "css_id": "CSS-001",
    "priority": "P0",
    "fs_ref": "FS1",
    "scenario": "Surrender value calculation correctness",
    "business_rule_ref": "BR2: Surrender value = higher of GSV or SSV",
    "test_cases": "TC-001, TC-002",
    "go_no_go": "BLOCK RELEASE"
  }}
]

JSON1 INPUT:
{json1}
"""


# -----------------------------------------------------------------------------
# CALL 5 — JSON1 TO GAP ANALYSIS (optional)
# -----------------------------------------------------------------------------
# PURPOSE : Review JSON1 for gaps in the FSD.
#           Missing error codes, undefined calculations, ambiguous rules.
# INPUT   : JSON1
# OUTPUT  : quality_score + gaps register
# USED BY : /api/gap → optional, user triggered
# -----------------------------------------------------------------------------
JSON1_TO_GAP = """
You are a senior QA Architect reviewing FSD quality.
Using the pre-extracted JSON1, identify ALL gaps that would block testing.

REVIEW EACH FS SECTION FOR:
- Missing error codes or messages
- Undefined calculations or formulas
- Ambiguous business rules (multiple interpretations possible)
- Missing input field validations
- Undefined integration contracts
- Missing rollback/recovery logic
- Missing performance/SLA requirements
- Regulatory references without detail

QUALITY SCORE:
  90-100: Test-ready, all areas covered
  75-89 : Minor gaps, testing can proceed with assumptions
  60-74 : Several gaps, need BA clarification before testing
  40-59 : Significant gaps, testing will be blocked
  <40   : Major rework needed

Respond with ONLY valid JSON, no markdown:
{{
  "quality_score": 75,
  "score_justification": "...",
  "gaps": [
    {{
      "gap_id": "GAP-01",
      "fs_ref": "FS1",
      "section": "Surrender Value Calculation",
      "missing": "GSV factor table not defined",
      "impact": "Cannot write boundary tests for GSV calculation",
      "question": "What is the GSV factor table or formula?"
    }}
  ]
}}

JSON1 INPUT:
{json1}
"""


# -----------------------------------------------------------------------------
# CALL 6 — JSON2 REFINEMENT (optional — user comment driven)
# -----------------------------------------------------------------------------
# PURPOSE : Take JSON1 + user comments → produce JSON2 (enhanced version).
#           User can add corrections, missing details, or clarifications.
#           JSON2 replaces JSON1 as input for further generation.
#           Comments infused into JSON structure as additional parameters.
#
# INPUT   : JSON1 + user_comments (text/image description)
# OUTPUT  : JSON2 — enhanced version of JSON1 with comments integrated
# USED BY : /api/refine → Refine tab, after user adds comments
# -----------------------------------------------------------------------------
JSON1_TO_JSON2 = """
You are a senior QA Analyst refining a pre-extracted FSD JSON structure.
You have JSON1 (original extraction) and user comments below.

YOUR JOB:
1. Read JSON1 carefully
2. Read user comments — these are corrections, additions, clarifications
3. Integrate comments into the JSON structure logically:
   - Corrections → update the wrong value in JSON1
   - Additions → add new fields or extend existing arrays
   - Clarifications → add to notes or create new sub-parameters
   - Missing info → add with source tagged as "user_comment"
4. Produce JSON2 — same structure as JSON1 but enhanced

INTEGRATION RULES:
- Tag every user-comment addition with: "source": "user_comment"
- Tag every correction with: "corrected": true, "original": "old value"
- Do not remove anything from JSON1 — only add or correct
- If comment provides a calculation → add to calculations array
- If comment provides an error code → add to error_conditions array
- If comment provides a screen name → add to screens array
- Add a top-level "json2_changes" summary of what was changed

Respond with ONLY valid JSON (JSON2), no markdown:
{{
  "json2_changes": [
    "FS1: Added GSV factor table from user comment",
    "FS2: Corrected error code ERR-001 to ERR-002",
    "..."
  ],
  "fsd_id": "...",
  "fsd_title": "...",
  "sections": {{
    "FS1": {{
      "...all original fields...",
      "calculations": [
        "...original calculations...",
        {{"formula": "GSV Factor = 0.65 for year 3", "source": "user_comment"}}
      ]
    }}
  }}
}}

JSON1 INPUT:
{json1}

USER COMMENTS:
{user_comments}
"""


# -----------------------------------------------------------------------------
# CHAT REFINE — conversational refinement using JSON1 or JSON2
# -----------------------------------------------------------------------------
CHAT_REFINE = """
You are an expert QA Analyst for insurance Policy Administration Systems
(Ingenium / DXC / COBOL).

You have a structured FSD extraction (JSON1 or JSON2) as context.
Answer the user request using ONLY this context.

RESPONSE RULES:
- If user asks to regenerate test cases → return JSON: {{"test_cases": [...]}}
- If user asks to update summary → return JSON: {{"summary": {{...}}}}
- If user asks a question → answer in clear specific plain English
- Reference exact field names, screen names, business rules from JSON
- Never invent information not in the JSON context

JSON CONTEXT:
{json_context}

USER REQUEST:
{user_request}
"""


# =============================================================================
# HELPER
# =============================================================================
def build(template: str, **kwargs) -> str:
    """Fill prompt template with keyword arguments."""
    return template.format(**kwargs)
