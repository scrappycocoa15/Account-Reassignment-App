"""
Client Sales Territory Redistribution Tool — Streamlit App

Data sources:
  • File upload  — drag-and-drop XLS/CSV, works without SFDC connection
  • Salesforce   — Analytics API with dynamic owner-filter override

Distribution rules:
  • Customer accounts      → greedy by ARR (equal total customer ARR per rep)
  • Non-customer accounts  → greedy by count (equal volume)
  • Open opps              → follow account assignment (same rep as account)
  • FY18 Sales Planning    → always append 'Prev Acct Owner: <name>' (preserves existing tags)

Output:
  • Full Excel  — all original columns + new owner cols + FY18 update
  • FS Excel    — required fields only: 18-digit ID, Account Name,
                  New Owner ID, New Owner Name, ARR
  • Email       — mailto: link to concur_fieldservices@sap.com
"""

import io, re, time, math, urllib.parse
from collections import defaultdict

import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Territory Redistribution",
    page_icon="🔄",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────────────────────────────────────
# SALESFORCE CONFIG
# ─────────────────────────────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# QUARTER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _current_and_next_quarter_end() -> tuple:
    """Return (current_quarter_start, next_quarter_end) as YYYY-MM-DD strings."""
    from datetime import date, timedelta
    today  = date.today()
    q      = (today.month - 1) // 3          # 0-indexed quarter
    cq_start_month = q * 3 + 1
    cq_start = date(today.year, cq_start_month, 1)
    # Next quarter end = first day of quarter after next - 1 day
    nq     = q + 2                           # quarter index two ahead
    nq_year  = today.year + nq // 4
    nq_month = (nq % 4) * 3 + 1
    nq_end   = date(nq_year, nq_month, 1) - timedelta(days=1)
    return str(cq_start), str(nq_end)

SF_INSTANCE        = "https://sapconcur.my.salesforce.com"
SF_API_VER         = "v59.0"
DEFAULT_ACCT_RPT   = "00OPg00000QkSbp"
DEFAULT_OPP_RPT    = "00OPg00000QkSf3"
FS_EMAIL           = "concur_fieldservices@sap.com"

# ─────────────────────────────────────────────────────────────────────────────
# COLOUR PALETTE  (SAP light)
# ─────────────────────────────────────────────────────────────────────────────
NAVY    = "00144A"
AMBER   = "F0AB00"
WHITE   = "FFFFFF"
BLUE_BG = "E8F4FF"
GREY_BG = "E1E2E6"
GREEN   = "107E3E"

# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', '72', sans-serif; }

.tool-header {
    background: linear-gradient(135deg, #0070F2 0%, #0134BF 100%);
    border-radius: 10px;
    padding: 22px 28px 18px 28px;
    margin-bottom: 28px;
}
.tool-header h1 { color: #fff !important; font-size: 1.45rem; font-weight: 700; margin: 0; }
.tool-header p  { color: #cfe6ff; font-size: 0.84rem; margin: 5px 0 0 0; }

.step-card {
    background: #F5F6F7;
    border: 1px solid #DCDCDC;
    border-radius: 8px;
    padding: 18px 22px 14px 22px;
    margin-bottom: 18px;
}
.step-num {
    display: inline-block;
    background: #0070F2;
    color: #fff;
    font-size: 0.72rem;
    font-weight: 700;
    border-radius: 50%;
    width: 22px; height: 22px;
    line-height: 22px;
    text-align: center;
    margin-right: 8px;
}
.step-title { font-size: 1rem; font-weight: 700; color: #32363A; }
.step-done  { opacity: 0.55; }

.info-box {
    background: #E1F4FF;
    border-left: 4px solid #4CB1FF;
    border-radius: 4px;
    padding: 10px 14px;
    font-size: 0.83rem;
    color: #32363A;
    margin: 8px 0 12px 0;
}
.warn-box {
    background: #FFF3CD;
    border-left: 4px solid #F0AB00;
    border-radius: 4px;
    padding: 10px 14px;
    font-size: 0.83rem;
    color: #32363A;
    margin: 8px 0 12px 0;
}
.success-box {
    background: #F1FAF5;
    border-left: 4px solid #107E3E;
    border-radius: 4px;
    padding: 10px 14px;
    font-size: 0.83rem;
    color: #32363A;
    margin: 8px 0 12px 0;
}
.metric-row { display: flex; gap: 12px; margin: 12px 0; flex-wrap: wrap; }
.metric-card {
    background: #fff;
    border: 1px solid #DCDCDC;
    border-radius: 6px;
    padding: 12px 18px;
    text-align: center;
    min-width: 110px;
}
.metric-card .label { font-size: 0.69rem; font-weight: 600; color: #6a6d70;
                      text-transform: uppercase; letter-spacing: 0.04em; }
.metric-card .value { font-size: 1.55rem; font-weight: 700; color: #0070F2; line-height: 1.1; }
.metric-card .value.green { color: #107E3E; }

.section-title {
    font-size: 0.9rem; font-weight: 700; color: #32363A;
    border-bottom: 2px solid #0070F2;
    padding-bottom: 4px; margin: 16px 0 10px 0;
}
footer { visibility: hidden; }
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────────────────────────────────────
# SESSION STATE INITIALISATION
# ─────────────────────────────────────────────────────────────────────────────
_DEFAULTS = {
    # ── shared ────────────────────────────────────────────────────────────────
    "mode":             "distribute",   # "distribute" | "return"
    "connected":        False,
    "conn_msg":         "",
    "sender_name":      "",
    # ── distribute mode ───────────────────────────────────────────────────────
    "roster_df":        None,
    "acct_df":          None,
    "opp_df":           None,
    "departing_name":   "",
    "departing_id":     "",
    "tag_name":         "",
    "receiving_reps":   [],    # list of {"name": .., "id": ..}
    "include_opps":     False,
    "result_full":      None,  # bytes
    "result_fs":        None,  # bytes
    "result_filename":  "",
    "run_done":         False,
    "dropped_acct":     [],
    "dropped_opp":      [],
    "dist_summary":     None,  # DataFrame
    "result_acct_df":   None,  # DataFrame — for preview
    "result_opp_df":    None,  # DataFrame — for preview
    # ── return mode ───────────────────────────────────────────────────────────
    "ret_rep_name":     "",
    "ret_rep_id":       "",
    "ret_tag_string":   "",    # exact string to match in Prev Acct Owner tag
    "ret_manager":      "",    # optional manager filter
    "ret_acct_df":      None,
    "ret_opp_df":       None,
    "ret_include_opps": False,
    "ret_run_done":     False,
    "ret_result_full":  None,
    "ret_result_fs":    None,
    "ret_result_filename": "",
    "ret_summary":      None,
    "ret_result_acct_df": None,  # DataFrame — for preview
    "ret_result_opp_df":  None,  # DataFrame — for preview
}
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

def reset_results():
    for k in ("result_full", "result_fs", "result_filename", "run_done",
              "dist_summary", "result_acct_df", "result_opp_df"):
        st.session_state[k] = _DEFAULTS[k]

def reset_ret_results():
    for k in ("ret_result_full", "ret_result_fs", "ret_result_filename",
              "ret_run_done", "ret_summary",
              "ret_result_acct_df", "ret_result_opp_df"):
        st.session_state[k] = _DEFAULTS[k]

# ─────────────────────────────────────────────────────────────────────────────
# SALESFORCE HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def sf_headers(sid: str) -> dict:
    return {"Authorization": f"Bearer {sid}",
            "Accept": "application/json",
            "Content-Type": "application/json"}

def test_connection(sid: str):
    try:
        r = requests.get(
            f"{SF_INSTANCE}/services/data/{SF_API_VER}/",
            headers=sf_headers(sid), timeout=15
        )
        if r.status_code == 401:
            return False, "Session ID invalid or expired."
        r.raise_for_status()
        # Get display name
        u = requests.get(
            f"{SF_INSTANCE}/services/oauth2/userinfo",
            headers=sf_headers(sid), timeout=15
        )
        name = u.json().get("name", "unknown") if u.ok else "unknown"
        return True, f"Connected as {name}"
    except Exception as e:
        return False, str(e)

def lookup_user_sfdc(sid: str, name_fragment: str) -> list:
    """Return list of {Id, Name, IsActive} matching the fragment."""
    safe = name_fragment.replace("'", "\\'")
    soql = (f"SELECT Id, Name, IsActive FROM User "
            f"WHERE Name LIKE '%{safe}%' AND IsActive = true LIMIT 20")
    url  = f"{SF_INSTANCE}/services/data/{SF_API_VER}/query"
    try:
        r = requests.get(url, headers=sf_headers(sid),
                         params={"q": soql}, timeout=20)
        r.raise_for_status()
        return r.json().get("records", [])
    except Exception:
        return []

# ── SOQL-based fetching (primary data path) ───────────────────────────────────

def _soql_query_all(sid: str, soql: str, status_fn=None) -> list:
    """
    Execute a SOQL query and return ALL records, automatically following
    nextRecordsUrl pagination.  No row-count cap.
    """
    url     = f"{SF_INSTANCE}/services/data/{SF_API_VER}/query"
    records = []
    params  = {"q": soql}
    while True:
        r = requests.get(url, headers=sf_headers(sid),
                         params=params, timeout=30)
        r.raise_for_status()
        data  = r.json()
        batch = data.get("records", [])
        records.extend(batch)
        if status_fn:
            status_fn(f"Fetched {len(records):,} records…")
        if data.get("done", True):
            break
        url    = f"{SF_INSTANCE}{data['nextRecordsUrl']}"
        params = {}
    return records


def _discover_account_fields(sid: str, owner_id: str = None,
                             status_fn=None) -> tuple:
    """
    Locates the Contractual ARR and FY18 Sales Planning field API names for
    Account.  Uses a multi-pass strategy; stops at first successful detection:

      Pass 0. FIELDS(CUSTOM) / FIELDS(ALL) SOQL probe — reads a live Account
              record and scans ALL accessible field keys for ARR-like names.
              Completely bypasses describe FLS restrictions.
      Pass 1. Exact normalized label match against a known-candidate list.
      Pass 2. Label contains both 'contractual' and 'arr'.
      Pass 3. API name contains both 'contractual' and 'arr'.
      Pass 4. Currency/number/double field whose API name contains 'arr'.
      Pass 5. Direct API-name guesses for common Contractual ARR patterns.
      Pass 6. Per-candidate SOQL probe — confirms field accessibility without
              requiring a non-null value (many accounts may legitimately
              have null ARR).

    If all passes fail, logs every numeric/currency field from describe and
    the best FIELDS(ALL) candidate (if any) so the user can copy the correct
    API name into the sidebar override box.

    Returns (arr_label, arr_api_name, fy18_label, fy18_api_name).
    Any element may be None if not found.
    """
    import unicodedata, re

    def _norm(s: str) -> str:
        """Normalize for comparison: NFKC unicode, collapse all whitespace,
        lowercase, strip.  Handles non-breaking spaces, double spaces, tabs."""
        s = unicodedata.normalize("NFKC", str(s))
        s = re.sub(r"\s+", " ", s)          # collapse runs of whitespace
        return s.lower().strip()

    url = f"{SF_INSTANCE}/services/data/{SF_API_VER}/sobjects/Account/describe"
    r   = requests.get(url, headers=sf_headers(sid), timeout=30)
    r.raise_for_status()
    all_fields = r.json().get("fields", [])

    if status_fn:
        status_fn(f"Account describe returned {len(all_fields)} fields.")

    by_label = {_norm(f["label"]): f for f in all_fields}
    by_name  = {_norm(f["name"]):  f for f in all_fields}

    # ── ARR field ────────────────────────────────────────────────────────────
    arr_result = None

    # Pass 0 — FIELDS(CUSTOM) / FIELDS(ALL) probe
    # SELECT FIELDS(CUSTOM) bypasses describe FLS entirely.  We scan the
    # returned record's keys for ARR-like names (numeric) and FY18-like names
    # (any type — it is a text field).  Candidates are used immediately if a
    # good match is found, or stored as fallbacks for later passes.
    _fields_all_candidate = None   # ARR best guess even if value is null
    _fy18_candidate       = None   # FY18 best guess from FIELDS probe
    if owner_id:
        for _fkw in ("FIELDS(CUSTOM)", "FIELDS(ALL)"):
            try:
                _fsql = (f"SELECT {_fkw} FROM Account "
                         f"WHERE OwnerId = '{owner_id}' LIMIT 5")
                _frecs = _soql_query_all(sid, _fsql)
                if _frecs:
                    # Collect all non-attribute keys from all sample records.
                    # Prefer keys where any record has a non-null numeric value.
                    _candidate_keys = {}  # key → best_value
                    for _rec in _frecs:
                        for _k, _v in _rec.items():
                            if _k == "attributes":
                                continue
                            if _k not in _candidate_keys and isinstance(_v, (int, float)):
                                _candidate_keys[_k] = _v
                            elif _k not in _candidate_keys:
                                _candidate_keys[_k] = _v

                    # Score keys: ARR (numeric) and FY18 (any type)
                    for _k, _v in _candidate_keys.items():
                        _kl = _k.lower()
                        if ("arr" in _kl or "contractual" in _kl):
                            if _v is not None and isinstance(_v, (int, float)):
                                # Non-null numeric hit — use immediately
                                arr_result = {"label": _k, "name": _k,
                                              "type": "currency"}
                                if status_fn:
                                    status_fn(
                                        f"ARR found (Pass 0 {_fkw}): "
                                        f"{_k!r} = {_v}")
                                break
                            elif _fields_all_candidate is None:
                                # Null but looks right — remember as fallback
                                _fields_all_candidate = _k
                        if "fy18" in _kl and _fy18_candidate is None:
                            _fy18_candidate = _k
                            if status_fn:
                                status_fn(f"FY18 candidate (Pass 0 {_fkw}): {_k!r}")
                    if arr_result:
                        break
                if arr_result:
                    break
            except Exception as _fe:
                if status_fn:
                    status_fn(f"Pass 0 {_fkw} failed: {_fe}")
                continue

    # Pass 1 — exact normalized label match
    for lbl in ["contractual arr (converted)", "contractual arr",
                "contractual arr usd", "contractual arr (usd)",
                "arr", "annual recurring revenue", "annual contract value"]:
        if lbl in by_label:
            arr_result = by_label[lbl]
            if status_fn:
                status_fn(f"ARR found (Pass 1 label): {arr_result['label']!r}")
            break

    # Pass 2 — label contains both 'contractual' and 'arr'
    if not arr_result:
        for f in all_fields:
            ll = _norm(f["label"])
            if "contractual" in ll and "arr" in ll:
                arr_result = f
                if status_fn:
                    status_fn(f"ARR found (Pass 2 partial label): {f['label']!r}")
                break

    # Pass 3 — API name contains both 'contractual' and 'arr'
    if not arr_result:
        for f in all_fields:
            nl = _norm(f["name"])
            if "contractual" in nl and "arr" in nl:
                arr_result = f
                if status_fn:
                    status_fn(f"ARR found (Pass 3 partial API name): {f['name']!r}")
                break

    # Pass 4 — any currency/double/number field whose API name contains 'arr'
    if not arr_result:
        for f in all_fields:
            if (f.get("type") in ("currency", "double", "percent", "number") and
                    "arr" in _norm(f["name"])):
                arr_result = f
                if status_fn:
                    status_fn(f"ARR found (Pass 4 type+name): {f['name']!r}")
                break

    # Pass 5 — try direct API-name guesses derived from known label variants
    if not arr_result:
        guess_names = [
            "contractual_arr__c", "contractual_arr_converted__c",
            "contractual_arr_usd__c", "contractualarr__c",
            "contractual_arr_value__c", "arr__c", "contractual_arr_amount__c",
        ]
        for nm in guess_names:
            if nm in by_name:
                arr_result = by_name[nm]
                if status_fn:
                    status_fn(f"ARR found (Pass 5 name guess): {arr_result['name']!r}")
                break

    # Pass 6 — live SOQL probe: try each candidate API name against a real
    # record.  Existence is confirmed when the query returns without a SOQL
    # error — we no longer require a non-null value (many accounts legitimately
    # have null ARR, which caused the old check to discard a valid field name).
    if not arr_result and owner_id:
        probe_names = [
            # Most common SAP Concur patterns first
            "Contractual_ARR__c",
            "Contractual_ARR_converted__c",
            "Contractual_ARR_USD__c",
            "ContractualARR__c",
            "Contractual_ARR_Value__c",
            "Contractual_ARR_Amount__c",
            "Contractual_Annual_Recurring_Revenue__c",
            "CARR__c",
            "C_ARR__c",
            "Contract_ARR__c",
            "ARR__c",
            "Annual_Recurring_Revenue__c",
            "Annual_Contract_Value__c",
        ]
        # If FIELDS(ALL) surfaced a likely candidate name, prepend it
        if _fields_all_candidate and _fields_all_candidate not in probe_names:
            probe_names.insert(0, _fields_all_candidate)
        for cand in probe_names:
            try:
                test_soql = (f"SELECT {cand} FROM Account "
                             f"WHERE OwnerId = '{owner_id}' LIMIT 1")
                test_recs = _soql_query_all(sid, test_soql)
                # If query ran without raising an exception the field exists
                # and is accessible — value may legitimately be null.
                val_sample = test_recs[0].get(cand) if test_recs else None
                arr_result = {"label": "Contractual ARR (converted)",
                              "name": cand, "type": "currency"}
                if status_fn:
                    status_fn(
                        f"ARR found (Pass 6 SOQL probe): {cand!r} "
                        f"(sample value: {val_sample!r})")
                break
            except Exception:
                pass   # field doesn't exist or not accessible — try next

    # ── If still not found — try FIELDS(ALL) null candidate, then dump list ──
    if not arr_result:
        # Last resort: if Pass 0 found a key that looks right but had null values,
        # use it anyway (field exists; accounts may just have no ARR data yet).
        if _fields_all_candidate:
            arr_result = {"label": "Contractual ARR (converted)",
                          "name": _fields_all_candidate, "type": "currency"}
            if status_fn:
                status_fn(
                    f"ARR field (all-null fallback from FIELDS probe): "
                    f"{_fields_all_candidate!r}"
                )
        elif status_fn:
            numeric = [f for f in all_fields
                       if f.get("type") in ("currency", "double", "percent",
                                            "number", "int")]
            status_fn(
                f"ARR field NOT found after all passes. "
                f"Listing all {len(numeric)} numeric/currency fields below — "
                f"copy the API Name of your ARR field into the sidebar override:"
            )
            for f in numeric:
                status_fn(f"  Label: {f['label']!r:50s}  API: {f['name']!r}")

    # ── FY18 field ───────────────────────────────────────────────────────────
    fy18_result = None
    # Pass A — exact label
    if _norm("fy18 sales planning") in by_label:
        fy18_result = by_label[_norm("fy18 sales planning")]
    # Pass B — known API name guesses in describe
    if not fy18_result:
        for nm in ["fy18_sales_planning__c", "fy18salesplanning__c",
                   "fy18_salesplanning__c", "fy18_sales_plan__c"]:
            if nm in by_name:
                fy18_result = by_name[nm]
                break
    # Pass C — any describe field with "fy18" in label or name
    if not fy18_result:
        for f in all_fields:
            ll = _norm(f["label"])
            nl = _norm(f["name"])
            if ("fy18" in ll and "planning" in ll) or \
               ("fy18" in nl and "planning" in nl):
                fy18_result = f
                break
    # Pass D — SOQL existence probe (same approach as ARR Pass 6)
    if not fy18_result and owner_id:
        fy18_probe_names = [
            "FY18_Sales_Planning__c",
            "FY18SalesPlanning__c",
            "FY18_SalesPlanning__c",
            "FY18_Sales_Plan__c",
            "FY18__c",
        ]
        if _fy18_candidate and _fy18_candidate not in fy18_probe_names:
            fy18_probe_names.insert(0, _fy18_candidate)
        for _cand in fy18_probe_names:
            try:
                _tsql = (f"SELECT {_cand} FROM Account "
                         f"WHERE OwnerId = '{owner_id}' LIMIT 1")
                _soql_query_all(sid, _tsql)
                # Query ran without error → field exists and is accessible
                fy18_result = {"label": "FY18 Sales Planning",
                               "name": _cand, "type": "string"}
                if status_fn:
                    status_fn(f"FY18 found (Pass D SOQL probe): {_cand!r}")
                break
            except Exception:
                pass
    # Pass E — use FIELDS(CUSTOM) candidate as last resort
    if not fy18_result and _fy18_candidate:
        fy18_result = {"label": "FY18 Sales Planning",
                       "name": _fy18_candidate, "type": "string"}
        if status_fn:
            status_fn(f"FY18 found (Pass E FIELDS candidate): {_fy18_candidate!r}")

    arr_label  = arr_result["label"]  if arr_result  else None
    arr_api    = arr_result["name"]   if arr_result  else None
    fy18_label = fy18_result["label"] if fy18_result else None
    fy18_api   = fy18_result["name"]  if fy18_result else None

    if status_fn:
        status_fn(
            f"Result — ARR: {arr_label!r} → {arr_api!r}  |  "
            f"FY18: {fy18_label!r} → {fy18_api!r}"
        )
    return arr_label, arr_api, fy18_label, fy18_api


def fetch_accounts_soql(sid: str, owner_id: str,
                        arr_field_override: str = None,
                        status_fn=None) -> tuple:
    """
    Fetch every Account owned by owner_id via SOQL (auto-paginated, no row cap).

    Custom fields (Contractual ARR, FY18 Sales Planning) are discovered from
    the Account object describe so the resulting column labels match what
    detect_arr_col() and detect_fy18_col() expect.

    Returns (list_of_row_dicts, warnings).
    """
    warnings_out = []

    if status_fn:
        status_fn("Describing Account object to locate ARR and FY18 fields…")

    arr_label, arr_field, fy18_label, fy18_field = _discover_account_fields(
        sid, owner_id=owner_id, status_fn=status_fn
    )
    # Apply manual override if the user supplied one
    if arr_field_override and arr_field_override.strip():
        arr_field = arr_field_override.strip()
        arr_label = "Contractual ARR (converted)"
        if status_fn:
            status_fn(f"ARR field overridden to: {arr_field}")

    select_parts = [
        "Id", "Name", "OwnerId", "Owner.Name", "Owner.Manager.Name",
        "Type", "Rating", "LastActivityDate",
    ]
    if arr_field:
        select_parts.append(arr_field)
    if fy18_field:
        select_parts.append(fy18_field)

    soql = (f"SELECT {', '.join(select_parts)} "
            f"FROM Account WHERE OwnerId = '{owner_id}'")

    if status_fn:
        status_fn(f"Running SOQL query for owner {owner_id}…")
    records = _soql_query_all(sid, soql, status_fn=status_fn)

    rows = []
    for rec in records:
        owner = rec.get("Owner") or {}
        mgr   = owner.get("Manager") or {}
        row   = {
            "18 Digit Account ID": rec.get("Id", ""),
            "Account Name":        rec.get("Name", ""),
            "Account Owner":       owner.get("Name", ""),
            "Account Owner ID":    rec.get("OwnerId", ""),
            "Manager":             mgr.get("Name", ""),
            "Account Type":        rec.get("Type",   "") or "",
            "Rating":              rec.get("Rating", "") or "",
            "Last Activity":       str(rec.get("LastActivityDate") or ""),
        }
        if arr_field:
            val = rec.get(arr_field)
            try:
                # Always store under "Contractual ARR (converted)" so detect_arr_col()
                # finds it regardless of how the field is labeled in this org.
                row["Contractual ARR (converted)"] = (
                    f"USD {float(val):,.2f}" if val is not None else ""
                )
            except (TypeError, ValueError):
                row["Contractual ARR (converted)"] = str(val) if val is not None else ""
        if fy18_field:
            val = rec.get(fy18_field)
            # Always store under the canonical label detect_fy18_col() expects.
            row["FY18 Sales Planning"] = str(val) if val is not None else ""
        rows.append(row)

    if not arr_field:
        warnings_out.append(
            "No Contractual ARR field found on Account — customers will be "
            "distributed by count only."
        )
    if not fy18_field:
        warnings_out.append(
            "FY18 Sales Planning field not found on Account — tagging will be skipped."
        )

    if status_fn:
        status_fn(f"Done — {len(rows):,} accounts loaded.")
    return rows, warnings_out


def fetch_opps_soql(sid: str, owner_id: str,
                    amount_field_override: str = None,
                    status_fn=None) -> tuple:
    """
    Fetch all open Opportunity records owned by owner_id via SOQL.
    Discovers the Forecast Amount field via Opportunity/describe so the
    pipeline totals are accurate even when the amount is a custom field.
    Returns (rows_as_list_of_dicts, warnings).
    """
    # ── Discover amount field ────────────────────────────────────────────────
    amt_api   = None
    amt_label = "Forecast Amount"

    if amount_field_override and amount_field_override.strip():
        amt_api = amount_field_override.strip()
        if status_fn:
            status_fn(f"Opp amount field overridden to: {amt_api}")
    else:
        if status_fn:
            status_fn("Describing Opportunity object to locate amount field…")
        try:
            _url = (f"{SF_INSTANCE}/services/data/{SF_API_VER}"
                    f"/sobjects/Opportunity/describe")
            _r   = requests.get(_url, headers=sf_headers(sid), timeout=30)
            _r.raise_for_status()
            opp_fields = _r.json().get("fields", [])
            by_lbl = {f["label"].lower(): f for f in opp_fields}
            by_nm  = {f["name"].lower():  f for f in opp_fields}

            # Pass 1 — exact label candidates
            for lbl in ["forecast amount", "amount", "deal amount",
                        "opportunity amount"]:
                if lbl in by_lbl:
                    amt_api   = by_lbl[lbl]["name"]
                    amt_label = by_lbl[lbl]["label"]
                    break
            # Pass 2 — partial label ("forecast" + "amount")
            if not amt_api:
                for f in opp_fields:
                    ll = f["label"].lower()
                    if "forecast" in ll and "amount" in ll:
                        amt_api   = f["name"]
                        amt_label = f["label"]
                        break
            # Pass 3 — partial API name
            if not amt_api:
                for f in opp_fields:
                    nl = f["name"].lower()
                    if "forecast" in nl and "amount" in nl:
                        amt_api   = f["name"]
                        amt_label = f["label"]
                        break

            # Pass 3b — FIELDS(CUSTOM) probe on a live opp record
            # Bypasses describe FLS; scans for numeric keys with "forecast",
            # "amount", "arr", or "revenue" in the API name.
            if not amt_api:
                _opp_fc_candidate = None
                for _fkw in ("FIELDS(CUSTOM)", "FIELDS(ALL)"):
                    try:
                        _osql = (
                            f"SELECT {_fkw} FROM Opportunity "
                            f"WHERE OwnerId = '{owner_id}' "
                            f"AND IsClosed = false LIMIT 5"
                        )
                        _orecs = _soql_query_all(sid, _osql)
                        if _orecs:
                            for _orec in _orecs:
                                for _k, _v in _orec.items():
                                    if _k == "attributes":
                                        continue
                                    _kl = _k.lower()
                                    if any(t in _kl for t in
                                           ("forecast", "amount", "arr", "revenue")):
                                        if _v is not None and isinstance(_v, (int, float)) and _v != 0:
                                            amt_api   = _k
                                            amt_label = _k
                                            if status_fn:
                                                status_fn(
                                                    f"Opp amount found "
                                                    f"({_fkw}): {_k!r} = {_v}")
                                            break
                                        elif _opp_fc_candidate is None and isinstance(_v, (int, float)):
                                            _opp_fc_candidate = _k
                                if amt_api:
                                    break
                        if amt_api:
                            break
                    except Exception:
                        continue
                # If FIELDS probe found only a null/zero candidate, record it
                if not amt_api and _opp_fc_candidate:
                    amt_api   = _opp_fc_candidate
                    amt_label = _opp_fc_candidate
                    if status_fn:
                        status_fn(
                            f"Opp amount field (null fallback from FIELDS probe): "
                            f"{_opp_fc_candidate!r}"
                        )

            # Pass 4 — SOQL probe: try common custom field API names directly.
            # A clean SOQL run (no exception) confirms the field is accessible,
            # regardless of whether the sample value is null or zero.
            if not amt_api:
                _opp_probe_names = [
                    "Forecast_Amount__c",
                    "ForecastAmount__c",
                    "Forecast_ARR__c",
                    "Expected_ARR__c",
                    "Expected_Amount__c",
                    "Total_Contract_Value__c",
                    "TCV__c",
                    "Deal_Amount__c",
                    "Opportunity_Amount__c",
                ]
                for _cand in _opp_probe_names:
                    try:
                        _osql = (
                            f"SELECT {_cand} FROM Opportunity "
                            f"WHERE OwnerId = '{owner_id}' "
                            f"AND IsClosed = false LIMIT 1"
                        )
                        _orecs = _soql_query_all(sid, _osql)
                        _v_sample = _orecs[0].get(_cand) if _orecs else None
                        amt_api   = _cand
                        amt_label = "Forecast Amount"
                        if status_fn:
                            status_fn(
                                f"Opp amount found (Pass 4 SOQL probe): "
                                f"{_cand!r} (sample: {_v_sample!r})"
                            )
                        break
                    except Exception:
                        pass

            # Pass 5 — fall back to standard Amount field
            if not amt_api and "amount" in by_nm:
                amt_api   = "Amount"
                amt_label = "Forecast Amount"

            if status_fn:
                status_fn(f"Opp amount field: {amt_label!r} → {amt_api!r}")
        except Exception as _e:
            # If describe fails, fall back to standard Amount
            amt_api   = "Amount"
            amt_label = "Forecast Amount"
            if status_fn:
                status_fn(f"Opp describe failed ({_e}); falling back to Amount.")

    # ── Build SOQL ───────────────────────────────────────────────────────────
    extra_field = f", {amt_api}" if amt_api and amt_api != "Amount" else ""
    cq_start, nq_end = _current_and_next_quarter_end()

    # Discover "Date Stamp Sales Working" field — filters to pipeline opps only
    sw_field = _probe_opp_field(
        sid,
        label_fragment="sales working",
        candidates=[
            "Date_Stamp_Sales_Working__c",
            "DateStampSalesWorking__c",
            "Date_Stamp_Working__c",
            "Sales_Working_Date__c",
            "Sales_Working__c",
        ],
        status_fn=status_fn,
    )
    sw_filter = f" AND {sw_field} != null" if sw_field else ""
    if not sw_field and status_fn:
        status_fn("WARNING: 'Date Stamp Sales Working' field not found — "
                  "all open opps in close-date range will be included.")

    soql = (
        f"SELECT Id, Name, AccountId, Account.Name, Type, "
        f"LeadSource, Amount{extra_field}, CloseDate, StageName, "
        f"Owner.Name, OwnerId "
        f"FROM Opportunity "
        f"WHERE OwnerId = '{owner_id}' AND IsClosed = false "
        f"AND CloseDate >= {cq_start} AND CloseDate <= {nq_end}"
        f"{sw_filter}"
    )
    if status_fn:
        status_fn(f"Querying open opportunities for owner {owner_id}…")
    records = _soql_query_all(sid, soql, status_fn=status_fn)

    rows = []
    for rec in records:
        acct  = rec.get("Account") or {}
        owner = rec.get("Owner")   or {}

        # Prefer the discovered/override amount field; fall back to Amount
        raw_amt = rec.get(amt_api) if amt_api else None
        if raw_amt is None:
            raw_amt = rec.get("Amount")
        try:
            amt_str = f"USD {float(raw_amt):,.2f}" if raw_amt is not None else ""
        except (TypeError, ValueError):
            amt_str = str(raw_amt) if raw_amt is not None else ""

        rows.append({
            "ID (18 Char)":        rec.get("Id", ""),
            "Opportunity Name":    rec.get("Name", ""),
            "18 Digit Account ID": rec.get("AccountId", ""),
            "Account Name":        acct.get("Name", ""),
            "Type":                rec.get("Type", "") or "",
            "Lead Source":         rec.get("LeadSource", "") or "",
            "Forecast Amount":     amt_str,
            "Close Date":          str(rec.get("CloseDate") or ""),
            "Stage":               rec.get("StageName", "") or "",
            "Opportunity Owner":   owner.get("Name", ""),
        })

    if status_fn:
        status_fn(f"Done — {len(rows):,} open opportunities loaded.")
    return rows, []


# ── Return-flow: fetch accounts by FY18 Prev Acct Owner tag ──────────────────

def _probe_opp_field(sid: str, label_fragment: str,
                     candidates: list, status_fn=None) -> str | None:
    """
    Discover a custom Opportunity field API name via describe then SOQL probe.
    label_fragment: lowercase substring to match against field labels.
    candidates: ordered list of API name guesses to try as SOQL probes.
    Returns the API name or None.
    """
    import unicodedata

    def _norm(s):
        s = unicodedata.normalize("NFKC", str(s))
        s = re.sub(r"\s+", " ", s)
        return s.lower().strip()

    try:
        url = (f"{SF_INSTANCE}/services/data/{SF_API_VER}"
               f"/sobjects/Opportunity/describe")
        r = requests.get(url, headers=sf_headers(sid), timeout=30)
        r.raise_for_status()
        all_fields = r.json().get("fields", [])
    except Exception:
        all_fields = []

    # Pass A — label contains the fragment
    for f in all_fields:
        if label_fragment in _norm(f["label"]):
            if status_fn:
                status_fn(f"Opp field '{label_fragment}' found via label: {f['name']!r}")
            return f["name"]

    # Pass B — SOQL existence probe (no WHERE filter needed, just confirms field exists)
    for cand in candidates:
        try:
            _soql_query_all(sid, f"SELECT {cand} FROM Opportunity LIMIT 1")
            if status_fn:
                status_fn(f"Opp field '{label_fragment}' found via probe: {cand!r}")
            return cand
        except Exception:
            continue

    if status_fn:
        status_fn(f"Opp field '{label_fragment}' not found — filter will be skipped.")
    return None
    """
    Discover the FY18 Sales Planning field API name without needing an owner_id.
    Uses describe-based passes first, then falls back to SOQL existence probes
    (SELECT <cand> FROM Account LIMIT 1) which confirm accessibility without
    requiring a specific owner to filter on.
    """
    import unicodedata

    def _norm(s):
        s = unicodedata.normalize("NFKC", str(s))
        s = re.sub(r"\s+", " ", s)
        return s.lower().strip()

    try:
        url = (f"{SF_INSTANCE}/services/data/{SF_API_VER}"
               f"/sobjects/Account/describe")
        r = requests.get(url, headers=sf_headers(sid), timeout=30)
        r.raise_for_status()
        all_fields = r.json().get("fields", [])
    except Exception as e:
        if status_fn:
            status_fn(f"Describe failed: {e}")
        all_fields = []

    by_label = {_norm(f["label"]): f for f in all_fields}
    by_name  = {_norm(f["name"]):  f for f in all_fields}

    # Pass A — exact label
    if _norm("fy18 sales planning") in by_label:
        found = by_label[_norm("fy18 sales planning")]["name"]
        if status_fn:
            status_fn(f"FY18 field found (label): {found!r}")
        return found

    # Pass B — known API name candidates
    for nm in ["fy18_sales_planning__c", "fy18salesplanning__c",
               "fy18_salesplanning__c", "fy18_sales_plan__c", "fy18__c"]:
        if nm in by_name:
            found = by_name[nm]["name"]
            if status_fn:
                status_fn(f"FY18 field found (API name): {found!r}")
            return found

    # Pass C — partial label/name match
    for f in all_fields:
        ll = _norm(f["label"])
        nl = _norm(f["name"])
        if ("fy18" in ll and "planning" in ll) or \
           ("fy18" in nl and "planning" in nl):
            if status_fn:
                status_fn(f"FY18 field found (partial match): {f['name']!r}")
            return f["name"]

    # Pass D — SOQL existence probe (no owner filter required)
    for cand in ["FY18_Sales_Planning__c", "FY18SalesPlanning__c",
                 "FY18_SalesPlanning__c", "FY18_Sales_Plan__c", "FY18__c"]:
        try:
            _soql_query_all(sid, f"SELECT {cand} FROM Account LIMIT 1")
            if status_fn:
                status_fn(f"FY18 field found (SOQL probe): {cand!r}")
            return cand
        except Exception:
            continue

    if status_fn:
        status_fn("FY18 field not found after all passes.")
    return None

def fetch_accounts_by_tag(sid: str, tag_string: str,
                          rep_ids: list,
                          manager_name: str = "",
                          arr_field_override: str = "",
                          fy18_field_override: str = "",
                          status_fn=None) -> tuple:
    """
    Fetch Account records whose FY18 Sales Planning field contains
    'Prev Acct Owner: <tag_string>' (case-insensitive).

    rep_ids: list of SFDC User IDs for the reps currently holding accounts.
             Resolved by the caller from the embedded/uploaded roster so no
             SFDC User lookup is needed here.

    Strategy:
      1. Query accounts WHERE OwnerId IN (rep_ids) — OwnerId is indexed, fast.
      2. Filter FY18 content in Python — avoids SOQL LIKE on Long Text Area.

    Returns (list_of_row_dicts, fy18_api_name, warnings).
    """
    warnings_out = []

    if not rep_ids:
        raise ValueError(
            "No rep IDs provided. Make sure the manager name matches entries "
            "in the roster."
        )

    # Discover field API names (use first rep ID to help with FIELDS probe)
    if status_fn:
        status_fn("Discovering ARR and FY18 field API names...")
    arr_label, arr_api, _unused, _unused2 = _discover_account_fields(
        sid, owner_id=rep_ids[0], status_fn=status_fn
    )
    if arr_field_override and arr_field_override.strip():
        arr_api = arr_field_override.strip()

    if fy18_field_override and fy18_field_override.strip():
        fy18_api = fy18_field_override.strip()
    else:
        fy18_api = _probe_fy18_field(sid, status_fn=status_fn)

    if not fy18_api:
        raise ValueError(
            "Could not auto-discover the FY18 Sales Planning field. "
            "Enter its API name in the sidebar 'FY18 field API name' box."
        )

    # Query accounts by OwnerId IN (indexed, no relationship traversal)
    select_parts = [
        "Id", "Name", "OwnerId", "Owner.Name", "Owner.Manager.Name",
        "Type", "Rating", "LastActivityDate",
    ]
    if arr_api:
        select_parts.append(arr_api)
    select_parts.append(fy18_api)

    BATCH = 200
    all_records = []
    for i in range(0, len(rep_ids), BATCH):
        batch   = rep_ids[i : i + BATCH]
        id_list = "', '".join(batch)
        soql    = (
            f"SELECT {', '.join(select_parts)} FROM Account "
            f"WHERE OwnerId IN ('{id_list}')"
        )
        if status_fn:
            status_fn(f"Fetching accounts batch {i//BATCH + 1} "
                      f"({min(i+BATCH, len(rep_ids))}/{len(rep_ids)} reps)...")
        all_records.extend(_soql_query_all(sid, soql, status_fn=status_fn))

    if status_fn:
        status_fn(
            f"Fetched {len(all_records):,} total accounts "
            f"— filtering for 'Prev Acct Owner: {tag_string}'..."
        )
        status_fn(f"FY18 field in use: {fy18_api!r}")
        # Sample a few FY18 values to confirm field is populated
        samples = [str(r.get(fy18_api) or "") for r in all_records[:5]]
        for i, s in enumerate(samples):
            status_fn(f"  Sample {i+1} FY18 value: {s[:120]!r}")

    # Filter FY18 content in Python (safe for Long Text Area fields)
    needle = f"prev acct owner: {tag_string}".lower()
    rows = []
    for rec in all_records:
        fy18_val = str(rec.get(fy18_api) or "")
        if needle not in fy18_val.lower():
            continue
        owner = rec.get("Owner") or {}
        mgr   = owner.get("Manager") or {}
        row   = {
            "18 Digit Account ID": rec.get("Id", ""),
            "Account Name":        rec.get("Name", ""),
            "Account Owner":       owner.get("Name", ""),
            "Account Owner ID":    rec.get("OwnerId", ""),
            "Manager":             mgr.get("Name", ""),
            "Account Type":        rec.get("Type",   "") or "",
            "Rating":              rec.get("Rating", "") or "",
            "Last Activity":       str(rec.get("LastActivityDate") or ""),
        }
        if arr_api:
            val = rec.get(arr_api)
            try:
                row["Contractual ARR (converted)"] = (
                    f"USD {float(val):,.2f}" if val is not None else ""
                )
            except (TypeError, ValueError):
                row["Contractual ARR (converted)"] = str(val) if val is not None else ""
        row["FY18 Sales Planning"] = fy18_val
        rows.append(row)

    if not arr_api:
        warnings_out.append("No ARR field found — ARR will show as blank.")
    if status_fn:
        status_fn(
            f"Done — {len(rows):,} matching accounts "
            f"(from {len(all_records):,} fetched across "
            f"{len(rep_ids)} rep(s) under '{manager_name or 'roster'}')."
        )
    return rows, fy18_api, warnings_out


def fetch_opps_by_account_ids(sid: str, account_ids: list,
                               amount_field_override: str = "",
                               status_fn=None) -> tuple:
    """
    Fetch all open Opportunities whose AccountId is in account_ids.
    Batches into groups of 500 to stay within SOQL IN-clause limits.
    Returns (rows_as_list_of_dicts, warnings).
    """
    if not account_ids:
        return [], []

    # ── Discover amount field (reuse opp describe logic) ────────────────────
    amt_api = amount_field_override.strip() if amount_field_override else None
    if not amt_api:
        try:
            _url = (f"{SF_INSTANCE}/services/data/{SF_API_VER}"
                    f"/sobjects/Opportunity/describe")
            _r = requests.get(_url, headers=sf_headers(sid), timeout=30)
            _r.raise_for_status()
            opp_fields = _r.json().get("fields", [])
            by_lbl = {f["label"].lower(): f for f in opp_fields}
            for lbl in ["forecast amount", "amount"]:
                if lbl in by_lbl:
                    amt_api = by_lbl[lbl]["name"]
                    break
            if not amt_api:
                amt_api = "Amount"
        except Exception:
            amt_api = "Amount"

    extra_field = f", {amt_api}" if amt_api and amt_api != "Amount" else ""
    BATCH = 500
    all_rows = []

    cq_start, nq_end = _current_and_next_quarter_end()

    sw_field = _probe_opp_field(
        sid,
        label_fragment="sales working",
        candidates=[
            "Date_Stamp_Sales_Working__c",
            "DateStampSalesWorking__c",
            "Date_Stamp_Working__c",
            "Sales_Working_Date__c",
            "Sales_Working__c",
        ],
        status_fn=status_fn,
    )
    sw_filter = f" AND {sw_field} != null" if sw_field else ""

    for i in range(0, len(account_ids), BATCH):
        batch   = account_ids[i : i + BATCH]
        id_list = "', '".join(batch)
        soql    = (
            f"SELECT Id, Name, AccountId, Account.Name, Type, "
            f"LeadSource, Amount{extra_field}, CloseDate, "
            f"StageName, Owner.Name, OwnerId "
            f"FROM Opportunity "
            f"WHERE IsClosed = false AND AccountId IN ('{id_list}') "
            f"AND CloseDate >= {cq_start} AND CloseDate <= {nq_end}"
            f"{sw_filter}"
        )
        if status_fn:
            status_fn(f"Fetching opps batch {i//BATCH + 1}…")
        records = _soql_query_all(sid, soql, status_fn=status_fn)

        for rec in records:
            acct  = rec.get("Account") or {}
            owner = rec.get("Owner")   or {}
            raw_amt = rec.get(amt_api) if amt_api else None
            if raw_amt is None:
                raw_amt = rec.get("Amount")
            try:
                amt_str = f"USD {float(raw_amt):,.2f}" if raw_amt is not None else ""
            except (TypeError, ValueError):
                amt_str = str(raw_amt) if raw_amt is not None else ""
            all_rows.append({
                "ID (18 Char)":        rec.get("Id", ""),
                "Opportunity Name":    rec.get("Name", ""),
                "18 Digit Account ID": rec.get("AccountId", ""),
                "Account Name":        acct.get("Name", ""),
                "Type":                rec.get("Type", "") or "",
                "Lead Source":         rec.get("LeadSource", "") or "",
                "Forecast Amount":     amt_str,
                "Close Date":          str(rec.get("CloseDate") or ""),
                "Stage":               rec.get("StageName", "") or "",
                "Opportunity Owner":   owner.get("Name", ""),
            })

    if status_fn:
        status_fn(f"Done — {len(all_rows):,} open opportunities found.")
    return all_rows, []


# ── Analytics API (kept as reference / fallback) ──────────────────────────────

def _describe_report(sid: str, report_id: str) -> dict:
    url = f"{SF_INSTANCE}/services/data/{SF_API_VER}/analytics/reports/{report_id}/describe"
    r   = requests.get(url, headers=sf_headers(sid), timeout=20)
    r.raise_for_status()
    return r.json()

def _run_report_instance(sid: str, report_id: str,
                         metadata_override: dict = None, timeout: int = 120) -> dict:
    """Launch async report instance (no 2000-row cap) and poll until done."""
    url  = f"{SF_INSTANCE}/services/data/{SF_API_VER}/analytics/reports/{report_id}/instances"
    hdrs = sf_headers(sid)
    body = {}
    if metadata_override:
        body["reportMetadata"] = metadata_override

    post = requests.post(url, headers=hdrs, json=body, timeout=30)
    post.raise_for_status()
    instance_id = post.json().get("id")

    inst_url = f"{SF_INSTANCE}/services/data/{SF_API_VER}/analytics/reports/{report_id}/instances/{instance_id}"
    deadline = time.time() + timeout
    while time.time() < deadline:
        poll = requests.get(inst_url, headers=hdrs, timeout=30)
        poll.raise_for_status()
        data   = poll.json()
        status = data.get("attributes", {}).get("status", "")
        if status == "Success":
            return data
        if status in ("Error", "Failed"):
            raise RuntimeError(f"Report instance failed: {data}")
        time.sleep(2)
    raise TimeoutError(f"Report {report_id} timed out after {timeout}s")

def _factmap_to_rows(data: dict) -> list[dict]:
    """Convert Analytics API factMap response into list of flat dicts."""
    meta         = data.get("reportMetadata", {})
    detail_cols  = meta.get("detailColumns", [])
    col_info     = (data.get("reportExtendedMetadata", {})
                        .get("detailColumnInfo", {}))
    # Build label map: API col key → human label
    label_map = {c: col_info.get(c, {}).get("label", c) for c in detail_cols}

    rows = []
    for key, section in data.get("factMap", {}).items():
        if not key.endswith("!T"):
            continue
        for row in section.get("rows", []):
            rec = {}
            for i, cell in enumerate(row.get("dataCells", [])):
                if i < len(detail_cols):
                    col_key = detail_cols[i]
                    label   = label_map.get(col_key, col_key)
                    rec[label] = cell.get("label", "") or cell.get("value", "")
            rows.append(rec)
    return rows

def _post_filter_rows(rows: list[dict], owner_id: str, owner_name: str) -> tuple[list[dict], str]:
    """
    Filter report rows to only those belonging to the departing rep.
    Matches on any column whose label contains 'owner id' (by Salesforce user ID,
    comparing first 15 chars to handle 15/18-char variations) or, failing that,
    any column whose label contains 'owner' (by name, case-insensitive).
    Returns (filtered_rows, method_description).
    """
    if not rows:
        return rows, "none"

    sample = rows[0]
    id_cols   = [k for k in sample if "owner" in k.lower() and "id" in k.lower()]
    name_cols = [k for k in sample if "owner" in k.lower() and "id" not in k.lower()]

    # Prefer ID match (more precise)
    if id_cols and owner_id:
        owner_id_15 = owner_id[:15]
        filtered = [r for r in rows
                    if any(str(r.get(c, ""))[:15] == owner_id_15 for c in id_cols)]
        if filtered:
            return filtered, f"ID match on '{id_cols[0]}'"

    # Fall back to name match
    if name_cols and owner_name:
        owner_lower = owner_name.strip().lower()
        filtered = [r for r in rows
                    if any(str(r.get(c, "")).strip().lower() == owner_lower
                           for c in name_cols)]
        if filtered:
            return filtered, f"name match on '{name_cols[0]}'"

    return rows, "unfiltered"  # couldn't identify owner column — return as-is


def fetch_report_with_owner(sid: str, report_id: str,
                            owner_id: str = None,
                            owner_name: str = None,
                            status_fn=None) -> tuple[list, list]:
    """
    Fetch a SFDC Analytics report and return only the rows belonging to the
    specified owner.

    The owner filter is applied entirely in Python after fetching the full report
    (the report's own saved filters — e.g. division scoping — are preserved).
    API-level owner filter injection has been removed: it was unreliable due to
    internal column key name variations across Salesforce org configurations.

    Python post-filter priority:
      1. Match on any column whose label contains 'owner' and 'id'  (15/18-char prefix)
      2. Match on any column whose label contains 'owner' but not 'id' (exact name match)
      3. No match → return all rows with a warning

    Returns (rows, warnings).
    """
    if status_fn:
        status_fn(f"Describing report {report_id}…")
    desc     = _describe_report(sid, report_id)
    meta     = desc.get("reportMetadata", {})
    warnings = []

    if status_fn:
        status_fn("Fetching full report (owner filter applied in Python)…")
    data = _run_report_instance(sid, report_id, metadata_override=meta)
    rows = _factmap_to_rows(data)
    if status_fn:
        status_fn(f"Report returned {len(rows)} rows. Filtering to owner…")

    if rows and (owner_id or owner_name):
        rows, method = _post_filter_rows(rows, owner_id or "", owner_name or "")
        if status_fn:
            status_fn(f"After owner filter: {len(rows)} rows ({method}).")
        if method == "unfiltered":
            warnings.append(
                "No owner column found in the report — results include all owners. "
                "Verify before distributing."
            )

    if status_fn:
        status_fn(f"Report complete — {len(rows)} rows loaded.")
    return rows, warnings

def parse_uploaded_file(uploaded) -> pd.DataFrame:
    """Parse a Streamlit UploadedFile (XLS, XLSX, or CSV) into a DataFrame."""
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        return pd.read_csv(uploaded, dtype=str).fillna("")
    # Try openpyxl first (xlsx), fall back to html (xls)
    try:
        return pd.read_excel(uploaded, dtype=str).fillna("")
    except Exception:
        return pd.read_html(uploaded.read())[0].fillna("").astype(str)

# ─────────────────────────────────────────────────────────────────────────────
# ROSTER HELPERS
# ─────────────────────────────────────────────────────────────────────────────
# Embedded roster — Global SMB Roster (Sep 23, 2026), 271 reps
# Format: (rep_id, rep_name, manager, division, region, team)
_ROSTER_DATA = [
    ('005Pg00000BCLTlIAP', 'Kylie Barrett', 'Samantha Young', 'US SMB Client Sales', 'US SMB CSE Strategic Southeast', 'US SMB Client Sales Strategic'),
    ('005Pg00000BQhHVIA1', 'Mary-Clare Dizon', 'Jack Schwartz', 'US Mid Market', 'US MM Mid Atlantic', 'US MM Mid Atlantic'),
    ('005Pg00000BrsjtIAB', 'Bailey Pesta', 'Eric Laliberte', 'US General Business', 'US GB Northeast', 'US GB East'),
    ('005Pg00000CDzy1IAD', 'Pao Thao', 'Caroline Kristek', 'US General Business', 'US GB DC Metro', 'US GB East'),
    ('005Pg00000CE0CXIA1', 'Katherine Oliver', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('005Pg00000Cu5irIAB', 'Wesley Bailey', 'Sarah Murray', 'US General Business', 'US GB South', 'US GB West'),
    ('005Pg00000D4jofIAB', 'Callum Shearer', 'Hamish Tebbutt', 'Australia Mid Market', 'Australia Mid Market', 'Australia Mid Market'),
    ('005Pg00000D4jS5IAJ', 'Kurt Iske', 'Justin Simon', 'US General Business', 'US GB Mid Atlantic', 'US GB West'),
    ('005Pg00000DwTzqIAF', 'David Jensen', 'Randi Kruger', 'US SMB Client Sales', 'US SMB CSE Key South', 'US SMB Client Sales Key'),
    ('005Pg00000E5dIQIAZ', 'Michael Grant', 'Danielle Walter', 'US General Business', 'US GB Southeast', 'US GB East'),
    ('005Pg00000FC7ZlIAL', 'Tyler Nuquay', 'Samantha Young', 'US SMB Client Sales', 'US SMB CSE Strategic Southeast', 'US SMB Client Sales Strategic'),
    ('005Pg00000FunJxIAJ', 'Karrah Manzanarez', 'Jake Rutenbar', 'US SMB Client Sales', 'US SMB CSE Key Mid Atlantic', 'US SMB Client Sales Key'),
    ('005Pg00000GwTrnIAF', 'Gabe Peizner', 'Philip Valle', 'US General Business', 'US GB Central North', 'US GB West'),
    ('005Pg00000HIPz3IAH', 'Alexis Null', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('005Pg00000Hmd8XIAR', 'Luke DeGrammont', 'Eric Laliberte', 'US General Business', 'US GB Northeast', 'US GB East'),
    ('005Pg00000IEYYvIAP', 'Jason Peterson', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('005Pg00000IH5fFIAT', 'Evan Smith', 'Megan Frodge', 'US SMB Client Sales', 'US SMB CSE Strategic Northeast', 'US SMB Client Sales Strategic'),
    ('005Pg00000IhCtBIAV', 'Staci Borowsky', 'Sarah Murray', 'US General Business', 'US GB South', 'US GB West'),
    ('005Pg00000Ihi6gIAB', 'Mercedes Nolte', 'Caroline Kristek', 'US General Business', 'US GB Central North', 'US GB West'),
    ('005Pg00000IMcHVIA1', 'Ben Drayton', 'Danielle Walter', 'US General Business', 'US GB Southeast', 'US GB East'),
    ('005Pg00000IujzQIAR', 'Cody Koeplin', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('005Pg00000IXctlIAD', 'Christian Conover', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('005Pg00000JdcvJIAR', 'Arshia Nazem', 'Lesley Nunes', 'Canada Mid Market', 'Canada MM East', 'Canada MM East'),
    ('005Pg00000KMidBIAT', 'Tom Evans', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('005Pg00000LBthpIAD', 'Maria Adragna', 'Sarah Murray', 'US General Business', 'US GB South', 'US GB West'),
    ('005Pg00000MR2AjIAL', 'Logan Holland', 'Caroline Kristek', 'US General Business', 'US GB DC Metro', 'US GB East'),
    ('005Pg00000Nbz45IAB', 'Kelsey Fredrickson', 'Jake Rutenbar', 'US SMB Client Sales', 'US SMB CSE Key Mid Atlantic', 'US SMB Client Sales Key'),
    ('005Pg00000ObrNhIAJ', 'Michael Monello', 'Blake Karnes', 'US SMB Client Sales', 'US SMB CSE Strategic Northwest', 'US SMB Client Sales Strategic'),
    ('005Pg00000P5ttVIAR', 'Austin Aghamirzai', 'Heather Lewis', 'US SMB Client Sales', 'US SMB CSE Key West', 'US SMB Client Sales Key'),
    ('005Pg00000PDl2jIAD', 'Brianna Basolo', 'Marissa Mock', 'US SMB Client Sales', 'US SMB CSE Key Great Lakes', 'US SMB Client Sales Key'),
    ('005Pg00000PMiqPIAT', 'Teylen Sheesley', 'Ronit Cohn', 'US SMB Client Sales', 'US SMB CSE Key North', 'US SMB Client Sales Key'),
    ('005Pg00000QSHi6IAH', 'Mackenzie Bowen', 'Jake Rutenbar', 'US SMB Client Sales', 'US SMB CSE Key Mid Atlantic', 'US SMB Client Sales Key'),
    ('005Pg00000R4GcvIAF', 'Tyler Witt', 'Sara Hurst', 'Canada SMB Client Sales', 'Canada SMB Client Sales Key', 'Canada SMB Client Sales Key'),
    ('005Pg00000RjfkDIAR', 'Calvin Nisban', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('005Pg00000Rx6mbIAB', 'Libby Hartnagel', 'Marissa Mock', 'US SMB Client Sales', 'US SMB CSE Key Great Lakes', 'US SMB Client Sales Key'),
    ('005Pg00000SDPrhIAH', 'Jacob Nickoloff', 'Marissa Mock', 'US SMB Client Sales', 'US SMB CSE Key Great Lakes', 'US SMB Client Sales Key'),
    ('005Pg00000SDSj7IAH', 'James Hirst', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('005Pg00000Sqsk9IAB', 'Margaret Hill', 'Mathees Karuna', 'Canada Mid Market', 'Canada MM West', 'Canada MM West'),
    ('005Pg00000Uhl0nIAB', 'Megan Menzuber', 'Justin Simon', 'US General Business', 'US GB Mid Atlantic', 'US GB West'),
    ('005Pg00000UnAD3IAN', 'Belen DeLuca', 'Brandon Schick', 'US National', 'US National Northeast', 'US National East'),
    ('005Pg00000UwEm5IAF', 'Nicole Breitenstein', 'Dave Elinger', 'US SMB Client Sales', 'US SMB CSE Strategic Southwest', 'US SMB Client Sales Strategic'),
    ('005Pg00000VzQq9IAF', 'Joe Ruddy', 'Nick Henney', 'UK Ireland', 'UK Ireland', 'UK Ireland'),
    ('005Pg00000WcEfTIAV', 'Manan Taneja', 'Hamish Tebbutt', 'Australia Mid Market', 'Australia Mid Market', 'Australia Mid Market'),
    ('005Pg00000WD54HIAT', 'Andrea Brown', 'Justin Simon', 'US General Business', 'US GB Mid Atlantic', 'US GB West'),
    ('005Pg00000WhLPAIA3', 'Courtney Melvin', 'Brandon Schick', 'US National', 'US National', 'US National'),
    ('005Pg00000Z2TOLIA3', 'Tom Osterberg', 'Christian Larson', 'US SMB Client Sales', 'US SMB CSE Strategic Midwest', 'US SMB Client Sales Strategic'),
    ('005Pg000008gsSPIAY', 'Elena Krischunas', 'Caroline Kristek', 'US General Business', 'US GB DC Metro', 'US GB East'),
    ('005Pg000008L0moIAC', 'Lily Shaw', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('005Pg000008sUu9IAE', 'Ryan Galvin', 'Philip Valle', 'US General Business', 'US GB Northwest', 'US GB West'),
    ('005Pg000008tZxHIAU', 'Megan Kirschenman', 'Danielle Walter', 'US General Business', 'US GB Southeast', 'US GB East'),
    ('005Pg0000094nbFIAQ', 'Sebastean Gonzalez-Johnson', 'Justin Simon', 'US General Business', 'US GB Mid Atlantic', 'US GB West'),
    ('0050e000006ACm1AAG', 'Ben Chandiram', 'Ryan Headington', 'UK National', 'UK Premier', 'UK Premier'),
    ('0050e000006Afl7AAC', 'Michael Torres', 'Holly Corser', 'US National', 'US National West', 'US National West'),
    ('0050e000006AXdtAAG', 'Julie Barter', 'Angie Koplan', 'US SMB Client Sales', 'US SMB CSE Premier West', 'US SMB Client Sales Premier'),
    ('0050e000006eyPUAAY', 'Ben Angelo', 'Jake Rutenbar', 'US SMB Client Sales', 'US SMB CSE Key Mid Atlantic', 'US SMB Client Sales Key'),
    ('0050e000006iRHfAAM', 'Joe Bellefeuille', 'Heather Lewis', 'US SMB Client Sales', 'US SMB CSE Key West', 'US SMB Client Sales Key'),
    ('0050e000006jfCWAAY', 'Breck Hansen', 'Kyle Loving', 'US SMB Client Sales', 'US SMB CSE Premier East', 'US SMB Client Sales Premier'),
    ('0050e000006k7r3AAA', 'Emily Shillock', 'Kyle Olson', 'US National', 'US National Northeast', 'US National East'),
    ('0050e000006L5jvAAC', 'Alan Donohoe', 'Bret Edis', 'UK SMB Client Sales', 'UK SMB Client Sales Premier', 'UK SMB Client Sales Premier'),
    ('0050e000006PhM3AAK', 'Montana Knell', 'Philip Valle', 'US General Business', 'US GB Northwest', 'US GB West'),
    ('0050e000006PkHzAAK', 'Allison Logan', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('0050e000006PkRpAAK', 'Deb Smith', 'Kyle Loving', 'US SMB Client Sales', 'US SMB CSE Premier East', 'US SMB Client Sales Premier'),
    ('0050e000006qcOsAAI', 'Garrett Schultz', 'Philip Valle', 'US General Business', 'US GB Northwest', 'US GB West'),
    ('0050e000006qcRwAAI', 'Deanna Rota', 'Sara Hurst', 'Canada SMB Client Sales', 'Canada SMB Client Sales Key', 'Canada SMB Client Sales Key'),
    ('0050e000006qJSqAAM', 'Lauren Pellowski', 'Dave Elinger', 'US SMB Client Sales', 'US SMB CSE Strategic Southwest', 'US SMB Client Sales Strategic'),
    ('0050e000006r1x9AAA', 'SMN Digital Commerce', '', 'US SMB Client Sales', '', ''),
    ('0050e000006rj2qAAA', 'Peter Mignin', 'Holly Corser', 'US National', 'US National West', 'US National West'),
    ('0050e000006vZFOAA2', 'Ryan Reese', 'Lauren Erickson', 'US General Business', 'US GB Central West', 'US GB West'),
    ('0050e000006x9OdAAI', 'Tyler Hazen', 'Samantha Young', 'US SMB Client Sales', 'US SMB CSE Strategic Southeast', 'US SMB Client Sales Strategic'),
    ('0050e000006YD7RAAW', 'Adrian Antony', 'Nick Henney', 'Netherlands SMB', 'Netherlands SMB', 'Netherlands SMB'),
    ('0050e000006Yp6MAAS', 'Jennifer Gernand', 'Eric Laliberte', 'US General Business', 'US GB Northeast', 'US GB East'),
    ('0050e000006Yp41AAC', 'Alyssa Wichman', 'Sarah Murray', 'US General Business', 'US GB South', 'US GB West'),
    ('0050e000006YpbZAAS', 'Ryan Doyle', 'Ronit Cohn', 'US SMB Client Sales', 'US SMB CSE Key North', 'US SMB Client Sales Key'),
    ('0050e000006YpchAAC', 'Joseph Silva', 'Samantha Young', 'US SMB Client Sales', 'US SMB CSE Strategic Southeast', 'US SMB Client Sales Strategic'),
    ('0050e000006yQwPAAU', 'Akeiro Lloyd', 'Philip Valle', 'US General Business', 'US GB Northwest', 'US GB West'),
    ('0050e000006YxyCAAS', 'Jun Park', 'Megan Frodge', 'US SMB Client Sales', 'US SMB CSE Strategic Northeast', 'US SMB Client Sales Strategic'),
    ('0050e000006yz1nAAA', 'Emily Norris', 'Megan Frodge', 'US SMB Client Sales', 'US SMB CSE Strategic Northeast', 'US SMB Client Sales Strategic'),
    ('0050e000006Z2BtAAK', 'Ben Goman', 'Chris Smith', 'US SMB Client Sales', 'US SMB CSE Premier North', 'US SMB Client Sales Premier'),
    ('0050e000007A1sCAAS', 'Michelle Starrett', 'Jack Schwartz', 'US National', 'US National Southeast', 'US National East'),
    ('0050e000007b5MdAAI', 'Alex Capeloto', 'Angie Koplan', 'US SMB Client Sales', 'US SMB CSE Premier West', 'US SMB Client Sales Premier'),
    ('0050e000007bFJ0AAM', 'Michael Shedd', 'Lauren Erickson', 'US General Business', 'US GB Central West', 'US GB West'),
    ('0050e000007eBInAAM', 'Abbie Lewis', 'Katie Brown', 'UK SMB Client Sales', 'UK SMB Client Sales Key I', 'UK SMB Client Sales Key'),
    ('0050e000007MdDBAA0', 'James Hooker', 'Ryan Headington', 'UK National', 'UK National', 'UK National'),
    ('0050e000007nl6qAAA', 'Jeffery Smith', 'Caroline Kristek', 'US General Business', 'US GB DC Metro', 'US GB East'),
    ('0050e000007o8FgAAI', 'Natalie Wahlers', 'Holly Corser', 'US National', 'US National West', 'US National West'),
    ('0050e000007o8JTAAY', 'Michael Madden', 'Holly Corser', 'US National', 'US National West', 'US National West'),
    ('0050e000007obCRAAY', 'Hannah White', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('0050e000007olKpAAI', 'Mike Antkowiak', 'Ronit Cohn', 'US SMB Client Sales', 'US SMB CSE Key North', 'US SMB Client Sales Key'),
    ('0050e000007olokAAA', 'Nick Bright', 'Nick Henney', 'UK Ireland', 'UK Ireland', 'UK Ireland'),
    ('0050e000007olUuAAI', 'James Pribble', 'Eric Laliberte', 'US General Business', 'US GB Northeast', 'US GB East'),
    ("0050e000007oSbDAAU", "Samantha O'Connell", 'Dave Elinger', 'US SMB Client Sales', 'US SMB CSE Strategic Southwest', 'US SMB Client Sales Strategic'),
    ('0050e000007oSffAAE', 'Laura Jungbauer', 'Christian Larson', 'US SMB Client Sales', 'US SMB CSE Strategic Midwest', 'US SMB Client Sales Strategic'),
    ('0050e000007oSgcAAE', 'Melissa Deutsch', 'Philip Valle', 'US General Business', 'US GB Northwest', 'US GB West'),
    ('0050e000007oSl3AAE', 'Amanda Player', 'Peter Soukos', 'Australia SMB Client Sales', 'Australia SMB Client Sales Premier', 'Australia SMB Client Sales Premier'),
    ('0050e000007oSobAAE', 'Bryan Baker', 'Lauren Erickson', 'US General Business', 'US GB Central North', 'US GB West'),
    ('0050e000007oStRAAU', 'John Hargreaves', 'Mathees Karuna', 'Canada National', 'Canada National West', 'Canada National West'),
    ('0050e000007oSwLAAU', 'Aghiles Benali', 'Kyle Loving', 'US SMB Client Sales', 'US SMB CSE Premier East', 'US SMB Client Sales Premier'),
    ('0050e000007oSwpAAE', 'Lucy Collins', 'Katie Brown', 'UK SMB Client Sales', 'UK SMB Client Sales Key I', 'UK SMB Client Sales Key'),
    ('0050e000007oT5IAAU', 'Chelsea Salonek', 'Sara Hurst', 'Canada SMB Client Sales', 'Canada SMB Client Sales Premier', 'Canada SMB Client Sales Premier'),
    ('0050e000007OtnYAAS', 'Peter Axon', 'Hamish Tebbutt', 'Australia Mid Market', 'Australia Mid Market', 'Australia Mid Market'),
    ('0050e000007ovW7AAI', 'Zack Smith', 'Brooke Nelson', 'US SMB Client Sales', 'US SMB CSE Strategic Northwest', 'US SMB Client Sales Strategic'),
    ('0050e000007P4l4AAC', 'Bridget Sands', 'Amber Peters', 'US National', 'US National Midwest', 'US National West'),
    ('0050e000007pdECAAY', 'Cristian Kawa', 'Sara Hurst', 'Canada SMB Client Sales', 'Canada SMB Client Sales Strategic', 'Canada SMB Client Sales Strategic'),
    ('0050e000007Qb2xAAC', 'Adam Sala', 'Angie Koplan', 'US SMB Client Sales', 'US SMB CSE Premier West', 'US SMB Client Sales Premier'),
    ('0050e000007qBXIAA2', 'Austin Parent', 'Danielle Walter', 'US General Business', 'US GB Southeast', 'US GB East'),
    ("0050e000007QgEFAA0", "Glen O'Brien", 'Ryan Headington', 'UK National', 'UK National', 'UK National'),
    ('0050e000007qoZwAAI', 'Meaghan Rodgers', 'Brooke Nelson', 'US SMB Client Sales', 'US SMB CSE Premier South', 'US SMB Client Sales Premier'),
    ('0050e000007qwOhAAI', 'Matt Knight', 'Megan Frodge', 'US SMB Client Sales', 'US SMB CSE Strategic Northeast', 'US SMB Client Sales Strategic'),
    ('0050e000007RAOeAAO', 'Jeffrey Danner', 'Christian Larson', 'US SMB Client Sales', 'US SMB CSE Strategic Midwest', 'US SMB Client Sales Strategic'),
    ('0050e000007SvxEAAS', 'Alexandra Parritt', 'Katie Brown', 'UK SMB Client Sales', 'UK SMB Client Sales Key I', 'UK SMB Client Sales Key'),
    ('0050e000007TLA3AAO', 'Brian Foster', 'Philip Valle', 'US General Business', 'US GB Northwest', 'US GB West'),
    ('0050e000008fyiMAAQ', 'Brad Holder', 'Sara Hurst', 'Canada SMB Client Sales', 'Canada SMB Client Sales Premier', 'Canada SMB Client Sales Premier'),
    ('0050e000008gvq0AAA', 'Olivia Allen', 'George Gregory', 'UK Ireland', 'UK Ireland', 'UK Ireland'),
    ('0050e000008gXTWAA2', 'Jessica Klein', 'Megan Frodge', 'US SMB Client Sales', 'US SMB CSE Strategic Northeast', 'US SMB Client Sales Strategic'),
    ('0050e000008hbtfAAA', 'Kate Meller', 'Hamish Tebbutt', 'Australia Mid Market', 'Australia Mid Market', 'Australia Mid Market'),
    ('0050e000008hDTYAA2', 'Kyle Flaherty', 'Kyle Olson', 'US National', 'US National Northeast', 'US National East'),
    ('0050e000008hkYDAAY', 'Dan Eagen', 'Christian Larson', 'US SMB Client Sales', 'US SMB CSE Strategic Midwest', 'US SMB Client Sales Strategic'),
    ('0050e000008HvISAA0', 'Scott Groshong', 'Kyle Olson', 'US National', 'US National Northeast', 'US National East'),
    ('0050e000008hWjkAAE', 'Steve Kavanagh', 'Peter Soukos', 'Australia SMB Client Sales', 'Australia SMB Client Sales Strategic', 'Australia SMB Client Sales Strategic'),
    ('0050e000008IdzLAAS', 'Lindsay Witt', 'Mathees Karuna', 'Canada Mid Market', 'Canada MM West', 'Canada MM West'),
    ('0050e000008IlPWAA0', 'Loreena Maguet', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('0050e000008J22iAAC', 'Hunter Hood', 'Amber Peters', 'US National', 'US National Midwest', 'US National West'),
    ('0050e000008JHkMAAW', 'Travis Jenks', 'Hamish Tebbutt', 'Australia Mid Market', 'Australia Mid Market', 'Australia Mid Market'),
    ('0050e0000070ozWAAQ', 'Eric Slezak', 'Amber Peters', 'US National', '', 'US National West'),
    ('0050e0000070PpvAAE', 'Thang Nguyen', 'Jake Rutenbar', 'US SMB Client Sales', 'US SMB CSE Key Mid Atlantic', 'US SMB Client Sales Key'),
    ('0050e0000077dNLAAY', 'Megan Barber', 'Amber Peters', 'US National', 'US National Midwest', 'US National West'),
    ('0050e0000077RXJAA2', 'Tyler Krob', 'Marissa Mock', 'US SMB Client Sales', 'US SMB CSE Key Great Lakes', 'US SMB Client Sales Key'),
    ('0050e0000078dYpAAI', 'Christina Monardo', 'Lesley Nunes', 'Canada Mid Market', 'Canada MM East', 'Canada MM East'),
    ('0050e0000078FkwAAE', 'John Krolicki', 'Lauren Erickson', 'US General Business', 'US GB Central West', 'US GB West'),
    ('0050e0000078FprAAE', 'Zack Scharf', 'Dave Elinger', 'US SMB Client Sales', 'US SMB CSE Strategic Southwest', 'US SMB Client Sales Strategic'),
    ('0050e0000078Fu3AAE', 'Jose Gamboa', 'Eric Laliberte', 'US General Business', 'US GB Central North', 'US GB West'),
    ('0050e0000078G6nAAE', 'Blair Sievert', 'Kyle Loving', 'US SMB Client Sales', 'US SMB CSE Premier East', 'US SMB Client Sales Premier'),
    ('0050e0000078GblAAE', 'Hung Do', 'Peter Soukos', 'Australia SMB Client Sales', 'Australia SMB Client Sales Strategic', 'Australia SMB Client Sales Strategic'),
    ('0050e0000078GC2AAM', 'Izabella Krawczyk Patel', 'Bret Edis', 'UK SMB Client Sales', 'UK SMB Client Sales Strategic', 'UK SMB Client Sales Strategic'),
    ('0050e0000078GcPAAU', 'Kate Hulmston', 'Peter Soukos', 'Australia SMB Client Sales', 'Australia SMB Client Sales Premier', 'Australia SMB Client Sales Premier'),
    ('0050e0000078GlRAAU', 'Megan Lucas', 'Lesley Nunes', 'Canada Mid Market', 'Canada MM East', 'Canada MM East'),
    ('0050e0000078GWvAAM', 'Karl Perkins', 'Bret Edis', 'UK SMB Client Sales', 'UK SMB Client Sales Strategic', 'UK SMB Client Sales Strategic'),
    ('0050e0000078MDVAA2', 'Alastair King', 'Ryan Headington', 'UK National', 'UK National', 'UK National'),
    ('0050e0000078Z5PAAU', 'Jordan Buri', 'Samantha Young', 'US SMB Client Sales', 'US SMB CSE Strategic Southeast', 'US SMB Client Sales Strategic'),
    ('0050e0000078Z6mAAE', 'Seth Johnson', 'Amber Peters', 'US National', 'US National Midwest', 'US National West'),
    ('0050e0000078zaxAAA', 'Caroline Morcom', 'Jack Schwartz', 'US National', 'US National Southeast', 'US National East'),
    ('0050e0000078Zf8AAE', 'Carol Murray', 'Sara Hurst', 'Canada SMB Client Sales', 'Canada SMB Client Sales Strategic', 'Canada SMB Client Sales Strategic'),
    ('0050e0000079k4yAAA', 'Zach Pesta', 'Justin Simon', 'US General Business', 'US GB Mid Atlantic', 'US GB West'),
    ('0050e0000079kDlAAI', 'Cory Eul', 'Amber Peters', 'US National', 'US National Midwest', 'US National West'),
    ('0050e0000086Fd2AAE', 'Bryn Cowling', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('0057V00000ASympQAD', 'Trevor Hecht', 'Samantha Young', 'US SMB Client Sales', 'US SMB CSE Strategic Southeast', 'US SMB Client Sales Strategic'),
    ('0057V00000AvIlTQAV', 'Jenny Sullivan', 'Danielle Walter', 'US General Business', 'US GB Southeast', 'US GB East'),
    ('0057V00000AZ8LHQA1', 'Lisanne Lemay', 'Lesley Nunes', 'Canada Mid Market', 'Canada MM East', 'Canada MM East'),
    ('0057V00000B57XbQAJ', 'Jack Zabel', 'Randi Kruger', 'US SMB Client Sales', 'US SMB CSE Key South', 'US SMB Client Sales Key'),
    ('0057V00000B57YZQAZ', 'Scott Bere', 'Randi Kruger', 'US SMB Client Sales', 'US SMB CSE Key South', 'US SMB Client Sales Key'),
    ('0057V00000BfS89QAF', 'Jonathan Barth', 'Ronit Cohn', 'US SMB Client Sales', 'US SMB CSE Key North', 'US SMB Client Sales Key'),
    ('0057V00000BfSJqQAN', 'Christopher Spencer', 'Marissa Mock', 'US SMB Client Sales', 'US SMB CSE Key Great Lakes', 'US SMB Client Sales Key'),
    ('0057V00000Bnu8cQAB', 'Katherine Boldt', 'Caroline Kristek', 'US General Business', 'US GB DC Metro', 'US GB East'),
    ('0057V00000C8evuQAB', 'Valerie Friedman', 'Sarah Murray', 'US General Business', 'US GB South', 'US GB West'),
    ('0057V00000C84mhQAB', 'Lydia Morrell', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('0057V00000CTWRPQA5', 'Vincent Costa', 'Sarah Murray', 'US General Business', 'US GB South', 'US GB West'),
    ('0057V00000CTWS3QAP', 'Paul Crampton', 'Sarah Murray', 'US General Business', 'US GB Central North', 'US GB West'),
    ('0057V00000CUwSNQA1', 'David Hicks', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('0057V00000CWed0QAD', 'Sierra Nardella', 'Lesley Nunes', 'Canada Mid Market', 'Canada MM East', 'Canada MM East'),
    ('0057V00000CWfsHQAT', 'Darcy Penman', 'Hamish Tebbutt', 'Australia Mid Market', 'Australia Mid Market', 'Australia Mid Market'),
    ('0057V00000CX6ksQAD', 'Caleb Boateng Bekyir', 'Nick Henney', 'Netherlands SMB', 'Netherlands SMB', 'Netherlands SMB'),
    ('0057V00000CX6ONQA1', 'Richard Vines', 'Ryan Headington', 'UK National', 'UK National', 'UK National'),
    ('0057V00000D0gTlQAJ', 'Giel De Muyt', 'Nick Henney', 'Netherlands SMB', 'Netherlands SMB', 'Netherlands SMB'),
    ('0057V000008hmPlQAI', 'Jack Morris', 'Ryan Headington', 'UK National', 'UK Premier', 'UK Premier'),
    ('0057V000008hq3yQAA', 'Lindsay Wilson', 'Brooke Nelson', 'US SMB Client Sales', 'US SMB CSE Premier South', 'US SMB Client Sales Premier'),
    ('0057V000008hwwfQAA', 'Demi Crossman', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('0057V000008iGMDQA2', 'Anna Christofaro', 'Christian Larson', 'US SMB Client Sales', 'US SMB CSE Strategic Midwest', 'US SMB Client Sales Strategic'),
    ('0057V000008iHOUQA2', 'Mark Hemmerle', 'Heather Lewis', 'US SMB Client Sales', 'US SMB CSE Key West', 'US SMB Client Sales Key'),
    ('0057V000008iME0QAM', 'Lindsay Paxton', 'Megan Frodge', 'US SMB Client Sales', 'US SMB CSE Strategic Northeast', 'US SMB Client Sales Strategic'),
    ('0057V000008QGIjQAO', 'Debbie Saysanavongphet', 'Blake Karnes', 'US SMB Client Sales', 'US SMB CSE Strategic Northwest', 'US SMB Client Sales Strategic'),
    ('0057V000008QVr5QAG', 'Joseph Zangel', 'Dave Elinger', 'US SMB Client Sales', 'US SMB CSE Strategic Southwest', 'US SMB Client Sales Strategic'),
    ('0057V000008RgcIQAS', 'Colin Kraker', 'Dave Elinger', 'US SMB Client Sales', 'US SMB CSE Strategic Southwest', 'US SMB Client Sales Strategic'),
    ('0057V000008RgeiQAC', 'Rachel Meyer', 'Randi Kruger', 'US SMB Client Sales', 'US SMB CSE Key South', 'US SMB Client Sales Key'),
    ('0057V000008RJ5zQAG', 'Jake Jenkins', 'Jessica Brown', 'UK Mid Market', 'UK Mid Market', 'UK Mid Market'),
    ('0057V000008Rm4XQAS', 'Alex Gormley', 'Kyle Olson', 'US National', 'US National Northeast', 'US National East'),
    ('0057V000008UgfXQAS', 'Alden Martinez', 'Ronit Cohn', 'US SMB Client Sales', 'US SMB CSE Key North', 'US SMB Client Sales Key'),
    ('0057V000008UkHZQA0', 'Natalie Rizk', 'Ronit Cohn', 'US SMB Client Sales', 'US SMB CSE Key North', 'US SMB Client Sales Key'),
    ('0057V000008V1DAQA0', 'Patrick Costello', 'Kyle Olson', 'US National', 'US National Northeast', 'US National East'),
    ('0057V000008VgOJQA0', 'Alexis Johnson', 'Lauren Erickson', 'US General Business', 'US GB Central West', 'US GB West'),
    ('0057V000008VIjhQAG', 'Meagan Sims', 'Kyle Olson', 'US National', 'US National Northeast', 'US National East'),
    ('0057V000008Vll3QAC', 'Chelsey Rosemann', 'Sarah Murray', 'US General Business', 'US GB South', 'US GB West'),
    ('0057V000008VrymQAC', 'Joe Dorey', 'Amanda Meek', 'US SMB Client Sales', 'US SMB CSE Strategic Mountain', 'US SMB Client Sales Strategic'),
    ('0057V000008VYA4QAO', 'Tom Wahl', 'Amanda Meek', 'US SMB Client Sales', 'US SMB CSE Strategic Mountain', 'US SMB Client Sales Strategic'),
    ('0057V000008W2pbQAC', 'Cole Hettick', 'Eric Laliberte', 'US General Business', 'US GB Northeast', 'US GB East'),
    ('0057V000008W9eLQAS', 'Andrea Flor', 'Amanda Meek', 'US SMB Client Sales', 'US SMB CSE Strategic Mountain', 'US SMB Client Sales Strategic'),
    ('0057V000008WtKUQA0', 'Cody Thelen', 'Danielle Walter', 'US General Business', 'US GB Southeast', 'US GB East'),
    ('0057V000008WtN1QAK', 'Alanna Parrott', 'Brooke Nelson', 'US SMB Client Sales', 'US SMB CSE Premier South', 'US SMB Client Sales Premier'),
    ('0057V000008WtOOQA0', 'Brooke Corcoran', 'Lauren Erickson', 'US General Business', 'US GB Central West', 'US GB West'),
    ('0057V000008XRzsQAG', 'Ryan Hale', 'Bret Edis', 'UK SMB Client Sales', 'UK SMB Client Sales Strategic', 'UK SMB Client Sales Strategic'),
    ('0057V000008XS0RQAW', 'Ashley McCue', 'Chris Smith', 'US SMB Client Sales', 'US SMB CSE Premier North', 'US SMB Client Sales Premier'),
    ('0057V000008XS1FQAW', 'Mara Obermeier', 'Amanda Meek', 'US SMB Client Sales', 'US SMB CSE Strategic Mountain', 'US SMB Client Sales Strategic'),
    ('0057V000009cmzYQAQ', 'Michaela Gormley', 'Heather Lewis', 'US SMB Client Sales', 'US SMB CSE Key West', 'US SMB Client Sales Key'),
    ('0057V000009e2B7QAI', 'Sara Bustad', 'Caroline Kristek', 'US General Business', 'US GB DC Metro', 'US GB East'),
    ('0057V000009fCq2QAE', 'Claire van der Vegt', 'Peter Soukos', 'Australia SMB Client Sales', 'Australia SMB Client Sales Strategic', 'Australia SMB Client Sales Strategic'),
    ('0057V000009fMovQAE', 'Christian Palutis', 'Philip Valle', 'US General Business', 'US GB Northwest', 'US GB West'),
    ('0057V000009fMtpQAE', 'Anthony Williams', 'Lauren Erickson', 'US General Business', 'US GB Central West', 'US GB West'),
    ('0057V000009fXzjQAE', 'Joe Vigil', 'Heather Lewis', 'US SMB Client Sales', 'US SMB CSE Key West', 'US SMB Client Sales Key'),
    ('0057V000009gD8SQAU', 'Sachin Wilde', 'Ryan Headington', 'UK National', 'UK Premier', 'UK Premier'),
    ('0057V000009HzTTQA0', 'Josh Crippen', 'Eric Laliberte', 'US General Business', 'US GB Northeast', 'US GB East'),
    ('0057V000009IGOpQAO', 'Alexander Spanton', 'Danielle Walter', 'US General Business', 'US GB Southeast', 'US GB East'),
    ('0057V000009IN8jQAG', 'Delaney Olguin', 'Lauren Erickson', 'US General Business', 'US GB Central West', 'US GB West'),
    ('0057V000009J4J1QAK', 'Richard Rostvig', 'Mathees Karuna', 'Canada Mid Market', 'Canada MM West', 'Canada MM West'),
    ('0057V000009J4JVQA0', 'Hannah Peach', 'Sara Hurst', 'Canada SMB Client Sales', 'Canada SMB Client Sales Strategic', 'Canada SMB Client Sales Strategic'),
    ('0057V000009JCz6QAG', 'Tyler Sanford', 'Randi Kruger', 'US SMB Client Sales', 'US SMB CSE Key South', 'US SMB Client Sales Key'),
    ('0057V000009JD82QAG', 'Kevin Mosier', 'Kyle Olson', 'US National', 'US National Northeast', 'US National East'),
    ('0057V000009Jf7lQAC', 'Dawson Watkins', 'Justin Simon', 'US General Business', 'US GB Mid Atlantic', 'US GB West'),
    ('0057V000009JfmAQAS', 'John Carmichael', 'Danielle Walter', 'US General Business', 'US GB Southeast', 'US GB East'),
    ('0057V000009JfmFQAS', 'Tyler Klein', 'Heather Lewis', 'US SMB Client Sales', 'US SMB CSE Key West', 'US SMB Client Sales Key'),
    ('0057V000009JfmPQAS', 'Jill Desjardine', 'Megan Frodge', 'US SMB Client Sales', 'US SMB CSE Strategic Northeast', 'US SMB Client Sales Strategic'),
    ('0057V000009JfmZQAS', 'Ryan Stanaway', 'Amber Peters', 'US National', 'US National Midwest', 'US National West'),
    ('0057V000009JrneQAC', 'Oliver Holdenson', 'Heather Lewis', 'US SMB Client Sales', 'US SMB CSE Key West', 'US SMB Client Sales Key'),
    ('0057V000009Ju27QAC', 'Brooke Mullis', 'Ronit Cohn', 'US SMB Client Sales', 'US SMB CSE Key North', 'US SMB Client Sales Key'),
    ('0057V000009K19AQAS', 'Robert Stoering', 'Brooke Nelson', 'US SMB Client Sales', 'US SMB CSE Premier South', 'US SMB Client Sales Premier'),
    ('0057V000009K77KQAS', 'Conor Tomlinson', 'Mathees Karuna', 'Canada Mid Market', 'Canada MM West', 'Canada MM West'),
    ('0057V000009KSbQQAW', 'Aaron Korus', 'Dave Elinger', 'US SMB Client Sales', 'US SMB CSE Strategic Southwest', 'US SMB Client Sales Strategic'),
    ('0057V000009QX2DQAW', 'Mark Workman', 'Blake Karnes', 'US SMB Client Sales', 'US SMB CSE Strategic Northwest', 'US SMB Client Sales Strategic'),
    ('0057V000009RcQ7QAK', 'Phoebe Sewrey', 'Cassie Petrie', 'UK SMB Client Sales', 'UK SMB Client Sales', 'UK SMB Client Sales'),
    ('0057V000009RdV3QAK', 'Chantell Paxton', 'Ryan Nelson', 'US General Business', 'US GB NY Metro', 'US GB East'),
    ('0057V000009TAMVQA4', 'Ashley Tribble', 'Meghan Letnes', 'US Small Business', 'US SB South', 'US SB West'),
    ('0057V000009TpAaQAK', 'Sara Stinson', 'Holly Corser', 'US National', 'US National West', 'US National West'),
    ('0057V0000090kTeQAI', 'Kurt Elden', 'Justin Simon', 'US General Business', 'US GB Mid Atlantic', 'US GB West'),
    ('0057V0000097cGxQAI', 'Sydney Clawson', 'Jack Schwartz', 'US National', 'US National Southeast', 'US National East'),
    ('00532000004kkFyAAI', 'Natalie Jamieson', 'Kyle Loving', 'US SMB Client Sales', 'US SMB CSE Premier East', 'US SMB Client Sales Premier'),
    ('00532000004KrX4AAK', 'Zak Jones', 'Katie Brown', 'UK SMB Client Sales', 'UK SMB Client Sales Key I', 'UK SMB Client Sales Key'),
    ('00532000004ngqXAAQ', 'Ross Greetham', 'Katie Brown', 'UK SMB Client Sales', 'UK SMB Client Sales Key I', 'UK SMB Client Sales Key'),
    ('00532000005CnFdAAK', 'Michael Benn', 'Bret Edis', 'UK SMB Client Sales', 'UK SMB Client Sales Strategic', 'UK SMB Client Sales Strategic'),
    ('00532000005DWEGAA4', 'Michael Landes', 'Brooke Nelson', 'US SMB Client Sales', 'US SMB CSE Premier South', 'US SMB Client Sales Premier'),
    ('00532000005EpKDAA0', 'Katrina Brock', 'Chris Smith', 'US SMB Client Sales', 'US SMB CSE Premier North', 'US SMB Client Sales Premier'),
    ('00532000005EQ6IAAW', 'Sam Angelo', 'Angie Koplan', 'US SMB Client Sales', 'US SMB CSE Premier West', 'US SMB Client Sales Premier'),
    ('00532000005EzQmAAK', 'Joel Segall', 'Christian Larson', 'US SMB Client Sales', 'US SMB CSE Strategic Midwest', 'US SMB Client Sales Strategic'),
    ('00532000005fD7fAAE', 'Amy Slezak', 'Jack Schwartz', 'US National', 'US National Southeast', 'US National East'),
    ('00532000005fQkvAAE', 'Rob Meek', 'Kyle Loving', 'US SMB Client Sales', 'US SMB CSE Premier East', 'US SMB Client Sales Premier'),
    ('00532000005FUakAAG', 'Danny Gloyne', 'Nick Henney', 'UK Ireland', 'UK Ireland', 'UK Ireland'),
    ('00532000005G3IZAA0', 'Kevin Chheang', 'Amanda Meek', 'US SMB Client Sales', 'US SMB CSE Strategic Mountain', 'US SMB Client Sales Strategic'),
    ('00532000005GsOxAAK', 'Evan Anderson', 'Blake Karnes', 'US SMB Client Sales', 'US SMB CSE Strategic Northwest', 'US SMB Client Sales Strategic'),
    ('00532000005LW0UAAW', 'Kim Koehn', 'Chris Smith', 'US SMB Client Sales', 'US SMB CSE Premier North', 'US SMB Client Sales Premier'),
    ('00532000005q3vnAAA', 'Candice Ward', 'Brooke Nelson', 'US SMN Client Sales', 'US SMN CSE Key Mid Atlantic', 'US SMN Client Sales Key'),
    ('00532000005QbBrAAK', 'Jacob Hautman', 'Jack Schwartz', 'US National', 'US National Southeast', 'US National East'),
    ('00532000005rs5dAAA', 'Ty Saathoff', 'Amanda Meek', 'US SMB Client Sales', 'US SMB CSE Strategic Mountain', 'US SMB Client Sales Strategic'),
    ('00532000005s5y9AAA', 'Michael Chamberlain', 'Tim Downs', 'US SMN Client Sales', 'US SMN CSE Strategic Southwest', 'US SMN Client Sales Strategic'),
    ('00532000005sE95AAE', 'Nathaniel Heussner', 'Samantha Young', 'US SMB Client Sales', 'US SMB CSE Strategic Southeast', 'US SMB Client Sales Strategic'),
    ('00532000005t0bbAAA', 'Mike Bosick', 'Amber Peters', 'US National', 'US National Midwest', 'US National West'),
    ('00532000005tCr2AAE', 'Emily Bott', 'Hamish Tebbutt', 'Australia Mid Market', 'Australia Mid Market', 'Australia Mid Market'),
    ('00532000005WIuVAAW', 'Lydia Holloway', 'Bret Edis', 'UK SMB Client Sales', 'UK SMB Client Sales Premier', 'UK SMB Client Sales Premier'),
    ('00532000005WNh9AAG', 'Jon Salmon', 'Angie Koplan', 'US SMB Client Sales', 'US SMB CSE Premier West', 'US SMB Client Sales Premier'),
    ('00532000006BQKXAA4', 'Tyler Mielnichuk', 'Lesley Nunes', 'Canada National', 'Canada National East', 'Canada National East'),
    ('00532000006DMbJAAW', 'Brittany Whims', 'Blake Karnes', 'US SMB Client Sales', 'US SMB CSE Strategic Northwest', 'US SMB Client Sales Strategic'),
    ('00560000001OH2mAAG', 'Adrian Sage', 'Katie Brown', 'UK SMB Client Sales', 'UK SMB Client Sales Key I', 'UK SMB Client Sales Key'),
    ('00560000001OKtNAAW', 'John Sarkie', 'Holly Corser', 'US National', 'US National West', 'US National West'),
    ('00560000001PLAdAAO', 'Tony Hagen', 'Eric Laliberte', 'US General Business', 'US GB Northeast', 'US GB East'),
    ('00560000001Pp2iAAC', 'Charlie Mason', 'Katie Brown', 'UK SMB Client Sales', 'UK SMB Client Sales Key I', 'UK SMB Client Sales Key'),
    ('00560000002IhbKAAS', 'Kammie Flitton', 'Bret Edis', 'UK SMB Client Sales', 'UK SMB Client Sales Premier', 'UK SMB Client Sales Premier'),
    ('00560000002TJb2AAG', 'Stephen Snediker', 'Chris Smith', 'US SMB Client Sales', 'US SMB CSE Premier North', 'US SMB Client Sales Premier'),
    ('00560000002ve4PAAQ', 'Andrew Cooksley', 'Peter Soukos', 'Australia SMB Client Sales', 'Australia SMB Client Sales Strategic', 'Australia SMB Client Sales Strategic'),
    ('00560000003e4voAAA', 'Rachel Burns', 'Christian Larson', 'US SMB Client Sales', 'US SMB CSE Strategic Midwest', 'US SMB Client Sales Strategic'),
    ('00560000003LX2sAAG', 'George Smith', 'Katie Brown', 'UK SMB Client Sales', 'UK SMB Client Sales Key I', 'UK SMB Client Sales Key'),
    ('00560000003M0SfAAK', 'Anders Halvorsen', 'Marissa Mock', 'US SMB Client Sales', 'US SMB CSE Key Great Lakes', 'US SMB Client Sales Key'),
    ('00560000004cwQgAAI', 'Joseph Baker', 'Bret Edis', 'UK SMB Client Sales', 'UK SMB Client Sales Strategic', 'UK SMB Client Sales Strategic'),
    ('00560000004da8wAAA', 'Ellen Bastian', 'Kyle Loving', 'US SMB Client Sales', 'US SMB CSE Premier East', 'US SMB Client Sales Premier'),
    ('005320000050NAzAAM', 'Jamie Stewart', 'Blake Karnes', 'US SMB Client Sales', 'US SMB CSE Strategic Northwest', 'US SMB Client Sales Strategic'),
    ('005320000056nUlAAI', 'Patrick Vestal', 'Angie Koplan', 'US SMN Client Sales', 'US SMN CSE Strategic Northwest', 'US SMN Client Sales Strategic'),
    ('005320000061QpuAAE', 'Amy Lawrence', 'Brooke Nelson', 'US SMB Client Sales', 'US SMB CSE Premier South', 'US SMB Client Sales Premier'),
    ('005600000037feSAAQ', 'Lindsay Bibb', 'Jack Schwartz', 'US National', 'US National Southeast', 'US National East'),
    ('005600000044u8yAAA', 'Gregory Windisch', 'Jack Schwartz', 'US National', 'US National Southeast', 'US National East'),
    ('005600000044z0zAAA', 'Katie Byrnes', 'Angie Koplan', 'US SMB Client Sales', 'US SMB CSE Premier West', 'US SMB Client Sales Premier'),
    ('005600000045u9FAAQ', 'Matt Warren', 'Holly Corser', 'US National', 'US National West', 'US National West'),
    ('005600000046EGpAAM', 'Steve Swift', 'Mathees Karuna', 'Canada National', 'Canada National West', 'Canada National West'),
    ('005600000046mvDAAQ', 'Mandy Gallanar', 'Chris Smith', 'US SMB Client Sales', 'US SMB CSE Premier North', 'US SMB Client Sales Premier'),
    ('005600000046txpAAA', 'Deb Tegan', 'Chris Smith', 'US SMB Client Sales', 'US SMB CSE Premier North', 'US SMB Client Sales Premier'),
    ('0053200000501hzAAA', 'Danny Lewis', 'Angie Koplan', 'US SMB Client Sales', 'US SMB CSE Premier West', 'US SMB Client Sales Premier'),
    ('0053200000609yTAAQ', 'Derek Rux', 'Ana Van de Hinz', 'US SMN Client Sales', 'US SMN CSE Key Great Lakes', 'US SMN Client Sales Key'),
]

_EMBEDDED_ROSTER_DF = pd.DataFrame(
    _ROSTER_DATA,
    columns=["rep_id", "rep_name", "manager", "division", "region", "team"]
)
_EMBEDDED_ROSTER_DF["manager"]  = _EMBEDDED_ROSTER_DF["manager"].fillna("")
_EMBEDDED_ROSTER_DF["region"]   = _EMBEDDED_ROSTER_DF["region"].fillna("")
_EMBEDDED_ROSTER_DF["team"]     = _EMBEDDED_ROSTER_DF["team"].fillna("")
# Drop utility / placeholder rows
_EMBEDDED_ROSTER_DF = _EMBEDDED_ROSTER_DF[
    ~_EMBEDDED_ROSTER_DF["rep_name"].isin(["SMN Digital Commerce"])
].reset_index(drop=True)

ROSTER_AS_OF = "Sep 23, 2026"

def get_embedded_roster() -> pd.DataFrame:
    """Return the embedded Global SMB Roster as a normalised DataFrame."""
    return _EMBEDDED_ROSTER_DF.copy()

# Expected roster columns (flexible matching — for uploaded overrides):
_ROSTER_ID_KEYS   = ["acctown - account owner id", "acctown - account owner id", "owner id", "rep id", "user id", "id"]
_ROSTER_NAME_KEYS = ["acctown - account owner name", "owner name", "rep name", "name"]
_ROSTER_MGR_KEYS  = ["acctown - manager name", "manager name", "manager", "leader", "rsd"]
_ROSTER_DIV_KEYS  = ["acctown - owner division", "owner division", "division"]
_ROSTER_REG_KEYS  = ["acctown - owner record region", "region", "owner record region"]
_ROSTER_TEAM_KEYS = ["acctown - owner team", "team", "owner team"]

def _find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols_lower = {c.lower(): c for c in df.columns}
    for k in candidates:
        if k in cols_lower:
            return cols_lower[k]
    return None

def parse_roster(df: pd.DataFrame) -> pd.DataFrame:
    """
    Normalise a roster DataFrame (uploaded override) into columns:
    rep_id | rep_name | manager | division | region | team
    """
    id_col   = _find_col(df, _ROSTER_ID_KEYS)
    name_col = _find_col(df, _ROSTER_NAME_KEYS)
    mgr_col  = _find_col(df, _ROSTER_MGR_KEYS)
    div_col  = _find_col(df, _ROSTER_DIV_KEYS)
    reg_col  = _find_col(df, _ROSTER_REG_KEYS)
    team_col = _find_col(df, _ROSTER_TEAM_KEYS)

    if not id_col or not name_col:
        raise ValueError(
            "Roster must contain Rep ID and Rep Name columns. "
            f"Detected columns: {list(df.columns)}"
        )

    out = pd.DataFrame()
    out["rep_id"]   = df[id_col].astype(str).str.strip()
    out["rep_name"] = df[name_col].astype(str).str.strip()
    out["manager"]  = df[mgr_col].astype(str).str.strip()  if mgr_col  else ""
    out["division"] = df[div_col].astype(str).str.strip()  if div_col  else ""
    out["region"]   = df[reg_col].astype(str).str.strip()  if reg_col  else ""
    out["team"]     = df[team_col].astype(str).str.strip() if team_col else ""
    # Drop any row where ID looks like a header or is too short
    out = out[out["rep_id"].str.len() > 5].reset_index(drop=True)
    # Drop utility placeholder rows
    out = out[~out["rep_name"].isin(["SMN Digital Commerce"])].reset_index(drop=True)
    return out

# ─────────────────────────────────────────────────────────────────────────────
# ARR / VALUE PARSING
# ─────────────────────────────────────────────────────────────────────────────
_ARR_CANDIDATES = [
    "Contractual ARR USD", "Contractual ARR (converted)", "Contractual ARR",
    "ARR", "Annual Revenue", "Forecast Amount", "Amount"
]

def detect_arr_col(df: pd.DataFrame) -> str | None:
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for cand in _ARR_CANDIDATES:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    # Partial match fallback — catches "Contractual ARR (USD)" etc.
    for col in df.columns:
        if "contractual arr" in col.lower() or ("arr" in col.lower() and "contractual" in col.lower()):
            return col
    return None

def parse_arr(val) -> float:
    """Parse an ARR value to float.

    Handles all formats the app may store or receive:
      • Plain float/int already:  55235.76         → 55235.76
      • SOQL formatted string:   "USD 55,235.76"   → 55235.76
      • Dollar-sign format:      "$55,235.76"      → 55235.76
      • Plain number string:     "55235.76"        → 55235.76
      • Empty / null / nan:                        → 0.0
    """
    if val is None:
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = str(val).strip()
    if s in ("", "nan", "None", "-"):
        return 0.0
    # Strip any leading 3-letter currency code (USD, EUR, GBP …) + whitespace
    s = re.sub(r'^[A-Za-z]{3}\s+', '', s)
    # Strip remaining symbols: commas, dollar signs, stray spaces
    s = re.sub(r"[,$\s]", "", s)
    try:
        return float(s)
    except ValueError:
        return 0.0

# ─────────────────────────────────────────────────────────────────────────────
# FY18 SALES PLANNING FIELD
# ─────────────────────────────────────────────────────────────────────────────
_FY18_COL_CANDIDATES = ["FY18 Sales Planning", "FY18_Sales_Planning__c",
                         "FY18", "Sales Planning"]

def detect_fy18_col(df: pd.DataFrame) -> str | None:
    cols_lower = {c.lower(): c for c in df.columns}
    for cand in _FY18_COL_CANDIDATES:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None

def apply_fy18_tag(val, tag: str) -> str:
    """
    Conditionally add 'Prev Acct Owner: <tag>' to the FY18 Sales Planning field.

    Rules:
      1. If the field already contains a 'Prev Acct Owner:' tag → leave it
         completely unchanged.  The account was temporarily on this rep's book
         and already has its real original owner recorded.
      2. If the field is empty → set it to 'Prev Acct Owner: <tag>'.
      3. If the field has other content but no Prev Acct Owner tag → append
         ', Prev Acct Owner: <tag>' after the existing text.
    """
    s = "" if (pd.isna(val) or str(val).strip() in ("", "nan")) else str(val).strip()

    # Rule 1 — already tagged from a prior redistribution, do not touch
    if "prev acct owner:" in s.lower():
        return s

    # Rule 2 — empty field
    if not s:
        return f"Prev Acct Owner: {tag}"

    # Rule 3 — has other content, append
    return f"{s}, Prev Acct Owner: {tag}"


def clear_fy18_prev_tag(val) -> str:
    """
    Remove ALL 'Prev Acct Owner: ...' and 'Prev Account Owner: ...' segments
    from the FY18 Sales Planning field, preserving all other tags.
    Handles stacked tags (no comma separator between them) cleanly.
    """
    if pd.isna(val):
        return val
    cleaned = re.sub(
        r',?\s*Prev Acct(?:ount)? Owner:[^,;]+',
        '', str(val), flags=re.IGNORECASE
    )
    cleaned = re.sub(r'^[\s,;]+|[\s,;]+$', '', cleaned)
    cleaned = re.sub(r',\s*,', ',', cleaned)
    return cleaned.strip()

# ─────────────────────────────────────────────────────────────────────────────
# DISTRIBUTION LOGIC
# ─────────────────────────────────────────────────────────────────────────────
def _is_customer(type_val: str) -> bool:
    t = str(type_val).strip().lower()
    return t == "customer"

def distribute(acct_df: pd.DataFrame,
               opp_df: pd.DataFrame | None,
               reps: list[dict],
               departing_name: str,
               tag_name: str,
               arr_col: str | None,
               fy18_col: str | None) -> tuple[pd.DataFrame, pd.DataFrame | None, pd.DataFrame]:
    """
    Core distribution engine.

    Customer accounts  → greedy by ARR (equal ARR per rep).
    Non-customer accts → greedy by count (equal volume per rep).
    Open opps          → follow their account's new owner.

    FY18 tag is always appended; pre-existing tags are preserved alongside it.

    Returns (acct_out, opp_out, summary_df).
    """
    rep_names = [r["name"] for r in reps]
    rep_ids   = {r["name"]: r["id"] for r in reps}

    arr_by_rep          = defaultdict(float)
    cust_count_by_rep   = defaultdict(int)   # customer accounts per rep
    count_by_rep        = defaultdict(int)   # non-customer accounts per rep
    opp_forecast_by_rep = defaultdict(float)
    opp_count_by_rep    = defaultdict(int)

    # Weights for multi-objective customer scoring (must sum to 1.0).
    # ARR is the primary axis; account count prevents a few large accounts
    # from piling onto one rep while opp load adds secondary balance.
    _W_ARR          = 0.50
    _W_ACCT_COUNT   = 0.25
    _W_OPP_FORECAST = 0.15
    _W_OPP_COUNT    = 0.10

    # ── 1. Apply FY18 tag ─────────────────────────────────────────────────
    df = acct_df.copy()
    if fy18_col and fy18_col in df.columns:
        # Column was fetched from Salesforce — apply conditional tag logic
        df[fy18_col] = df[fy18_col].apply(lambda v: apply_fy18_tag(v, tag_name))
    # If fy18_col is None the field wasn't fetched; skip tagging entirely
    # rather than risk overwriting an existing Prev Acct Owner tag we can't see.

    # ── 2. Split customers vs others ─────────────────────────────────────
    type_col = _find_col(df, ["type", "account type"])
    if type_col:
        is_cust  = df[type_col].apply(_is_customer)
    else:
        is_cust  = pd.Series([False] * len(df), index=df.index)

    cust_df  = df[is_cust].copy()
    other_df = df[~is_cust].copy()

    # ── 3. Sort customers by ARR desc ─────────────────────────────────────
    if arr_col and arr_col in df.columns:
        cust_df["_arr"] = cust_df[arr_col].apply(parse_arr)
    else:
        cust_df["_arr"] = 0.0
    cust_df = cust_df.sort_values("_arr", ascending=False)

    total_arr = cust_df["_arr"].sum() if len(cust_df) else 0.0

    # If no ARR data is available, collapse customers into the count pool —
    # avoids the greedy algorithm degenerate case where target_arr == 0 causes
    # every iteration to pick the same (first) rep.
    if total_arr == 0 and len(cust_df) > 0:
        other_df = pd.concat([cust_df, other_df])
        cust_df  = cust_df.iloc[:0]   # empty

    total_count      = len(other_df)
    total_cust_count = len(cust_df)
    total_all        = len(df)
    n                = len(rep_names)
    target_arr        = total_arr        / n if n else 1.0
    target_cust_count = total_cust_count / n if n else 1.0
    target_count      = total_count      / n if n else 1.0
    target_all        = total_all        / n if n else 1.0

    # ── Pre-compute per-account opp forecast & count (customers only) ────────
    # Used to factor opp load into the customer greedy scoring.
    acct_opp_forecast: dict[str, float] = {}
    acct_opp_count:    dict[str, int]   = {}
    if opp_df is not None and len(opp_df) > 0:
        _opp_id_col  = _find_col(opp_df, ["18 digit account id", "account id", "accountid"])
        _opp_amt_col = _find_col(opp_df, ["forecast amount", "amount"])
        _acct_id_col = _find_col(df,     ["18 digit account id", "account id", "id"])
        if _opp_id_col and _acct_id_col:
            for _, orow in opp_df.iterrows():
                aid = str(orow.get(_opp_id_col, "") or "")
                amt = parse_arr(orow.get(_opp_amt_col, 0) if _opp_amt_col else 0)
                acct_opp_forecast[aid] = acct_opp_forecast.get(aid, 0.0) + amt
                acct_opp_count[aid]    = acct_opp_count.get(aid, 0) + 1

    # Totals across customer accounts only
    _cust_id_col = _find_col(cust_df, ["18 digit account id", "account id", "id"])
    total_opp_forecast = sum(
        acct_opp_forecast.get(str(cust_df.at[i, _cust_id_col]), 0.0)
        for i in cust_df.index if _cust_id_col
    ) if _cust_id_col else 0.0
    total_opp_count_cust = sum(
        acct_opp_count.get(str(cust_df.at[i, _cust_id_col]), 0)
        for i in cust_df.index if _cust_id_col
    ) if _cust_id_col else 0

    target_opp_forecast   = total_opp_forecast   / n if (n and total_opp_forecast)   else 1.0
    target_opp_count_cust = total_opp_count_cust / n if (n and total_opp_count_cust) else 1.0

    assignments = {}

    # ── Customer greedy: ARR + account count + opp forecast + opp count ──────
    # Sorting by ARR desc first gives the greedy the best chance to spread
    # high-value accounts before filling in smaller ones.
    for idx, row in cust_df.iterrows():
        arr     = row.get("_arr", 0.0)
        acct_id = str(row.get(_cust_id_col, "")) if _cust_id_col else ""
        opp_fc  = acct_opp_forecast.get(acct_id, 0.0)
        opp_ct  = acct_opp_count.get(acct_id, 0)
        best = min(
            rep_names,
            key=lambda r: (
                _W_ARR        * (arr_by_rep[r]          / target_arr) +
                _W_ACCT_COUNT * (cust_count_by_rep[r]   / (target_cust_count or 1.0)) +
                _W_OPP_FORECAST * (opp_forecast_by_rep[r] / target_opp_forecast) +
                _W_OPP_COUNT    * (opp_count_by_rep[r]    / target_opp_count_cust)
            ),
        )
        assignments[idx] = best
        arr_by_rep[best]          += arr
        cust_count_by_rep[best]   += 1
        opp_forecast_by_rep[best] += opp_fc
        opp_count_by_rep[best]    += opp_ct

    # ── Non-customer greedy: balance TOTAL accounts (customers + non-customers)
    # This compensates for any count skew introduced by the customer pass above.
    for idx, _ in other_df.iterrows():
        best = min(
            rep_names,
            key=lambda r: (cust_count_by_rep[r] + count_by_rep[r]) / target_all,
        )
        assignments[idx] = best
        count_by_rep[best] += 1

    df["New Account Owner Name"] = pd.Series(assignments)
    df["New Account Owner ID"]   = df["New Account Owner Name"].map(rep_ids)
    if "_arr" in df.columns:
        df = df.drop(columns=["_arr"])

    # ── 4. Insert new owner cols after Account Owner ──────────────────────
    df = _insert_after(df, "Account Owner",
                       ["New Account Owner Name", "New Account Owner ID"])

    # ── 5. Align opps to account assignment ──────────────────────────────
    opp_out = None
    if opp_df is not None and len(opp_df) > 0:
        opp_out = opp_df.copy()
        # Build account-ID → new owner map
        id_col_acct = _find_col(df, ["18 digit account id", "account id", "id"])
        id_col_opp  = _find_col(opp_df, ["18 digit account id", "account id",
                                          "accountid"])
        owner_map = {}
        id_map_id = {}
        if id_col_acct and id_col_opp:
            owner_map = dict(zip(df[id_col_acct].astype(str),
                                 df["New Account Owner Name"]))
            id_map_id = dict(zip(df[id_col_acct].astype(str),
                                 df["New Account Owner ID"]))
        else:
            # Fall back: match by Account Name
            name_col_acct = _find_col(df, ["account name"])
            name_col_opp  = _find_col(opp_df, ["account name"])
            if name_col_acct and name_col_opp:
                owner_map = dict(zip(df[name_col_acct].astype(str),
                                     df["New Account Owner Name"]))
                id_map_id = dict(zip(df[name_col_acct].astype(str),
                                     df["New Account Owner ID"]))

        match_key_opp  = id_col_opp or _find_col(opp_df, ["account name"])
        match_key_map  = "id" if id_col_opp else "name"

        if match_key_opp:
            opp_out["New Opp Owner Name"] = (
                opp_out[match_key_opp].astype(str).map(owner_map)
            )
            opp_out["New Opp Owner ID"] = (
                opp_out[match_key_opp].astype(str).map(id_map_id)
            )
        else:
            opp_out["New Opp Owner Name"] = ""
            opp_out["New Opp Owner ID"]   = ""

        opp_out = _insert_after(opp_out, "Opportunity Owner",
                                ["New Opp Owner Name", "New Opp Owner ID"])

    # ── 6. Build summary ──────────────────────────────────────────────────
    summary_rows = []
    for r in rep_names:
        mask  = df["New Account Owner Name"] == r
        n_acc = mask.sum()
        arr_v = (df.loc[mask, arr_col].apply(parse_arr).sum()
                 if arr_col and arr_col in df.columns else 0.0)
        n_opp = 0
        if opp_out is not None and "New Opp Owner Name" in opp_out.columns:
            n_opp = (opp_out["New Opp Owner Name"] == r).sum()
        summary_rows.append({
            "Rep Name":         r,
            "Owner ID":         rep_ids[r],
            "Accounts Assigned": int(n_acc),
            "Account ARR":      round(arr_v, 2),
            "Opps Assigned":    int(n_opp),
        })
    summary_rows.append({
        "Rep Name":         "TOTAL",
        "Owner ID":         "",
        "Accounts Assigned": sum(r["Accounts Assigned"] for r in summary_rows),
        "Account ARR":      sum(r["Account ARR"]       for r in summary_rows),
        "Opps Assigned":    sum(r["Opps Assigned"]     for r in summary_rows),
    })
    summary_df = pd.DataFrame(summary_rows)

    return df, opp_out, summary_df


def reassign(acct_df: pd.DataFrame,
             opp_df: pd.DataFrame | None,
             rep_name: str,
             rep_id: str,
             fy18_col: str | None,
             arr_col: str | None) -> tuple:
    """
    Return-flow engine: assign every account to a single rep and clear the
    Prev Acct Owner tag from the FY18 Sales Planning field.

    Returns (acct_out, opp_out, summary_df).
    """
    df = acct_df.copy()

    # Preserve original owner for FS reference
    if "Account Owner" in df.columns:
        df.insert(df.columns.get_loc("Account Owner"), "Original Account Owner",
                  df["Account Owner"])

    df["New Account Owner Name"] = rep_name
    df["New Account Owner ID"]   = rep_id

    # Clear the Prev Acct Owner tag; preserve all other FY18 content
    if fy18_col and fy18_col in df.columns:
        df[fy18_col] = df[fy18_col].apply(clear_fy18_prev_tag)

    df = _insert_after(df, "Account Owner",
                       ["New Account Owner Name", "New Account Owner ID"])

    # Align opps to the returning rep
    opp_out = None
    if opp_df is not None and len(opp_df) > 0:
        opp_out = opp_df.copy()
        opp_out["New Opp Owner Name"] = rep_name
        opp_out["New Opp Owner ID"]   = rep_id
        opp_out = _insert_after(opp_out, "Opportunity Owner",
                                ["New Opp Owner Name", "New Opp Owner ID"])

    arr_v = (df[arr_col].apply(parse_arr).sum()
             if arr_col and arr_col in df.columns else 0.0)
    summary_df = pd.DataFrame([{
        "Rep Name":          rep_name,
        "Owner ID":          rep_id,
        "Accounts Assigned": len(df),
        "Account ARR":       round(arr_v, 2),
        "Opps Assigned":     len(opp_out) if opp_out is not None else 0,
    }])
    return df, opp_out, summary_df


def _insert_after(df: pd.DataFrame, after_col: str,
                  new_cols: list[str]) -> pd.DataFrame:
    """Insert new_cols immediately after after_col (if present)."""
    cols = [c for c in df.columns if c not in new_cols]
    try:
        pos = cols.index(after_col) + 1
    except ValueError:
        pos = len(cols)
    for i, nc in enumerate(new_cols):
        cols.insert(pos + i, nc)
    # Ensure new cols exist in df
    for nc in new_cols:
        if nc not in df.columns:
            df[nc] = ""
    return df[cols]

# ─────────────────────────────────────────────────────────────────────────────
# FS SLIM DATAFRAME  (required fields only)
# ─────────────────────────────────────────────────────────────────────────────
_FS_ACCT_REQUIRED = [
    "18 Digit Account ID", "Account Name", "Account Type",
    "New Account Owner Name", "New Account Owner ID",
]
_FS_OPP_REQUIRED  = [
    "ID (18 Char)", "18 Digit Account ID", "Opportunity Name", "Account Name",
    "Stage", "Close Date", "Forecast Amount",
    "New Opp Owner Name", "New Opp Owner ID",
]

def _arr_col_for_fs(df: pd.DataFrame) -> str | None:
    for c in _ARR_CANDIDATES:
        if c in df.columns:
            return c
    return None

def make_fs_acct_df(acct_df: pd.DataFrame) -> pd.DataFrame:
    arr_col  = _arr_col_for_fs(acct_df)
    fy18_col = detect_fy18_col(acct_df)
    cols = [c for c in _FS_ACCT_REQUIRED if c in acct_df.columns]
    if arr_col and arr_col not in cols:
        cols.append(arr_col)
    if fy18_col and fy18_col not in cols:
        cols.append(fy18_col)
    # Preserve original owner for Field Services reference
    if "Account Owner" in acct_df.columns and "Account Owner" not in cols:
        cols.insert(0, "Account Owner")
    return acct_df[cols].copy()

def make_fs_opp_df(opp_df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in _FS_OPP_REQUIRED if c in opp_df.columns]
    # Include Opportunity Owner so FS can see who currently holds it
    if "Opportunity Owner" in opp_df.columns and "Opportunity Owner" not in cols:
        cols.insert(0, "Opportunity Owner")
    return opp_df[cols].copy()

# ─────────────────────────────────────────────────────────────────────────────
# EXCEL GENERATION
# ─────────────────────────────────────────────────────────────────────────────
_BORDER_THIN = Border(
    bottom=Side(style="thin", color="D0D0D0")
)

def _write_ws(ws, df: pd.DataFrame,
              highlight: set[str],
              row_bg: str | None = None,
              num_fmt_cols: set[str] | None = None):
    """Write a DataFrame to an openpyxl worksheet with SAP styling."""
    num_fmt_cols = num_fmt_cols or set()
    cols = df.columns.tolist()

    # Header row
    for ci, col in enumerate(cols, 1):
        c          = ws.cell(row=1, column=ci, value=col)
        c.font     = Font(bold=True, color=WHITE, name="Calibri", size=10)
        c.fill     = PatternFill("solid",
                                 fgColor=AMBER if col in highlight else NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center",
                                wrap_text=True)
        ws.column_dimensions[get_column_letter(ci)].width = min(
            max(14, len(str(col)) + 4), 46
        )
    ws.row_dimensions[1].height = 30

    # Data rows
    fill = PatternFill("solid", fgColor=row_bg) if row_bg else None
    for ri, (_, row) in enumerate(df.iterrows(), 2):
        for ci, col in enumerate(cols, 1):
            val = row[col]
            if isinstance(val, float) and math.isnan(val):
                val = None
            cell = ws.cell(row=ri, column=ci, value=val)
            if fill:
                cell.fill = fill
            cell.font      = Font(name="Calibri", size=10)
            cell.alignment = Alignment(vertical="top",
                                       wrap_text=(col in {"FY18 Sales Planning",
                                                          "FY18_Sales_Planning__c"}))
            cell.border    = _BORDER_THIN
            if col in num_fmt_cols and val is not None:
                cell.number_format = '#,##0.00'

    ws.freeze_panes = "A2"

def _write_summary_ws(ws, summary_df: pd.DataFrame, include_opps: bool):
    cols = ["Rep Name", "Owner ID", "Accounts Assigned", "Account ARR"]
    if include_opps:
        cols.append("Opps Assigned")
    ws.append([])
    for ci, h in enumerate(cols, 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font      = Font(bold=True, color=WHITE, name="Calibri", size=10)
        c.fill      = PatternFill("solid", fgColor=NAVY)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(ci)].width = [28, 22, 20, 18, 16][ci-1]
    ws.row_dimensions[2].height = 26

    grey = PatternFill("solid", fgColor=GREY_BG)
    for i, (_, row) in enumerate(summary_df.iterrows(), start=3):
        is_total = row["Rep Name"] == "TOTAL"
        for ci, col in enumerate(cols, 1):
            val  = row.get(col, "")
            cell = ws.cell(row=i, column=ci, value=val)
            cell.font      = Font(bold=is_total, name="Calibri", size=10)
            cell.fill      = grey if is_total else PatternFill()
            cell.alignment = Alignment(vertical="top")
            cell.border    = _BORDER_THIN
            if col == "Account ARR" and val:
                cell.number_format = '#,##0.00'
    ws.freeze_panes = "A3"

def build_excel(acct_df: pd.DataFrame,
                opp_df:  pd.DataFrame | None,
                summary_df: pd.DataFrame,
                include_opps: bool,
                full: bool = True) -> bytes:
    """
    Build the output Excel file.
    full=True  → all original columns (field ops version)
    full=False → required fields only (FS version)
    """
    wb = Workbook()

    acct_out    = acct_df if full else make_fs_acct_df(acct_df)
    hl_acct     = {"New Account Owner Name", "New Account Owner ID",
                   "FY18 Sales Planning", "FY18_Sales_Planning__c"}
    arr_col     = detect_arr_col(acct_out)
    num_cols_a  = {arr_col} if arr_col else set()

    ws_acct         = wb.active
    ws_acct.title   = "Account Reassignments"
    _write_ws(ws_acct, acct_out, hl_acct, BLUE_BG, num_cols_a)

    if include_opps and opp_df is not None and len(opp_df) > 0:
        opp_out  = opp_df if full else make_fs_opp_df(opp_df)
        hl_opp   = {"New Opp Owner Name", "New Opp Owner ID"}
        arr_opp  = detect_arr_col(opp_out)
        num_o    = {arr_opp} if arr_opp else set()
        ws_opp   = wb.create_sheet("Opp Reassignments")
        _write_ws(ws_opp, opp_out, hl_opp, BLUE_BG, num_o)

    ws_sum = wb.create_sheet("Distribution Summary")
    _write_summary_ws(ws_sum, summary_df, include_opps)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ─────────────────────────────────────────────────────────────────────────────
# EMAIL HELPER
# ─────────────────────────────────────────────────────────────────────────────
def compose_mailto(departing_name: str, n_accts: int,
                   n_opps: int, include_opps: bool,
                   filename: str,
                   sender_name: str = "") -> str:
    """Build a mailto: URI for the Field Services email."""
    subject = f"Territory Reassignment — {departing_name}"

    # Build the conditional numbered action list
    actions = [
        "Reassign the accounts",
        "Update the FY18 Sales Planning field",
    ]
    if include_opps and n_opps > 0:
        actions.append("Reassign the opportunities")

    numbered = "\n".join(f"{i + 1}. {a}" for i, a in enumerate(actions))
    sig = sender_name.strip() if sender_name.strip() else "[Your name]"

    body = (
        "Hi team,\n"
        "\n"
        "Please make the following updates using the attached file:\n"
        f"{numbered}\n"
        "\n"
        "Thanks,\n"
        f"{sig}"
    )
    return (f"mailto:{FS_EMAIL}"
            f"?subject={urllib.parse.quote(subject)}"
            f"&body={urllib.parse.quote(body)}")


def compose_mailto_return(rep_name: str, n_accts: int,
                          n_opps: int, include_opps: bool,
                          filename: str,
                          sender_name: str = "") -> str:
    """Build a mailto: URI for the return-territory Field Services email."""
    subject = f"Territory Return — {rep_name}"
    actions = [
        f"Reassign the {n_accts:,} accounts back to {rep_name}",
        "Update the FY18 Sales Planning field (Prev Acct Owner tag removed)",
    ]
    if include_opps and n_opps > 0:
        actions.append(f"Reassign the {n_opps:,} open opportunities back to {rep_name}")
    numbered = "\n".join(f"{i + 1}. {a}" for i, a in enumerate(actions))
    sig = sender_name.strip() if sender_name.strip() else "[Your name]"
    body = (
        "Hi team,\n\n"
        "Please make the following updates using the attached file:\n"
        f"{numbered}\n\n"
        "Thanks,\n"
        f"{sig}"
    )
    return (f"mailto:{FS_EMAIL}"
            f"?subject={urllib.parse.quote(subject)}"
            f"&body={urllib.parse.quote(body)}")

# ─────────────────────────────────────────────────────────────────────────────
# UI HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def _metric(label: str, value, green: bool = False) -> str:
    cls = "value green" if green else "value"
    return (f'<div class="metric-card">'
            f'<div class="label">{label}</div>'
            f'<div class="{cls}">{value}</div>'
            f'</div>')

def _metrics_row(*cards: str):
    st.markdown(
        '<div class="metric-row">' + "".join(cards) + '</div>',
        unsafe_allow_html=True
    )

def _step_header(num: int, title: str, done: bool = False):
    cls = "step-card step-done" if done else "step-card"
    st.markdown(
        f'<div class="{cls}">'
        f'<span class="step-num">{num}</span>'
        f'<span class="step-title">{title}</span>'
        f'</div>',
        unsafe_allow_html=True
    )

def _info(msg):  st.markdown(f'<div class="info-box">{msg}</div>', unsafe_allow_html=True)
def _warn(msg):  st.markdown(f'<div class="warn-box">{msg}</div>', unsafe_allow_html=True)
def _ok(msg):    st.markdown(f'<div class="success-box">{msg}</div>', unsafe_allow_html=True)

def _preview_tables(acct_df, opp_df, include_opps,
                    acct_key="prev_acct", opp_key="prev_opp"):
    """Render collapsible account + opp preview tables."""
    # ── Account preview ───────────────────────────────────────────────────────
    ACCT_PREFER = [
        "Account Name", "Account Type", "Type",
        "Contractual ARR (converted)",
        "Original Account Owner", "Account Owner",
        "New Account Owner Name",
        "FY18 Sales Planning", "FY18_Sales_Planning__c",
        "Rating", "Last Activity",
    ]
    acct_show = [c for c in ACCT_PREFER if c in acct_df.columns]
    # Fallback: show all if none of the preferred cols are present
    if not acct_show:
        acct_show = list(acct_df.columns)

    with st.expander(f"Preview accounts ({len(acct_df):,})", expanded=False):
        st.dataframe(
            acct_df[acct_show],
            use_container_width=True,
            hide_index=True,
            key=acct_key,
        )

    # ── Opp preview ───────────────────────────────────────────────────────────
    if include_opps and opp_df is not None and len(opp_df) > 0:
        OPP_PREFER = [
            "ID (18 Char)", "Opportunity Name", "Account Name",
            "Stage", "Close Date", "Forecast Amount",
            "Opportunity Owner", "New Opp Owner Name", "New Opp Owner ID",
        ]
        opp_show = [c for c in OPP_PREFER if c in opp_df.columns]
        if not opp_show:
            opp_show = list(opp_df.columns)
        with st.expander(f"Preview opportunities ({len(opp_df):,})",
                         expanded=False):
            st.dataframe(
                opp_df[opp_show],
                use_container_width=True,
                hide_index=True,
                key=opp_key,
            )

# ─────────────────────────────────────────────────────────────────────────────
# MAIN APP
# ─────────────────────────────────────────────────────────────────────────────
def main():
    # ── Header ───────────────────────────────────────────────────────────────
    st.markdown("""
    <div class="tool-header">
        <h1>Territory Redistribution Tool</h1>
        <p>Distribute a departing rep's accounts and open pipeline to their team.
           Generates a field-ops Excel and a Field Services email attachment.</p>
    </div>
    """, unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════════════════════
    # SIDEBAR — Auth + Config
    # ═══════════════════════════════════════════════════════════════════════
    with st.sidebar:
        st.markdown("### Salesforce Connection")
        sid = st.text_input("Session ID (sid cookie)", type="password",
                            help="Paste your SF session ID from browser DevTools → "
                                 "Application → Cookies → sid")
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("Test", use_container_width=True):
                ok, msg = test_connection(sid)
                st.session_state.connected = ok
                st.session_state.conn_msg  = msg
        with col_b:
            if st.button("Clear", use_container_width=True):
                st.session_state.connected = False
                st.session_state.conn_msg  = ""

        if st.session_state.connected:
            st.success(st.session_state.conn_msg)
        elif st.session_state.conn_msg:
            st.error(st.session_state.conn_msg)





    # ── Mode toggle (placed after sidebar so `sid` is already bound) ─────────
    mode_col, _ = st.columns([3, 5])
    with mode_col:
        mode_label = st.radio(
            "Mode",
            ["Distribute Territory", "Return Territory"],
            index=0 if st.session_state.mode == "distribute" else 1,
            horizontal=True,
            help=(
                "**Distribute Territory** — spread a departing/leaving rep's "
                "accounts across their team and tag with Prev Acct Owner.\n\n"
                "**Return Territory** — pull accounts back to a returning or "
                "newly hired rep by matching the Prev Acct Owner tag, and "
                "clear the tag."
            ),
        )
    new_mode = "distribute" if mode_label == "Distribute Territory" else "return"
    if new_mode != st.session_state.mode:
        st.session_state.mode = new_mode
        st.rerun()

    st.divider()

    # ═══════════════════════════════════════════════════════════════════════
    # RETURN TERRITORY FLOW
    # ═══════════════════════════════════════════════════════════════════════
    if st.session_state.mode == "return":

        # ── Step R1 — Returning / Incoming Rep ──────────────────────────────
        ret_step1_done = bool(st.session_state.ret_rep_name and
                              st.session_state.ret_tag_string and
                              st.session_state.ret_manager)

        with st.expander("Step 1 — Returning / Incoming Rep",
                         expanded=not ret_step1_done):
            c1, c2 = st.columns(2)
            with c1:
                ret_name = st.text_input(
                    "Full name",
                    value=st.session_state.ret_rep_name,
                    placeholder="e.g. Jordan Breisacher",
                    key="ret_name_input",
                )
            with c2:
                ret_id = st.text_input(
                    "Salesforce User ID (18-char)",
                    value=st.session_state.ret_rep_id,
                    placeholder="0057V000…",
                    key="ret_id_input",
                    help="For a new hire taking over the territory, enter their "
                         "SFDC User ID here.",
                )

            c3, c4 = st.columns(2)
            with c3:
                ret_tag = st.text_input(
                    "Original rep's name",
                    value=st.session_state.ret_tag_string or ret_name,
                    placeholder="e.g. Jordan Breisacher",
                    key="ret_tag_input",
                    help="The name that appears in the Prev Acct Owner tag — "
                         "i.e. whose accounts are being returned. For a rep coming "
                         "back from leave this is the same as the name above. For a "
                         "new hire taking over a territory, enter the departed rep's "
                         "name here instead.",
                )
            with c4:
                ret_mgr = st.text_input(
                    "Manager (required)",
                    value=st.session_state.ret_manager,
                    placeholder="e.g. Jake Rutenbar",
                    key="ret_mgr_input",
                    help="The manager whose team currently holds these accounts. "
                         "Used to scope the Salesforce query — accounts are fetched "
                         "by manager and then filtered by the original rep's name.",
                )

            # Live SFDC lookup
            lookup_shown_r = False
            if st.session_state.connected and ret_name and not ret_id:
                results = lookup_user_sfdc(sid, ret_name)
                if results:
                    lookup_shown_r = True
                    options = {f"{r['Name']} ({r['Id']})": r for r in results}
                    choice  = st.selectbox("Select rep from Salesforce",
                                           list(options), key="ret_sfdc_sel")
                    if st.button("Use this rep", key="ret_use_rep"):
                        picked = options[choice]
                        st.session_state.ret_rep_name   = picked["Name"]
                        st.session_state.ret_rep_id     = picked["Id"]
                        st.session_state.ret_tag_string = ret_tag.strip() or picked["Name"]
                        st.session_state.ret_manager    = ret_mgr.strip()
                        reset_ret_results()
                        st.rerun()

            if not lookup_shown_r:
                if st.button("Confirm Rep", key="ret_confirm_rep"):
                    if ret_name and ret_tag and ret_mgr:
                        st.session_state.ret_rep_name   = ret_name.strip()
                        st.session_state.ret_rep_id     = ret_id.strip()
                        st.session_state.ret_tag_string = ret_tag.strip()
                        st.session_state.ret_manager    = ret_mgr.strip()
                        reset_ret_results()
                        _ok(f"Rep set: **{ret_name}** | Matching tag: "
                            f"*Prev Acct Owner: {ret_tag}* | Manager: *{ret_mgr}*")
                    else:
                        _warn("Please enter name, original rep's name, and manager.")

        if ret_step1_done:
            _ok(
                f"Rep: **{st.session_state.ret_rep_name}** &nbsp;|&nbsp; "
                f"Tag match: *Prev Acct Owner: {st.session_state.ret_tag_string}*"
                + (f" &nbsp;|&nbsp; Manager: *{st.session_state.ret_manager}*"
                   if st.session_state.ret_manager else "")
            )

        # ── Step R2 — Fetch Accounts ─────────────────────────────────────────
        ret_step2_done = st.session_state.ret_acct_df is not None

        with st.expander("Step 2 — Account Data",
                         expanded=ret_step1_done and not ret_step2_done):
            if not ret_step1_done:
                _info("Complete Step 1 first.")
            else:
                src_r = st.radio(
                    "Data source",
                    ["Upload file", "Fetch from Salesforce"],
                    horizontal=True,
                    key="ret_acct_src",
                )

                if src_r == "Upload file":
                    _info(
                        "Upload the account file (XLS, XLSX, CSV). The app will "
                        "filter rows whose <strong>FY18 Sales Planning</strong> field "
                        f"contains <strong>Prev Acct Owner: "
                        f"{st.session_state.ret_tag_string}</strong>"
                        + (f", currently held by reps under "
                           f"<strong>{st.session_state.ret_manager}</strong>."
                           if st.session_state.ret_manager else ".")
                    )
                    f_r = st.file_uploader(
                        "Account file (XLS, XLSX, CSV)",
                        type=["xls", "xlsx", "csv"],
                        key="ret_acct_upload",
                    )
                    if f_r and st.button("Load accounts", key="ret_load_file"):
                        try:
                            raw = parse_uploaded_file(f_r)
                            fy18_col_up = detect_fy18_col(raw)
                            if not fy18_col_up:
                                st.error(
                                    "No FY18 Sales Planning column found in the file. "
                                    "Make sure the file includes that field."
                                )
                            else:
                                needle = (
                                    f"prev acct owner: "
                                    f"{st.session_state.ret_tag_string}".lower()
                                )
                                mask = raw[fy18_col_up].astype(str).str.lower().str.contains(
                                    needle, na=False
                                )
                                # Optional manager filter
                                if st.session_state.ret_manager:
                                    mgr_col = _find_col(raw, ["manager"])
                                    if mgr_col:
                                        mgr_mask = raw[mgr_col].astype(str).str.lower().str.contains(
                                            st.session_state.ret_manager.lower(), na=False
                                        )
                                        mask = mask & mgr_mask
                                df = raw[mask].reset_index(drop=True)
                                if len(df) == 0:
                                    _warn(
                                        f"No accounts found with tag "
                                        f"'Prev Acct Owner: "
                                        f"{st.session_state.ret_tag_string}'."
                                    )
                                else:
                                    st.session_state.ret_acct_df = df
                                    reset_ret_results()
                                    st.rerun()
                        except Exception as e:
                            st.error(f"Could not parse file: {e}")

                else:  # Fetch from Salesforce
                    _info(
                        "Queries Salesforce for all accounts whose <strong>FY18 Sales "
                        "Planning</strong> field contains "
                        f"<strong>Prev Acct Owner: "
                        f"{st.session_state.ret_tag_string}</strong>."
                        + (f" Filtered to accounts currently held by reps under "
                           f"<strong>{st.session_state.ret_manager}</strong>."
                           if st.session_state.ret_manager else "")
                    )
                    if not st.session_state.connected:
                        _warn("Connect to Salesforce first (sidebar).")
                    elif st.button("Fetch accounts from Salesforce",
                                   key="ret_fetch_accts"):
                        with st.spinner("Querying accounts…"):
                            try:
                                roster = (
                                    st.session_state.roster_df
                                    if st.session_state.roster_df is not None
                                    else get_embedded_roster()
                                )
                                mgr_filter = st.session_state.ret_manager.strip().lower()
                                team = roster[
                                    roster["manager"].str.lower().str.contains(
                                        mgr_filter, na=False
                                    )
                                ]
                                if len(team) == 0:
                                    st.error(
                                        f"No reps found under '{st.session_state.ret_manager}' "
                                        "in the roster. Check the spelling or upload an "
                                        "updated roster in the Distribute tab."
                                    )
                                else:
                                    rep_ids = team["rep_id"].tolist()
                                    msgs = []
                                    rows, _fy18_api, warns = fetch_accounts_by_tag(
                                        sid,
                                        tag_string=st.session_state.ret_tag_string,
                                        rep_ids=rep_ids,
                                        manager_name=st.session_state.ret_manager,
                                        status_fn=lambda m: msgs.append(m),
                                    )
                                    if not rows:
                                        _warn(
                                            f"No accounts found with tag "
                                            f"'Prev Acct Owner: "
                                            f"{st.session_state.ret_tag_string}'. "
                                            "Check the tag string or manager name."
                                        )
                                    else:
                                        df = pd.DataFrame(rows).fillna("")
                                        st.session_state.ret_acct_df = df
                                        reset_ret_results()
                                        for w in warns:
                                            _warn(w)
                                        if msgs:
                                            with st.expander("Fetch details",
                                                             expanded=True):
                                                for m in msgs:
                                                    st.caption(m)
                                        st.rerun()
                            except ValueError as ve:
                                st.error(str(ve))
                            except Exception as e:
                                st.error(f"Account fetch failed: {e}")

        if ret_step2_done:
            df = st.session_state.ret_acct_df
            arr_col_r  = detect_arr_col(df)
            type_col_r = _find_col(df, ["account type", "type"])
            n_cust_r   = int(df[type_col_r].apply(_is_customer).sum()) if type_col_r else 0
            n_other_r  = len(df) - n_cust_r
            total_arr_r = df[arr_col_r].apply(parse_arr).sum() if arr_col_r else 0.0
            _metrics_row(
                _metric("Accounts", len(df)),
                _metric("Customers", n_cust_r, green=True),
                _metric("Non-Customers", n_other_r),
                _metric("Total ARR", f"${total_arr_r:,.0f}")
                if arr_col_r else _metric("ARR col", "Not found"),
            )

        # ── Step R3 — Open Pipeline (optional) ──────────────────────────────
        with st.expander("Step 3 — Open Pipeline (optional)",
                         expanded=ret_step2_done and not st.session_state.ret_run_done):
            if not ret_step2_done:
                _info("Complete Step 2 first.")
            else:
                include_r = st.checkbox(
                    "Include open opportunities in this return",
                    value=st.session_state.ret_include_opps,
                    key="ret_include_opps_cb",
                )
                st.session_state.ret_include_opps = include_r

                if include_r:
                    if not st.session_state.connected:
                        _warn("Connect to Salesforce first (sidebar).")
                    else:
                        _info(
                            "Fetches all open opportunities linked to the matched "
                            "accounts, regardless of current opp owner."
                        )
                        if st.button("Fetch open opps from Salesforce",
                                     key="ret_fetch_opps"):
                            with st.spinner("Querying opportunities…"):
                                try:
                                    id_col_r = _find_col(
                                        st.session_state.ret_acct_df,
                                        ["18 digit account id", "id"]
                                    )
                                    acct_ids = (
                                        st.session_state.ret_acct_df[id_col_r]
                                        .astype(str).tolist()
                                        if id_col_r else []
                                    )
                                    msgs2 = []
                                    rows2, warns2 = fetch_opps_by_account_ids(
                                        sid, acct_ids,
                                        status_fn=lambda m: msgs2.append(m),
                                    )
                                    st.session_state.ret_opp_df = (
                                        pd.DataFrame(rows2).fillna("")
                                        if rows2 else None
                                    )
                                    reset_ret_results()
                                    for w in warns2:
                                        _warn(w)
                                    if msgs2:
                                        with st.expander("Fetch details",
                                                         expanded=False):
                                            for m in msgs2:
                                                st.caption(m)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Opp fetch failed: {e}")

                    if st.session_state.ret_opp_df is not None:
                        n_opps_r   = len(st.session_state.ret_opp_df)
                        opp_arr_r  = detect_arr_col(st.session_state.ret_opp_df)
                        total_pipe_r = (
                            st.session_state.ret_opp_df[opp_arr_r]
                            .apply(parse_arr).sum() if opp_arr_r else 0.0
                        )
                        _metrics_row(
                            _metric("Open Opps", n_opps_r, green=True),
                            _metric("Forecast Amount",
                                    f"${total_pipe_r:,.0f}"
                                    if opp_arr_r else "—"),
                        )
                else:
                    st.session_state.ret_opp_df = None

        # ── Step R4 — Run Return ─────────────────────────────────────────────
        with st.expander("Step 4 — Run Return",
                         expanded=ret_step2_done and not st.session_state.ret_run_done):
            if not ret_step2_done:
                _info("Complete Steps 1–2 first.")
            else:
                include_opps_r = st.session_state.ret_include_opps
                opp_ready_r    = (not include_opps_r) or (
                    st.session_state.ret_opp_df is not None)
                id_missing     = not st.session_state.ret_rep_id

                if id_missing:
                    ret_id_late = st.text_input(
                        "Salesforce User ID (18-char) — required to run",
                        value="",
                        placeholder="0057V000…",
                        key="ret_id_late",
                        help="Enter the SFDC User ID of the returning or incoming rep. "
                             "For a new hire, use the ID provisioned in Salesforce.",
                    )
                    if ret_id_late.strip():
                        st.session_state.ret_rep_id = ret_id_late.strip()
                        st.rerun()
                    else:
                        _warn("Enter the rep's Salesforce User ID above before running.")

                if not opp_ready_r:
                    _warn("You checked 'Include open opportunities' but "
                          "haven't loaded the opp data yet.")

                col_run_r, _ = st.columns([2, 5])
                run_r = col_run_r.button(
                    "Run Return",
                    disabled=(not opp_ready_r or id_missing),
                    type="primary",
                    use_container_width=True,
                    key="ret_run_btn",
                )

                if run_r:
                    with st.spinner("Preparing return file…"):
                        acct_df_r = st.session_state.ret_acct_df.copy()
                        opp_df_r  = (st.session_state.ret_opp_df.copy()
                                     if include_opps_r and
                                     st.session_state.ret_opp_df is not None
                                     else None)
                        arr_col_r  = detect_arr_col(acct_df_r)
                        fy18_col_r = detect_fy18_col(acct_df_r)

                        acct_out_r, opp_out_r, summary_r = reassign(
                            acct_df_r, opp_df_r,
                            st.session_state.ret_rep_name,
                            st.session_state.ret_rep_id,
                            fy18_col_r, arr_col_r,
                        )

                        rep = st.session_state.ret_rep_name
                        fname_r = f"{rep} Territory Return.xlsx"

                        full_r = build_excel(acct_out_r, opp_out_r, summary_r,
                                             include_opps_r, full=True)
                        fs_r   = build_excel(acct_out_r, opp_out_r, summary_r,
                                             include_opps_r, full=False)

                        st.session_state.ret_result_full     = full_r
                        st.session_state.ret_result_fs       = fs_r
                        st.session_state.ret_result_filename = fname_r
                        st.session_state.ret_summary         = summary_r
                        st.session_state.ret_result_acct_df  = acct_out_r
                        st.session_state.ret_result_opp_df   = opp_out_r
                        st.session_state.ret_run_done        = True
                        st.rerun()

        # ── Return results panel ─────────────────────────────────────────────
        if st.session_state.ret_run_done:
            st.divider()
            st.markdown('<div class="section-title">Results</div>',
                        unsafe_allow_html=True)
            summary_r = st.session_state.ret_summary
            row_r = summary_r.iloc[0]
            _metrics_row(
                _metric("Accounts Returned", int(row_r["Accounts Assigned"]),
                        green=True),
                _metric("Account ARR",
                        f"${row_r['Account ARR']:,.0f}"),
                _metric("Open Opps",
                        int(row_r["Opps Assigned"])
                        if st.session_state.ret_include_opps
                        else "Not in scope"),
            )

            _preview_tables(
                st.session_state.ret_result_acct_df,
                st.session_state.ret_result_opp_df,
                st.session_state.ret_include_opps,
                acct_key="ret_prev_acct",
                opp_key="ret_prev_opp",
            )

            st.divider()
            st.markdown('<div class="section-title">Downloads &amp; Email</div>',
                        unsafe_allow_html=True)
            dl1_r, dl2_r, dl3_r = st.columns(3)
            with dl1_r:
                st.download_button(
                    label="Download — Full File (Field Ops)",
                    data=st.session_state.ret_result_full,
                    file_name=st.session_state.ret_result_filename,
                    mime="application/vnd.openxmlformats-officedocument"
                         ".spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                    key="ret_dl_full",
                )
            with dl2_r:
                fs_fname_r = st.session_state.ret_result_filename.replace(
                    ".xlsx", " — FS Attachment.xlsx"
                )
                st.download_button(
                    label="Download — FS Attachment",
                    data=st.session_state.ret_result_fs,
                    file_name=fs_fname_r,
                    mime="application/vnd.openxmlformats-officedocument"
                         ".spreadsheetml.sheet",
                    use_container_width=True,
                    key="ret_dl_fs",
                )
            with dl3_r:
                sender_r = st.text_input(
                    "Your name (email signature)",
                    value=st.session_state.sender_name,
                    placeholder="e.g. Naman",
                    key="ret_sender_input",
                )
                st.session_state.sender_name = sender_r
                mailto_r = compose_mailto_return(
                    st.session_state.ret_rep_name,
                    int(row_r["Accounts Assigned"]),
                    int(row_r["Opps Assigned"]),
                    st.session_state.ret_include_opps,
                    fs_fname_r,
                    sender_name=sender_r,
                )
                st.link_button(
                    "Compose Email to Field Services",
                    url=mailto_r,
                    use_container_width=True,
                    key="ret_email_btn",
                )
            _info(
                "Download the <strong>FS Attachment</strong> first, then click "
                "<strong>Compose Email</strong> — your email client will open with "
                "To/Subject/Body pre-filled. Attach the downloaded file manually."
            )

    # ═══════════════════════════════════════════════════════════════════════
    # DISTRIBUTE TERRITORY FLOW
    # ═══════════════════════════════════════════════════════════════════════
    if st.session_state.mode == "distribute":

        # ═══════════════════════════════════════════════════════════════════════
        # STEP 1 — Departing Rep
        # ═══════════════════════════════════════════════════════════════════════
        step1_done = bool(st.session_state.departing_name and
                          st.session_state.departing_id)

        with st.expander("Step 1 — Departing Rep", expanded=not step1_done):
            c1, c2, c3 = st.columns([3, 3, 2])
            with c1:
                dep_name = st.text_input(
                    "Full name", value=st.session_state.departing_name,
                    placeholder="e.g. Sydney Clawson"
                )
            with c2:
                dep_id = st.text_input(
                    "Salesforce User ID (18-char)",
                    value=st.session_state.departing_id,
                    placeholder="0057V000…"
                )
            with c3:
                tag_suffix = st.text_input(
                    "FY18 tag suffix (optional)",
                    value="",
                    placeholder='e.g. "26" → Sydney Clawson 26',
                    help='Appended after the name in the Prev Acct Owner tag. '
                         'Leave blank for no suffix.'
                )

            # Live SFDC lookup
            lookup_shown = False
            if st.session_state.connected and dep_name and not dep_id:
                results = lookup_user_sfdc(sid, dep_name)
                if results:
                    lookup_shown = True
                    options = {f"{r['Name']} ({r['Id']})": r for r in results}
                    choice  = st.selectbox("Select rep from Salesforce", list(options))
                    if st.button("Use this rep"):
                        picked = options[choice]
                        st.session_state.departing_name = picked["Name"]
                        st.session_state.departing_id   = picked["Id"]
                        sfx = f" {tag_suffix.strip()}" if tag_suffix.strip() else ""
                        st.session_state.tag_name = picked["Name"] + sfx
                        reset_results()
                        st.rerun()

            if not lookup_shown:
                if st.button("Confirm Rep", key="confirm_rep"):
                    if dep_name:
                        st.session_state.departing_name = dep_name.strip()
                        st.session_state.departing_id   = dep_id.strip()
                        sfx = f" {tag_suffix.strip()}" if tag_suffix.strip() else ""
                        st.session_state.tag_name = dep_name.strip() + sfx
                        reset_results()
                        _ok(f"Departing rep set: {dep_name} | FY18 tag: "
                            f"Prev Acct Owner: {st.session_state.tag_name}")
                    else:
                        _warn("Please enter the rep's name.")

        if step1_done:
            _ok(f"Departing rep: **{st.session_state.departing_name}** "
                f"| FY18 tag: *Prev Acct Owner: {st.session_state.tag_name}*")

        # ═══════════════════════════════════════════════════════════════════════
        # STEP 2 — Load Account Data
        # ═══════════════════════════════════════════════════════════════════════
        step2_done = st.session_state.acct_df is not None
        with st.expander("Step 2 — Account Data", expanded=step1_done and not step2_done):
            if not step1_done:
                _info("Complete Step 1 first.")
            else:
                src = st.radio("Data source", ["Upload file", "Fetch from Salesforce"],
                               horizontal=True, key="acct_src")
                if src == "Upload file":
                    f = st.file_uploader("Account file (XLS, XLSX, CSV)",
                                         type=["xls", "xlsx", "csv"], key="acct_upload")
                    if f and st.button("Load accounts"):
                        try:
                            df = parse_uploaded_file(f)
                            # Filter to departing rep if Account Owner column exists
                            oc = _find_col(df, ["account owner"])
                            if oc:
                                before = len(df)
                                df = df[df[oc].str.strip() == st.session_state.departing_name]
                                df = df.reset_index(drop=True)
                                if len(df) == 0:
                                    _warn(f"No rows matched owner '{st.session_state.departing_name}'. "
                                          "Loaded all rows instead.")
                                    df = parse_uploaded_file(f)
                            st.session_state.acct_df = df
                            reset_results()
                            st.rerun()
                        except Exception as e:
                            st.error(f"Could not parse file: {e}")

                else:  # Fetch from Salesforce
                    if not st.session_state.connected:
                        _warn("Connect to Salesforce first (sidebar).")
                    else:
                        if not st.session_state.departing_id:
                            _warn("Enter the departing rep's Salesforce User ID in Step 1.")
                        else:
                            _info(
                                f"Queries Salesforce directly via SOQL — all accounts "
                                f"owned by <strong>{st.session_state.departing_name}</strong> "
                                f"({st.session_state.departing_id}). "
                                f"No row-count limit. Custom fields auto-discovered."
                            )
                            if st.button("Fetch accounts from Salesforce"):
                                with st.spinner("Querying accounts…"):
                                    try:
                                        msgs = []
                                        rows, warns = fetch_accounts_soql(
                                            sid,
                                            owner_id=st.session_state.departing_id,
                                            status_fn=lambda m: msgs.append(m)
                                        )
                                        df = pd.DataFrame(rows).fillna("")
                                        st.session_state.acct_df      = df
                                        st.session_state.dropped_acct = warns
                                        st.session_state.fetch_msgs   = msgs
                                        reset_results()
                                        for w in warns:
                                            _warn(w)
                                        if msgs:
                                            with st.expander("Fetch details",
                                                             expanded=True):
                                                for m in msgs:
                                                    st.caption(m)
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Account fetch failed: {e}")

        if step2_done:
            df = st.session_state.acct_df
            arr_col  = detect_arr_col(df)
            fy18_col = detect_fy18_col(df)
            type_col = _find_col(df, ["account type", "type"])
            n_cust   = int(df[type_col].apply(_is_customer).sum()) if type_col else 0
            n_other  = len(df) - n_cust
            total_arr = df[arr_col].apply(parse_arr).sum() if arr_col else 0.0

            _metrics_row(
                _metric("Accounts", len(df)),
                _metric("Customers", n_cust, green=True),
                _metric("Non-Customers", n_other),
                _metric("Total ARR", f"${total_arr:,.0f}") if arr_col else _metric("ARR col", "Not found"),
            )
            if not arr_col:
                _warn("No ARR column detected — customers will be distributed by count only.")
            if not fy18_col:
                _warn("No FY18 Sales Planning column detected — tagging will be skipped.")

        # ═══════════════════════════════════════════════════════════════════════
        # STEP 3 — Receiving Team
        # ═══════════════════════════════════════════════════════════════════════
        step3_done = len(st.session_state.receiving_reps) > 0
        with st.expander("Step 3 — Receiving Team", expanded=step2_done and not step3_done):
            if not step2_done:
                _info("Complete Step 2 first.")
            else:
                # ── Roster source ─────────────────────────────────────────────
                st.markdown('<div class="section-title">Roster</div>',
                            unsafe_allow_html=True)
                _info(
                    f"Built-in roster: **Global SMB Roster ({ROSTER_AS_OF})** — "
                    f"{len(_EMBEDDED_ROSTER_DF)} reps across "
                    f"{_EMBEDDED_ROSTER_DF['division'].nunique()} divisions.  "
                    "Upload a newer file below to override."
                )
                roster_file = st.file_uploader(
                    "Upload updated roster (XLS, XLSX, CSV) — optional",
                    type=["xls", "xlsx", "csv"], key="roster_upload"
                )
                if roster_file:
                    try:
                        raw    = parse_uploaded_file(roster_file)
                        parsed = parse_roster(raw)
                        st.session_state.roster_df = parsed
                        _ok(f"Uploaded roster loaded: {len(parsed)} reps.")
                    except Exception as e:
                        st.error(f"Could not parse roster: {e}")
                        st.session_state.roster_df = None

                # Use uploaded override if available, otherwise fall back to embedded
                roster: pd.DataFrame = (
                    st.session_state.roster_df
                    if st.session_state.roster_df is not None
                    else get_embedded_roster()
                )

                # ── Division filter ───────────────────────────────────────────
                st.markdown('<div class="section-title">Filter Roster</div>',
                            unsafe_allow_html=True)
                divisions = sorted(roster["division"].dropna().unique().tolist())
                divisions = [d for d in divisions if d not in ("", "nan")]
                all_div_label = "— All Divisions —"
                div_options   = [all_div_label] + divisions

                selected_division = st.selectbox(
                    "Division", div_options, key="division_sel"
                )
                if selected_division == all_div_label:
                    div_filtered = roster
                else:
                    div_filtered = roster[roster["division"] == selected_division]

                # ── Leader filter (scoped to chosen division) ─────────────────
                leaders = sorted(div_filtered["manager"].dropna().unique().tolist())
                leaders = [l for l in leaders if l not in ("", "nan")]
                all_mgr_label  = "— All Leaders —"
                leader_options = [all_mgr_label] + leaders

                selected_leader = st.selectbox(
                    "Leader / Manager", leader_options, key="leader_sel"
                )
                if selected_leader == all_mgr_label:
                    filtered = div_filtered
                else:
                    filtered = div_filtered[div_filtered["manager"] == selected_leader]

                # Exclude the departing rep from the receiving list
                filtered = filtered[
                    filtered["rep_name"] != st.session_state.departing_name
                ].reset_index(drop=True)

                if len(filtered) == 0:
                    _warn("No eligible reps found after filtering.")
                else:
                    st.markdown(
                        f'<div class="section-title">'
                        f'Select Receiving Reps ({len(filtered)} eligible)</div>',
                        unsafe_allow_html=True
                    )
                    selected = []
                    cols_per_row = 2
                    rows_needed  = math.ceil(len(filtered) / cols_per_row)
                    for row_i in range(rows_needed):
                        row_cols = st.columns(cols_per_row)
                        for ci in range(cols_per_row):
                            idx = row_i * cols_per_row + ci
                            if idx >= len(filtered):
                                break
                            rep   = filtered.iloc[idx]
                            label = (
                                rep["rep_name"]
                                + (f"  ·  {rep['region']}" if rep.get("region") else "")
                            )
                            if row_cols[ci].checkbox(label, value=True,
                                                     key=f"rep_{rep['rep_id']}"):
                                selected.append({"name": rep["rep_name"],
                                                 "id":   rep["rep_id"]})

                    if st.button("Confirm Team", key="confirm_team"):
                        if len(selected) < 1:
                            _warn("Select at least one receiving rep.")
                        else:
                            st.session_state.receiving_reps = selected
                            reset_results()
                            _ok(f"{len(selected)} reps confirmed.")
                            st.rerun()

        if step3_done:
            names = [r["name"] for r in st.session_state.receiving_reps]
            _ok(f"Receiving team ({len(names)}): " + " · ".join(names))

        # ═══════════════════════════════════════════════════════════════════════
        # STEP 4 — Open Pipeline (optional)
        # ═══════════════════════════════════════════════════════════════════════
        with st.expander("Step 4 — Open Pipeline (optional)",
                         expanded=step3_done and not st.session_state.run_done):
            if not step3_done:
                _info("Complete Step 3 first.")
            else:
                include = st.checkbox(
                    "Include open opportunities in this redistribution",
                    value=st.session_state.include_opps,
                    key="include_opps_cb"
                )
                st.session_state.include_opps = include

                if include:
                    src2 = st.radio("Opp data source",
                                    ["Upload file", "Fetch from Salesforce"],
                                    horizontal=True, key="opp_src")
                    if src2 == "Upload file":
                        f2 = st.file_uploader("Open pipeline file (XLS, XLSX, CSV)",
                                              type=["xls", "xlsx", "csv"],
                                              key="opp_upload")
                        if f2 and st.button("Load opps"):
                            try:
                                st.session_state.opp_df = parse_uploaded_file(f2)
                                reset_results()
                                st.rerun()
                            except Exception as e:
                                st.error(f"Could not parse file: {e}")
                    else:
                        if not st.session_state.connected:
                            _warn("Connect to Salesforce first (sidebar).")
                        else:
                            _info(
                                f"Queries Salesforce directly via SOQL — all open "
                                f"opportunities owned by "
                                f"<strong>{st.session_state.departing_name}</strong>."
                            )
                            if st.button("Fetch opps from Salesforce"):
                                with st.spinner("Querying open opportunities…"):
                                    try:
                                        msgs2 = []
                                        rows, warns = fetch_opps_soql(
                                            sid,
                                            owner_id=st.session_state.departing_id,
                                            status_fn=lambda m: msgs2.append(m)
                                        )
                                        st.session_state.opp_df      = pd.DataFrame(rows).fillna("")
                                        st.session_state.dropped_opp = warns
                                        reset_results()
                                        for w in warns:
                                            _warn(w)
                                        if msgs2:
                                            with st.expander("Fetch details",
                                                             expanded=True):
                                                for m in msgs2:
                                                    st.caption(m)
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"Opp fetch failed: {e}")

                    if st.session_state.opp_df is not None:
                        n_opps = len(st.session_state.opp_df)
                        opp_arr_col = detect_arr_col(st.session_state.opp_df)
                        total_pipe  = (st.session_state.opp_df[opp_arr_col]
                                       .apply(parse_arr).sum()
                                       if opp_arr_col else 0.0)
                        _metrics_row(
                            _metric("Open Opps", n_opps, green=True),
                            _metric("Forecast Amount",
                                    f"${total_pipe:,.0f}" if opp_arr_col else "—"),
                        )
                        _info(
                            "Opps will be assigned to the same rep as their account. "
                            "No separate opp distribution algorithm is applied."
                        )
                else:
                    st.session_state.opp_df = None

        # ═══════════════════════════════════════════════════════════════════════
        # STEP 5 — Run + Results
        # ═══════════════════════════════════════════════════════════════════════
        with st.expander("Step 5 — Run Distribution",
                         expanded=step3_done and not st.session_state.run_done):
            if not step3_done:
                _info("Complete Steps 1–3 first.")
            else:
                include_opps = st.session_state.include_opps
                opp_ready    = (not include_opps) or (st.session_state.opp_df is not None)

                if not opp_ready:
                    _warn("You checked 'Include open opportunities' but haven't loaded the opp file yet.")

                col_run, _ = st.columns([2, 5])
                run_clicked = col_run.button(
                    "Run Distribution",
                    disabled=(not opp_ready),
                    type="primary",
                    use_container_width=True
                )

                if run_clicked:
                    with st.spinner("Distributing accounts…"):
                        acct_df  = st.session_state.acct_df.copy()
                        opp_df   = (st.session_state.opp_df.copy()
                                    if include_opps and st.session_state.opp_df is not None
                                    else None)
                        reps     = st.session_state.receiving_reps
                        arr_col  = detect_arr_col(acct_df)
                        fy18_col = detect_fy18_col(acct_df)
                        tag      = st.session_state.tag_name

                        acct_out, opp_out, summary = distribute(
                            acct_df, opp_df, reps,
                            st.session_state.departing_name, tag,
                            arr_col, fy18_col
                        )

                        dep = st.session_state.departing_name
                        fname = f"{dep} Territory Distribution.xlsx"

                        full_bytes = build_excel(acct_out, opp_out, summary,
                                                 include_opps, full=True)
                        fs_bytes   = build_excel(acct_out, opp_out, summary,
                                                 include_opps, full=False)

                        st.session_state.result_full     = full_bytes
                        st.session_state.result_fs       = fs_bytes
                        st.session_state.result_filename = fname
                        st.session_state.dist_summary    = summary
                        st.session_state.result_acct_df  = acct_out
                        st.session_state.result_opp_df   = opp_out
                        st.session_state.run_done        = True
                        st.rerun()

        # ── Results panel ────────────────────────────────────────────────────────
        if st.session_state.run_done:
            st.divider()
            st.markdown('<div class="section-title">Results</div>',
                        unsafe_allow_html=True)

            summary = st.session_state.dist_summary
            reps_only = summary[summary["Rep Name"] != "TOTAL"]
            total_row = summary[summary["Rep Name"] == "TOTAL"].iloc[0]

            # Metrics row
            _metrics_row(
                _metric("Total Accounts", int(total_row["Accounts Assigned"]), green=True),
                _metric("Receiving Reps", len(reps_only)),
                _metric("Open Opps", int(total_row["Opps Assigned"]))
                if st.session_state.include_opps
                else _metric("Open Opps", "Not in scope"),
            )

            # Summary table
            display_cols = ["Rep Name", "Accounts Assigned", "Account ARR"]
            if st.session_state.include_opps:
                display_cols.append("Opps Assigned")
            st.dataframe(
                reps_only[display_cols].style.format({"Account ARR": "${:,.0f}"}),
                use_container_width=True, hide_index=True
            )

            _preview_tables(
                st.session_state.result_acct_df,
                st.session_state.result_opp_df,
                st.session_state.include_opps,
                acct_key="dist_prev_acct",
                opp_key="dist_prev_opp",
            )

            st.divider()
            st.markdown('<div class="section-title">Downloads & Email</div>',
                        unsafe_allow_html=True)

            # Two download buttons + email button
            dl1, dl2, dl3 = st.columns(3)

            with dl1:
                st.download_button(
                    label="Download — Full File (Field Ops)",
                    data=st.session_state.result_full,
                    file_name=st.session_state.result_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                )
            with dl2:
                fs_fname = st.session_state.result_filename.replace(
                    ".xlsx", " — FS Attachment.xlsx"
                )
                st.download_button(
                    label="Download — FS Attachment",
                    data=st.session_state.result_fs,
                    file_name=fs_fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            with dl3:
                sender_input = st.text_input(
                    "Your name (email signature)",
                    value=st.session_state.sender_name,
                    placeholder="e.g. Naman",
                    key="sender_name_input",
                )
                st.session_state.sender_name = sender_input
                n_accts = int(total_row["Accounts Assigned"])
                n_opps  = int(total_row["Opps Assigned"])
                mailto  = compose_mailto(
                    st.session_state.departing_name,
                    n_accts, n_opps,
                    st.session_state.include_opps,
                    fs_fname,
                    sender_name=st.session_state.sender_name,
                )
                st.link_button(
                    "Compose Email to Field Services",
                    url=mailto,
                    use_container_width=True,
                )

            _info(
                "Download the <strong>FS Attachment</strong> first, then click "
                "<strong>Compose Email</strong> — your email client will open with "
                "To/Subject/Body pre-filled. Attach the downloaded file manually."
            )


if __name__ == "__main__":
    main()
