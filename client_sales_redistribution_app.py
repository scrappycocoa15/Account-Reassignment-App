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
    "connected":        False,
    "conn_msg":         "",
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
}
for k, v in _DEFAULTS.items():
    if k not in st.session_state:
        st.session_state[k] = v

def reset_results():
    for k in ("result_full", "result_fs", "result_filename", "run_done",
              "dist_summary"):
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

# ── Report fetching ───────────────────────────────────────────────────────────

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
    Fetch a SFDC Analytics report filtered to a single owner.

    Owner column resolution (describe response only — no pre-run needed):
      1. Existing reportFilters  → column key already known
      2. reportTypeMetadata.categories columns → label-based search ("Account Owner" etc.)
         This is the reliable path: labels are consistent, keys vary by org/version.
      3. detailColumns keys     → catches keys that happen to contain 'owner'
      4. Hard-coded fallback    → last resort

    Filter value: owner NAME is used (Salesforce User fields in report filters
    match on name, not ID). Python post-filter then matches by ID as confirmation.

    If the API filter returns 0 rows (wrong column name or no match), the function
    automatically retries without the owner filter and applies the Python post-filter
    on the full result set.

    Returns (rows, warnings).
    """
    if status_fn:
        status_fn(f"Describing report {report_id}…")
    desc         = _describe_report(sid, report_id)
    meta         = desc.get("reportMetadata", {})
    warnings     = []
    api_filtered = False

    import copy
    meta_original = copy.deepcopy(meta)   # preserve unfiltered meta for fallback

    if owner_id or owner_name:
        existing  = meta.get("reportFilters", [])
        owner_col = None

        # 1. Saved report filters
        for f in existing:
            if "owner" in f.get("column", "").lower():
                owner_col = f["column"]
                break

        # 2. reportTypeMetadata.categories — available in describe response, label-based
        #    Most reliable: "Account Owner" / "Opportunity Owner" labels are always consistent
        if not owner_col:
            for cat in desc.get("reportTypeMetadata", {}).get("categories", []):
                for key, col_def in cat.get("columns", {}).items():
                    label       = col_def.get("label", "")
                    filterable  = col_def.get("filterable", True)
                    if filterable and "owner" in label.lower() and "id" not in label.lower():
                        owner_col = key
                        break
                if owner_col:
                    break

        # 3. detailColumns keys
        if not owner_col:
            for col in meta.get("detailColumns", []):
                if "owner" in col.lower():
                    owner_col = col
                    break

        # 4. Hard-coded fallback
        if not owner_col:
            report_type = meta.get("reportType", {}).get("type", "").upper()
            owner_col   = "OPP_OWNER_NAME" if "OPP" in report_type else "ACCOUNT_OWNER"
            warnings.append(
                f"Owner column not found in report metadata — using fallback '{owner_col}'."
            )

        # Use owner NAME as filter value: Salesforce User fields in report filters
        # match on the user's display name, not their ID.
        filter_value = owner_name if owner_name else owner_id
        new_filters  = [f for f in existing if f.get("column") != owner_col]
        new_filters.append({"column": owner_col, "operator": "equals", "value": filter_value})
        meta["reportFilters"] = new_filters
        api_filtered = True
        if status_fn:
            status_fn(f"Applying owner filter: {owner_col} = '{filter_value}'…")

    if status_fn:
        status_fn("Running report — polling for results…")
    data = _run_report_instance(sid, report_id, metadata_override=meta)
    rows = _factmap_to_rows(data)
    if status_fn:
        status_fn(f"Report returned {len(rows)} rows.")

    # If API filter returned 0 rows, retry without it and rely on Python post-filter
    if api_filtered and len(rows) == 0:
        if status_fn:
            status_fn("0 rows from API filter — fetching full report and filtering in Python…")
        warnings.append(
            "API owner filter returned 0 rows (column name mismatch). "
            "Fetched full report and applied Python filter — results are correct."
        )
        data = _run_report_instance(sid, report_id, metadata_override=meta_original)
        rows = _factmap_to_rows(data)
        if status_fn:
            status_fn(f"Full report: {len(rows)} rows. Filtering to owner…")

    # Python post-filter — always applied as final guarantee
    if rows and (owner_id or owner_name):
        rows, method = _post_filter_rows(rows, owner_id or "", owner_name or "")
        if status_fn:
            status_fn(f"After owner filter: {len(rows)} rows ({method}).")
        if method == "unfiltered":
            warnings.append(
                "No owner column found in report data — results include all owners. "
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
    "Contractual ARR USD", "Contractual ARR", "ARR", "Annual Revenue",
    "Forecast Amount", "Amount"
]

def detect_arr_col(df: pd.DataFrame) -> str | None:
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for cand in _ARR_CANDIDATES:
        if cand.lower() in cols_lower:
            return cols_lower[cand.lower()]
    return None

def parse_arr(val) -> float:
    if val is None:
        return 0.0
    s = str(val).strip()
    if s in ("", "nan", "None", "-"):
        return 0.0
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
    Always append 'Prev Acct Owner: <tag>' to the FY18 Sales Planning field.
    Pre-existing tags (e.g. Prev Acct Owner: Bridget Sands) are left untouched;
    the new tag is added alongside them.
    """
    if pd.isna(val) or str(val).strip() in ("", "nan"):
        return f"Prev Acct Owner: {tag}"
    s = str(val).strip()
    return f"{s}, Prev Acct Owner: {tag}"

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

    arr_by_rep   = defaultdict(float)
    count_by_rep = defaultdict(int)

    # ── 1. Apply FY18 tag ─────────────────────────────────────────────────
    df = acct_df.copy()
    if fy18_col and fy18_col in df.columns:
        df[fy18_col] = df[fy18_col].apply(lambda v: apply_fy18_tag(v, tag_name))

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

    total_arr    = cust_df["_arr"].sum() if len(cust_df) else 0.0
    total_count  = len(other_df)
    n            = len(rep_names)
    target_arr   = total_arr   / n if n else 1.0
    target_count = total_count / n if n else 1.0

    assignments = {}

    # Greedy by ARR for customers
    for idx, row in cust_df.iterrows():
        arr  = row.get("_arr", 0.0)
        best = min(rep_names, key=lambda r: arr_by_rep[r] / target_arr if target_arr else 0)
        assignments[idx] = best
        arr_by_rep[best] += arr

    # Greedy by count for non-customers
    for idx, _ in other_df.iterrows():
        best = min(rep_names, key=lambda r: count_by_rep[r] / target_count if target_count else 0)
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
    "18 Digit Account ID", "Account Name",
    "New Account Owner Name", "New Account Owner ID",
]
_FS_OPP_REQUIRED  = [
    "18 Digit Account ID", "Opportunity Name", "Account Name",
    "New Opp Owner Name",  "New Opp Owner ID",
]

def _arr_col_for_fs(df: pd.DataFrame) -> str | None:
    for c in _ARR_CANDIDATES:
        if c in df.columns:
            return c
    return None

def make_fs_acct_df(acct_df: pd.DataFrame) -> pd.DataFrame:
    arr_col = _arr_col_for_fs(acct_df)
    cols    = [c for c in _FS_ACCT_REQUIRED if c in acct_df.columns]
    if arr_col and arr_col not in cols:
        cols.append(arr_col)
    # Account Owner original
    if "Account Owner" in acct_df.columns and "Account Owner" not in cols:
        cols.insert(0, "Account Owner")
    return acct_df[cols].copy()

def make_fs_opp_df(opp_df: pd.DataFrame) -> pd.DataFrame:
    cols = [c for c in _FS_OPP_REQUIRED if c in opp_df.columns]
    for extra in ["Forecast Amount", "Forecast Amount Currency",
                  "Stage", "Close Date"]:
        if extra in opp_df.columns and extra not in cols:
            cols.append(extra)
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
                   filename: str) -> str:
    subject = f"Territory Reassignment — {departing_name}"
    body_lines = [
        f"Hi Field Services team,",
        "",
        f"Please find attached the territory reassignment file for {departing_name}.",
        "",
        f"  Accounts to reassign : {n_accts}",
    ]
    if include_opps:
        body_lines.append(f"  Open opps to reassign: {n_opps}")
    body_lines += [
        "",
        f"Attachment: {filename}",
        "",
        "Please process the ownership changes at your earliest convenience.",
        "",
        "Thank you,",
    ]
    body = "\n".join(body_lines)
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

        st.divider()
        st.markdown("### Report IDs")
        acct_rpt = st.text_input("Account report", value=DEFAULT_ACCT_RPT)
        opp_rpt  = st.text_input("Open pipeline report", value=DEFAULT_OPP_RPT)

        st.divider()
        st.markdown(
            "<small>Reports are fetched via the Analytics API with a dynamic "
            "owner filter — no need to edit the saved report.</small>",
            unsafe_allow_html=True
        )

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
        if st.session_state.connected and dep_name and not dep_id:
            results = lookup_user_sfdc(sid, dep_name)
            if results:
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

        if st.button("Confirm Rep", key="confirm_rep"):
            if dep_name and dep_id:
                st.session_state.departing_name = dep_name.strip()
                st.session_state.departing_id   = dep_id.strip()
                sfx = f" {tag_suffix.strip()}" if tag_suffix.strip() else ""
                st.session_state.tag_name = dep_name.strip() + sfx
                reset_results()
                _ok(f"Departing rep set: {dep_name} | FY18 tag: "
                    f"Prev Acct Owner: {st.session_state.tag_name}")
            else:
                _warn("Please enter both the rep name and Salesforce User ID.")

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
                    _info(
                        f"Will fetch report **{acct_rpt}** filtered to "
                        f"**{st.session_state.departing_name}**. "
                        f"If the API filter fails, the full report will be fetched "
                        f"and filtered in Python automatically."
                    )
                    if st.button("Fetch accounts from Salesforce"):
                        with st.spinner("Running report…"):
                            try:
                                msgs = []
                                rows, warns = fetch_report_with_owner(
                                    sid, acct_rpt,
                                    owner_id=st.session_state.departing_id,
                                    owner_name=st.session_state.departing_name,
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
                                    with st.expander("Fetch details", expanded=False):
                                        for m in msgs:
                                            st.caption(m)
                                st.rerun()
                            except Exception as e:
                                st.error(f"Report fetch failed: {e}")

    if step2_done:
        df = st.session_state.acct_df
        arr_col  = detect_arr_col(df)
        fy18_col = detect_fy18_col(df)
        type_col = _find_col(df, ["type"])
        n_cust   = int(df[type_col].apply(_is_customer).sum()) if type_col else 0
        n_other  = len(df) - n_cust
        total_arr = df[arr_col].apply(parse_arr).sum() if arr_col else 0.0

        _metrics_row(
            _metric("Accounts", len(df)),
            _metric("Customers", n_cust, green=True),
            _metric("Non-Customers", n_other),
            _metric("Total ARR", f"${total_arr:,.0f}") if total_arr else _metric("ARR col", "Not found"),
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
                            f"Will fetch report **{opp_rpt}** with the same "
                            "owner filter override."
                        )
                        if st.button("Fetch opps from Salesforce"):
                            with st.spinner("Running opp report…"):
                                try:
                                    rows, warns = fetch_report_with_owner(
                                        sid, opp_rpt,
                                        owner_id=st.session_state.departing_id,
                                        owner_name=st.session_state.departing_name
                                    )
                                    st.session_state.opp_df    = pd.DataFrame(rows).fillna("")
                                    st.session_state.dropped_opp = warns
                                    reset_results()
                                    for w in warns:
                                        _warn(w)
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"Opp report fetch failed: {e}")

                if st.session_state.opp_df is not None:
                    n_opps = len(st.session_state.opp_df)
                    opp_arr_col = detect_arr_col(st.session_state.opp_df)
                    total_pipe  = (st.session_state.opp_df[opp_arr_col]
                                   .apply(parse_arr).sum()
                                   if opp_arr_col else 0.0)
                    _metrics_row(
                        _metric("Open Opps", n_opps, green=True),
                        _metric("Total Pipeline",
                                f"${total_pipe:,.0f}") if total_pipe else _metric("Pipeline", "No ARR col"),
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
            n_accts = int(total_row["Accounts Assigned"])
            n_opps  = int(total_row["Opps Assigned"])
            mailto  = compose_mailto(
                st.session_state.departing_name,
                n_accts, n_opps,
                st.session_state.include_opps,
                fs_fname
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
