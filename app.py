"""Company Builds – CoE team build tracker and QC checklist (no sign-in).

Run locally:  streamlit run app.py
Data: Supabase (SUPABASE_URL / SUPABASE_KEY in secrets), or a local SQLite file when those are not set.
"""
import base64
import calendar
import hashlib
import hmac
import html
import os
import re
from datetime import date, datetime
from io import BytesIO
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st

TZ = ZoneInfo("Asia/Kolkata")     # dates and times follow India time

# The CoE team, used for the Built By and QC dropdowns. Edit this list to add or remove people.
TEAM_MEMBERS = ["Fiza", "Aman", "Harsh", "Akash", "Madhu", "Siddharth",
                "Anubhav", "Harshith", "Salman", "Manish", "Rakesh"]


st.set_page_config(page_title="Company Builds", page_icon="🏗️", layout="wide")


# ─────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────
def now_local():
    return datetime.now(TZ)


def today_local():
    return now_local().date()


def month_start(d):
    return d.replace(day=1)


def month_end(d):
    return d.replace(day=calendar.monthrange(d.year, d.month)[1])


def add_months(d, n):
    m = d.month - 1 + n
    return date(d.year + m // 12, m % 12 + 1, 1)


def py(v):
    """Convert numpy/pandas values to plain Python for openpyxl."""
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    return v.item() if hasattr(v, "item") else v


def flash(msg):
    st.session_state["_flash"] = msg


# ── Proper casing ───────────────────────────────────────────
# Words always written in capitals, and small words kept lower case inside a name.
ACRONYMS = {"LLC", "LLP", "PLLC", "LP", "PC", "PA", "DBA", "USA", "US", "GMC", "YMCA", "YWCA", "CDJR", "RV",
            "HVAC", "BBQ", "CPA", "DDS", "MD", "II", "III", "IV", "NX", "SF", "CC", "QC", "PEPM", "AI", "CEO",
            "EAS", "OPCO", "TRT", "PHMG", "KFHS", "BHM", "IPS", "IAB", "EIM", "ATJ", "ABQFF", "ABQTPS", "AGNW",
            "KTOSA", "MDMILLER", "TMS", "SSR", "LBP", "RPB", "MVSP", "PHX", "JTT", "DC"}
SMALL_WORDS = {"of", "and", "the", "for", "at", "in", "on", "a", "an", "to", "by"}
SUFFIXES = {"inc": "Inc", "inc.": "Inc.", "co": "Co", "co.": "Co.", "corp": "Corp", "corp.": "Corp.",
            "ltd": "Ltd", "ltd.": "Ltd."}


def _cap_word(w):
    """Capitalise a word, also after - and ' (e.g. o'neil -> O'Neil, smith-jones -> Smith-Jones)."""
    out, up = [], True
    for ch in w.lower():
        out.append(ch.upper() if up and ch.isalpha() else ch)
        if ch.isalpha():
            up = False
        if ch in "-'(/":
            up = True
    return "".join(out)


def proper_case(text, person=False):
    """Proper casing for company names and people's names.

    ALL-CAPS or all-lower names are rewritten word by word ("BUTLER DD IT NC" -> "Butler DD IT NC").
    Mixed-case names are kept as typed, except for LLC/Inc-style words and words typed all in lower case.
    """
    text = _clean_text(text)
    if not text:
        return text
    text = " ".join(text.split())
    # Judge the casing on the main name only, so "ACME LLC (bi-weekly)" still counts as all caps.
    main = re.split(r"\(|\s-|-\s", text)[0]
    letters = [c for c in main if c.isalpha()] or [c for c in text if c.isalpha()]
    shouting_or_lower = letters and (all(c.isupper() for c in letters) or all(c.islower() for c in letters))
    words = text.split(" ")
    # A name that is just one all-caps word plus LLC/Inc (e.g. "YPO LLC", "AGNW") is an abbreviation: keep it.
    legal = {"LLC", "LLP", "PLLC", "LP", "PC", "PA", "INC", "CO", "CORP", "LTD"}
    name_words = [w for w in main.split() if w.strip(".,()").upper() not in legal]
    if shouting_or_lower and len(name_words) <= 1 and main.isupper() and not person:
        return text                  # (people's names, e.g. SALMAN -> Salman, are always fixed)
    out = []
    for i, w in enumerate(words):
        core = w.strip(".,()")
        if not core or any(c.isdigit() for c in core):
            out.append(w)                                   # codes like H20X2, 551- keep as typed
        elif core.upper() in ACRONYMS:
            out.append(w.replace(core, core.upper()))
        elif w.lower() in SUFFIXES:
            out.append(SUFFIXES[w.lower()])
        elif i > 0 and core.lower() in SMALL_WORDS and (shouting_or_lower or core.islower()):
            out.append(w.lower())
        elif shouting_or_lower:
            # In an all-caps name, 2-letter words (DD, IT, NC) and words with no vowels (BHM) are most likely
            # initials, so keep them in capitals.
            initials = w.isupper() and core not in ("ST", "DR", "MT", "FT") and (
                ("&" in core and len(core) <= 5) or core.isalpha() and (
                    len(core) <= 2 or not set(core.lower()) & set("aeiouy")))
            out.append(w.upper() if initials else _cap_word(w))
        elif core.islower():
            out.append(_cap_word(w))
        else:
            out.append(w)
    return " ".join(out)



# ─────────────────────────────────────────────────────────────
# Data layer (Supabase in production, SQLite locally)
# ─────────────────────────────────────────────────────────────
def _secrets_get(name):
    try:
        return st.secrets.get(name, None)
    except Exception:
        return os.getenv(name)


SUPABASE_URL = _secrets_get("SUPABASE_URL")
SUPABASE_KEY = _secrets_get("SUPABASE_KEY")
USE_SUPABASE = bool(SUPABASE_URL and SUPABASE_KEY)

@st.cache_resource
def get_supabase():
    from supabase import create_client
    return create_client(SUPABASE_URL, SUPABASE_KEY)


def sqlite_conn():
    import sqlite3
    path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "company_builds.db")
    conn = sqlite3.connect(path, timeout=30, check_same_thread=False)
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA busy_timeout=30000")
    conn.execute("""CREATE TABLE IF NOT EXISTS builds (
        id INTEGER PRIMARY KEY AUTOINCREMENT, company_name TEXT NOT NULL, company_code TEXT, hierarchy TEXT,
        status TEXT, completion_date TEXT, due_date TEXT, built_by TEXT, assigned_by TEXT, qc_by TEXT,
        qc_status TEXT)""")
    conn.execute("""CREATE TABLE IF NOT EXISTS build_checks (
        build_id INTEGER NOT NULL, item_id TEXT NOT NULL, built INTEGER DEFAULT 0, qc INTEGER DEFAULT 0,
        na INTEGER DEFAULT 0, PRIMARY KEY(build_id, item_id))""")
    cols = [r[1] for r in conn.execute("PRAGMA table_info(build_checks)")]
    if "comments" not in cols:
        conn.execute("ALTER TABLE build_checks ADD COLUMN comments TEXT")
    if "notes" not in cols:
        conn.execute("ALTER TABLE build_checks ADD COLUMN notes TEXT")
    bcols = [r[1] for r in conn.execute("PRAGMA table_info(builds)")]
    for col in ("ua_status", "balance_status", "qc_completion_date", "build_month"):
        if col not in bcols:
            conn.execute(f"ALTER TABLE builds ADD COLUMN {col} TEXT")
    conn.commit()
    return conn


def _fetch_all(table, cols):
    """Supabase returns at most 1,000 rows per request, so page through everything."""
    sb = get_supabase()
    rows, start, size = [], 0, 1000
    while True:
        batch = sb.table(table).select(cols).range(start, start + size - 1).execute().data or []
        rows.extend(batch)
        if len(batch) < size:
            return rows
        start += size



# ─────────────────────────────────────────────────────────────
# Company Builds
# ─────────────────────────────────────────────────────────────
NOT_REQUIRED = "Not Required"
BUILD_STATUSES = ["Pending", "In Progress", "Completed", NOT_REQUIRED]
WAIT_CLIENT, WAIT_INTERNAL = "Waiting on Client", "Waiting on Internal Resource"
STATUS_LABEL = {"Pending": "⚪ Pending", "In Progress": "🟡 In Progress", "Completed": "🟢 Completed",
                NOT_REQUIRED: "⚫ Not Required", WAIT_CLIENT: "🔵 Waiting on Client",
                WAIT_INTERNAL: "🟠 Waiting on Internal Resource"}
# Extra tracking columns: shown and editable in the tables, but NOT counted in any of the totals.
UA_STATUSES = ["Pending", "In Progress", "Completed", NOT_REQUIRED, WAIT_CLIENT, WAIT_INTERNAL]
BALANCE_STATUSES = ["Pending", "In Progress", "Completed", WAIT_CLIENT, WAIT_INTERNAL]
EXTRA_STATUS = {"ua_status": UA_STATUSES, "balance_status": BALANCE_STATUSES}
EXTRA_COLUMNS = {"enabled": True}   # turns off if the database doesn't have these columns yet
# QC Completed date: filled in automatically when QC Status is set to Completed.
QC_DATE = "qc_completion_date"
QC_DATE_COLUMN = {"enabled": True}  # turns off if the database doesn't have this column yet
# The month a build was added under in Monthly Builds. Keeps a build without a due date in that month
# (a due date, once added, decides the month instead). Not shown as a column.
BUILD_MONTH = "build_month"
BUILD_MONTH_COLUMN = {"enabled": True}  # turns off if the database doesn't have this column yet
# Payroll Leads who assign builds (US team). Add names here as needed.
PAYROLL_LEADS = ["Chelsi", "Dakota", "Jessikah", "Kristin", "Lauren", "Macy", "Tiffany"]

BUILD_COLS = ["id", "company_name", "company_code", "hierarchy", "status", "completion_date",
              "due_date", "built_by", "assigned_by", "qc_by", "qc_status", "ua_status", "balance_status",
              "qc_completion_date", "build_month"]
CHECK_COLS = ["build_id", "item_id", "built", "qc", "na", "comments", "notes"]
BUILD_FIELDS = {  # column shown in the tracker -> database field, in the order the tables show them
    "Company Name": "company_name", "Company Code": "company_code", "Hierarchy": "hierarchy",
    "Status": "status", "QC Status": "qc_status",          # the two statuses side by side
    "UA Status": "ua_status", "Balance Status": "balance_status",   # extra tracking, not counted in totals
    "Built By": "built_by", "QC": "qc_by",                  # then who did each, in the same order
    "Assigned By": "assigned_by", "Due Date": "due_date", "Build Completed": "completion_date",
    "QC Completed": "qc_completion_date",
}
DATE_COLUMNS = ("Build Completed", "QC Completed", "Due Date")
DATE_FIELDS = ("completion_date", "qc_completion_date", "due_date")
# Column order of the Monthly Builds table.
BUILD_COLUMN_ORDER = ["Company Name", "Company Code", "Hierarchy", "Due Date", "Built By", "Status", "QC",
                      "QC Status", "Assigned By", "Build Completed", "QC Completed", "UA Status", "Balance Status"]
PROPER_CASE_FIELDS = ("company_name", "hierarchy", "built_by", "assigned_by", "qc_by")
PERSON_FIELDS = ("built_by", "assigned_by", "qc_by")
# Implementation Checklist - Direct Sales Clients (updated Nov 2026: the new checklist merged with the
# first version, keeping the more detailed wording where both had the same step).
# (section, note shown under the section, [(sub-heading, [(item id, item text)])]); "★ " marks key items.
# Saved ticks and comments are stored by item id, so:
#   - never change or reuse an id; to reword an item, keep its id and change only the text
#   - give a new item a new id (next free N-number); a removed item's id is simply left unused
# Ids like "S01-05" come from the first version of the checklist, so builds ticked then keep their ticks.
CHECKLIST = [
    ("Salesforce - Details", "", [
        ("Account Information", [
            ("S01-01", "Under the Account Information section, change Account Record Type from Prospect to Client Account"),
            ("S01-04", "If Copy Company: verify account is tied to Parent Account"),
            ("S01-03", "Verify Account Name matches legal name on PIN (if not, update it so it matches)"),
            ("S01-05", "Enter DBA Name (if applicable)"),
            ("S01-06", "Add Company Code to Client Company Code field"),
            ("S01-07", "★ Verify Federal ID matches PIN"),
            ("S01-08", "Update Phone Number (if applicable)"),
            ("S01-09", "Change TYPE from Prospect to Netchex Client"),
            ("S01-10", "Choose Pricing Model — PEPM"),
            ("S01-11", "Choose Package — DealerTrack clients only (leave blank for everyone else)"),
            ("S01-12", "Choose Yes/No for Distributed Invoicing"),
            ("S01-13", "Choose Yes/No for Seasonal Client"),
            ("S01-14", "★ Check Currently in Implementation box"),
        ]),
        ("Account Financial Status", [
            ("S01-02", "★ Enter Average Payroll Amount from the analysis document (not known yet? enter 0.00 and update before the first direct deposit payroll)"),
        ]),
        ("Address Information", [("S01-15", "Enter Shipping Address")]),
        ("Designated Points of Contact", [
            ("S01-16", "★ Select Payroll Contact"),
            ("S01-17", "★ Select Tax Contact"),
            ("S01-18", "★ Select Billing/ACH Contact"),
            ("S01-19", "★ Select The Work Number Contact (only if TWN is on PEPM proposal)"),
            ("S01-20", "Select Human Resources Contact (usually same as Work Number Contact)"),
            ("S01-21", "Select Benefits Contact (only if Benefits is on PEPM proposal)"),
        ]),
        ("Client Information", [
            ("S01-22", "Enter Netchex Annual Revenue (monthly PEPM total x 12)"),
            ("S01-23", "Check AZ/HI Daylight Savings Time (only if client is in Arizona or Hawaii)"),
        ]),
        ("Implementation Details", [
            ("S01-24", "★ Select Yes or No for Direct Deposit Waiting Period Required"),
            ("S01-25", "★ Select Proof of Direct Deposit Received (once proof is on file)"),
        ]),
        ("Employer Services Information", [
            ("S01-26", "Update Primary Payroll Frequency"),
            ("S01-27", "Update Secondary Payroll Frequency (if applicable)"),
            ("S01-28", "Verify client's Payroll Vendor is populated and correct"),
            ("S01-29", "401k Offered — Yes or No"),
            ("S01-30", "If Yes: enter Current 401k Plan Provider (if applicable)"),
        ]),
    ]),
    ("Salesforce - Contacts", "", [
        ("Contacts", [
            ("S02-01", "Add/Update all Contacts (First Name, Last Name, Title, Phone Number, Email)"),
            ("S02-02", "Check 'Able to Request User Access' within the contact ONLY if user is on the Client Admin Form"),
        ]),
    ]),
    ("Salesforce - Assets",
     "⚠️ Activate only the PAYROLL assets below that are on the proposal (Status Pending > Implemented/In Use; "
     "missing? click New). Leave HR, Benefits and E-Verify assets to the HR Onboarding Lead / Benefit Leads: "
     "Onboarding & Workflows, Recruit, Performance, NetLearn, Insights, ExpensePath, AskHR, Employee Navigator, "
     "NetBenefits, NetEnroll, NetCobra, ACA Central, Carrier Solutions, EDI Feed or any forms, NetEngage, "
     "NetScheduler, Compensation Management, Fillable Forms and E-Verify.", [
        ("Assets", [
            ("N01", "Payroll/Tax"),
            ("N02", "ESS/MSS"),
            ("N03", "The Work Number"),
            ("N04", "Time & Attendance"),
            ("N05", "EE Engagement/Recognition Portal — Basic/Advanced (Community Core)"),
        ]),
    ]),
    ("Salesforce - Onboarding Details", "", [
        ("Onboarding Details", [
            ("S04-01", "Select Project Manager (Owner)"),
            ("S04-02", "Select Data Specialist"),
            ("S04-03", "Select Payroll Onboarding Lead"),
            ("S04-04", "Select HR Onboarding Lead"),
            ("S04-05", "Select Payroll Survey Contact (under Payroll Information)"),
            ("S04-06", "Select Survey Contact (under Client Contacts)"),
        ]),
    ]),
    ("Netchex - Tax Codes, Worker's Comp & Job Codes", "", [
        ("Maint > Payroll > Tax Codes", [
            ("S05-01", "Update or Add State Tax Codes"),
            ("N06", "Update or Add SUTA Tax Codes"),
            ("S05-02", "Add Local Tax Codes (if applicable)"),
        ]),
        ("Maint > Payroll > Worker's Compensation Codes", [
            ("S05-03", "Build Worker's Comp Codes for each worked-in state/WC code combination"),
            ("S05-04", "Update Worker's Comp Rates (move decimal 2 places left to convert %)"),
        ]),
        ("Maint > HR > Job Codes", [
            ("S05-05", "Build Hourly Job Code for each state/WC code combination"),
            ("N07", "Build Salary Job Code for each state/WC code combination"),
        ]),
    ]),
    ("Netchex - Widgets & The Work Number", "", [
        ("Maint > HR > Dashboard Configuration > Add/Remove Widgets", [
            ("S06-01", "My Profile"),
            ("S06-02", "Company Directory"),
            ("S06-03", "My Pay"),
            ("S06-04", "People"),
            ("S06-05", "Favorite Reports"),
            ("S06-06", "My Accrual Attendance"),
            ("S06-07", "Calendar"),
            ("S06-08", "Company Info"),
            ("S06-09", "Employee Search"),
            ("S06-10", "Payroll"),
            ("N16", "Secure Act 2.0"),
        ]),
        ("Reports > File Feeds", [
            ("S06-12", "Set up The Work Number (Company Info + Client Contact + Partner Contact — Rene Roy)"),
        ]),
    ]),
    ("Netchex - Client Billing",
     "Fee – Premium negative amounts: Weekly: Base -$18.00, Per Check -$1.35 · Bi-Weekly/Semi-Monthly: "
     "Base -$23.00, Per Check -$2.10 · Monthly: Base -$43.00, Per Check -$3.65. Premium Pay steps can only be done "
     "once Netchex Premium Pay Client is checked in System Tables.", [
        ("Back Office > Enterprise > Client Billing Schedule", [
            ("S07-01", "Choose Payroll Frequency"),
            ("S07-02", "Check Sales Tax Box (Texas companies only)"),
            ("S07-03", "Choose Sales Agent — Addie Kelley"),
            ("S07-04", "Choose Payroll Coordinator (Project Lead)"),
            ("S07-05", "Choose Sales Tax Code (TX = Dallas 8.25%; all others = State Sales Tax)"),
            ("S07-06", "★ Enter Fee – Premium negative amounts (Base + Per Check) based on pay frequency"),
        ]),
        ("Premium Pay clients only", [
            ("S07-07", "Check Premium Pay"),
            ("S07-08", "Uncheck Money Movement"),
            ("S07-09", "Zero out $25.00 Stop Payment Fee"),
        ]),
        ("Maint > Company Structure > Client Billing View", [
            ("S07-10", "View Client Billing View to confirm accuracy — no errors showing"),
        ]),
    ]),
    ("Netchex - Company Information (System Tables)", "", [
        ("Company Demographic Information", [
            ("S08-01", "★ Verify Company Name (Trade Name and W2 Name), Federal ID, and Address match the PIN (ALL CAPS)"),
            ("S08-02", "Enter Client Phone Number"),
            ("S08-03", "Add Client Email Address"),
        ]),
        ("Banking Information", [
            ("S08-04", "★ Verify Banking Information — Routing & Account Numbers"),
            ("S08-05", "★ Enter EFTPS PIN under ACH Contact field"),
            ("S08-06", "Check Minimum Wage Makeup if Restaurant or Hotel"),
        ]),
        ("Payroll Information", [
            ("S08-07", "Confirm Minimum Wage $ Amount"),
            ("S08-08", "Update Default Payroll Group"),
            ("S08-09", "Update Default Job Code"),
            ("S08-10", "Populate Units/Pay Period"),
            ("S08-11", "Change Default Employee Classification to W-2 Tipped if Restaurant/Hotel"),
            ("S08-12", "Check Allow Manager Self Service (only if ESS/MSS is on proposal)"),
            ("S08-13", "PREMIUM PAY CLIENTS ONLY: Check the Netchex Premium Pay Client box (see Client Billing)"),
        ]),
        ("Direct Deposit", [
            ("N08", "★ Direct Deposit Effective Date — enter onboarding start date (bypassing 90-day) "
                    "or 90 days post 1st check date"),
            ("N09", "Opt-Out — check ONLY if client will use live checks only (no DD at all)"),
        ]),
    ]),
    ("Netchex - Claims (activate based on PEPM proposal)", "", [
        ("Claims", [
            ("N10", "Organization Chart (if on proposal)"),
            ("N11", "Onboarding and Workflows (if on proposal)"),
            ("S08-17", "Performance / NetPerform (if on proposal)"),
            ("S08-14", "Recruit - HigherMe / NetRecruiter (if on proposal)"),
            ("S08-18", "NetScheduler (if on proposal)"),
            ("S08-16", "Fillable Forms (if on proposal)"),
            ("S08-19", "Engagement Basic (activate for ALL clients — has feature flags)"),
            ("S08-20", "Engagement Core (if on proposal)"),
            ("S08-21", "Engagement Advanced (if on proposal)"),
        ]),
    ]),
    ("Netchex - Feature Flags",
     "Maintenance > Security > New Feature Access. Follows the live Feature Flags Index (11/16/2026); "
     "if the two differ, the Index wins.", [
        ("All clients", [
            ("S09-06", "★ New Payroll Engine — NOT auto-on for new clients. If missing: uncheck Multi-Company "
                       "Payroll, check New Payroll Engine, then re-check Multi-Company Payroll"),
            ("S09-07", "★ Multi-Company Payroll (3012) + Payroll Dashboard — turn on together (MCP alone shows "
                       "mock data; ~15 min to apply, no payroll in progress)"),
            ("S09-08", "Payroll Review and Submit (3000) + Review and Submit UI Refresh (3033) — needs New Payroll "
                       "Engine; claims Employee Check Preview + Payroll Comparison"),
            ("S09-09", "EditTimesheets (3035), Payroll Status (3036), VoidRecreateCheck (3034), "
                       "Start Payroll Redesign (3037), Automated 401k Catch-up (3040)"),
            ("S09-10", "★ Locals Tax, Payroll Process Timesheet Service, Invoice PDF Attachment"),
            ("S09-11", "Verify the defaults are on: Payroll Schedules, Suggested Taxes, Secure Act 2.0, Recognition "
                       "Settings, Minimum Hours Worked, Accrual Expiration, Badge Punch Decoupled, Who's In "
                       "Dashboard, Time Off Policies, Automatic Holidays, Payday Notification, Direct Deposit "
                       "Authorization, New Scheduler"),
            ("S09-02", "Engagement Phase One, with the Engagement Basic + Engagement Core claims "
                       "(claim names still being confirmed)"),
            ("S09-03", "Verify Earned Wage Access is on (not yet confirmed for every client; for legacy Rain "
                       "clients check before changing it)"),
            ("S09-04", "Earned Wage Access: Product Suite > Earned Wage Access, activate the slider for the company"),
            ("S09-12", "Punch Request — clients with Time & Attendance"),
            ("S09-13", "Employee Profile (user level) — verify every employee has it after the import"),
        ]),
        ("If on the PEPM proposal", [
            ("S09-01", "Add the flags for every product on the PEPM proposal (claims in Company Info page)"),
            ("S09-14", "Scheduler: Scheduler claim + New Scheduler + Who's In Dashboard; user access case for admins"),
            ("S09-15", "AskHR: AskHR flag + Community Basic and Community Core claims + security 5320; "
                       "docs uploaded, then Tanya Jain enables AI"),
            ("S09-16", "Netchex AI: Netchex AI flag + NetchexAI claim, then Zoom Tanya Jain with the company code"),
            ("S09-17", "Insights: Insights flag + security groups 5090 and 5110"),
            ("S09-18", "NetLearn: LMS + NetLearn - Manage Users (after BizLibrary setup)"),
            ("S09-19", "Recruit: Recruit - HigherMe claim + HigherMe Recruiting UI Access (user level) — "
                       "admin permissions first, flag last"),
            ("S09-20", "Communications flag (Announcements / Conversations) + Workplace Communication claim"),
            ("S09-21", "Benefits clients: Benefits Enablement Experience"),
            ("S09-22", "ExpensePath (employee level) + ExpensePath API integration"),
            ("S09-23", "Onboarding & Workflows clients: Effective Dating - Company Structure, "
                       "WorkflowsLicenseAndCertificationTasks"),
            ("S09-05", "Activate QBO (ONLY for QBO Integration clients)"),
        ]),
        ("Leave off unless requested", [
            ("S09-24", "By client request only: Effective Dating - Rates (now off by default), CA Break Waived "
                       "(4029), Arrears deductions, Shift Preview Notes, Position attributes"),
            ("S09-25", "Never for clients: [DemoOnly] Conversation Dashboard, Employee list MUI grid. "
                       "Fillable Forms Bulk Download only for existing NetGuide clients"),
        ]),
    ]),
    ("Netchex - Department Structure & Payroll Maintenance",
     "Payroll Group hours: Weekly = 40 · Bi-Weekly = 80 · Semi-Monthly = 86.6667 · Monthly = 173.3334", [
        ("Maint > More > Company Structure", [
            ("S10-01", "Set up Div/Business Unit/Department Structure"),
            ("S10-03", "Set up Locations (all physical work locations)"),
        ]),
        ("Maint > HR", [
            ("S10-04", "Set up Positions (mandatory — mimic departments if client does not provide positions)"),
        ]),
        ("Payroll > Utilities", [
            ("S10-05", "★ Build Payroll Schedule(s)"),
            ("S10-06", "Verify Payroll Schedule with client after building"),
            ("S10-07", "★ Close the Year"),
            ("S10-08", "Add Company Note — PEPM Package/New Proposal + initials"),
            ("N13", "Add Company Note — DD Waiting Period (if applicable)"),
            ("S10-09", "Build Payroll Groups (Hourly + Salary for each frequency used)"),
            ("N14", "Delete unused default payroll groups (e.g., remove Bi-Weekly if not applicable)"),
        ]),
    ]),
    ("Netchex - Payroll Codes & General Ledger", "", [
        ("Payroll Codes & GL", [
            ("S11-01", "Update / Add Deduction Codes"),
            ("N15", "Update / Add Earnings Codes"),
            ("S11-02", "Delete unused Earnings/Deduction Codes"),
            ("S11-03", "Build GL Base Account for each Earnings/Deduction/Tax code"),
            ("S11-04", "★ Map GL Base Accounts — Deductions, Earnings, and Taxes tabs (no unmapped codes)"),
        ]),
    ]),
    ("Netchex - Maintenance & Report Runner",
     "Shipping: Zone 2 $31 · Zone 3 $36.50 · Zone 4 $44 · Zone 5 $47 · Zone 6 $48 · Zone 7 $52.50 · "
     "Zone 8 $55 · Garnishments $1.25.", [
        ("Maint > HR > Email/Text Event Configuration", [
            ("S12-01", "Payroll Invoiced — list client as To: recipient"),
            ("S12-02", "Payroll Submitted — list client and Project Lead as To: recipients"),
            ("S12-03", "Wire Transfer Initiated — list client and Project Lead as To: recipients"),
            ("S12-04", "Wire Transfer Received — list client and Project Lead as To: recipients"),
        ]),
        ("Maint > HR", [
            ("S12-05", "Set up Shipping Information (correct zone fee, notes, garnishments)"),
            ("S12-06", "Add Payroll and Shipping Contacts (physical address — no PO Box)"),
            ("S12-08", "Add W-2 Contact"),
        ]),
        ("Maint > System Tables > Company Structure", [
            ("S12-09", "Enter Check Header (Fields 1, 2, 3)"),
            ("S12-10", "Check Box Details 1, 3, 6, 7 (Dept Info, Accrual Info, Saved Signature, Employer Ded. Amounts)"),
            ("S12-11", "Premium Pay only: add \"(C/O Netchex Payroll)\" to the second line of the Check Header"),
        ]),
        ("Maint > Company Information", [
            ("N17", "Upload Client Signature (JPEG via Company Image Upload Utility; confirm in Check Signatures "
                    "report; not for Premium Pay)"),
        ]),
        ("Back Office > More > RR", [("S12-12", "Verify the company is set up in Report Runner (now automatic for new companies; "
                                               "if missing, copy from CIW/OHE/parent)")]),
    ]),
    ("Netchex - Time & Attendance", "", [
        ("T&A > More > Settings > Web Terminal Setup", [
            ("S13-01", "Verify Web Terminal WEB01 (exists by default) — Time Clock or Time Sheet, enable/disable per client"),
        ]),
        ("T&A > Company Rules", [
            ("S13-02", "Set up Payroll Group Rules (all payroll groups)"),
            ("S13-03", "Shift Setup — OT > 40 on all hourly groups; remove OT from Salary Exempt groups"),
            ("S13-04", "Verify Terminal Group T01 is linked to WEB01 (exists by default)"),
            ("S13-05", "Add Employees to Terminal Group — after the employee load, not during the build"),
        ]),
    ]),
    ("Salesforce - Cases", "⚠️ All cases below must be created after every shell build without exception.", [
        ("Cases", [
            ("S14-01", "★ Pre-Note Case — let flow to Prenote Queue"),
            ("S14-02", "★ Tax Case — assign to your Tax Rep (I'Neka Martin or James Brunson)"),
            ("S14-03", "★ User Access Case — assign to your HR Lead"),
            ("S14-04", "Signature Case — assign to Macy, Jessica D or Jamie"),
            ("S14-05", "E-Verify Case — assign to your HR Lead (only if on proposal)"),
            ("S14-06", "Pay Card Case (only if client is interested in Pay Cards)"),
        ]),
    ]),
    ("GuideCX", "", [
        ("GuideCX", [("S15-01", "Update completed tasks accordingly")]),
    ]),
]


def _checklist_items():
    items, seen = [], set()
    for section, _note, groups in CHECKLIST:
        for sub, entries in groups:
            for item_id, text in entries:
                assert item_id not in seen, f"Duplicate checklist id {item_id}"
                seen.add(item_id)
                key = text.startswith("★ ")
                items.append({"id": item_id, "section": section, "sub": sub,
                              "item": text[2:] if key else text, "key": key})
    return items


CHECK_ITEMS = _checklist_items()
CHECK_ITEM_IDS = {it["id"] for it in CHECK_ITEMS}


def _to_date(v):
    if v is None:
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    try:
        return date.fromisoformat(str(v)[:10])
    except ValueError:
        return None


def _clean_text(v):
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    text = str(v).strip()
    return text or None


def load_builds():
    if USE_SUPABASE:
        # If the database is missing newer columns (Build month, QC Completed, UA/Balance), load without
        # them: they're shown blank and not saved until the column is added.
        optional = [((BUILD_MONTH,), BUILD_MONTH_COLUMN), ((QC_DATE,), QC_DATE_COLUMN),
                    (tuple(EXTRA_STATUS), EXTRA_COLUMNS)]
        try:
            df = pd.DataFrame(_fetch_all("builds", ",".join(BUILD_COLS)), columns=BUILD_COLS)
            for _cols, flag in optional:
                flag["enabled"] = True
        except Exception:
            sb = get_supabase()
            for cols_, flag in optional:     # find out which of the newer columns exist
                try:
                    sb.table("builds").select(",".join(cols_)).limit(1).execute()
                    flag["enabled"] = True
                except Exception:
                    flag["enabled"] = False
            missing = {c for cols_, flag in optional if not flag["enabled"] for c in cols_}
            cols = [c for c in BUILD_COLS if c not in missing]
            df = pd.DataFrame(_fetch_all("builds", ",".join(cols)), columns=cols)
            for c in missing:
                df[c] = None
    else:
        conn = sqlite_conn()
        try:
            df = pd.read_sql_query(f"SELECT {', '.join(BUILD_COLS)} FROM builds", conn)
        finally:
            conn.close()
    for col in (*DATE_FIELDS, BUILD_MONTH):
        df[col] = df[col].map(_to_date).astype(object)
    for col in ("company_name", "company_code", "hierarchy", "status", "built_by", "assigned_by", "qc_by", "qc_status"):
        df[col] = df[col].map(_clean_text).astype(object)
    df["company_code"] = df["company_code"].map(lambda v: v.upper() if isinstance(v, str) else None).astype(object)
    # A build (or its QC) with no status yet counts as Pending.
    for col in ("status", "qc_status"):
        df[col] = df[col].map(lambda v: v if v in BUILD_STATUSES else "Pending").astype(object)
    # UA / Balance: blank until someone picks one (no default, since they aren't counted anywhere).
    for col, allowed in EXTRA_STATUS.items():
        lookup = {a.lower(): a for a in allowed}
        df[col] = df[col].map(lambda v: lookup.get(v.strip().lower()) if isinstance(v, str) else None).astype(object)
    for col in PROPER_CASE_FIELDS:
        df[col] = df[col].map(lambda v, p=col in PERSON_FIELDS: proper_case(v, person=p)
                              if isinstance(v, str) else None).astype(object)
    df["id"] = df["id"].astype(int) if not df.empty else df["id"]
    return df


CHECK_COMMENTS = {"enabled": True}   # turns off if the database has no comments column yet
CHECK_NOTES = {"enabled": True}      # turns off if the database has no notes column yet (run supabase_add_notes.sql)


def load_build_checks():
    if USE_SUPABASE:
        # Older databases may lack the notes (and comments) columns: still show the checklist without them.
        for drop in ((), ("notes",), ("notes", "comments")):
            cols = [c for c in CHECK_COLS if c not in drop]
            try:
                df = pd.DataFrame(_fetch_all("build_checks", ",".join(cols)), columns=cols)
            except Exception:
                if len(drop) == 2:
                    raise
                continue
            for c in drop:
                df[c] = None
            CHECK_NOTES["enabled"] = "notes" not in drop
            CHECK_COMMENTS["enabled"] = "comments" not in drop
            break
    else:
        conn = sqlite_conn()
        try:
            df = pd.read_sql_query(f"SELECT {', '.join(CHECK_COLS)} FROM build_checks", conn)
        finally:
            conn.close()
    for col in ("built", "qc", "na"):
        df[col] = df[col].fillna(False).astype(bool)
    df["comments"] = df["comments"].map(_clean_text).astype(object)
    df["notes"] = df["notes"].map(_clean_text).astype(object)
    return df


def delete_builds(ids):
    """Removes builds and their checklist ticks."""
    ids = [int(i) for i in ids]
    if not ids:
        return
    if USE_SUPABASE:
        sb = get_supabase()
        sb.table("build_checks").delete().in_("build_id", ids).execute()
        res = sb.table("builds").delete().in_("id", ids).execute()
        if not res.data:
            raise RuntimeError("The database did not allow the delete. Run supabase_update_nov.sql in Supabase "
                               "to turn on deleting.")
        return
    conn = sqlite_conn()
    marks = ",".join("?" * len(ids))
    conn.execute(f"DELETE FROM build_checks WHERE build_id IN ({marks})", ids)
    conn.execute(f"DELETE FROM builds WHERE id IN ({marks})", ids)
    conn.commit(); conn.close()


def _db_value(v):
    return v.isoformat() if isinstance(v, date) else v


def _db_fields(fields):
    """Leave out UA/Balance and QC Completed when the database doesn't have those columns yet."""
    return {k: _db_value(v) for k, v in fields.items()
            if (EXTRA_COLUMNS["enabled"] or k not in EXTRA_STATUS) and (QC_DATE_COLUMN["enabled"] or k != QC_DATE)
            and (BUILD_MONTH_COLUMN["enabled"] or k != BUILD_MONTH)}


def insert_builds(rows):
    rows = [_db_fields(r) for r in rows]
    if USE_SUPABASE:
        get_supabase().table("builds").insert(rows).execute()
        return
    cols = list(BUILD_FIELDS.values())
    conn = sqlite_conn()
    conn.executemany(f"INSERT INTO builds({','.join(cols)}) VALUES({','.join(':' + c for c in cols)})",
                     [{c: r.get(c) for c in cols} for r in rows])
    conn.commit(); conn.close()


def insert_build(row):
    """Adds one build and returns its new id."""
    row = _db_fields(row)
    if USE_SUPABASE:
        res = get_supabase().table("builds").insert(row).execute()
        return int(res.data[0]["id"])
    cols = [*BUILD_FIELDS.values(), BUILD_MONTH]
    conn = sqlite_conn()
    cur = conn.execute(f"INSERT INTO builds({','.join(cols)}) VALUES({','.join(':' + c for c in cols)})",
                       {c: row.get(c) for c in cols})
    conn.commit(); new_id = cur.lastrowid; conn.close()
    return int(new_id)


def update_build(build_id, fields):
    fields = _db_fields(fields)
    if not fields:
        return
    if USE_SUPABASE:
        get_supabase().table("builds").update(fields).eq("id", int(build_id)).execute()
        return
    conn = sqlite_conn()
    conn.execute(f"UPDATE builds SET {', '.join(f'{k}=:{k}' for k in fields)} WHERE id=:_id",
                 {**fields, "_id": int(build_id)})
    conn.commit(); conn.close()


def save_build_checks(rows):
    if not rows:
        return
    if USE_SUPABASE:
        if not CHECK_COMMENTS["enabled"]:
            rows = [{k: v for k, v in r.items() if k != "comments"} for r in rows]
        if not CHECK_NOTES["enabled"]:
            rows = [{k: v for k, v in r.items() if k != "notes"} for r in rows]
        get_supabase().table("build_checks").upsert(rows, on_conflict="build_id,item_id").execute()
        return
    conn = sqlite_conn()
    if any("notes" in r for r in rows):
        rows = [{"notes": None, **r} for r in rows]
        conn.executemany("""INSERT INTO build_checks(build_id,item_id,built,qc,na,comments,notes)
            VALUES(:build_id,:item_id,:built,:qc,:na,:comments,:notes)
            ON CONFLICT(build_id,item_id) DO UPDATE SET built=excluded.built, qc=excluded.qc, na=excluded.na,
            comments=excluded.comments, notes=excluded.notes""", rows)
    else:
        conn.executemany("""INSERT INTO build_checks(build_id,item_id,built,qc,na,comments)
            VALUES(:build_id,:item_id,:built,:qc,:na,:comments)
            ON CONFLICT(build_id,item_id) DO UPDATE SET built=excluded.built, qc=excluded.qc, na=excluded.na,
            comments=excluded.comments""", rows)
    conn.commit(); conn.close()


# "Checklist uploaded" tick for a build. Stored as one extra row in build_checks (built = uploaded),
# so it needs no database change; it isn't a checklist item, so the counts and the Excel ignore it.
UPLOAD_ITEM = "_UPLOADED"


def _uploaded_ids(checks):
    """IDs of the builds whose checklist is ticked as uploaded."""
    return {int(i) for i in checks[(checks.item_id == UPLOAD_ITEM) & checks.built].build_id}


def _progress(checks, build_id):
    c = checks[(checks.build_id == build_id) & checks.item_id.isin(CHECK_ITEM_IDS)]   # ignore removed items
    na = int(c.na.sum())
    applicable = len(CHECK_ITEMS) - na
    return int((c.built & ~c.na).sum()), int((c.qc & ~c.na).sum()), applicable


def _builds_table(frame):
    rows = []
    for r in frame.itertuples():
        rows.append({
            "id": r.id, "Company Name": r.company_name, "Company Code": r.company_code, "Hierarchy": r.hierarchy,
            "Status": r.status, "Build Completed": r.completion_date, "QC Completed": r.qc_completion_date,
            "Due Date": r.due_date,
            "Built By": r.built_by, "Assigned By": r.assigned_by, "QC": r.qc_by, "QC Status": r.qc_status,
            "UA Status": r.ua_status, "Balance Status": r.balance_status,
        })
    cols = ["id", *BUILD_FIELDS]
    table = pd.DataFrame(rows, columns=cols)
    # Keep every column as plain values with blanks as None. A column that is blank for the whole month
    # (e.g. Hierarchy) would otherwise turn into a number column and the table refuses to edit it as text.
    for col in BUILD_FIELDS:
        table[col] = pd.Series([None if _clean_text(v) is None else v for v in table[col]],
                               index=table.index, dtype=object)
    return table


def _norm(field, value):
    if field in DATE_FIELDS:
        return _to_date(value)
    text = _clean_text(value)
    if field == "company_code" and text:
        return text.upper()          # company codes are always upper case, e.g. 1kx -> 1KX
    if field in PROPER_CASE_FIELDS:
        return proper_case(text, person=field in PERSON_FIELDS)     # e.g. BUTLER DD IT NC -> Butler DD IT NC
    return text


def _write_sheet(ws, title, headers, rows, date_cols=(), status_cols=()):
    """One formatted sheet: a title, a yellow header row, coloured statuses and MM/DD/YYYY dates."""
    from openpyxl.styles import Font, PatternFill

    ws.append([title])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    ws.append(headers)
    for c in ws[3]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFE699")
    fills = {"Completed": "C6EFCE", "In Progress": "FFEB9C"}
    for row in rows:
        ws.append(row)
        for col_name in status_cols:
            cell = ws.cell(row=ws.max_row, column=headers.index(col_name) + 1)
            if cell.value in fills:
                cell.fill = PatternFill("solid", fgColor=fills[cell.value])
        for col_name in date_cols:
            ws.cell(row=ws.max_row, column=headers.index(col_name) + 1).number_format = "mm/dd/yyyy"
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(max(len(str(x.value or "")) for x in col[2:]) + 2, 45)
    ws.freeze_panes = "A4"


def _builds_rows(table):
    headers = [*BUILD_FIELDS]
    return headers, [[py(r[h]) if h not in DATE_COLUMNS else r[h] for h in headers]
                     for _, r in table.iterrows()]


def export_builds_excel(table, label):
    from openpyxl import Workbook

    wb = Workbook()
    ws = wb.active
    ws.title = "Company Builds"
    headers, rows = _builds_rows(table)
    _write_sheet(ws, f"Company Builds – {label}", headers, rows,
                 date_cols=DATE_COLUMNS, status_cols=("Status", "QC Status"))
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


EXCEL_ICON = ('<svg width="22" height="22" viewBox="0 0 24 24" aria-hidden="true">'
              '<rect x="7" y="2" width="15" height="20" rx="2" fill="#21A366"/>'
              '<rect x="14" y="2" width="8" height="20" rx="2" fill="#33C481"/>'
              '<rect x="2" y="6" width="12" height="12" rx="2" fill="#107C41"/>'
              '<path d="M5 9l2.2 3L5 15h1.7l1.3-2 1.3 2H11l-2.2-3L11 9H9.3L8 11 6.7 9z" fill="#fff"/></svg>')


def _excel_link(data, file_name, label, tip, align="right", big=False):
    """An 'Excel' download link, right-aligned by default (a plain link, so it can carry the Excel icon).
    big=True makes it a larger button with a little space below it."""
    b64 = base64.b64encode(data).decode()
    mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    icon_px, pad, font = (19, "6px 14px", "0.92rem") if big else (15, "3px 10px", "0.8rem")
    icon = EXCEL_ICON.replace('width="22" height="22"', f'width="{icon_px}" height="{icon_px}"')
    lift = "margin-top:-14px" if align == "right" else ""
    below = "margin-bottom:10px" if big else ""
    return (f'<div style="text-align:{align};{lift};{below}"><a download="{html.escape(file_name)}" '
            f'title="{html.escape(tip)}" href="data:{mime};base64,{b64}" style="display:inline-flex;'
            f'align-items:center;gap:8px;padding:{pad};font-size:{font};border:1px solid rgba(128,128,128,.35);'
            f'border-radius:6px;text-decoration:none;color:inherit;font-weight:500;white-space:nowrap">'
            f'{icon}{html.escape(label)}</a></div>')


def export_checklist_excel(b, saved):
    """One build's QC checklist as a printable Excel file: build details on top, then every section and item
    with its Built / QC / N/A ticks and comments."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    def txt(v):
        return _clean_text(v) or ""

    def day(v):
        d = _to_date(v)
        return f"{d:%m/%d/%Y}" if d else ""

    wb = Workbook()
    ws = wb.active
    ws.title = "QC Checklist"
    bold = Font(bold=True)
    thin = Side(style="thin", color="BFBFBF")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    center = Alignment(horizontal="center", vertical="top")

    name = txt(b.company_name)
    ws.append([f"QC Checklist – {name}"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:G1")
    status, qc_status = txt(b.status) or "Pending", txt(b.qc_status) or "Pending"
    applicable = [it for it in CHECK_ITEMS if not (saved.get(it["id"]) is not None and saved[it["id"]].na)]
    n_built = sum(1 for it in applicable if saved.get(it["id"]) is not None and saved[it["id"]].built)
    n_qc = sum(1 for it in applicable if saved.get(it["id"]) is not None and saved[it["id"]].qc)
    details = [("Company Code", txt(b.company_code)), ("Hierarchy", txt(b.hierarchy)),
               ("Built By", txt(b.built_by)), ("Status", status),
               ("QC'd By", txt(b.qc_by)), ("QC Status", qc_status),
               ("Due Date", day(b.due_date)), ("Build Completed", day(b.completion_date)),
               ("QC Completed", day(getattr(b, QC_DATE, None))),
               ("Progress", f"Built {n_built}/{len(applicable)} · QC {n_qc}/{len(applicable)}"),
               ("Downloaded", f"{now_local():%m/%d/%Y %I:%M %p} IST")]
    ws.append([])
    for i in range(0, len(details), 2):     # two label/value pairs per row
        row = []
        for label, value in details[i:i + 2]:
            row += [label, value, ""]
        ws.append(row[:6])
        for col in (1, 4):
            ws.cell(row=ws.max_row, column=col).font = bold
    ws.append([])

    headers = ["Area", "Item", "Key", "N/A", "Built", "Notes", "QC", "Errors"]
    ws.append(headers)
    head_row = ws.max_row
    for c in ws[head_row]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
        c.border = box
        c.alignment = center
    ws.freeze_panes = f"A{head_row + 1}"

    section_fill = PatternFill("solid", fgColor="D9E1F2")
    done_fill = PatternFill("solid", fgColor="C6EFCE")
    na_fill = PatternFill("solid", fgColor="EDEDED")
    for section, note, _groups in CHECKLIST:
        items = [it for it in CHECK_ITEMS if it["section"] == section]
        app_items = [it for it in items if not (saved.get(it["id"]) is not None and saved[it["id"]].na)]
        sb = sum(1 for it in app_items if saved.get(it["id"]) is not None and saved[it["id"]].built)
        sq = sum(1 for it in app_items if saved.get(it["id"]) is not None and saved[it["id"]].qc)
        ws.append([f"{section}  —  Built {sb}/{len(app_items)} · QC {sq}/{len(app_items)}"])
        r = ws.max_row
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
        ws.cell(row=r, column=1).font = bold
        for col in range(1, 9):
            ws.cell(row=r, column=col).fill = section_fill
            ws.cell(row=r, column=col).border = box
        if note:
            ws.append([note])
            r = ws.max_row
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
            ws.cell(row=r, column=1).font = Font(italic=True, size=9, color="595959")
            ws.cell(row=r, column=1).alignment = wrap
            ws.row_dimensions[r].height = 15 * max(1, len(note) // 140 + 1)
        for it in items:
            s = saved.get(it["id"])
            built, qc, na = (bool(s.built), bool(s.qc), bool(s.na)) if s is not None else (False, False, False)
            ws.append([it["sub"], it["item"], "★" if it["key"] else "", "N/A" if na else "",
                       "✓" if built else "", txt(getattr(s, "notes", None)) if s is not None else "",
                       "✓" if qc else "", txt(s.comments) if s is not None else ""])
            r = ws.max_row
            for col in range(1, 9):
                cell = ws.cell(row=r, column=col)
                cell.border = box
                cell.alignment = center if col in (3, 4, 5, 7) else wrap
            if na:
                for col in range(1, 9):
                    ws.cell(row=r, column=col).fill = na_fill
            else:
                for col, ok in ((5, built), (7, qc)):
                    if ok:
                        ws.cell(row=r, column=col).fill = done_fill

    for col, width in zip("ABCDEFGH", (30, 70, 6, 7, 8, 40, 8, 40)):
        ws.column_dimensions[col].width = width
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = f"{head_row}:{head_row}"
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def export_all_builds_excel(builds, raines):
    """Every build to date: a Summary sheet, all company builds, and the Raines builds."""
    from openpyxl import Workbook
    from openpyxl.styles import Font

    table = _builds_table(builds.sort_values("id"))     # in the order they were added
    done = int((builds.status == "Completed").sum())
    qc_done = int(((builds.status == "Completed") & (builds.qc_status == "Completed")).sum())
    stamp = f"{today_local():%m/%d/%Y}"

    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    summary = [
        ("Company builds", len(builds)), ("Raines builds", len(raines)),
        ("Total builds", len(builds) + len(raines)),
        ("Completed (incl. Raines)", done + len(raines)),
        ("In progress", int((builds.status == "In Progress").sum())),
        ("Pending", int((builds.status == "Pending").sum())),
        ("Not Required", int((builds.status == NOT_REQUIRED).sum())),
        ("QC completed (company builds)", qc_done),
    ]
    _write_sheet(ws, f"All builds to date – as of {stamp}", ["", "Builds"], [list(r) for r in summary])
    ws.cell(row=6, column=1).font = Font(bold=True)    # Total builds
    ws.cell(row=6, column=2).font = Font(bold=True)

    headers, rows = _builds_rows(table)
    _write_sheet(wb.create_sheet("Company Builds"), f"Company Builds – all months (as of {stamp})", headers, rows,
                 date_cols=DATE_COLUMNS, status_cols=("Status", "QC Status"))

    r = raines.sort_values(["completion_date", "company_name"])
    _write_sheet(wb.create_sheet("Raines Builds"), "Raines Builds",
                 ["Company Code", "Company Name", "Built By", "Status", "Build Completed"],
                 [[x.company_code, x.company_name, x.built_by, "Completed", x.completion_date] for x in r.itertuples()],
                 date_cols=("Build Completed",), status_cols=("Status",))
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _build_months(builds):
    """The month each build belongs to: its due date's month, else the month it was added under in
    Monthly Builds, else the current month."""
    current = month_start(today_local())
    added = builds[BUILD_MONTH] if BUILD_MONTH in builds.columns else pd.Series(None, index=builds.index)

    def pick(due, mo):
        due, mo = _to_date(due), _to_date(mo)
        return month_start(due) if due else (month_start(mo) if mo else current)
    return pd.Series([pick(d, m) for d, m in zip(builds.due_date, added)], index=builds.index, dtype=object)


def _month_frame(builds, choice):
    """Builds in the chosen month (see _build_months)."""
    return builds[_build_months(builds) == choice]


def _filter_builds(table, key):
    """A row of filters above a builds table: search by name/code, and pick statuses or people.
    Empty filters show everything. Returns only the matching rows."""
    def people_in(col):
        return sorted({p for v in table[col] for p in _people(v)})

    def values_in(col):
        return sorted({v for v in table[col] if isinstance(v, str) and v.strip()}, key=str.lower)

    # One clean row: a search bar, a "Filters" button that opens a small panel, and Clear.
    keys = {"Status": f"flt_st_{key}", "QC Status": f"flt_qs_{key}", "UA Status": f"flt_ua_{key}",
            "Balance Status": f"flt_bal_{key}", "Built By": f"flt_bb_{key}",
            "QC": f"flt_qb_{key}", "Assigned By": f"flt_ab_{key}",
            "Company Name": f"flt_cn_{key}", "Company Code": f"flt_cc_{key}"}
    q_key = f"flt_q_{key}"
    n_active = sum(bool(st.session_state.get(k)) for k in keys.values())

    def clear():
        for k in keys.values():
            st.session_state[k] = []
        st.session_state[q_key] = ""

    c1, c2, c3 = st.columns([5, 1.2, 0.8], vertical_alignment="bottom")
    q = c1.text_input("Search", key=q_key, placeholder="🔎  Search company name or code",
                      label_visibility="collapsed")
    with c2.popover(f"Filters · {n_active}" if n_active else "Filters", icon=":material/filter_list:",
                    width="stretch"):
        picks = {
            "Status": st.multiselect("Status", BUILD_STATUSES, key=keys["Status"],
                                     format_func=STATUS_LABEL.get, placeholder="All"),
            "QC Status": st.multiselect("QC Status", BUILD_STATUSES, key=keys["QC Status"],
                                        format_func=STATUS_LABEL.get, placeholder="All"),
            "UA Status": st.multiselect("UA Status", UA_STATUSES, key=keys["UA Status"],
                                        format_func=STATUS_LABEL.get, placeholder="All"),
            "Balance Status": st.multiselect("Balance Status", BALANCE_STATUSES, key=keys["Balance Status"],
                                             format_func=STATUS_LABEL.get, placeholder="All"),
            "Built By": st.multiselect("Built By", people_in("Built By"), key=keys["Built By"], placeholder="All"),
            "QC": st.multiselect("QC", people_in("QC"), key=keys["QC"], placeholder="All"),
            "Assigned By": st.multiselect("Assigned By", people_in("Assigned By"), key=keys["Assigned By"],
                                          placeholder="All"),
            "Company Name": st.multiselect("Company Name", values_in("Company Name"), key=keys["Company Name"],
                                           placeholder="All"),
            "Company Code": st.multiselect("Company Code", values_in("Company Code"), key=keys["Company Code"],
                                           placeholder="All"),
        }
    c3.button("Clear", key=f"flt_clear_{key}", on_click=clear, width="stretch",
              disabled=not (n_active or st.session_state.get(q_key)))
    keep = pd.Series(True, index=table.index)
    if q.strip():
        text = q.strip().lower()
        keep &= table.apply(lambda r: text in str(r["Company Name"] or "").lower()
                            or text in str(r["Company Code"] or "").lower(), axis=1)
    for col, chosen in picks.items():
        if not chosen:
            continue
        if col in ("Status", "QC Status", "UA Status", "Balance Status", "Company Name", "Company Code"):
            keep &= table[col].isin(chosen)
        else:
            keep &= table[col].map(lambda v: any(p in chosen for p in _people(v)))
    if not keep.all():
        on = [f"{col}: {', '.join(v)}" for col, v in picks.items() if v] + ([f'"{q.strip()}"'] if q.strip() else [])
        st.caption(f"Showing **{int(keep.sum())} of {len(table)}** builds · " + " · ".join(on))
    return table[keep].reset_index(drop=True)     # a plain 0..n index, so adding rows still works


def _builds_editor(builds, frame, key, label, month=None):
    """Editable builds table: click a cell to edit, add a row at the bottom, or select rows and delete them.
    Edits and new rows save automatically in the background; deleting asks for confirmation first."""
    frame = frame.sort_values("id")      # in the order they were added; new builds go at the bottom
    table = _filter_builds(_builds_table(frame), key)
    fresh = table
    st.caption("Click a cell to edit (saves automatically, keep going) · add a build in the empty last row · "
               "to delete, tick the box at the left of a row and click 🗑️")
    save_bar = st.container()           # save status and delete confirmation show here
    if USE_SUPABASE and not QC_DATE_COLUMN["enabled"]:
        st.warning("QC Completed dates can't be saved yet: the database has no qc_completion_date column. "
                   "In Supabase → SQL Editor, run: alter table builds add column if not exists "
                   "qc_completion_date date; notify pgrst, 'reload schema';  — then refresh this page.")

    def with_extras(base, *cols):
        """Dropdown choices: the standard list, plus any other names already saved (e.g. older builds)."""
        extra = sorted({v for c in cols for v in builds[c] if isinstance(v, str) and v not in base})
        return list(base) + extra

    team = with_extras(TEAM_MEMBERS, "built_by", "qc_by")
    leads = with_extras(PAYROLL_LEADS, "assigned_by")
    # Each change saves in the background while the table stays as it is, so you can keep making
    # entries one after another. The table only reloads from the database when it needs to show
    # something new: a build was added or deleted, a date was filled in, or the filters changed.
    version = st.session_state.get("builds_editor_version", 0)
    filt_sig = hashlib.md5(repr(sorted((k, str(v)) for k, v in st.session_state.items()
                                       if isinstance(k, str) and k.startswith("flt_") and k.endswith(f"_{key}")))
                           .encode()).hexdigest()[:8]
    tag = f"{key}_{version}_{filt_sig}"
    for old_k in [k for k in st.session_state if isinstance(k, str)
                  and k.startswith((f"builds_snap_{key}_", f"builds_last_{key}_")) and not k.endswith(tag)]:
        del st.session_state[old_k]
    table = st.session_state.setdefault(f"builds_snap_{tag}", table)   # what the table shows until it reloads
    _keep_scroll_position()     # so a reload doesn't jump you back to the top of the page or the table
    keep_box = st.container(key=f"cbkeep_builds_{key}")
    edited = keep_box.data_editor(
        table, key=f"builds_editor_{tag}", hide_index=True, width="stretch", num_rows="dynamic",
        column_order=BUILD_COLUMN_ORDER,
        height=min(35 * len(table) + 75, 700),
        column_config={
            "id": None,
            "Company Name": st.column_config.TextColumn("Company Name", width=220, required=True, pinned=True),
            "Company Code": st.column_config.TextColumn("Company Code", width=115,
                                                        help="Enter after you create the company"),
            "Hierarchy": st.column_config.TextColumn("Hierarchy", width=130),
            "Status": st.column_config.SelectboxColumn("Status", options=BUILD_STATUSES, default="Pending", width=125,
                                                       format_func=lambda s: STATUS_LABEL.get(s, s)),
            "Build Completed": st.column_config.DateColumn("Build Completed", format="MM/DD/YYYY", width=120),
            "QC Completed": st.column_config.DateColumn(
                "QC Completed", format="MM/DD/YYYY", width=120,
                help="Filled in automatically when QC Status is set to Completed. You can also edit it."),
            "Due Date": st.column_config.DateColumn("Due Date", format="MM/DD/YYYY", width=100),
            "Built By": st.column_config.SelectboxColumn("Built By", options=team, width=100),
            "Assigned By": st.column_config.SelectboxColumn("Assigned By", options=leads, width=105),
            "QC": st.column_config.SelectboxColumn("QC", options=team, width=100),
            "QC Status": st.column_config.SelectboxColumn("QC Status", options=BUILD_STATUSES, width=125,
                                                          format_func=lambda s: STATUS_LABEL.get(s, s)),
            "UA Status": st.column_config.SelectboxColumn(
                "UA Status", options=UA_STATUSES, width=150, format_func=lambda s: STATUS_LABEL.get(s, s),
                help="User Access status. For tracking only, not counted in the totals above"),
            "Balance Status": st.column_config.SelectboxColumn(
                "Balance Status", options=BALANCE_STATUSES, width=150, format_func=lambda s: STATUS_LABEL.get(s, s),
                help="For tracking only, not counted in the totals above"),
        })

    originals = {int(r["id"]): r for _, r in table.iterrows()}
    # What each build looked like when it was last saved, so only real changes are written.
    last = st.session_state.setdefault(f"builds_last_{tag}", {
        rid: {f: _norm(f, r[c]) for c, f in BUILD_FIELDS.items()} for rid, r in originals.items()})
    inserts, updates, problems = [], [], []
    for _, r in edited.iterrows():
        values = {f: _norm(f, r[c]) for c, f in BUILD_FIELDS.items()}
        rid = r.get("id")
        is_new = rid is None or pd.isna(rid)
        if is_new and not any(values.values()):
            continue
        if not values["company_name"]:
            problems.append("Every build needs a Company Name. That row saves once it has one.")
            continue
        if is_new:
            if month is not None:
                values[BUILD_MONTH] = month      # keeps it in this month, even without a due date
                # Without the build_month column in the database, a build with no due date would land in
                # the current month and vanish from this one, so it waits for its Due Date instead.
                if (not BUILD_MONTH_COLUMN["enabled"] and month != month_start(today_local())
                        and not values["due_date"]):
                    problems.append(f"Add a **Due Date** in {month:%B %Y} to the new build "
                                    f"({values['company_name']}) to save it here.")
                    continue
            values["status"] = values["status"] or "Pending"
            inserts.append(values)
        else:
            diff = {k: v for k, v in values.items() if v != last[int(rid)][k]}
            if diff:
                updates.append((int(rid), diff, values))

    kept = {int(i) for i in edited["id"] if i is not None and not pd.isna(i)}
    removed = [rid for rid in originals if rid not in kept]

    def reload_table():
        st.session_state["builds_editor_version"] = version + 1
        st.rerun()

    with save_bar:
        for msg in dict.fromkeys(problems):
            st.warning(msg)

        # Deleting: rows removed in the table are only deleted from the database after confirming.
        if removed:
            names = ", ".join(str(originals[rid]["Company Name"]) for rid in removed[:5]) + ("…" if len(removed) > 5 else "")
            st.error(f"Delete {len(removed)} build{'s' if len(removed) != 1 else ''}: **{names}**? "
                     "This also removes their QC checklist and can't be undone.")
            d1, d2, _ = st.columns([1, 1, 3])
            if d1.button("🗑️ Yes, delete", type="primary", key=f"del_yes_{key}_{version}"):
                try:
                    delete_builds(removed)
                    st.session_state["_toast"] = f"🗑️ Deleted {len(removed)} build{'s' if len(removed) != 1 else ''}"
                except Exception as e:
                    st.session_state["_flash_error"] = str(e)
                reload_table()
            if d2.button("Cancel", key=f"del_no_{key}_{version}"):
                reload_table()

    # Save in the background: each change is written as soon as it's made, without reloading the table.
    if inserts or updates:
        today = today_local()
        needs_reload = bool(inserts)        # a new build needs its id before it can be edited further
        try:
            for rid, diff, values in updates:
                # A date fills in only at the moment its status is changed to Completed.
                if "status" in diff and values["status"] == "Completed" and not values["completion_date"]:
                    diff["completion_date"] = values["completion_date"] = today
                    needs_reload = True     # show the date that was just filled in
                if "qc_status" in diff and values["qc_status"] == "Completed" and not values[QC_DATE]:
                    diff[QC_DATE] = values[QC_DATE] = today
                    needs_reload = True
                update_build(rid, diff)
                last[rid] = values
            for values in inserts:
                if values["status"] == "Completed" and not values["completion_date"]:
                    values["completion_date"] = today
                if values["qc_status"] == "Completed" and not values[QC_DATE]:
                    values[QC_DATE] = today
                insert_build(values)
            st.session_state[f"builds_saved_at_{key}"] = now_local()
        except Exception as e:
            st.error(f"Could not save your last change, it will retry on your next edit. ({e})")
            needs_reload = False
        if needs_reload:
            if inserts:
                st.session_state["_toast"] = f"✅ Added {len(inserts)} new build{'s' if len(inserts) != 1 else ''}"
            reload_table()
    saved_at = st.session_state.get(f"builds_saved_at_{key}")
    save_bar.caption(f"✓ Saved automatically at {saved_at:%I:%M:%S %p}" if saved_at else "Changes save automatically.")
    st.download_button(f"📥 Download Excel · {label}", data=export_builds_excel(fresh, label),
                       file_name=f"Company_Builds_{label.replace(' ', '_')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       key=f"builds_download_{key}")




# ── Reference sheets (live Excel files on SharePoint) ──
# The live Excel sheet kept by the US team (SharePoint). The app only links to it, so everyone always
# sees the latest version; opening it needs a Netchex Microsoft sign-in. Change the link here if it moves.
FEATURE_FLAGS_URL = ("https://netchex849-my.sharepoint.com/:x:/r/personal/ckrebs_netchexonline_com/Documents/"
                     "Feature%20Flags%20Index.xlsx?d=wd0129a38e09943edb0c01076a4ab323d&csf=1&web=1&e=eV6ddh")
FF_HELP = "Opens the live Feature Flags Index (Excel) in a new tab. Sign in with your Netchex account."
SUTA_RATES_URL = ("https://netchex849-my.sharepoint.com/:x:/r/personal/fali_netchexonline_com/Documents/Documents/"
                  "Important%20Documents/2026_State_Unemployment_New_Employer_Rates.xlsx"
                  "?d=w62feec7e76fc41e4a2f67b41d3fbf5d2&csf=1&web=1&e=mxviGy")

# Links shown at the top right of a checklist section: section name starts with -> (label, url, tooltip)
SECTION_LINKS = {
    "Netchex - Tax Codes": [("2026 SUTA New Employer Rates", SUTA_RATES_URL,
                             "Opens the 2026 State Unemployment new employer rates (Excel) in a new tab. "
                             "Sign in with your Netchex account.")],
    "Netchex - Claims": [("Feature Flags Index", FEATURE_FLAGS_URL, FF_HELP)],
    "Netchex - Feature Flags": [("Feature Flags Index", FEATURE_FLAGS_URL, FF_HELP)],
}

# "How to" notes, opened from a small button at the top right of a checklist section so the table stays clean.
# From the Company Build Checklist instructions (v1.0, Oct 2026). ⚠️ = not confirmed yet.
SECTION_HOWTO = {
    "Netchex - Tax Codes": """
**Default state codes** · Maintenance > Payroll > Tax Codes
- Never delete FED EIC, FED INCOME TAX, FUTA - 940, MEDICARE or SOCIAL SECURITY.
- No Louisiana employees? Delete LA State and LA SUTA. Louisiana employees? Keep and update them.
- *Why: unused states lead to employees coded to the wrong state, often found only at W-2 time.*

**State withholding code (each state)** · Salesforce Knowledge > Tax Setup Guide > [state]
- Tax Code `GA` · Description "GA State" · Tax Table = state income tax · Tax Mapping from the list.
- W-2 ID = state ID (FEIN if none; never blank). SUTA ID blank. Employer and SDI rates 0.000.
- Payee MasterTax · Type State Tax · Reference `GA0000-001` · default deposit frequency.
- No-income-tax states (e.g. TX) still get a code. Ask your lead about long local lists (PA, OH).

**SUTA code (each state)** · Maintenance > Payroll > Tax Codes > Add
- Tax Code `GAUI` · Description "GA SUTA" · Tax Table = state unemployment insurance.
- SUTA ID = UI account number; blank if unknown, never "Applied For".
- Employer rate as a decimal (2.62% = 0.0262); state default until the client's notice.
- Payee MasterTax · Reference `GA0000-010` · End date 12/31/9999 · rate note "Initial Rate Load".
- ⚠️ *Not confirmed yet: SUTA W-2 ID = FEIN or UI account number.*

**Workers' comp codes** · Maintenance > Payroll > Workers' Compensation Codes > Add
- One per state + class, always state-prefixed: `LA8810`, `GA8810`. Description "LA Clerical".
- Cost Factor and Bill Rate 1 only if the client gave rates (decimal); Bill Rates 2–6 = 0.
- Payee = the company · Payable GL = Payable - Workers' Comp · Expense GL = Employer Workers' Comp Expense · Hazard Group 2.

**Hourly + salary job codes** · Maintenance > Human Resources > Job Codes
- For every WC code in every state, build Hourly and Salary. Pick Salary Grade first (Hourly Grade 1 = H; Exempt Grade 1 = E).
- Code `LA8810` · "LA Clerical Hourly" / "LA Clerical Salary" · Employee Type Hourly/Exempt · Job Group Office/Clerical · link the WC code.
- Update or delete the default LA job codes.
- *Why: the job code makes an employee hourly or exempt.*
""",
    "Salesforce - Details": """
**Record type** · Salesforce > Account > Details > Account Record Type > Change
- Click Change next to Prospect Account > Client Account > Next. Click Edit on any field to open the form.

**Account information** · Details > Account Information
- Account Name and Federal ID match the PIN exactly. DBA Name only if they use one.
- Client Company Code = the new Netchex code (between DBA and Federal ID).
- Parent Account only for a child code (build the parent first; check Ultimate Parent Account on the parent).
- Type = Netchex Client · Pricing Model = PEPM · Package only for DealerTrack.
- Distributed Invoicing and Seasonal Client = No unless your lead says Yes.
- ⚠️ *Not confirmed yet: whether the Package field still appears on newer accounts.*

**Average payroll** · Details > Account Financial Status
- Average Payroll Amount from the analysis document. Not known yet? Enter 0.00 and update before the first direct deposit payroll.
- *Why: Banking calculates the risk limit from it. Payrolls over it need a debit date override, a wire, or an earlier submission.*

**Points of contact + address** · Details > Designated Points of Contact
- Payroll, ACH/Banking, Tax, Billing, Human Resources and The Work Number contacts (Benefits if purchased).
- One person per field; a client email group can be a contact.
- Shipping Address from the analysis document's Delivery tab.
- ⚠️ *Not confirmed yet: who fills the Billing Address and from which document.*

**Revenue + implementation details** · Details > Client Information / Implementation Details
- Netchex Annual Revenue = monthly PEPM from the proposal × 12 (not Hierarchy ARR).
- Direct Deposit Waiting Period: No + Proof of DD Received if proof given; Yes for a startup.
- AZ/HI Daylight Savings: check only for Arizona or Hawaii.
- Check Currently in Implementation (routes cases to the Implementation queue).

**Employer services** · Details > Employer Services Information
- Primary (and secondary) Payroll Frequency.
- Payroll Vendor = the client's previous provider; Startup if none. Broker and 401k details if known.
- Copy/child company: match the parent's services. Click Save.
""",
    "Salesforce - Contacts": """
**Contacts + user access** · Account > Contacts > View All
- Add every admin on the User Access form plus any your lead names: New Contact > Client Contact.
- Last Name and Title are required; get the email.
- Contact already exists on another account? Use Add Relationship instead of a new contact.
- Check **Able to Request User Access** for Client Admin Form users. At least one contact must have it.
- *Why: Netchex holds SSNs, 941s, 940s and state account numbers. Support needs to know who can approve access.*
""",
    "Salesforce - Assets": """
**Payroll assets** · Account > Assets > View All > [asset] > Edit
- For each payroll asset on the proposal: Status Pending > Implemented/In Use > Save.
- Asset missing? Click New and add it.
- Yours: Payroll/Tax, ESS/MSS, Time & Attendance, The Work Number, Engagement/Recognition portal.
- Leave HR assets (Onboarding & Workflows, Recruit, Performance, NetLearn, Insights, ExpensePath, AskHR, benefits products, E-Verify) for the HR Onboarding Lead.
""",
    "Salesforce - Onboarding Details": """
**Onboarding team** · Account > Onboarding > Onboarding record > Onboarding Details
- Owner (Change Owner) = Project Manager, if the account has one.
- Data Specialist = Project Coordinator.
- Payroll Information > Payroll Onboarding Lead > Save.
- HR Details tab > HR Onboarding Lead > Save.
- ⚠️ *Not confirmed yet: field labels are being renamed, and whether survey contacts are still used.*
""",
    "Netchex - Client Billing": """
**Client billing schedule** · Back Office > Enterprise > Client Billing Schedules
- Payroll Frequency. Collect Taxes checked unless a No-Tax client.
- Sales Tax box: Texas only. Sales Tax Code: TX = Dallas 8.25%; others = their state.
- Sales Agent = Addie Kelley · "Payroll Coordinator" = the Project Lead.
- Check Fee – Premium and enter negatives (Base / Per Check), then Add/Update:
  - Weekly −18.00 / −1.35
  - Bi-Weekly & Semi-Monthly −23.00 / −2.10
  - Monthly −43.00 / −3.65

**Premium Pay billing** (Premium Pay only) · Client Billing Schedules
- Check Premium Pay, uncheck Netchex Money Movement, zero the $25 stop payment fee.

**Check for errors** · Maintenance > Company Structure > General > Client Billing View
- Correct entries show no data here. Anything listed is an error to fix.
""",
    "Netchex - Company Information": """
**Demographics** · Maintenance > Company Information > code + Enter > company
- Company, Trade and W-2 Company Name = legal name from the PIN, ALL CAPS (only the display name may show a DBA).
- FEIN and address match the PIN exactly. Primary contact, phone and email entered.
- ⚠️ *Not confirmed yet: which name field must match the PIN.*

**Banking + PIN** · Company Information > Banking Information
- ACH, Live Check, Operating and Billing Fee accounts match the Funding Authorization Agreement.
- ACH Contact = the 4-digit EFTPS PIN.
- Minimum Wage Makeup: restaurants and hotels only.

**Payroll information** · Company Information > Payroll Information
- Minimum Wage $ = state (one state) or federal (more than one); look it up in Back Office > Minimum Wage.
- Default Payroll Group (BiWklyHrly, WklyHrly, SemiHrly, MnlyHrly) and Default Job Code.
- Default Employee Classification: W-2 Tipped (restaurant/hotel), 1099 (contractor-only code), otherwise W-2.
- Units/Pay Period 40 / 80 / 86.6667 / 173.3334. Allow Manager Self Service checked.
- Premium Pay only: check Netchex Premium Pay Client.

**Direct deposit date** · Company Information > Direct Deposit
- Proof of prior DD: Effective Date = implementation start date.
- No proof or startup: Effective Date = 90 days after the first check date.
- Live checks only, forever: Opt-Out.
- Section missing? Confirm the Direct Deposit Authorization feature flag is on.
- ⚠️ *Not confirmed yet: when Opt-Out is used.*
""",
    "Netchex - Claims": """
**Claims per proposal** · Company Information > Company Claims
- Engagement Basic for every client. Engagement builds up: Basic > Core > Advanced.
- Recruit = "Recruit – HigherMe" (never GetHired).
- Other claims per the Feature Flags Index (link at the top right of this section).
- Click **Update** at the top of the page.
- Demo companies: turn on almost everything; skip Daily Pay, NetEngage, Delegate, NetAssist HR.
- ⚠️ *Not confirmed yet: which claims the builder sets vs the HR Onboarding Lead.*
""",
    "Netchex - Feature Flags": """
*Source: the live Feature Flags Index (link at the top right), updated 11/16/2026. If anything here differs, the Index wins.*

**All clients** · Maintenance > Security > New Feature Access > [company]
- **New Payroll Engine** is NOT auto-on for new clients. If it's missing: uncheck Multi-Company Payroll, check New Payroll Engine, then re-check Multi-Company Payroll. Review & Submit won't show without it.
- **Multi-Company Payroll (3012)** with **Payroll Dashboard**: MCP alone shows mock data. Takes ~15 min; make sure no payroll is in progress and the client reviews the payroll schedules.
- **Payroll Review and Submit (3000)** + **Review and Submit UI Refresh (3033)**, keep both for now. Claims: Employee Check Preview, Payroll Comparison. If it blocks a payroll, turn it off first.
- New for all clients: **EditTimesheets (3035)**, **Payroll Status (3036)**, **VoidRecreateCheck (3034)**, **Start Payroll Redesign (3037)**, **Automated 401k Catch-up (3040)**.
- **Locals Tax** (confirmed for all new companies), **Payroll Process Timesheet Service** (needed for FlexPay), **Invoice PDF Attachment**.
- Already on by default, just verify: Payroll Schedules, Suggested Taxes, Secure Act 2.0, Recognition Settings - Organization, Minimum Hours Worked, Accrual Expiration, Badge Punch Decoupled, Who's In Dashboard, Time Off Policies, Automatic Holidays, Payday Notification, Direct Deposit Authorization, New Scheduler.
- **Engagement Phase One** with claims Engagement Basic + Engagement Core. ⚠️ *Claim names not confirmed yet (SOP says Community Basic + Community Core).*
- **Earned Wage Access**: on by default. ⚠️ *Not confirmed for every client yet; check legacy Rain clients before changing it.* FlexPay claim only if on the PEPM and not on the SF opt-out list.
- **Punch Request** for Time & Attendance clients · **Employee Profile** (user level): verify every employee has it after the import.
- Turn a flag on: toggle > Pending Changes > Confirm Changes.

**If on the PEPM proposal**
- Scheduler: Scheduler claim (System Tables > company) + New Scheduler + Who's In Dashboard; open a user access case for admin NetScheduler access.
- AskHR: AskHR flag + Community Basic, then Community Core claims + security 5320. Upload docs, client confirms access, then email/Zoom Tanya Jain to enable AI (1–3 days). ⚠️ *AskHR - Engagement Dashboard: not confirmed yet.*
- Netchex AI: Netchex AI flag + NetchexAI claim, then Zoom Tanya Jain with the company code.
- Insights: Insights flag + security groups 5090 and 5110.
- NetLearn: LMS + NetLearn - Manage Users, after the BizLibrary setup (3–7 days).
- Recruit: assign Org Admin + Configure Recruit first, Recruit - HigherMe claim, then HigherMe - Recruiting UI Access **last**.
- Communications flag (Announcements, Conversations) + Workplace Communication claim for Conversations.
- Benefits Enablement Experience for new Benefits clients · ExpensePath (employee level) + API integration.
- Onboarding & Workflows is included with Core for every new client; flags are no longer needed to turn it on. With O&W: Effective Dating - Company Structure, WorkflowsLicenseAndCertificationTasks. ⚠️ *WorkflowsOnboardingDashboard / NewHireDashboard (2028): not confirmed yet.*

**Leave off unless the client asks**
- **Effective Dating - Rates is now OFF by default** (problems for multi-rate and in-implementation clients). On only by request.
- By request: CA Break Waived (4029, CA employees), Arrears deductions, Shift Preview Notes, Position attributes.
- Never for clients: [DemoOnly] Conversation Dashboard, Employee list MUI grid. Fillable Forms Bulk Download is for existing NetGuide clients only (new builds use the O&W form builder).
""",
    "Netchex - Department Structure": """
**Div / Business Unit / Department** · Maintenance > Company Structure > Div/Bus/Dept > Add/Edit/Delete
- Expand All. Add Division > + Business Unit > + Department (codes have character limits).
- No structure given: 100 – [Company] / 100 – [Company] / 100 – Staff.
- Large structure: Import with the spec file (Add, div, BU, dept, cost center if asked); CSV without headers.

**Locations** · Maintenance > Company Structure > General > Locations
- Main Location 000 carries the main address. Add every other physical work location.

**Positions** · Maintenance > Human Resources > Positions
- Add or Import (position name + SOC code). Send the spec file to the client to fill in.
- No positions given: create one per department.
- *Why: positions drive Insights and salary benchmarking.*

**Payroll schedule** · Payroll > Payroll Settings > Payroll Schedules > Add Schedule
- Name exactly: Weekly Payroll / Bi-Weekly Payroll / Semi-Monthly Payroll / Monthly Payroll.
- Frequency, next period start and end, next check date, hours (40 / 80 / 86.6667 / 173.3334).
- Federal Holiday Banking Rule (usually previous business day) > select company > Save Schedule > Confirm and Save.
- One schedule per frequency. Send the dates to your lead to confirm with the client.

**Close the year** · Payroll > Utilities > Close the Year
- Click Close the Year > OK, once. The current year shows; that's expected.

**Payroll group codes** · Maintenance > Payroll > General > Payroll Group Codes
- Not bi-weekly? Delete the default bi-weekly groups and add the right ones (e.g. WklyHrly + salary group).
- Description spelled out, frequency, hours, last pay period start/end (current dates), GL account for wages = Regular.
- M3 group # only for M3 clients; leave the OT fields.
- Bi-weekly? Still update the defaults' last pay period dates.
""",
    "Netchex - Payroll Codes": """
**Deduction codes** · Maintenance > Payroll > Deductions > +Add
- Add only what the client needs beyond CIW's defaults. Template or Custom.
- Decide: pre/post-tax, Box 14, calculation, payroll cycle, reduction (partial/full), Netchex remit/payee.

**Earnings codes** · Maintenance > Payroll > Earnings > +Add
- Add what's missing; use Custom if a template label looks wrong.
- Set taxability, Box 14, calculation and options.
- ⚠️ *Not confirmed yet: whether the Earnings Code email template still goes to the client.*

**Unused codes** · Deductions / Earnings
- Delete codes the client won't use.

**GL base accounts** · Maintenance > Payroll > General Ledger > Edit Base GL Accounts
- Add or Import the client's account numbers (number, description, liability/expense/asset). Skip if the client gave none.

**GL mapping** · General Ledger > Manage Expense and Payable Accounts
- Map the Deductions, Earnings and Taxes tabs; use "only show unmapped". Finish with zero unmapped codes.
""",
    "Netchex - Maintenance": """
**Email/text events** · Maintenance > Human Resources > Email/Text Event Configuration
- Payroll Invoiced: client in To (Other E-Mail Address).
- Payroll Submitted, Wire Transfer Initiated, Wire Transfer Received: client + Project Lead.

**Shipping information** · Maintenance > Human Resources > Shipping Information > Edit
- Shipping Notes: PEPM (note a $0 special hierarchy). FedEx Standard unless local pickup.
- Shipping fee by FedEx zone (origin ZIP 70433); use the June 2026 zone table in the OneNote.
- Garnishments: blank if mailed to client, or Mail USPS to Agency + $1.25. Paperless/DD stubs/reports = No. Save.
- ⚠️ *Not confirmed yet: whether Special Notes stays blank.*

**Payroll, Shipping + W-2 contacts** · Maintenance > Human Resources > Contacts > Add
- Shipping contact must use a physical address; FedEx can't deliver to a PO Box.
- PO Box only? Loop in your manager and the sales rep.

**Check header** · Company Information > code + Enter > Paycheck Parameters icon
- Don't click the company name; click the last icon. Re-select the company code at the top or it won't save.
- Field 1 name · Field 2 address or DBA · Field 3 city, state, ZIP.
- Check Display Dept Info, Accrual Info, Saved Signature, Employer Deduction Amounts > Submit.
- Premium Pay: line 2 "C/O Netchex Payroll" and no signature upload.

**Upload signature** · Maintenance > Company Information > wand icon
- *Capture:* open the signature agreement attached to the case. Pick the clearest of the client's three signatures (single or double boxes) and screenshot it without the box lines.
- Double signature with a line through it? In the screenshot, use the pen tool in white to draw over the line and any other marks.
- Save As > Implementation > Scanned Signatures (so everyone has access), file type JPEG, named with the company code.
- *Upload:* go to the company in Netchex > Maintenance > Company Information, type the company code and click the wand icon on its row. The Company Image Upload Utility opens.
- Check the company drop-down is the right company (it defaults to the code you're in). Choose File > your JPEG > Upload.
- *Confirm:* Back Office > Reports > Check Signatures. Search by company name (not code) and generate; the signature shows next to the company's bank accounts.

**Report Runner** · Back Office > Report Runner > Create/Edit
- New companies now appear to be set by default. If yours is in the copy list, copy from CIW/OHE/parent (Edge; never edit the settings page).
- ⚠️ *Not confirmed yet: that this step is now automatic.*
""",
    "Netchex - Time & Attendance": """
**Web terminal** · Time & Attendance > Settings > Web Terminal Setup
- WEB01 now exists by default. Time clock (punch buttons) or time sheet; time zone offset from Central.
- Lunch buttons Yes; breaks, department or client/project transfer only if used.
- Terminal Enabled Yes (No if they don't punch in Netchex).
- Restrictions per client: IP address, geofence, mobile-only, desktop-only, or none.

**Payroll group rules** · Time & Attendance > Company Rules > Payroll Group Rules
- Every group: period start date (first Netchex pay period) and cycle (14 days bi-weekly, 7 for the others).
- Use Current Pay Rates During Pay Confirm: all groups. Auto Decrement: salary exempt only.
- Unlock timecards after processing: semi-monthly only. Rounding, auto lunch and alerts only if requested.
- ⚠️ *Not confirmed yet: the midnight rule (day of IN punch).*

**Overtime** · Time & Attendance > Company Rules > Shift Setup
- Hourly groups: OT after 40 (now automatic). Salary exempt: no OT.
- Client wants no OT for hourly? Create and close a Salesforce case documenting the request.
- ⚠️ *Not confirmed yet: that OT is removed from exempt automatically.*

**Terminal group** · Time & Attendance > Company Rules > Terminal Group Setup
- T01 now exists linked to WEB01. Employees are added after the employee load, not now.
""",
    "Salesforce - Cases": """
All cases: Salesforce > Contact > Cases > New

**Prenote case**
- Call Segment Banking · Call Topic Client Account Prenote · Origin Internal · Subject = company code.
- Takes up to 3 business days, so open it early.

**Tax case**
- Status New · Segment Tax · Topic New Client Set-up · Origin Internal · Subject "New Client Setup – [first check date]".
- Description: prior balances or none, state abbreviations, docs attached.
- Assign to your team's tax rep; attach the PIN, funding form and tax docs in Files & Notes.

**User Access case**
- Segment User Access · Topic New User or Edit User · Subject "User Access – [code]" · list each admin's name and email.
- Assign to the HR Onboarding Lead (uncheck Assign case to me); attach the User Access form and proposal.

**Check signature case**
- Assign to the signature team.
- ⚠️ *Not confirmed yet: fields and current assignees.*

**E-Verify case** (E-Verify only)
- Segment E-Verify · Topic Activate/Set-up · Subject "[code] – E-Verify" · "Form attached, thank you!"
- Assign to the HR Onboarding Lead; attach the signed E-Verify form.

**Pay card case**
- Only if the client wants pay cards.
- ⚠️ *Not confirmed yet: fields.*
""",
}


# Pushes the red "N comments" badge on a checklist section bar to the far right.
CHECKLIST_CSS = """<style>
[data-testid="stExpander"] summary:has(.stMarkdownBadge) > span { flex: 1 1 auto; width: 100%; }
[data-testid="stExpander"] summary:has(.stMarkdownBadge) > span > div { flex: 1 1 auto; min-width: 0; }
[data-testid="stExpander"] summary:has(.stMarkdownBadge) [data-testid="stMarkdownContainer"] { width: 100%; }
[data-testid="stExpander"] summary:has(.stMarkdownBadge) [data-testid="stMarkdownContainer"] p {
    display: flex; align-items: center; gap: 8px; width: 100%; }
[data-testid="stExpander"] summary .stMarkdownBadge { margin-left: auto; }
</style>"""


def _render_team_overview(builds, people, checks):
    """'All team members': one row per person. Click a row to open that person's page and their checklists."""
    status = builds.status.map(lambda v: _clean_text(v) or "Pending")
    qc_status = builds.qc_status.map(lambda v: _clean_text(v) or "Pending")
    uploaded = builds.id.astype(int).isin(_uploaded_ids(checks))
    rows = []
    for p in people:
        mine = builds.built_by.map(lambda v, p=p: p in _people(v))
        done = mine & (status == "Completed")
        qcd = mine & (qc_status == "Completed")
        rows.append({"Team member": p, "Builds": int(mine.sum()), "Completed": int(done.sum()),
                     "In progress": int((mine & (status == "In Progress")).sum()),
                     "Pending": int((mine & (status == "Pending")).sum()),
                     "QC'd": int(qcd.sum()),
                     "Awaiting QC": int((done & ~qcd & (qc_status != NOT_REQUIRED)).sum()),
                     "Checklist uploaded": int((mine & uploaded).sum())})
    table = pd.DataFrame(rows)

    def open_person():
        sel = st.session_state.get("chk_team_table")
        picked = (sel.selection.rows or [c[0] for c in sel.selection.cells]) if sel else []
        if picked:
            st.session_state["chk_person"] = table.at[picked[0], "Team member"]

    st.markdown("#### Team overview · click a team member to open their builds and checklists")
    st.dataframe(table, key="chk_team_table", hide_index=True, width="stretch", on_select=open_person,
                 selection_mode=["single-row", "single-cell"], height=35 * len(table) + 40,
                 column_config={
                     "QC'd": st.column_config.NumberColumn(help="Their builds with QC status Completed"),
                     "Awaiting QC": st.column_config.NumberColumn(help="Completed builds whose QC isn't completed yet"),
                     "Checklist uploaded": st.column_config.NumberColumn(
                         help="Their builds whose checklist is ticked as uploaded")})
    st.caption("Counts are by Built By. You can also pick a team member from the dropdown above.")


def _render_build_checklist(builds, checks):
    if builds.empty:
        st.info("Add a build in the 📅 Monthly Builds section first.")
        return

    def txt(v):
        return _clean_text(v) or ""

    def due(v):
        d = _to_date(v)
        return f" · due {d:%m/%d/%Y}" if d else ""

    # ── Pick a team member to open their page, then a build to open its checklist ──
    # The whole CoE team is listed (even with no builds yet), plus past builders such as Jyoti and Supriya.
    past = sorted({p for v in builds.built_by for p in _people(v)} - set(TEAM_MEMBERS))
    ALL = "All team members"
    who = st.selectbox("Team member", [ALL, *TEAM_MEMBERS, *past], key="chk_person")
    if who == ALL:
        _render_team_overview(builds, [*TEAM_MEMBERS, *past], checks)
        return
    order = builds.sort_values("id")     # in the order they were added
    if who != ALL:
        built_mask = order.built_by.map(lambda v: who in _people(v))
        # Only the builds they built (Built By). Builds they QC are opened from the builder's page.
        order = order[built_mask]
        st.markdown(f"#### {who}")
        qc_complete = order.reindex(built_mask.index).qc_status == "Completed"
        c = st.columns(3)
        c[0].metric("Builds", int(built_mask.sum()), help="Builds where they're the Built By")
        c[1].metric("Builds completed", int((built_mask & (order.reindex(built_mask.index).status == "Completed")).sum()))
        c[2].metric("QC'd", int((built_mask & qc_complete).sum()),
                    help="Their builds where the QC Status is Completed, whoever did the QC")
        if order.empty:
            st.info(f"{who} has no builds yet.")
            return
    labels = {int(r.id): txt(r.company_name) + (f" ({txt(r.company_code)})" if txt(r.company_code) else "")
              + due(r.due_date) for r in order.itertuples()}

    # Their builds as a table: click a row to open that build's checklist below.
    # Streamlit tables can't be clickable and editable at once, so editing is behind a switch.
    pick_key = f"chk_pick_{who}"
    # Read-only here: builds are edited in 📅 Monthly Builds.
    own = order
    uploaded_ids = _uploaded_ids(checks)
    all_qc = load_qc_reports() if QC_REPORTS["enabled"] else None

    def show_table(subset, key):
        """A read-only builds table; clicking a row (or any cell in it) opens that build's checklist below."""
        table = _builds_table(subset).reset_index(drop=True)

        def on_pick():
            sel = st.session_state[key].selection
            picked = sel.rows or [c[0] for c in sel.cells]
            if picked:
                st.session_state[pick_key] = int(table.at[picked[0], "id"])

        shown = table.copy()
        for col in ("Status", "QC Status", "UA Status", "Balance Status"):
            shown[col] = shown[col].map(lambda s: STATUS_LABEL.get(s, s) if s else "")
        shown.insert(shown.columns.get_loc("QC Status") + 1, "Checklist Uploaded",
                     shown["id"].map(lambda i: int(i) in uploaded_ids))
        if all_qc is not None:
            after = "Checklist Uploaded"
            for kind in ("sf", "netchex"):
                latest_k = all_qc[all_qc.report_type == kind].drop_duplicates("build_id", keep="first")
                pct = {int(r.build_id): float(r.pct_pass) / 100 for r in latest_k.itertuples() if pd.notna(r.build_id)}
                col = f"{QC_KINDS[kind]['label']} % Pass"
                shown.insert(shown.columns.get_loc(after) + 1, col, shown["id"].map(lambda i: pct.get(int(i))))
                after = col
        st.dataframe(
            shown, key=key, hide_index=True, width="stretch", on_select=on_pick,
            selection_mode=["single-row", "single-cell"], height=min(35 * len(shown) + 40, 400),
            column_config={"id": None,
                           "Company Name": st.column_config.TextColumn("Company Name", pinned=True),
                           "Checklist Uploaded": st.column_config.CheckboxColumn(
                               "Checklist Uploaded", help="Ticked on the build's checklist below once it's uploaded"),
                           "SF QC % Pass": st.column_config.ProgressColumn(
                               "SF QC % Pass", format="percent", min_value=0, max_value=1,
                               help="From the latest SF QC report uploaded on the build's page below"),
                           "Netchex QC % Pass": st.column_config.ProgressColumn(
                               "Netchex QC % Pass", format="percent", min_value=0, max_value=1,
                               help="From the latest Netchex QC report uploaded on the build's page below "
                                    "(N/A rows left out)"),
                           "Build Completed": st.column_config.DateColumn("Build Completed", format="MM/DD/YYYY"),
                           "QC Completed": st.column_config.DateColumn("QC Completed", format="MM/DD/YYYY"),
                           "Due Date": st.column_config.DateColumn("Due Date", format="MM/DD/YYYY")})

    st.markdown(f"**{who}'s builds ({len(own)})** · click a row to open its checklist")
    show_table(own, f"chk_table_{who}")

    build_id = st.session_state.get(pick_key)
    if build_id not in labels:
        build_id = list(labels)[-1]       # nothing clicked yet: their most recently added build
    b = builds[builds.id == build_id].iloc[0]
    saved = {r.item_id: r for r in checks[checks.build_id == build_id].itertuples()}
    st.markdown(f"#### ✅ Checklist · {labels[build_id]}")
    code = txt(b.company_code) or str(build_id)
    safe = re.sub(r"[^A-Za-z0-9]+", "_", txt(b.company_name)).strip("_")[:40] or "Build"
    download_link = _excel_link(export_checklist_excel(b, saved), f"QC_Checklist_{code}_{safe}.xlsx",
                                "Download checklist",
                                "This build's checklist as an Excel file: details, every item with its Built / QC / "
                                "N/A ticks, errors and notes. Ready to print or save to the client folder.", align="left", big=True)
    status, qc_status = txt(b.status) or "Pending", txt(b.qc_status) or "Pending"

    st.markdown(f"**Status:** {STATUS_LABEL.get(status, status)} &nbsp;·&nbsp; "
                f"**QC status:** {STATUS_LABEL.get(qc_status, qc_status)}")
    st.caption("Ticks, N/A, errors, notes and the names below save automatically. A section's counts "
               "update when you close it.")
    # The name/status boxes get fresh keys after each save so they show the saved values.
    ver = st.session_state.get(f"chk_ver_{build_id}", 0)

    st.markdown(CHECKLIST_CSS, unsafe_allow_html=True)

    def _refresh_section(si):
        """Reloads one section from the database (fresh ticks and counts) with a full page rerun."""
        st.session_state[f"chk_ver_{build_id}_{si}"] = st.session_state.get(f"chk_ver_{build_id}_{si}", 0) + 1
        st.rerun()

    @st.fragment
    def _section(si, section, note, rows):
        """One checklist section. A tick, N/A or comment reruns only this section, which saves it in the
        background without reloading the rest of the page, so you can keep working straight away."""
        # Each section has its own reload counter, so refreshing one never disturbs another.
        sv = st.session_state.get(f"chk_ver_{build_id}_{si}", 0)
        # While a section is open its table and heading stay exactly as they were when it was opened, so
        # quiet auto-saves never reload it under your cursor. Closing the section refreshes it.
        snap_key = f"chk_snap_{build_id}_{si}_{sv}"
        if snap_key not in st.session_state:
            st.session_state[snap_key] = rows
        df = pd.DataFrame(st.session_state[snap_key])
        df["Note"] = (df.Comments != "").map(lambda c: "🔴" if c else "")    # red dot only on rows with an error
        n_notes = int((df.Comments != "").sum())                    # badge counts items with an error logged
        sec_app = int((~df["N/A"]).sum())
        sec_built = int((df.Built & ~df["N/A"]).sum())
        sec_qc = int((df.QC & ~df["N/A"]).sum())
        done = "✅ " if sec_app and sec_built == sec_app and sec_qc == sec_app else ""
        # Shown as a red badge, pushed to the far right of the section bar by CHECKLIST_CSS.
        notes = f" :red-badge[🔴 {n_notes} error{'s' if n_notes != 1 else ''}]" if n_notes else ""
        # Remember which sections are open, so a section stays open when its counts refresh.
        open_key = f"chk_open_{build_id}_{si}"
        was_open = st.session_state.get(open_key, False)
        # A fixed key keeps it the same section bar when its counts or badge change, so one click opens or
        # closes it (without a key, a changed label made it a new bar and the first click was lost).
        exp = st.expander(f"{done}{section}  —  Built {sec_built}/{sec_app} · QC {sec_qc}/{sec_app}{notes}",
                          key=f"chk_exp_{build_id}_{si}", on_change="rerun")
        st.session_state[open_key] = bool(exp.open)
        if was_open and not exp.open:
            _refresh_section(si)             # just closed: reload it with the saved ticks and counts
        with exp:
            if note:
                st.caption(note)
            links = [l for start, ls in SECTION_LINKS.items() if section.startswith(start) for l in ls]
            howto = next((h for start, h in SECTION_HOWTO.items() if section.startswith(start)), None)
            if links or howto:
                with st.container(horizontal=True, horizontal_alignment="right"):
                    if howto:
                        with st.popover("How to", icon=":material/menu_book:", type="tertiary",
                                        help="Step-by-step notes for this section"):
                            st.markdown(howto)
                    for link_label, url, tip in links:
                        st.link_button(link_label, url, type="tertiary", icon=":material/open_in_new:", help=tip)
            view = df.reset_index(drop=True)
            t1, t2, _ = st.columns([1, 1, 4])
            want_all = {}
            for col, field, label in ((t1, "Built", "Tick all Built"), (t2, "QC", "Tick all QC")):
                applicable = ~view["N/A"]
                all_now = bool(applicable.any() and view.loc[applicable, field].all())
                want = col.checkbox(label, value=all_now, key=f"chk_all{field}_{build_id}_{si}_{sv}",
                                    help=f"Ticks {field} for every item in this section (not N/A ones). "
                                         "Untick to clear them all.")
                if want != all_now:
                    want_all[field] = want
            ed = st.data_editor(
                view, key=f"chk_{build_id}_{si}_{sv}", hide_index=True, width="stretch",
                num_rows="fixed", disabled=["Area", "Item", "Note"],
                column_order=["Area", "Item", "N/A", "Built", "Notes", "QC", "Comments", "Note"],
                column_config={"item_id": None,
                               "Area": st.column_config.TextColumn("Area", width=200),
                               "Item": st.column_config.TextColumn("Item", width=520),
                               "Built": st.column_config.CheckboxColumn("Built", width=60),
                               "QC": st.column_config.CheckboxColumn("QC", width=60),
                               "Comments": st.column_config.TextColumn("Errors ✏️", width=260,
                                                                       help="Click to log an error found on this item"),
                               "Notes": st.column_config.TextColumn("Notes ✏️", width=260,
                                                                    help="Click to add a note for this item",
                                                                    disabled=not CHECK_NOTES["enabled"]),
                               "N/A": st.column_config.CheckboxColumn("N/A", width=55),
                               "Note": st.column_config.TextColumn("💬", width=45,
                                                                   help="🔴 = this item has an error")})
            ed = ed.copy()
            refresh_after = False
            for field, want in want_all.items():
                ed.loc[~ed["N/A"].astype(bool), field] = want
            if want_all:
                refresh_after = True         # Tick all: reload the section so the table shows the new ticks
            # Compare with what this section last saved, and save only what's new.
            last_key = f"chk_last_{build_id}_{si}_{sv}"
            last = st.session_state.setdefault(last_key, {
                r["item_id"]: (r["Built"], r["QC"], r["N/A"], r["Comments"], r["Notes"])
                for r in st.session_state[snap_key]})
            sec_changes = {}
            for i in range(len(ed)):
                new = (bool(ed.at[i, "Built"]), bool(ed.at[i, "QC"]), bool(ed.at[i, "N/A"]),
                       txt(ed.at[i, "Comments"]), txt(ed.at[i, "Notes"]))
                if new != last[ed.at[i, "item_id"]]:
                    sec_changes[ed.at[i, "item_id"]] = new
            if sec_changes:
                try:
                    save_build_checks([{"build_id": int(build_id), "item_id": k, "built": v[0], "qc": v[1],
                                        "na": v[2], "comments": v[3] or None, "notes": v[4] or None}
                                       for k, v in sec_changes.items()])
                    last.update(sec_changes)
                    st.session_state[f"chk_saved_at_{build_id}_{si}"] = now_local()
                except Exception as e:
                    st.error(f"Could not save your last change, it will retry on your next click. ({e})")
            if refresh_after:
                # Tick all: reload just this section (it stays open) so the table shows the new ticks
                st.session_state[f"chk_ver_{build_id}_{si}"] = sv + 1
                st.rerun(scope="fragment")
            saved_at = st.session_state.get(f"chk_saved_at_{build_id}_{si}")
            st.caption(f"✓ Saved automatically at {saved_at:%I:%M:%S %p}" if saved_at
                       else "Changes save automatically.")

    for si, (section, note, _groups) in enumerate(CHECKLIST, 1):
        items = [it for it in CHECK_ITEMS if it["section"] == section]
        rows = []
        for it in items:
            s = saved.get(it["id"])
            rows.append({"item_id": it["id"], "Area": it["sub"], "Item": ("★ " if it["key"] else "") + it["item"],
                         "Built": bool(s.built) if s else False, "QC": bool(s.qc) if s else False,
                         "N/A": bool(s.na) if s else False, "Comments": (txt(s.comments) if s else ""),
                         "Notes": (txt(getattr(s, "notes", None)) if s else "")})
        _section(si, section, note, rows)

    # ── Who built it and who QC'd it, editable here ──
    team = [*TEAM_MEMBERS, *sorted({v for c in ("built_by", "qc_by") for v in builds[c]
                                    if isinstance(v, str) and v not in TEAM_MEMBERS})]
    NONE = "—"

    def pick(col, label, options, current, k):
        opts = [NONE, *options] + ([current] if current and current not in options else [])
        return col.selectbox(label, opts, index=opts.index(current) if current else 0, key=f"{k}_{build_id}_{ver}")

    with st.container(border=True):
        n1, n2, n3, n4 = st.columns(4)
        new_fields = {
            "built_by": pick(n1, "Built by", team, txt(b.built_by), "chk_bb"),
            "status": pick(n2, "Status", BUILD_STATUSES, status, "chk_st"),
            "qc_by": pick(n3, "QC'd by", team, txt(b.qc_by), "chk_qb"),
            "qc_status": pick(n4, "QC status", BUILD_STATUSES, qc_status, "chk_qs"),
        }
    old_fields = {"built_by": txt(b.built_by), "status": status, "qc_by": txt(b.qc_by), "qc_status": qc_status}
    field_changes = {k: (None if v == NONE else v) for k, v in new_fields.items()
                     if (v if v != NONE else "") != old_fields[k]}

    # Names and statuses save as soon as they change.
    if field_changes:
        try:
            if field_changes.get("status") == "Completed" and not _to_date(b.completion_date):
                field_changes["completion_date"] = today_local()
            if field_changes.get("qc_status") == "Completed" and not _to_date(getattr(b, QC_DATE, None)):
                field_changes[QC_DATE] = today_local()
            update_build(build_id, field_changes)
            st.session_state["_toast"] = f"✅ Saved for {txt(b.company_name)}"
        except Exception as e:
            st.session_state["_flash_error"] = f"Could not save your change, please try again. ({e})"
        st.session_state[f"chk_ver_{build_id}"] = ver + 1
        st.rerun()

    built, qc, applicable = _progress(checks, build_id)
    if applicable and built == applicable and status != "Completed":
        st.info("Every checklist item is built. Set **Status** to Completed above.")
    if applicable and qc == applicable and qc_status != "Completed":
        st.info("Every checklist item is QC'd. Set **QC status** to Completed above.")

    # ── Last check: the finished checklist is uploaded. Saves straight away. ──
    up = saved.get(UPLOAD_ITEM)
    is_up = bool(up is not None and up.built)
    up_key = f"chk_uploaded_{build_id}"

    def _save_uploaded():
        ticked = bool(st.session_state.get(up_key))
        try:
            save_build_checks([{"build_id": int(build_id), "item_id": UPLOAD_ITEM, "built": ticked, "qc": False,
                                "na": False, "comments": f"{today_local():%m/%d/%Y}" if ticked else None}])
            st.session_state["_toast"] = ("Checklist marked as uploaded" if ticked
                                          else "Checklist marked as not uploaded")
        except Exception as e:
            st.session_state["_flash_error"] = f"The upload status could not be saved. Please try again. ({e})"

    st.session_state[up_key] = is_up       # always show the saved value
    # Download the checklist, then tick that it's uploaded.
    st.space("small")
    st.markdown(download_link, unsafe_allow_html=True)
    with st.container(border=True):
        st.checkbox("Checklist uploaded", key=up_key, on_change=_save_uploaded,
                    help="Shown in the Checklist Uploaded column of the team member's builds table.")
        if is_up and _clean_text(up.comments):
            st.caption(f"Uploaded on {_clean_text(up.comments)}")
        else:
            st.caption("Confirm once the completed checklist has been downloaded and uploaded to the S drive.")

    # ── SF QC report (Salesforce QC workbook) and Netchex QC report for this build ──
    _render_qc_for_build(builds, b, "sf")
    _render_qc_for_build(builds, b, "netchex")




def _month_options(builds):
    """Months to choose from, newest first: from the earliest due date up to next month (or a later due date)."""
    nxt = add_months(month_start(today_local()), 1)
    due_months = list(_build_months(builds)) if not builds.empty else []
    first = min([month_start(today_local()), *due_months])
    last = max([nxt, *due_months])
    opts, mo = [], first
    while mo <= last:
        opts.append(mo)
        mo = add_months(mo, 1)
    return opts[::-1]


def _render_month(builds, pick_month):
    """The month picked in the sidebar: the numbers, progress, and every build in that month."""
    today = today_local()
    current = month_start(today)

    df = builds.copy()
    df["status"] = df.status.map(lambda v: _clean_text(v) or "Pending")
    df["qc_status"] = df.qc_status.map(lambda v: _clean_text(v) or "")
    df["due"] = df.due_date.map(_to_date)
    # Same builds as the Builds table (due in this month, or added under it without a due date).
    g = df[_build_months(df) == pick_month]

    st.subheader(f"{pick_month:%B %Y} Company Builds")
    if g.empty:
        st.info(f"No builds for {pick_month:%B %Y} yet. Add them in the table below.")
    else:
        not_req = g.status == NOT_REQUIRED
        total = len(g)
        needed = total - int(not_req.sum())     # Not Required builds don't count towards progress
        done = g.status == "Completed"
        qc_done = done & g.qc_status.isin(["Completed", NOT_REQUIRED])
        overdue = (~done) & (~not_req) & g.due.map(lambda d: d is not None and d < today)

        m = st.columns(7)
        m[0].metric("Builds", needed, help="Builds due this month, plus any without a due date. "
                    "N/A (Not Required) builds are not included")
        m[1].metric("Completed", int(done.sum()))
        m[2].metric("In progress", int((g.status == "In Progress").sum()))
        m[3].metric("Pending", int((g.status == "Pending").sum()))
        m[4].metric("N/A", int(not_req.sum()), help="Builds with Status Not Required")
        m[5].metric("Awaiting QC", int((done & ~qc_done).sum()), help="Built, but QC not completed yet")
        m[6].metric("Overdue", int(overdue.sum()), help="Past the due date and not completed")

        p1, p2 = st.columns(2)
        p1.progress(done.sum() / needed if needed else 1.0,
                    text=f"Builds completed: {int(done.sum())} of {needed} ({done.sum() / needed if needed else 1:.0%})")
        qc_base = int(done.sum())
        p2.progress(qc_done.sum() / qc_base if qc_base else 0.0,
                    text=f"QC completed: {int(qc_done.sum())} of {qc_base} completed builds")

    # ── Every build in the month: edit, add or delete right here ──
    _month_builds_table(pick_month)


@st.fragment
def _month_builds_table(pick_month):
    """The Monthly Builds table, like a checklist section: an edit reruns only this table, which saves
    it in the background without reloading (or fading) the rest of the page."""
    builds = load_builds()
    _builds_editor(builds, _month_frame(builds, pick_month), key=f"month_{pick_month}", label=f"{pick_month:%B %Y}",
                   month=pick_month)


# Raines builds: a separate project of 52 companies, all completed 14-18 May 2026.
# Shown on the Dashboard only (not in Monthly Builds or QC). (company code, company name, built by, completed on)
RAINES_BUILDS = [
    ('0YW', 'Atlantic Colony Venture I Llc. - Hampton Inn & Suites Colony', 'Supriya', "2026-05-15"),
    ('0YX', 'Atlantic Colony Venture II Llc. - Homewood Suites Colony', 'Supriya', "2026-05-16"),
    ('0ZJ', 'Atlantic Downtown Dallas Venture - Element Gaston', 'Supriya', "2026-05-16"),
    ('0Z4', 'Atlantic Grapevine Inn', 'Supriya', "2026-05-16"),
    ('0ZL', 'Atlantic McKinney Hotel Ventures Llc. - Home2 McKinney', 'Harshith', "2026-05-17"),
    ('0ZK', 'Atlantic McKinney Hotel Ventures Llc. - Springhill Suites McKinney', 'Supriya', "2026-05-17"),
    ('0YM', 'CAI Atlantic Waco Express LLC - Holiday Inn Express Waco', 'Harshith', "2026-05-17"),
    ('0YS', 'CAI Atlantic Waco Suites Llc. - Towneplace Waco', 'Harshith', "2026-05-18"),
    ('0YY', 'Cartersville Paradise Group - Hilton Garden Inn Cartersville', 'Supriya', "2026-05-18"),
    ('0ZM', 'M&M Hotel Venture Llc. - Aloft Element Love Field', 'Supriya', "2026-05-18"),
    ('0Z7', 'Mohit Developers Llc. - Towneplace Suites DeSoto', 'Salman', "2026-05-18"),
    ('0YZ', 'Naresh and Asha - Hilton Garden inn Gainesville', 'Akash', "2026-05-18"),
    ('0ZE', 'Raines HP Operations, Llc.', 'Harshith', "2026-05-18"),
    ('0Z8', 'Raines HP Workforce, Llc.', 'Harshith', "2026-05-18"),
    ('0Z1', 'Raines HP Workforce, Llc. Courtyard Banner Elk', 'Salman', "2026-05-16"),
    ('0Z2', 'Raines HP Workforce, Llc. Courtyard Columbus', 'Harshith', "2026-05-16"),
    ('0Z3', 'Raines HP Workforce, Llc. Fairfield Whitsett', 'Fiza', "2026-05-16"),
    ('0Z5', 'Raines HP Workforce, Llc. Hilton Garden Inn Conway', 'Harshith', "2026-05-16"),
    ('0Z6', 'Raines HP Workforce, Llc. Hilton Garden Inn Greenville', 'Harshith', "2026-05-16"),
    ('0Z9', 'Raines HP Workforce, Llc. Hilton Garden Inn West Little Rock', 'Fiza', "2026-05-15"),
    ('0ZA', 'Raines HP Workforce, Llc. Holiday Inn Roswell', 'Akash', "2026-05-18"),
    ('0ZB', 'Raines HP Workforce, Llc. Home 2 Conway', 'Fiza', "2026-05-15"),
    ('0ZC', 'Raines HP Workforce, Llc. Home2 Charlotte University', 'Salman', "2026-05-18"),
    ('0ZD', 'Raines HP Workforce, Llc. Hotel Theo', 'Harshith', "2026-05-16"),
    ('0ZF', 'Raines HP Workforce, Llc. Sea Palms', 'Harshith', "2026-05-16"),
    ('0ZG', 'Raines HP Workforce, Llc. Springhill Suites Cheraw', 'Salman', "2026-05-15"),
    ('0ZH', 'Raines HP Workforce, Llc. The Lantern Columbia', 'Salman', "2026-05-16"),
    ('0ZN', 'Raines Staffing Llc. Cambria Hotel Mount Pleasant-Charleston', 'Supriya', "2026-05-15"),
    ('0ZO', 'Raines Staffing Llc. Comfort Florence', 'Salman', "2026-05-14"),
    ('0ZP', 'Raines Staffing Llc. DoubleTree Greenville', 'Fiza', "2026-05-15"),
    ('0ZQ', 'Raines Staffing Llc. DoubleTree Riverview', 'Fiza', "2026-05-15"),
    ('0ZR', 'Raines Staffing Llc. Doubletree Summerville', 'Supriya', "2026-05-15"),
    ('0ZS', 'Raines Staffing Llc. Fairfield Inn & Suites Florence', 'Akash', "2026-05-14"),
    ('0ZT', 'Raines Staffing Llc. Hampton Inn & Suites Columbia', 'Akash', "2026-05-15"),
    ('0ZU', 'Raines Staffing Llc. Hilton Garden Inn Mt Pleasant', 'Fiza', "2026-05-15"),
    ('0ZV', 'Raines Staffing Llc. Hotel Florence', 'Fiza', "2026-05-14"),
    ('0ZW', 'Raines Staffing Llc. Hyatt Place Augusta', 'Harshith', "2026-05-16"),
    ('0ZX', 'Raines Staffing Llc. Hyatt Place Florence', 'Akash', "2026-05-15"),
    ('0ZY', 'Raines Staffing Llc. Hyatt Place Mt Pleasant', 'Harshith', "2026-05-15"),
    ('0ZZ', 'Raines Staffing Llc. Residence Inn Florence', 'Salman', "2026-05-15"),
    ('0YO', 'Raines Staffing Llc. Springhill Suites Charleston Mount Pleasant', 'Akash', "2026-05-15"),
    ('0YP', 'Raines Staffing Llc. The Foundry Hotel Asheville, Curio Collection', 'Supriya', "2026-05-14"),
    ('0YQ', 'Raines Staffing Llc. The Spa at Hotel Florence', 'Salman', "2026-05-15"),
    ('0YR', 'Raines Staffing Llc. Towneplace Florence', 'Harshith', "2026-05-15"),
    ('0YT', 'Raines Staffing Llc. Tru North Augusta', 'Supriya', "2026-05-15"),
    ('0YV', 'Raines Staffing Llc. Waynesville Inn & Golf', 'Salman', "2026-05-15"),
    ('0YN', 'Renaissance Group Cartersville Courtyard Cartersville', 'Harshith', "2026-05-18"),
    ('10A', 'Renaissance Group Shrinathji Llc. Fairfield Inn & Suites Woodstock', 'Harshith', "2026-05-18"),
    ('0Z0', 'Renaissance Group Woodstock Inc. Hampton Inn Woodstock', 'Salman', "2026-05-18"),
    ('10B', 'Smithcurry Hotel GP Manchester - Panos Smith Hotel Group Manchester Villa - Hilton Garden Inn Rock Hill', 'Harshith', "2026-05-16"),
    ('10C', 'Smithcurry Hotel Group Kings Grant - Panos Smith Hotel Group Kings Grant - Hampton Inn & Suites Speedway', 'Danielle', "2026-05-14"),
    ('0YU', 'TT&C Hotel Group Llc. - DoubleTree Columbia', 'Supriya', "2026-05-15"),
]

STATUS_COLORS = ["#A0A7B4", "#E8B33A", "#2E9E5B", "#4A5568"]   # Pending, In Progress, Completed, Not Required


def _people(value):
    """'Fiza (SF) / Harshith (NX)' -> ['Fiza', 'Harshith']"""
    if not isinstance(value, str):
        return []
    return [re.sub(r"\s*\(.*?\)", "", p).strip() for p in value.split("/") if p.strip()]


def _render_overview(builds):
    """The main dashboard: totals for every month and the charts."""
    import altair as alt

    if builds.empty:
        st.info("No builds yet. Add builds in the 📅 Monthly Builds section and the totals will appear here.")
        return
    df = builds.copy()
    df["status"] = df.status.map(lambda v: _clean_text(v) or "Pending")
    df["qc_status"] = df.qc_status.map(lambda v: _clean_text(v) or "")
    # Month for the chart: the due date's month, else the month the build was added under.
    df["month"] = [month_start(_to_date(d)) if _to_date(d) else (month_start(_to_date(m)) if _to_date(m) else None)
                   for d, m in zip(df.due_date, df[BUILD_MONTH])]

    raines = _raines_frame()
    n_raines = len(raines)
    n_not_req = int((df.status == NOT_REQUIRED).sum())
    total = len(df) - n_not_req + n_raines      # Not Required (N/A) builds aren't counted in the total
    done = df.status == "Completed"
    n_done = int(done.sum()) + n_raines          # every Raines build is completed
    qc_done = done & (df.qc_status == "Completed")
    dated = df[df.month.map(lambda v: v is not None)]
    months = sorted(set(dated.month.unique()) | set(raines.month))
    first = f"{months[0]:%B %Y}" if months else ""
    this_month = int((dated.month == month_start(today_local())).sum())

    h1, h2 = st.columns([3, 1], vertical_alignment="center")
    h1.subheader(f"All builds · {first} to date" if first else "All builds")
    h2.markdown(_excel_link(export_all_builds_excel(df, raines),
                            f"Company_Builds_All_to_{today_local():%Y-%m-%d}.xlsx",
                            "Download Excel",
                            f"All {total:,} builds to date: Summary, every company build (all months) "
                            "and the Raines builds, one sheet each"),
                unsafe_allow_html=True)
    m = st.columns(7)
    m[0].metric("Total builds", f"{total:,}", help=f"Includes the {n_raines} Raines builds. "
                f"N/A (Not Required) builds are not included" + (f" ({n_not_req})" if n_not_req else ""))
    m[1].metric("Completed", f"{n_done:,}", help=f"Includes the {n_raines} Raines builds")
    m[2].metric("In progress", int((df.status == "In Progress").sum()))
    m[3].metric("Pending", int((df.status == "Pending").sum()))
    m[4].metric("N/A", n_not_req, help="Builds with Status Not Required")
    m[5].metric("QC completed", f"{int(qc_done.sum()):,}", help="Raines builds aren't QC'd in this app")
    m[6].metric("This month", this_month, help=f"Builds due in {today_local():%B %Y}")
    p1, p2 = st.columns(2)
    p1.progress(n_done / total if total else 1.0, text=f"Builds completed: {n_done / total if total else 1:.0%}")
    p2.progress(qc_done.sum() / done.sum() if done.sum() else 0.0,
                text=f"QC completed: {qc_done.sum() / done.sum():.0%} of completed builds (not counting Raines)"
                if done.sum() else "QC")

    if not months:
        st.info("Add due dates to builds to see the monthly chart.")
        return
    order = [f"{mo:%b %Y}" for mo in months]

    # ── Builds per month, split by status (Raines builds as their own colour) ──
    st.markdown("#### Builds per month")
    counts = dated.groupby(["month", "status"]).size().reset_index(name="Builds")
    raines_counts = raines.groupby("month").size().reset_index(name="Builds").assign(status=RAINES_LABEL)
    counts = pd.concat([counts, raines_counts], ignore_index=True)
    counts["Month"] = counts.month.map(lambda mo: f"{mo:%b %Y}")
    status_domain = [*BUILD_STATUSES, RAINES_LABEL]
    counts["rank"] = counts.status.map({s: i for i, s in enumerate(status_domain)})
    bars = alt.Chart(counts).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X("Month:N", sort=order, title=None, axis=alt.Axis(labelAngle=0)),
        y=alt.Y("sum(Builds):Q", title="Builds", axis=alt.Axis(tickMinStep=1, format="d")),
        color=alt.Color("status:N", title=None,
                        scale=alt.Scale(domain=status_domain, range=[*STATUS_COLORS, RAINES_COLOR]),
                        legend=alt.Legend(orient="top")),
        order=alt.Order("rank:Q"),
        tooltip=["Month", alt.Tooltip("status:N", title="Status"), "Builds"])
    totals = counts.groupby("Month", as_index=False)["Builds"].sum()
    labels = alt.Chart(totals).mark_text(dy=-9, fontWeight="bold", fontSize=13).encode(
        x=alt.X("Month:N", sort=order), y="Builds:Q", text="Builds:Q")
    st.altair_chart((bars + labels).properties(height=340), width="stretch")

    # ── Who built and who QC'd (a shared build counts for each person) ──
    def per_person(col, mask):
        names = [p for v in df.loc[mask, col] for p in _people(v)]
        return pd.Series(names, dtype=object).value_counts().rename_axis("Person").reset_index(name="Builds")

    def person_chart(data):
        """Completed builds per person, with their Raines builds stacked on the end of the bar."""
        totals = data.groupby("Person", as_index=False)["Builds"].sum()
        order = totals.sort_values(["Builds", "Person"], ascending=[False, True]).Person.tolist()
        data = data.assign(rank=data.Type.map({"Company builds": 0, RAINES_LABEL: 1}))
        bars = alt.Chart(data).mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3).encode(
            y=alt.Y("Person:N", sort=order, title=None),
            x=alt.X("sum(Builds):Q", title=None, axis=alt.Axis(tickMinStep=1, format="d"),
                    scale=alt.Scale(domain=[0, int(totals.Builds.max() * 1.15) + 1])),   # room for the labels
            color=alt.Color("Type:N", title=None, legend=alt.Legend(orient="top"),
                            scale=alt.Scale(domain=["Company builds", RAINES_LABEL], range=["#2F5D9A", RAINES_COLOR])),
            order=alt.Order("rank:Q"),
            tooltip=["Person", "Type", "Builds"])
        labels = alt.Chart(totals).mark_text(align="left", dx=4, fontWeight="bold").encode(
            y=alt.Y("Person:N", sort=order), x="Builds:Q", text="Builds:Q")
        return (bars + labels).properties(height=30 * max(len(totals), 1) + 80)

    st.markdown("#### Builds by team member")
    data = pd.concat([
        per_person("built_by", done).assign(Type="Company builds"),
        raines.built_by.value_counts().rename_axis("Person").reset_index(name="Builds").assign(Type=RAINES_LABEL),
    ], ignore_index=True)
    if data.empty:
        st.caption("No completed builds with a Built By name yet.")
    else:
        st.altair_chart(person_chart(data), width="stretch")
    undated = df[df.month.map(lambda v: v is None)]
    if len(undated):
        names = ", ".join(_clean_text(n) or "(no name)" for n in undated.company_name)
        st.caption(f"No due date yet: **{names}**. {'It is' if len(undated) == 1 else 'They are'} in the totals "
                   "but not the monthly chart. Add the due date in 📅 Monthly Builds (current month).")

    _render_raines(raines)


RAINES_LABEL = "Raines"
RAINES_COLOR = "#8E6CC9"


def _raines_frame():
    df = pd.DataFrame(RAINES_BUILDS, columns=["company_code", "company_name", "built_by", "completion_date"])
    df["completion_date"] = df.completion_date.map(date.fromisoformat)
    df["month"] = df.completion_date.map(month_start)
    return df


def _render_raines(raines):
    """Raines builds: done as a separate project, so they get their own section."""
    st.divider()
    st.subheader("🏨 Raines Builds")
    if raines.empty:
        st.caption("No Raines builds.")
        return
    start, end = raines.completion_date.min(), raines.completion_date.max()
    when = f"{start:%B} {start.day}–{end.day}, {end:%Y}" if (start.year, start.month) == (end.year, end.month) \
        else f"{start:%b %d, %Y} – {end:%b %d, %Y}"
    m = st.columns(3)
    m[0].metric("Raines builds", len(raines))
    m[1].metric("Completed", len(raines))
    m[2].metric("Completed on", when)

    left, right = st.columns([1, 2])
    by_person = raines.built_by.value_counts().rename_axis("Built By").reset_index(name="Company Codes")
    by_person = pd.concat([by_person, pd.DataFrame([{"Built By": "Grand Total", "Company Codes": len(raines)}])],
                          ignore_index=True)
    left.dataframe(by_person, hide_index=True, width="stretch")
    with right.expander(f"All {len(raines)} Raines builds"):
        table = raines.sort_values(["completion_date", "company_name"]).rename(columns={
            "company_code": "Company Code", "company_name": "Company Name", "built_by": "Built By",
            "completion_date": "Build Completed"})[["Company Code", "Company Name", "Built By", "Build Completed"]]
        st.dataframe(table, hide_index=True, width="stretch",
                     column_config={"Build Completed": st.column_config.DateColumn(format="MM/DD/YYYY")})


@st.fragment(run_every="30s")
def _live(render, *args):
    """Re-reads the database and redraws the page every 30 seconds."""
    try:
        builds = load_builds()
    except Exception as e:
        st.error("Could not refresh the data. It will try again in 30 seconds.")
        with st.expander("Error details"):
            st.code(str(e))
        return
    render(builds, *args)
    st.caption(f"🟢 Live: updates automatically every 30 seconds. Last updated {now_local():%I:%M:%S %p} IST.")


# Team passcode: set APP_PASSCODE in the app's Streamlit secrets to require it on the welcome page.
# Leave it out to open the app with no passcode. Change it there at any time; the old one stops working at once.
APP_PASSCODE = _secrets_get("APP_PASSCODE")


LOGIN_COOKIE = "cb_login"
LOGIN_DAYS = 30   # stay signed in on this browser for 30 days after the last visit


def _login_token():
    # Derived from the passcode, so changing APP_PASSCODE signs every browser out.
    return hmac.new(str(APP_PASSCODE).strip().encode(), b"company-builds-login", "sha256").hexdigest()


def _valid_login_token(saved):
    return isinstance(saved, str) and bool(saved) and hmac.compare_digest(saved.encode(), _login_token().encode())


def _needs_passcode():
    if not APP_PASSCODE or st.session_state.get("authed"):
        return False
    # A browser that signed in before carries the login cookie: let it straight in.
    # (Read here where the server sees cookies, e.g. locally; on Streamlit Cloud _sync_login_cookie reads it.)
    try:
        saved = st.context.cookies.get(LOGIN_COOKIE, "")
    except Exception:
        saved = ""
    if _valid_login_token(saved):
        st.session_state["authed"] = True
        st.session_state.setdefault("started", True)
        return False
    return True


# Small invisible helper that runs in the browser: it reads the login cookie and hands it to the app,
# and once you're signed in it saves (and renews) the cookie. Streamlit Cloud doesn't pass browser
# cookies on to the app, so the browser has to send it this way.
_LOGIN_BRIDGE_JS = """
export default function(component) {
    const { data, setStateValue } = component;
    const name = data.name;
    if (data.set) {
        const secure = window.location.protocol === "https:" ? "; Secure" : "";
        document.cookie = name + "=" + data.set + "; Max-Age=" + data.maxAge + "; Path=/; SameSite=Lax" + secure;
        return;
    }
    const found = document.cookie.split("; ").find(c => c.startsWith(name + "="));
    const value = found ? found.slice(name.length + 1) : "";
    if (value && value !== data.sent) {
        setStateValue("token", value);
    }
}
"""
try:
    _login_bridge = st.components.v2.component("cb_login_bridge", js=_LOGIN_BRIDGE_JS)
except Exception:          # older Streamlit versions: the login then lasts for this visit only
    _login_bridge = None


# Small invisible helper that keeps your place: the builds table reloads after every save, which
# would jump the page (and the table's own scroll) back to the top. It remembers where you had
# scrolled to and puts you back there once the table has reloaded.
_SCROLL_KEEPER_JS = """
export default function(component) {
    if (window.__cbKeep) return;
    const keep = window.__cbKeep = { page: null, grids: {}, nodes: {}, active: 0, mouse: false, until: 0 };
    const WRAP = '[class*="st-key-cbkeep_"]';
    const wrapName = el => { const w = el.closest(WRAP); return w ? [...w.classList].find(c => c.startsWith("st-key-cbkeep_")) : null; };
    const touch = () => { keep.active = Date.now(); keep.until = 0; };
    ["wheel", "touchmove", "keydown"].forEach(ev => document.addEventListener(ev, touch, true));
    document.addEventListener("mousedown", () => { keep.mouse = true; touch(); }, true);
    document.addEventListener("mouseup", () => { keep.mouse = false; }, true);
    const userScrolling = () => keep.mouse || Date.now() - keep.active < 500;
    const gridsPresent = () => [...document.querySelectorAll(WRAP)].every(w => w.querySelector(".dvn-scroller"));
    document.addEventListener("scroll", e => {
        if (!userScrolling() || Date.now() < keep.until) return;
        const t = e.target === document ? document.scrollingElement : e.target;
        if (!(t instanceof Element)) return;
        if (t.classList.contains("dvn-scroller")) {
            const name = wrapName(t);
            if (name) keep.grids[name] = { left: t.scrollLeft, top: t.scrollTop };
        } else if (!t.closest(".dvn-scroller") && gridsPresent()) {
            keep.page = { el: t, top: t.scrollTop };
        }
    }, true);
    const restore = () => {
        if (keep.page && keep.page.el.isConnected && Math.abs(keep.page.el.scrollTop - keep.page.top) > 2)
            keep.page.el.scrollTop = keep.page.top;
        document.querySelectorAll(WRAP).forEach(w => {
            const name = wrapName(w), g = keep.grids[name], s = w.querySelector(".dvn-scroller");
            if (g && s) { s.scrollLeft = g.left; s.scrollTop = g.top; }
        });
    };
    new MutationObserver(() => {
        let fresh = false;
        document.querySelectorAll(WRAP).forEach(w => {
            const name = wrapName(w), s = w.querySelector(".dvn-scroller");
            if (s && keep.nodes[name] !== s) { keep.nodes[name] = s; fresh = true; }
        });
        if (!fresh) return;
        const start = Date.now();
        keep.until = start + 1500;
        const loop = () => {
            if (keep.active > start || Date.now() > start + 1500) return;
            restore();
            requestAnimationFrame(loop);
        };
        requestAnimationFrame(loop);
    }).observe(document.body, { childList: true, subtree: true });
}
"""
try:
    _scroll_keeper = st.components.v2.component("cb_scroll_keeper", js=_SCROLL_KEEPER_JS)
except Exception:
    _scroll_keeper = None


def _keep_scroll_position():
    if _scroll_keeper is None:
        return
    try:
        st.html("<style>.st-key-cb_scroll_keeper_box { display: none; }</style>")
        with st.container(key="cb_scroll_keeper_box"):
            _scroll_keeper(key="cb_scroll_keeper", data={})
    except Exception:
        pass


def _sync_login_cookie():
    """Signs in a browser that carries a valid login cookie, and saves the cookie once signed in."""
    if not APP_PASSCODE or _login_bridge is None:
        return
    authed = bool(st.session_state.get("authed"))
    data = {"name": LOGIN_COOKIE, "sent": st.session_state.get("login_cookie_seen", "")}
    if authed:
        data.update(set=_login_token(), maxAge=LOGIN_DAYS * 86400)
    try:
        st.html("<style>.st-key-cb_login_bridge_box { display: none; }</style>")
        with st.container(key="cb_login_bridge_box"):
            res = _login_bridge(key="cb_login_bridge", data=data, on_token_change=lambda: None)
    except Exception:
        return
    token = getattr(res, "token", None)
    if authed or not token:
        return
    st.session_state["login_cookie_seen"] = token     # so the browser doesn't keep re-sending it
    if _valid_login_token(token):
        st.session_state["authed"] = True
        st.session_state.setdefault("started", True)


def _start():
    if _needs_passcode():
        typed = str(st.session_state.get("passcode_input", "")).strip()
        if not hmac.compare_digest(typed.encode(), str(APP_PASSCODE).strip().encode()):
            st.session_state["passcode_error"] = True
            return
        st.session_state["authed"] = True       # also saved as a cookie, see _sync_login_cookie
        st.session_state.pop("passcode_input", None)
    st.session_state["started"] = True


def _go_home():
    st.session_state["started"] = False


def render_landing_page():
    """Welcome page shown before the app. 'Get started' opens the dashboard."""
    now = now_local()
    greeting = "Good morning" if now.hour < 12 else "Good afternoon" if now.hour < 17 else "Good evening"

    st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;800&family=Michroma&display=swap');
.lp { font-family: 'Manrope', system-ui, sans-serif; padding-top: 1rem; }
.lp-logo { font-family: 'Michroma', 'Eurostile', system-ui, sans-serif; font-weight: 400;
           font-size: clamp(1.5rem, 2.4vw, 2rem); letter-spacing: 0.07em; line-height: 1;
           white-space: nowrap; margin: 0 0 4.5rem clamp(-2.5rem, -2.5vw, -0.75rem);
           -webkit-text-stroke: 0.4px currentColor; }
.lp-logo .sc { font-size: 0.82em; }
.lp-logo sup { font-size: 0.3em; letter-spacing: 0; vertical-align: 0.25em; margin-left: 0.15em;
               -webkit-text-stroke: 0; }
.lp-brand { font-weight: 800; font-size: clamp(1.4rem, 2.5vw, 2.1rem); line-height: 1.1;
            letter-spacing: -0.02em; white-space: nowrap; margin: 0 0 2.2rem; }
.lp-greet { font-size: 1.05rem; opacity: 0.8; margin: 0 0 0.9rem; }
.lp-title { font-family: 'Manrope', system-ui, sans-serif;
            font-weight: 800; font-size: clamp(2.4rem, 5vw, 3.7rem); line-height: 1.04;
            letter-spacing: -0.025em; margin: 0 0 1.2rem; }
.lp-sub { font-size: 1.08rem; line-height: 1.6; max-width: 32rem; opacity: 0.85; margin: 0 0 1.8rem; }
</style>
""", unsafe_allow_html=True)

    left, right = st.columns([1.15, 1], gap="large")
    with left:
        st.markdown(f"""
<div class="lp">
  <div class="lp-logo" aria-label="Netchex">N<span class="sc">ETCHEX</span><sup>®</sup></div>
  <div class="lp-brand" role="heading" aria-level="1">🏗️ Company Builds</div>
  <p class="lp-greet">{greeting}. It's {now:%A}, {now.day} {now:%B}.</p>
  <div class="lp-title" role="heading" aria-level="2">Build it.<br>QC it.<br>Hand it over.</div>
  <p class="lp-sub">Track every company build and its QC checklist in one place,
  from the first Salesforce update to the last GuideCX task.</p>
</div>
""", unsafe_allow_html=True)
        b, _ = st.columns([1, 1.4])
        if _needs_passcode():
            # Team passcode: Enter or the button both submit.
            with b.form("passcode_form", border=False):
                st.text_input("Team passcode", type="password", key="passcode_input",
                              placeholder="Enter the team passcode")
                st.form_submit_button("Get started  →", type="primary", width="stretch", on_click=_start)
            if st.session_state.pop("passcode_error", False):
                b.error("That passcode isn't right. Ask your team lead for the current one.")
        else:
            b.button("Get started  →", type="primary", width="stretch", on_click=_start, key="get_started")

    # Illustration: a month's build list, a sticky note from the QC checklist and today's date.
    sample = [("Riverbend Dental", "Completed"), ("Oak & Pine Bistro", "Completed"),
              ("Harbor Logistics", "In Progress"), ("Summit Clinics", "Pending")]
    pill = {"Completed": "done", "In Progress": "prog", "Pending": "pend"}
    rows = "".join(f'<div class="row"><span class="co">{n}</span><span class="pill {pill[s]}">{s}</span></div>'
                   for n, s in sample)
    note = "".join(f'<div class="nrow"><span>{t}</span><b>{m}</b></div>'
                   for t, m in [("Salesforce", "✓"), ("Tax codes", "✓"), ("Billing", "✓"), ("Feature flags", "…")])
    with right:
        embed = getattr(st, "iframe", None)
        if embed is None:                      # older Streamlit versions
            import streamlit.components.v1 as components
            embed = components.html
        embed(f"""
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@400;600;800&display=swap" rel="stylesheet">
<style>
  html, body {{ margin: 0; background: transparent; font-family: 'Manrope', system-ui, sans-serif; overflow: hidden; }}
  .wrap {{ height: 560px; display: flex; align-items: center; justify-content: center; }}
  .stage {{ position: relative; width: 440px; height: 545px; }}
  @media (max-width: 440px) {{ .stage {{ transform: scale(0.82); transform-origin: center; }} }}

  .card {{ position: absolute; left: 30px; top: 122px; width: 300px; background: #FCFCFA; color: #13233A;
           border-radius: 8px; padding: 22px 24px 16px; transform: rotate(-3deg); z-index: 1;
           box-shadow: 0 22px 45px rgba(19, 35, 58, 0.30), 0 2px 6px rgba(19, 35, 58, 0.15); }}
  .head {{ display: flex; justify-content: space-between; align-items: baseline;
           border-bottom: 2px solid #13233A; padding-bottom: 8px; margin-bottom: 4px; }}
  .head b {{ font-size: 19px; font-weight: 800; }}
  .head span {{ font-size: 12px; color: #5B6B80; }}
  .row {{ display: flex; justify-content: space-between; align-items: center; font-size: 14px;
          padding: 11px 0; border-bottom: 1px dashed #CDD5DF; }}
  .row:last-child {{ border-bottom: 0; }}
  .co {{ font-weight: 600; }}
  .pill {{ font-size: 11px; font-weight: 700; padding: 3px 9px; border-radius: 999px; }}
  .done {{ background: #D7F0E0; color: #1E7A45; }}
  .prog {{ color: #8A5A00; background: repeating-linear-gradient(-45deg, #FFE3B0 0 6px, #FFD48A 6px 12px);
           background-size: 17px 17px; animation: slide 1.2s linear infinite; }}
  @keyframes slide {{ to {{ background-position: 17px 0; }} }}
  .pend {{ background: #E6EBF1; color: #5B6B80; }}
  .bar {{ height: 8px; border-radius: 4px; background: #E6EBF1; overflow: hidden; margin: 12px 0 4px; }}
  .bar i {{ display: block; height: 100%; width: 50%; background: #2E9E5B; border-radius: 4px; }}
  .foot {{ font-size: 11px; color: #5B6B80; }}

  .note {{ position: absolute; right: -4px; top: 0; width: 160px; background: #FFE9A8; color: #3D3212;
           padding: 14px 16px 12px; transform: rotate(6deg); z-index: 2; border-radius: 2px;
           box-shadow: 0 14px 28px rgba(0, 0, 0, 0.28); }}
  .note::before {{ content: ""; position: absolute; top: -9px; left: 50%; width: 54px; height: 18px;
                   transform: translateX(-50%) rotate(-4deg); background: rgba(255, 255, 255, 0.55); }}
  .note h4 {{ margin: 0 0 8px; font-size: 14px; font-weight: 800; }}
  .nrow {{ display: flex; justify-content: space-between; font-size: 13px; padding: 3px 0;
           border-bottom: 1px solid rgba(61, 50, 18, 0.15); }}
  .nrow:last-child {{ border-bottom: 0; }}

  .cal {{ position: absolute; left: 4px; bottom: 0; width: 118px; background: #FFFFFF; color: #13233A;
          border-radius: 10px; overflow: hidden; transform: rotate(-7deg); z-index: 2; text-align: center;
          box-shadow: 0 14px 28px rgba(0, 0, 0, 0.28); }}
  .cal .m {{ background: #2F5D9A; color: #FFFFFF; font-size: 13px; font-weight: 700; padding: 5px 0; }}
  .cal .n {{ font-size: 38px; font-weight: 800; line-height: 1; padding: 10px 0 2px; }}
  .cal .t {{ font-size: 12px; color: #2F5D9A; font-weight: 600; padding: 0 0 10px; }}
</style>
<div class="wrap"><div class="stage">
  <div class="card">
    <div class="head"><b>{now:%B} builds</b><span>sample</span></div>
    {rows}
    <div class="bar"><i></i></div><div class="foot">2 of 4 completed</div>
  </div>
  <div class="note"><h4>QC checklist</h4>{note}</div>
  <div class="cal"><div class="m">{now:%B}</div><div class="n">{now.day}</div><div class="t">{now:%A}</div></div>
</div></div>
""", height=570)


# ─────────────────────────────────────────────────────────────
# SF QC reports: the Excel QC file for a build, stored in Supabase Storage (bucket "qc-reports")
# with one row per upload in the qc_reports table. Run supabase_qc_reports.sql once to set it up.
# ─────────────────────────────────────────────────────────────
QC_BUCKET = "qc-reports"
QC_RESULTS = ("Pass", "Fail", "Needs eyes", "No source", "N/A")
QC_REPORT_COLS = ["id", "build_id", "company_code", "file_name", "file_path", "pct_pass", "pass_count",
                  "fail_count", "needs_eyes_count", "no_source_count", "total_checks", "uploaded_by", "uploaded_at"]
QC_EXTRA_COLS = ["report_type", "na_count"]     # added by supabase_netchex_qc_reports.sql
QC_SETUP_HINT = ("SF QC reports aren't set up yet. In Supabase open SQL Editor → New query, paste "
                 "**supabase_qc_reports.sql**, click Run, then refresh this page.")
NX_SETUP_HINT = ("Netchex QC reports need one more setup step. In Supabase open SQL Editor → New query, paste "
                 "**supabase_netchex_qc_reports.sql**, click Run, then refresh this page.")
# Two kinds of QC report share the qc_reports table, told apart by report_type.
QC_KINDS = {
    "sf": {"label": "SF QC", "sheet": "SF QC", "file": "'<Client> (<CODE>) - SF QC.xlsx'"},
    "netchex": {"label": "Netchex QC", "sheet": "Netchex QC", "file": "'<Client> (<CODE>) - Netchex QC.xlsx'"},
}


def parse_qc_workbook(data, kind="sf"):
    """Reads an SF QC or Netchex QC workbook (new or old layout) and counts its results straight from the
    Result column, so it works even when Excel hasn't saved the formula values. Returns a dict, or raises
    ValueError. % Pass leaves out N/A rows (same as the % PASS tile in the Netchex QC file)."""
    from openpyxl import load_workbook
    try:
        wb = load_workbook(BytesIO(data), data_only=True, read_only=True)
    except Exception:
        raise ValueError("This isn't an Excel (.xlsx) file.")
    want, other = QC_KINDS[kind]["sheet"], QC_KINDS["netchex" if kind == "sf" else "sf"]["sheet"]
    if want not in wb.sheetnames and other in wb.sheetnames:
        raise ValueError(f"This is the {other} workbook. Upload it in the {other} report box instead.")
    ws = wb[want] if want in wb.sheetnames else wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    title = str(rows[0][0]) if rows and rows[0] and rows[0][0] else ""
    hdr = next((i for i, r in enumerate(rows[:15]) if r and "Result" in [str(v).strip() if v else "" for v in r]), None)
    if hdr is None:
        raise ValueError(f"No Result column found. Upload the {QC_KINDS[kind]['label']} workbook made by the QC.")
    col = [str(v).strip() if v else "" for v in rows[hdr]].index("Result")
    counts = {k: 0 for k in QC_RESULTS}
    for r in rows[hdr + 1:]:
        v = str(r[col]).strip() if r and len(r) > col and r[col] else ""
        if v in counts:
            counts[v] += 1
    total = sum(counts.values())
    scored = total - counts["N/A"]
    if not scored:
        raise ValueError("No Pass / Fail / Needs eyes results found in this file.")
    m = re.search(r"\(([A-Z0-9]{2,5})\)", title.upper())
    return {"title": title, "company_code": m.group(1) if m else None, "counts": counts, "total": total,
            "pct_pass": round(100 * counts["Pass"] / scored, 1), "kind": kind}


QC_REPORTS = {"enabled": True, "netchex": True}


def load_qc_reports():
    """All uploaded SF QC reports, newest first. Empty (and setup hint shown) if the table isn't there yet."""
    cols = QC_REPORT_COLS + QC_EXTRA_COLS
    try:
        if USE_SUPABASE:
            try:
                df = pd.DataFrame(_fetch_all("qc_reports", ",".join(cols)), columns=cols)
                QC_REPORTS["netchex"] = True
            except Exception:
                # Netchex QC columns not added yet: SF QC keeps working, Netchex QC shows the setup hint.
                df = pd.DataFrame(_fetch_all("qc_reports", ",".join(QC_REPORT_COLS)), columns=QC_REPORT_COLS)
                QC_REPORTS["netchex"] = False
        else:
            conn = sqlite_conn()
            conn.execute(f"""CREATE TABLE IF NOT EXISTS qc_reports (id INTEGER PRIMARY KEY AUTOINCREMENT,
                build_id INTEGER, company_code TEXT, file_name TEXT, file_path TEXT, pct_pass REAL,
                pass_count INTEGER, fail_count INTEGER, needs_eyes_count INTEGER, no_source_count INTEGER,
                total_checks INTEGER, uploaded_by TEXT, uploaded_at TEXT)""")
            have = {r[1] for r in conn.execute("PRAGMA table_info(qc_reports)")}
            if "report_type" not in have:
                conn.execute("ALTER TABLE qc_reports ADD COLUMN report_type TEXT DEFAULT 'sf'")
            if "na_count" not in have:
                conn.execute("ALTER TABLE qc_reports ADD COLUMN na_count INTEGER DEFAULT 0")
            conn.commit()
            df = pd.read_sql_query(f"SELECT {', '.join(cols)} FROM qc_reports", conn)
            conn.close()
            QC_REPORTS["netchex"] = True
        QC_REPORTS["enabled"] = True
    except Exception:
        QC_REPORTS["enabled"] = False
        QC_REPORTS["netchex"] = False
        df = pd.DataFrame(columns=cols)
    for c in QC_EXTRA_COLS:
        if c not in df.columns:
            df[c] = None
    df["report_type"] = df["report_type"].fillna("sf")
    df["na_count"] = pd.to_numeric(df["na_count"], errors="coerce").fillna(0)
    if not df.empty:
        df["uploaded_at"] = pd.to_datetime(df["uploaded_at"], errors="coerce", utc=True).dt.tz_convert(TZ)
        df = df.sort_values("uploaded_at", ascending=False)
    return df


def _qc_local_dir():
    d = os.path.join(os.path.dirname(os.path.abspath(__file__)), "qc_reports")
    os.makedirs(d, exist_ok=True)
    return d


def save_qc_report(build_id, company_code, file_name, data, parsed, uploaded_by):
    kind = parsed.get("kind", "sf")
    stamp = now_local().strftime("%Y%m%d-%H%M%S")
    safe = re.sub(r"[^A-Za-z0-9._() -]", "_", file_name)
    path = f"{(company_code or 'NO-CODE').upper()}/{'netchex/' if kind == 'netchex' else ''}{stamp}_{safe}"
    c = parsed["counts"]
    row = {"build_id": int(build_id) if build_id is not None else None, "company_code": company_code,
           "file_name": file_name, "file_path": path, "pct_pass": parsed["pct_pass"],
           "pass_count": c["Pass"], "fail_count": c["Fail"], "needs_eyes_count": c["Needs eyes"],
           "no_source_count": c["No source"], "total_checks": parsed["total"], "uploaded_by": uploaded_by,
           "uploaded_at": datetime.now(TZ).isoformat()}
    if kind == "netchex":       # SF rows leave these to the column defaults, so SF works before the Netchex setup
        row.update({"report_type": "netchex", "na_count": c["N/A"]})
    if USE_SUPABASE:
        sb = get_supabase()
        sb.storage.from_(QC_BUCKET).upload(path, data, {
            "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "upsert": "true"})
        sb.table("qc_reports").insert(row).execute()
        return
    with open(os.path.join(_qc_local_dir(), path.replace("/", "__")), "wb") as f:
        f.write(data)
    conn = sqlite_conn()
    conn.execute(f"INSERT INTO qc_reports({','.join(row)}) VALUES({','.join(':' + k for k in row)})", row)
    conn.commit(); conn.close()


@st.cache_data(ttl=600, show_spinner=False)
def download_qc_report(path):
    if USE_SUPABASE:
        return get_supabase().storage.from_(QC_BUCKET).download(path)
    with open(os.path.join(_qc_local_dir(), path.replace("/", "__")), "rb") as f:
        return f.read()


def delete_qc_report(report):
    if USE_SUPABASE:
        sb = get_supabase()
        sb.storage.from_(QC_BUCKET).remove([report.file_path])
        sb.table("qc_reports").delete().eq("id", int(report.id)).execute()
        return
    try:
        os.remove(os.path.join(_qc_local_dir(), report.file_path.replace("/", "__")))
    except OSError:
        pass
    conn = sqlite_conn()
    conn.execute("DELETE FROM qc_reports WHERE id=?", (int(report.id),))
    conn.commit(); conn.close()


def _qc_score_text(r):
    return (f"{r.pct_pass:.0f}% Pass · {int(r.fail_count)} Fail · {int(r.needs_eyes_count)} Needs eyes"
            + (f" · {int(r.no_source_count)} No source" if int(r.no_source_count or 0) else "")
            + (f" · {int(r.na_count)} N/A" if int(getattr(r, "na_count", 0) or 0) else ""))


def _qc_upload_box(builds, key, build=None, kind="sf"):
    """Upload widget. With `build`, the file is attached to that build; otherwise the build is matched by the
    company code in the file's title (e.g. 'Abilene Country Club (1M7)') or picked from a list."""
    k = QC_KINDS[kind]
    up = st.file_uploader(f"Upload the {k['label']} workbook (.xlsx)", type=["xlsx"], key=f"{key}_file",
                          help=f"The {k['file']} file from the QC. Its score is read from the file.")
    if not up:
        return
    data = up.getvalue()
    try:
        parsed = parse_qc_workbook(data, kind)
    except ValueError as e:
        st.error(str(e)); return
    c = parsed["counts"]
    tail = f"{c['N/A']} N/A" if kind == "netchex" else f"{c['No source']} No source"
    st.caption(f"Read from the file: **{parsed['pct_pass']:.1f}% Pass** · {c['Pass']} Pass · {c['Fail']} Fail · "
               f"{c['Needs eyes']} Needs eyes · {tail} · {parsed['total']} checks")
    if build is None:
        code = parsed["company_code"]
        match = builds[builds.company_code == code] if code else builds.iloc[0:0]
        options = {int(r.id): f"{r.company_name} ({r.company_code or 'no code'})"
                   for r in builds.sort_values("company_name").itertuples()}
        default = int(match.iloc[0].id) if not match.empty else None
        build_id = st.selectbox("Build", list(options), format_func=options.get, key=f"{key}_build",
                                index=list(options).index(default) if default in options else None,
                                placeholder="Pick the build this QC is for")
        if code and match.empty:
            st.warning(f"No build has company code {code} yet. Pick the build, or add the code to it in Monthly Builds.")
    else:
        build_id = int(build.id)
        code = _clean_text(build.company_code)
        if parsed["company_code"] and code and parsed["company_code"] != code.upper():
            st.warning(f"This file is for {parsed['company_code']}, but this build's code is {code}. Check it's the right file.")
    who = st.selectbox("Uploaded by", TEAM_MEMBERS, key=f"{key}_who",
                       index=TEAM_MEMBERS.index("Fiza") if "Fiza" in TEAM_MEMBERS else 0)
    if st.button(f"⬆️ Save {k['label']} report", key=f"{key}_save", type="primary", disabled=build_id is None):
        b = builds[builds.id == build_id].iloc[0]
        try:
            save_qc_report(build_id, _clean_text(b.company_code) or parsed["company_code"], up.name, data, parsed, who)
        except Exception as e:
            hint = NX_SETUP_HINT if kind == "netchex" else QC_SETUP_HINT
            st.error(f"The {k['label']} report could not be saved. {hint} ({e})"); return
        st.session_state["_toast"] = f"✅ {k['label']} report saved for {b.company_name}: {parsed['pct_pass']:.0f}% Pass"
        st.session_state.pop(f"{key}_file", None)
        st.rerun()


def _qc_download(r, key):
    try:
        data = download_qc_report(r.file_path)
    except Exception as e:
        st.caption(f"Download unavailable ({e})"); return
    st.download_button("⬇️ Download", data=data, file_name=r.file_name, key=key,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


def _render_qc_for_build(builds, build, kind="sf"):
    """SF QC or Netchex QC report box on a build's checklist page."""
    label = QC_KINDS[kind]["label"]
    st.space("small")
    with st.container(border=True):
        st.markdown(f"**{label} report**")
        if not QC_REPORTS["enabled"]:
            st.info(QC_SETUP_HINT); return
        reports = load_qc_reports()
        if kind == "netchex" and not QC_REPORTS["netchex"]:
            st.info(NX_SETUP_HINT); return
        mine = reports[(reports.build_id == int(build.id)) & (reports.report_type == kind)]
        if not mine.empty:
            r = mine.iloc[0]
            c = st.columns([1, 1, 1, 1, 1.4])
            c[0].metric("% Pass", f"{r.pct_pass:.0f}%")
            c[1].metric("Fail", int(r.fail_count))
            c[2].metric("Needs eyes", int(r.needs_eyes_count))
            c[3].metric("Checks", int(r.total_checks))
            with c[4]:
                st.caption(f"{r.file_name}\n\nUploaded by {r.uploaded_by or '—'}")
                _qc_download(r, f"{kind}dl_build_{int(r.id)}")
            st.caption("Changed some Needs eyes to Pass in the file? Upload it again: the new file and score replace this one.")
        else:
            st.caption(f"No {label} report yet for this build.")
        prefix = "nxqc" if kind == "netchex" else "qc"
        _qc_upload_box(builds, f"{prefix}up_{int(build.id)}", build=build, kind=kind)


def _render_qc_reports(builds):
    """📑 SF QC Reports section: every build's latest QC score and file, plus upload."""
    if not QC_REPORTS["enabled"]:
        st.info(QC_SETUP_HINT); return
    reports = load_qc_reports()
    reports = reports[reports.report_type == "sf"]
    with st.expander("⬆️ Upload an SF QC report", expanded=reports.empty):
        _qc_upload_box(builds, "qcup_main")
    if reports.empty:
        st.info("No SF QC reports uploaded yet."); return
    latest = reports.drop_duplicates("build_id", keep="first").copy()
    names = builds.set_index("id")
    latest["Company Name"] = latest.build_id.map(lambda i: names.company_name.get(int(i), "") if pd.notna(i) else "")
    latest["QC By"] = latest.build_id.map(lambda i: names.qc_by.get(int(i), "") if pd.notna(i) else "")
    latest["Built By"] = latest.build_id.map(lambda i: names.built_by.get(int(i), "") if pd.notna(i) else "")
    latest["Code"] = [(_clean_text(names.company_code.get(int(i))) if pd.notna(i) and int(i) in names.index else None)
                      or _clean_text(c) for i, c in zip(latest.build_id, latest.company_code)]
    c = st.columns(4)
    c[0].metric("Builds with a QC report", len(latest))
    c[1].metric("Average % Pass", f"{latest.pct_pass.mean():.0f}%")
    c[2].metric("Builds with Fails", int((latest.fail_count > 0).sum()))
    c[3].metric("Builds with Needs eyes", int((latest.needs_eyes_count > 0).sum()))
    table = pd.DataFrame({
        "Company Name": latest["Company Name"], "Company Code": latest["Code"],
        "% Pass": latest.pct_pass / 100, "Pass": latest.pass_count, "Fail": latest.fail_count,
        "Needs eyes": latest.needs_eyes_count, "No source": latest.no_source_count,
        "Checks": latest.total_checks, "Built By": latest["Built By"], "QC By": latest["QC By"],
        "Uploaded by": latest.uploaded_by, "Versions": latest.build_id.map(reports.build_id.value_counts()),
    }).reset_index(drop=True)
    st.markdown("**Latest QC per build** · click a row to download its file")
    event = st.dataframe(
        table, key="qc_reports_table", hide_index=True, width="stretch", on_select="rerun",
        selection_mode="single-row", height=min(35 * len(table) + 40, 520),
        column_config={"Company Name": st.column_config.TextColumn(pinned=True),
                       "% Pass": st.column_config.ProgressColumn("% Pass", format="percent", min_value=0, max_value=1)})
    if event.selection.rows:
        r = latest.iloc[event.selection.rows[0]]
        with st.container(border=True):
            st.markdown(f"**{r['Company Name'] or r.company_code}** — {_qc_score_text(r)}")
            a, b2, _ = st.columns([1, 1, 3])
            with a:
                _qc_download(r, f"qcdl_main_{int(r.id)}")
            with b2:
                if st.button("🗑️ Delete this version", key=f"qcdel_{int(r.id)}"):
                    try:
                        delete_qc_report(r)
                        st.session_state["_toast"] = "🗑️ QC report deleted"
                    except Exception as e:
                        st.session_state["_flash_error"] = f"Could not delete the report. ({e})"
                    st.rerun()


SECTIONS = {"📊 Dashboard": "Build counts and charts, all months",
            "📅 Monthly Builds": "View, add, edit and delete builds by month",
            "✅ Quality Check": "Checklist by team member and build"}
MONTH_VIEWS = ("📅 Monthly Builds",)    # sections that use the sidebar Month picker
LIVE_VIEWS = ("📊 Dashboard",)          # sections that refresh themselves every 30 seconds


def render_builds_page():
    try:
        builds = load_builds()
        checks = load_build_checks()
        load_qc_reports()          # only to know whether the SF QC reports table exists yet
    except Exception as e:
        st.error("Company Builds could not load its data. Check that the builds and build_checks "
                 "tables exist in Supabase.")
        with st.expander("Error details"):
            st.code(str(e))
        st.stop()

    if "_flash" in st.session_state:
        st.success(st.session_state.pop("_flash"))
    if "_toast" in st.session_state:
        st.toast(st.session_state.pop("_toast"))
    if "_flash_error" in st.session_state:
        st.error(st.session_state.pop("_flash_error"))

    with st.sidebar:
        st.markdown("""<style>
@import url('https://fonts.googleapis.com/css2?family=Michroma&display=swap');
.sb-logo { font-family: 'Michroma', system-ui, sans-serif; font-size: 1.15rem; letter-spacing: 0.07em;
           -webkit-text-stroke: 0.3px currentColor; margin: 0 0 0.3rem; }
.sb-logo .sc { font-size: 0.82em; }
</style><div class="sb-logo">N<span class="sc">ETCHEX</span></div>""", unsafe_allow_html=True)
        st.markdown("## 🏗️ Company Builds")
        st.caption("Builds and QC for the CoE team")
        st.divider()
        st.markdown("**Sections**")
        view = st.radio("Sections", list(SECTIONS), captions=list(SECTIONS.values()),
                        label_visibility="collapsed", key="builds_view")
        pick_month = None
        if view in MONTH_VIEWS:
            st.divider()
            opts = _month_options(builds)
            current = month_start(today_local())
            # One month choice shared by the Monthly page and the Builds table.
            if st.session_state.get("month") not in opts:
                st.session_state["month"] = current if current in opts else opts[0]
            pick_month = st.selectbox("Month", opts, key="month", format_func=lambda o: f"{o:%B %Y}")
        st.divider()
        if view in LIVE_VIEWS:
            st.caption("🟢 This page refreshes on its own every 30 seconds.")
        else:
            if st.button("🔄 Refresh data", key="builds_refresh", width="stretch",
                         help="Load the latest changes saved by others. Save your own edits first, "
                              "because refreshing clears anything not yet saved."):
                st.rerun()
            st.caption(f"Data as of {now_local():%I:%M %p} IST. Click Refresh to see changes others have saved.")
        st.button("🏠 Home", key="go_home", width="stretch", on_click=_go_home)

    st.title(view)
    st.caption(SECTIONS[view])
    if view == "📊 Dashboard":
        _live(_render_overview)
    elif view == "📅 Monthly Builds":
        _render_month(builds, pick_month)
    else:
        _render_build_checklist(builds, checks)


# ─────────────────────────────────────────────────────────────
# App
# ─────────────────────────────────────────────────────────────
_sync_login_cookie()
if not _needs_passcode() and st.session_state.get("started"):
    render_builds_page()
else:
    render_landing_page()
