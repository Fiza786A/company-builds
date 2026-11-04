"""Company Builds – CoE team build tracker and QC checklist (no sign-in).

Run locally:  streamlit run app.py
Data: Supabase (SUPABASE_URL / SUPABASE_KEY in secrets), or a local SQLite file when those are not set.
"""
import calendar
import os
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
BUILD_STATUSES = ["Pending", "In Progress", "Completed"]
STATUS_LABEL = {"Pending": "⚪ Pending", "In Progress": "🟡 In Progress", "Completed": "🟢 Completed"}
# Payroll Leads who assign builds (US team). Add names here as needed.
PAYROLL_LEADS = ["Chelsi", "Dakota", "Jessikah", "Kristin", "Lauren", "Macy", "Tiffany"]

BUILD_COLS = ["id", "company_name", "company_code", "hierarchy", "status", "completion_date",
              "due_date", "built_by", "assigned_by", "qc_by", "qc_status"]
CHECK_COLS = ["build_id", "item_id", "built", "qc", "na"]
BUILD_FIELDS = {  # column shown in the tracker -> database field
    "Company Name": "company_name", "Company Code": "company_code", "Hierarchy": "hierarchy",
    "Status": "status", "Date of Completion": "completion_date", "Due Date": "due_date",
    "Built By": "built_by", "Assigned By": "assigned_by", "QC": "qc_by", "QC Status": "qc_status",
}
BUILD_POINTERS = """
- **Company Name** – Please update the correct company name.
- **Company Code** – Enter the company code once the import is completed.
- **Status** – Please update the current status as: Pending, In Progress, or Completed.
- **Date of Completion** – Please follow the format: MM/DD/YYYY.
- **Built By** – Name of the person who completed the company/shell build.
- **Assigned By** – Name of the Payroll Lead who assigned the build.
- **QC** – Name of the person who verified the company build.
"""

# 2025 Implementation Checklist - Direct Sales Clients.
# (section, note shown under the section, [(sub-heading, [items])]); "★ " marks key items
# (highlighted green on the QC page). Add new items at the END of a section so saved ticks stay aligned.
CHECKLIST = [
    ("Salesforce - Details", "", [
        ("Details page", ["Under the Account Information section, change Account Record Type from Prospect to Client"]),
        ("Account Financial Status", ["★ Enter Average Payroll Amount (enter 0.00 until Proof of DD is received, then update)"]),
        ("Account Information", [
            "Verify Account Name matches Legal Name on PIN (if not, update it so it matches)",
            "If Copy Company, verify account is tied to Parent Account",
            "Enter DBA Name (if applicable)",
            "Add Company Code to Client Company Code field",
            "Verify Federal ID matches PIN",
            "Update Phone Number (if applicable)",
            "Change TYPE from Prospect to Netchex Client",
            "Choose Pricing Model (PEPM or Legacy)",
            "Choose Package (Core, Complete, Legacy) – if PEPM, leave blank",
            "Choose Yes/No for Distributed Invoicing",
            "Choose Yes/No for Seasonal Client",
            "★ Check Currently in Implementation",
        ]),
        ("Address Information", ["Enter Shipping Address"]),
        ("Designated Points of Contact", [
            "★ Select Payroll Contact",
            "★ Select Tax Contact",
            "★ Select Billing/ACH Contact",
            "★ Select The Work Number Contact (if TWN is listed on PEPM proposal)",
            "Select Human Resources Contact (usually same as Work Number Contact)",
            "Select Benefits Contact (if Benefits is listed on PEPM proposal)",
        ]),
        ("Client Information", [
            "Enter the Netchex Annual Revenue (monthly total on PEPM proposal × 12)",
            "Check AZ/HI Daylight Savings Time (if client is in Arizona or Hawaii)",
        ]),
        ("Implementation Details", [
            "★ Select Yes or No for Direct Deposit Waiting Period Required",
            "★ Select Proof of Direct Deposit Received (once proof of DD is received)",
        ]),
        ("Employer Services Information", [
            "Update Primary Payroll Frequency",
            "Update Secondary Payroll Frequency (if applicable)",
            "Verify Client's Payroll Vendor is populated and correct",
            "401k Offered – Yes or No",
            "If yes, enter Current 401k Plan Provider (if applicable)",
        ]),
    ]),
    ("Salesforce - Contacts", "", [
        ("Contacts", [
            "Add/Update all Contacts (First Name, Last Name, Title, Phone Number, Email)",
            "Check Able to Request User Access within the contact (ONLY if the user is on the Client Admin Form)",
        ]),
    ]),
    ("Salesforce - Assets",
     "Do NOT activate Employee Navigator, NetBenefits, NetEnroll, NetCobra, ACA Central, Carrier Solutions, EDI Feed "
     "or any forms. Also do NOT activate these HR assets: NetInsight, NetEngage, NetGuide, NetLearn, NetPerform, "
     "NetRecruiter, NetScheduler, NetExpense/ExpensePath, Org Charts, Compensation Management, Fillable Forms, "
     "AskHR+AI or E-Verify. (HR Leads / Benefit Leads activate HR/Benefits assets.)", [
        ("Assets", ["Activate PAYROLL assets ONLY (see the list of categorized assets)"]),
    ]),
    ("Salesforce - Onboarding Details", "", [
        ("Onboarding Details", [
            "Select Project Manager (Owner)",
            "Select Data Specialist",
            "Select Payroll Onboarding Lead",
            "Select HR Onboarding Lead",
            "Select Payroll Survey Contact under Payroll Information",
            "Select Survey Contact under Client Contacts",
        ]),
    ]),
    ("Netchex - Tax, Worker's Comp, Job Codes", "", [
        ("Maint > Payroll > Tax Codes", [
            "Update or add State and SUTA Tax Codes",
            "Update or add Local Tax Codes (if applicable)",
        ]),
        ("Maint > Payroll > Workers Compensation Codes", [
            "Set up Worker's Comp Codes for each worked-in state built",
            "Update Worker's Comp Rates (move decimal 2 places to the left if converting the rate)",
        ]),
        ("Maint > HR > Job Codes", ["Set up Job Codes for each state built"]),
    ]),
    ("Netchex - Widgets, TWN", "", [
        ("Maint > HR > Dashboard Configuration > Add/Remove Widgets", [
            "My Profile", "Company Directory", "My Pay", "People", "Favorite Reports",
            "My Accrual Attendance", "Calendar", "Company Info", "Employee Search", "Payroll",
            "NetGuide (ONLY if on proposal – activate in System Claims first, then add widget)",
        ]),
        ("Reports > File Feeds", ["Set up The Work Number"]),
    ]),
    ("Netchex - Client Billing",
     "Fee - Premium (negative amounts, Base / Per Check): Weekly -18.00 / -1.35 · Bi-Weekly -23.00 / -2.10 · "
     "Semi-Monthly -23.00 / -2.10 · Monthly -43.00 / -3.65. Premium Pay steps can only be done once "
     "Netchex Premium Pay Client is checked in System Tables.", [
        ("Back Office > Enterprise > Client Billing Schedule", [
            "Choose Payroll Frequency",
            "Check Sales Tax box if Texas company",
            "Choose Sales Agent",
            "Choose Payroll Coordinator (Project Lead)",
            "Choose Sales Tax Code (if TX, choose Dallas 8.25%)",
            "In Fee - Premium, enter the appropriate negative amounts in Base and Per Check",
        ]),
        ("Premium Pay clients only", [
            "Check Premium Pay",
            "Uncheck Money Movement",
            "Zero out $25.00 Stop Payment Fee",
        ]),
        ("Maint > Company Structure > Client Billing View", ["View the Client Billing View to ensure accuracy"]),
    ]),
    ("Netchex - Company Information (FKA System Tables)", "", [
        ("Company Demographic Information", [
            "Verify Trade Name and W2 Name, Federal ID and Address match PIN",
            "Enter Client Phone Number",
            "Add Client Email Address",
        ]),
        ("Banking Information", [
            "Verify Banking Information – Routing & Account Numbers",
            "Enter PIN under ACH Contact – PIN #",
            "Check Minimum Wage Makeup if Restaurant/Hotel",
        ]),
        ("Payroll Information", [
            "Confirm Minimum Wage $ Amount",
            "Update Default Payroll Group",
            "Update Default Job Code",
            "Populate Units/Pay Period",
            "Change Default Employee Classification to W-2 Tipped if Restaurant/Hotel",
            "Check Allow Manager Self Service (if ESS/MSS is listed on proposal)",
            "Check Netchex Premium Pay Client? (Premium Pay clients only – see Billing)",
        ]),
        ("Claims – activate products based on PEPM proposal", [
            "NetRecruiter", "NetGuide", "Fillable Forms", "NetPerform", "NetScheduler",
            "Engagement Basic (activate for ALL clients – has feature flags)",
            "Engagement Core", "Engagement Advanced",
        ]),
    ]),
    ("Netchex - Feature Flags", "", [
        ("Maint > Security > New Feature Access", [
            "Activate Feature Flags based on products listed on PEPM proposal",
            "Activate Employee Engagement Basic (Community Basic) for ALL clients",
            "(Post 1st payroll) Activate Earned Wage Access for clients using T&A only",
            "Earned Wage Access: Product Suite > Earned Wage Access, activate the slider for the company",
            "Activate QBO (ONLY for QBO Integration clients)",
        ]),
    ]),
    ("Netchex - Department Structure, Payroll Maintenance",
     "Payroll Group hours: Weekly 40 · Bi-Weekly 80 · Semi-Monthly 86.6667 · Monthly 173.3334.", [
        ("Maint > More > Company Structure", [
            "Set up Div/BU/Dept structure",
            "Hide Main Department",
            "Set up Locations (all physical work locations)",
        ]),
        ("Maint > HR", ["Set up Positions (mandatory – if client gives none, mimic departments)"]),
        ("Payroll > Utilities", [
            "Build Payroll Schedule(s)",
            "Verify Payroll Schedule with client after building",
            "Close the Year",
            "Add Company Note (PEPM Package or PEPM - New Proposal + Initials // DD Waiting Period)",
            "Build Payroll Groups",
        ]),
    ]),
    ("Netchex - Payroll Codes & General Ledger", "", [
        ("Payroll Codes & GL", [
            "Update/Add Deduction & Earning Codes",
            "Delete unused Earning/Deduction Codes",
            "Build GL Base Account for each Earning/Deduction/Tax built",
            "Map GL Base Accounts",
        ]),
    ]),
    ("Netchex - Maintenance & Report Runner",
     "Shipping: Zone 2 $31 · Zone 3 $36.50 · Zone 4 $44 · Zone 5 $47 · Zone 6 $48 · Zone 7 $52.50 · "
     "Zone 8 $55 · Garnishments $1.25.", [
        ("Maint > HR > Email/Text Event Configuration", [
            "List client as \"To:\" recipient for Payroll Invoiced",
            "List client and project lead as \"To:\" recipients for Payroll Submitted",
            "List client and project lead as \"To:\" recipients for Wire Transfer Initiated",
            "List client and project lead as \"To:\" recipients for Wire Transfer Received",
        ]),
        ("Maint > HR", [
            "Set up Shipping Information",
            "Update Shipping Contact",
            "Update Payroll Contact",
            "Update W-2 Contact",
        ]),
        ("Maint > System Tables > Company Structure", [
            "Enter Check Header",
            "Check box Details 1, 3, 6, 7",
            "Premium Pay only: add \"(C/O Netchex Payroll)\" to the second line of the Check Header",
        ]),
        ("Back Office > More > RR", ["Assign company to Report Runner"]),
    ]),
    ("Netchex - Time and Attendance", "", [
        ("T&A > More > Settings > Web Terminal Setup", ["Web Terminal Setup"]),
        ("T&A > Company Rules", [
            "Set up Payroll Group Rules",
            "Shift Setup",
            "Terminal Group Setup",
            "Add Employees to Terminal Group",
        ]),
    ]),
    ("Salesforce - Cases", "", [
        ("Cases", [
            "Pre-Note Case – let flow to Prenote Queue",
            "Tax Case – assign to your Tax Rep (I'Neka Martin or James Brunson)",
            "User Access Case – let flow to User Access Queue",
            "Signature Case – assign to Macy, Jessica D or Jamie",
            "E-Verify Case (if on proposal) – assign to your HR Lead",
            "Pay Card Case (if client is interested in Pay Cards)",
        ]),
    ]),
    ("GuideCX", "", [
        ("GuideCX", ["Update completed tasks accordingly"]),
    ]),
]


def _checklist_items():
    items = []
    for si, (section, _note, groups) in enumerate(CHECKLIST, 1):
        n = 0
        for sub, texts in groups:
            for text in texts:
                n += 1
                key = text.startswith("★ ")
                items.append({"id": f"S{si:02d}-{n:02d}", "section": section, "sub": sub,
                              "item": text[2:] if key else text, "key": key})
    return items


CHECK_ITEMS = _checklist_items()


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
        df = pd.DataFrame(_fetch_all("builds", ",".join(BUILD_COLS)), columns=BUILD_COLS)
    else:
        conn = sqlite_conn()
        try:
            df = pd.read_sql_query(f"SELECT {', '.join(BUILD_COLS)} FROM builds", conn)
        finally:
            conn.close()
    for col in ("completion_date", "due_date"):
        df[col] = df[col].map(_to_date).astype(object)
    for col in ("company_name", "company_code", "hierarchy", "status", "built_by", "assigned_by", "qc_by", "qc_status"):
        df[col] = df[col].map(_clean_text).astype(object)
    df["id"] = df["id"].astype(int) if not df.empty else df["id"]
    return df


def load_build_checks():
    if USE_SUPABASE:
        df = pd.DataFrame(_fetch_all("build_checks", ",".join(CHECK_COLS)), columns=CHECK_COLS)
    else:
        conn = sqlite_conn()
        try:
            df = pd.read_sql_query(f"SELECT {', '.join(CHECK_COLS)} FROM build_checks", conn)
        finally:
            conn.close()
    for col in ("built", "qc", "na"):
        df[col] = df[col].fillna(False).astype(bool)
    return df


def _db_value(v):
    return v.isoformat() if isinstance(v, date) else v


def insert_builds(rows):
    rows = [{k: _db_value(v) for k, v in r.items()} for r in rows]
    if USE_SUPABASE:
        get_supabase().table("builds").insert(rows).execute()
        return
    cols = list(BUILD_FIELDS.values())
    conn = sqlite_conn()
    conn.executemany(f"INSERT INTO builds({','.join(cols)}) VALUES({','.join(':' + c for c in cols)})",
                     [{c: r.get(c) for c in cols} for r in rows])
    conn.commit(); conn.close()


def update_build(build_id, fields):
    fields = {k: _db_value(v) for k, v in fields.items()}
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
        get_supabase().table("build_checks").upsert(rows, on_conflict="build_id,item_id").execute()
        return
    conn = sqlite_conn()
    conn.executemany("""INSERT INTO build_checks(build_id,item_id,built,qc,na) VALUES(:build_id,:item_id,:built,:qc,:na)
        ON CONFLICT(build_id,item_id) DO UPDATE SET built=excluded.built, qc=excluded.qc, na=excluded.na""", rows)
    conn.commit(); conn.close()


def _progress(checks, build_id):
    c = checks[checks.build_id == build_id]
    na = int(c.na.sum())
    applicable = len(CHECK_ITEMS) - na
    return int((c.built & ~c.na).sum()), int((c.qc & ~c.na).sum()), applicable


def _builds_table(frame, checks):
    rows = []
    for r in frame.itertuples():
        built, qc, applicable = _progress(checks, r.id)
        rows.append({
            "id": r.id, "Company Name": r.company_name, "Company Code": r.company_code, "Hierarchy": r.hierarchy,
            "Status": r.status, "Date of Completion": r.completion_date, "Due Date": r.due_date,
            "Built By": r.built_by, "Assigned By": r.assigned_by, "QC": r.qc_by, "QC Status": r.qc_status,
            "Checklist": f"{built}/{applicable} built, {qc}/{applicable} QC" if built or qc else "Not started",
        })
    cols = ["id", *BUILD_FIELDS, "Checklist"]
    table = pd.DataFrame(rows, columns=cols)
    for col in ("Date of Completion", "Due Date"):
        table[col] = table[col].astype(object)
    return table


def _norm(field, value):
    if field in ("completion_date", "due_date"):
        return _to_date(value)
    return _clean_text(value)


def export_builds_excel(table, label):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill

    wb = Workbook()
    ws = wb.active
    ws.title = "Company Builds"
    ws.append([f"Company Builds – {label}"])
    ws["A1"].font = Font(bold=True, size=14)
    ws.append([])
    headers = [*BUILD_FIELDS, "Checklist"]
    ws.append(headers)
    for c in ws[3]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="FFE699")
    fills = {"Completed": "C6EFCE", "In Progress": "FFEB9C"}
    for _, r in table.iterrows():
        ws.append([py(r[h]) if h not in ("Date of Completion", "Due Date") else r[h] for h in headers])
        for col_name in ("Status", "QC Status"):
            cell = ws.cell(row=ws.max_row, column=headers.index(col_name) + 1)
            if cell.value in fills:
                cell.fill = PatternFill("solid", fgColor=fills[cell.value])
        for col_name in ("Date of Completion", "Due Date"):
            ws.cell(row=ws.max_row, column=headers.index(col_name) + 1).number_format = "mm/dd/yyyy"
    for col in ws.columns:
        ws.column_dimensions[col[0].column_letter].width = min(max(len(str(x.value or "")) for x in col[2:]) + 2, 45)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _render_build_tracker(builds, checks):
    with st.container(border=True):
        st.markdown("**📌 Important pointers**")
        st.markdown(BUILD_POINTERS)

    current = month_start(today_local())
    months = {month_start(d) for d in builds.due_date if d} | {current, add_months(current, 1)}
    options = ["open"] + sorted(months, reverse=True)
    choice = st.selectbox("Show", options, index=options.index(current), key="builds_month",
                          format_func=lambda o: "Open builds (all months)" if o == "open"
                          else f"{o:%B %Y} Company Builds")

    if choice == "open":
        frame, label = builds[(builds.status != "Completed") | (builds.qc_status != "Completed")], "Open builds"
    else:
        start, end = choice, month_end(choice)
        in_month = builds.due_date.map(lambda d: d is not None and start <= d <= end)
        no_due = builds.due_date.map(lambda d: d is None) & (choice == current)
        frame, label = builds[in_month | no_due], f"{choice:%B %Y}"
    frame = frame.sort_values(["due_date", "company_name"], key=lambda s: s.map(
        lambda v: (v is None, v) if not isinstance(v, str) else (False, v.lower())))

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Builds", len(frame))
    c2.metric("Pending", int((frame.status.fillna("Pending") == "Pending").sum()))
    c3.metric("In progress", int((frame.status == "In Progress").sum()))
    c4.metric("Completed", int((frame.status == "Completed").sum()))
    c5.metric("Awaiting QC", int(((frame.status == "Completed") & (frame.qc_status != "Completed")).sum()))

    st.caption("Click a cell to edit. Add a build in the empty row at the bottom, then click **Save changes**. "
               "Builds can't be deleted here.")
    table = _builds_table(frame, checks)
    team = TEAM_MEMBERS
    edited = st.data_editor(
        table, key=f"builds_editor_{choice}", hide_index=True, width="stretch", num_rows="add",
        disabled=["Checklist"],
        column_config={
            "id": None,
            "Company Name": st.column_config.TextColumn("Company Name", width=220, required=True),
            "Company Code": st.column_config.TextColumn("Company Code", width=115,
                                                        help="Enter after you create the company"),
            "Hierarchy": st.column_config.TextColumn("Hierarchy", width=130),
            "Status": st.column_config.SelectboxColumn("Status", options=BUILD_STATUSES, default="Pending", width=125,
                                                       format_func=lambda s: STATUS_LABEL.get(s, s)),
            "Date of Completion": st.column_config.DateColumn("Completed On", format="MM/DD/YYYY", width=120),
            "Due Date": st.column_config.DateColumn("Due Date", format="MM/DD/YYYY", width=100),
            "Built By": st.column_config.SelectboxColumn("Built By", options=team, width=100),
            "Assigned By": st.column_config.SelectboxColumn("Assigned By", options=PAYROLL_LEADS, width=105),
            "QC": st.column_config.SelectboxColumn("QC", options=team, width=100),
            "QC Status": st.column_config.SelectboxColumn("QC Status", options=BUILD_STATUSES, width=125,
                                                          format_func=lambda s: STATUS_LABEL.get(s, s)),
            "Checklist": st.column_config.TextColumn("Checklist", width=165,
                                                     help="Progress from the QC section"),
        })

    originals = {int(r["id"]): r for _, r in table.iterrows()}
    inserts, updates, problems = [], [], []
    for _, r in edited.iterrows():
        values = {f: _norm(f, r[c]) for c, f in BUILD_FIELDS.items()}
        rid = r.get("id")
        is_new = rid is None or pd.isna(rid)
        if is_new and not any(values.values()):
            continue
        if not values["company_name"]:
            problems.append("Every build needs a Company Name.")
            continue
        if is_new:
            values["status"] = values["status"] or "Pending"
            inserts.append(values)
        else:
            old = {f: _norm(f, originals[int(rid)][c]) for c, f in BUILD_FIELDS.items()}
            diff = {k: v for k, v in values.items() if v != old[k]}
            if diff:
                updates.append((int(rid), diff, values))

    pending = len(inserts) + len(updates)
    for msg in dict.fromkeys(problems):
        st.warning(msg)
    s1, s2 = st.columns([1, 1])
    if s1.button(f"💾 Save changes ({pending})" if pending else "💾 Save changes", type="primary",
                 width="stretch", disabled=pending == 0 or bool(problems), key="builds_save"):
        today = today_local()
        for values in inserts:
            if values["status"] == "Completed" and not values["completion_date"]:
                values["completion_date"] = today
        insert_builds(inserts)
        for rid, diff, values in updates:
            if values["status"] == "Completed" and not values["completion_date"]:
                diff["completion_date"] = today
            update_build(rid, diff)
        flash(f"Saved: {len(inserts)} new build(s), {len(updates)} updated.")
        st.rerun()
    s2.download_button(f"📥 Download Excel · {label}", data=export_builds_excel(table, label),
                       file_name=f"Company_Builds_{label.replace(' ', '_')}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch", key="builds_download")


def _render_build_checklist(builds, checks):
    if builds.empty:
        st.info("Add a build in the 📋 Builds section first.")
        return
    order = builds.sort_values("id", ascending=False)
    def txt(v):
        return _clean_text(v) or ""

    def due(v):
        d = _to_date(v)
        return f" · due {d:%m/%d/%Y}" if d else ""

    labels = {int(r.id): txt(r.company_name) + (f" ({txt(r.company_code)})" if txt(r.company_code) else "")
              + due(r.due_date) for r in order.itertuples()}
    build_id = st.selectbox("Company build", list(labels), format_func=labels.get, key="chk_build")
    b = builds[builds.id == build_id].iloc[0]
    status, qc_status = txt(b.status), txt(b.qc_status)

    st.markdown(f"**Built by:** {txt(b.built_by) or '—'} &nbsp;·&nbsp; **QC by:** {txt(b.qc_by) or '—'} &nbsp;·&nbsp; "
                f"**Status:** {STATUS_LABEL.get(status, status or '—')} &nbsp;·&nbsp; "
                f"**QC status:** {STATUS_LABEL.get(qc_status, qc_status or '—')}")
    st.caption("Set Built By, QC and the statuses in the 📋 Builds section. ★ = key items. "
               "Tick N/A for items that don't apply to this client; they're left out of the counts.")

    built, qc, applicable = _progress(checks, build_id)
    p1, p2 = st.columns(2)
    p1.progress(built / applicable if applicable else 0.0, text=f"Built: {built} of {applicable}")
    p2.progress(qc / applicable if applicable else 0.0, text=f"QC: {qc} of {applicable}")

    show = st.radio("Show", ["All items", "Not built yet", "Not QC'd yet"], horizontal=True, key="chk_show")
    saved = {r.item_id: r for r in checks[checks.build_id == build_id].itertuples()}
    save_top = st.container()

    changes = {}
    for si, (section, note, _groups) in enumerate(CHECKLIST, 1):
        items = [it for it in CHECK_ITEMS if it["section"] == section]
        rows = []
        for it in items:
            s = saved.get(it["id"])
            rows.append({"item_id": it["id"], "Area": it["sub"], "Item": ("★ " if it["key"] else "") + it["item"],
                         "Built": bool(s.built) if s else False, "QC": bool(s.qc) if s else False,
                         "N/A": bool(s.na) if s else False})
        df = pd.DataFrame(rows)
        sec_app = int((~df["N/A"]).sum())
        sec_built = int((df.Built & ~df["N/A"]).sum())
        sec_qc = int((df.QC & ~df["N/A"]).sum())
        if show == "Not built yet":
            df = df[~df.Built & ~df["N/A"]]
        elif show == "Not QC'd yet":
            df = df[~df.QC & ~df["N/A"]]
        if df.empty:
            continue
        done = "✅ " if sec_app and sec_built == sec_app and sec_qc == sec_app else ""
        with st.expander(f"{done}{section}  —  Built {sec_built}/{sec_app} · QC {sec_qc}/{sec_app}"):
            if note:
                st.caption(note)
            ed = st.data_editor(
                df.reset_index(drop=True), key=f"chk_{build_id}_{si}_{show}", hide_index=True, width="stretch",
                num_rows="fixed", disabled=["Area", "Item"],
                column_config={"item_id": None,
                               "Area": st.column_config.TextColumn("Area", width=230),
                               "Item": st.column_config.TextColumn("Item", width=640),
                               "Built": st.column_config.CheckboxColumn("Built", width=70),
                               "QC": st.column_config.CheckboxColumn("QC", width=70),
                               "N/A": st.column_config.CheckboxColumn("N/A", width=70)})
            base = df.reset_index(drop=True)
            for i in range(len(ed)):
                new = (bool(ed.at[i, "Built"]), bool(ed.at[i, "QC"]), bool(ed.at[i, "N/A"]))
                old = (bool(base.at[i, "Built"]), bool(base.at[i, "QC"]), bool(base.at[i, "N/A"]))
                if new != old:
                    changes[base.at[i, "item_id"]] = new

    def save():
        save_build_checks([{"build_id": int(build_id), "item_id": k, "built": v[0], "qc": v[1], "na": v[2]}
                           for k, v in changes.items()])
        flash(f"Saved {len(changes)} checklist change(s) for {txt(b.company_name)}.")
        st.rerun()

    label = f"💾 Save checklist ({len(changes)})" if changes else "💾 Save checklist"
    with save_top:
        if st.button(label, type="primary", disabled=not changes, key="chk_save_top"):
            save()
    if st.button(label, type="primary", disabled=not changes, key="chk_save_bottom"):
        save()

    if applicable and built == applicable and status != "Completed":
        st.info("Every checklist item is built. Set **Status** to Completed in the 📋 Builds section.")
    if applicable and qc == applicable and qc_status != "Completed":
        st.info("Every checklist item is QC'd. Set **QC Status** to Completed in the 📋 Builds section.")


STATUS_COLORS = ["#A0A7B4", "#E8B33A", "#2E9E5B"]   # Pending, In Progress, Completed
DASHBOARD_START = date(2026, 11, 1)                  # the dashboard counts builds from this month on


def _render_builds_dashboard(builds):
    import altair as alt

    if builds.empty:
        st.info("No builds yet. Add builds in the 📋 Builds section and the totals will appear here.")
        return

    df = builds.copy()
    df["status"] = df.status.map(lambda v: _clean_text(v) or "Pending")
    df["qc_status"] = df.qc_status.map(lambda v: _clean_text(v) or "")
    df["built_by"] = df.built_by.map(_clean_text)
    df["qc_by"] = df.qc_by.map(_clean_text)
    df["month"] = df.due_date.map(lambda d: month_start(_to_date(d)) if _to_date(d) else None)
    # The dashboard starts in November 2026; builds due earlier are left out (undated builds stay in).
    df = df[df.month.map(lambda mo: mo is None or mo >= DASHBOARD_START)]

    # ── One month at a glance ──
    current = month_start(today_local())
    last = max([current, *[mo for mo in df.month if mo is not None]])
    month_opts, mo = [], DASHBOARD_START
    while mo <= last:
        month_opts.append(mo)
        mo = add_months(mo, 1)
    default = month_opts.index(current) if current in month_opts else 0
    head, pick = st.columns([3, 1], vertical_alignment="bottom")
    pick_month = pick.selectbox("Month", month_opts, index=default, key="dash_month",
                                format_func=lambda o: f"{o:%B %Y}")
    head.subheader(f"{pick_month:%B %Y} builds")
    g = df[df.month == pick_month]
    g_total, g_done = len(g), int((g.status == "Completed").sum())
    m = st.columns(5)
    m[0].metric("Builds due", g_total)
    m[1].metric("Completed", g_done)
    m[2].metric("In progress", int((g.status == "In Progress").sum()))
    m[3].metric("Pending", int((g.status == "Pending").sum()))
    m[4].metric("Awaiting QC", int(((g.status == "Completed") & (g.qc_status != "Completed")).sum()))
    if g_total:
        st.progress(g_done / g_total, text=f"{g_done / g_total:.0%} of {pick_month:%B} builds completed")
        by_person = (g.assign(person=g.built_by.fillna("Unassigned"))
                     .groupby(["person", "status"]).size().reset_index(name="Builds"))
        by_person["rank"] = by_person.status.map({s: i for i, s in enumerate(BUILD_STATUSES)})
        people = by_person.groupby("person")["Builds"].sum().sort_values(ascending=False).index.tolist()
        st.markdown("**Builds by team member**")
        chart = alt.Chart(by_person).mark_bar(cornerRadiusTopRight=3, cornerRadiusBottomRight=3).encode(
            y=alt.Y("person:N", title=None, sort=people),
            x=alt.X("sum(Builds):Q", title="Builds", axis=alt.Axis(tickMinStep=1, format="d")),
            color=alt.Color("status:N", title=None, scale=alt.Scale(domain=BUILD_STATUSES, range=STATUS_COLORS),
                            legend=alt.Legend(orient="bottom")),
            order=alt.Order("rank:Q"),
            tooltip=[alt.Tooltip("person:N", title="Built by"), alt.Tooltip("status:N", title="Status"), "Builds"])
        # Height covers the axis and legend too, so add room for them on top of ~40px per bar.
        st.altair_chart(chart.properties(height=40 * len(people) + 110), width="stretch")
    else:
        st.info(f"No builds are due in {pick_month:%B %Y} yet.")

    st.divider()
    total = len(df)
    completed = int((df.status == "Completed").sum())
    st.subheader(f"Totals since {DASHBOARD_START:%B %Y}")
    m = st.columns(6)
    m[0].metric("Total builds", total)
    m[1].metric("Completed", completed)
    m[2].metric("In progress", int((df.status == "In Progress").sum()))
    m[3].metric("Pending", int((df.status == "Pending").sum()))
    m[4].metric("QC completed", int((df.qc_status == "Completed").sum()))
    m[5].metric("Awaiting QC", int(((df.status == "Completed") & (df.qc_status != "Completed")).sum()))
    if total:
        st.progress(completed / total, text=f"{completed / total:.0%} of all builds completed")

    dated = df[df.month.map(lambda v: v is not None)]
    if dated.empty:
        st.info("Add due dates to builds to see the month-by-month charts.")
        return
    months = sorted(dated.month.unique())
    order = [f"{mo:%b %Y}" for mo in months]

    st.divider()
    left, right = st.columns(2)
    with left:
        st.markdown("**Builds per month** (by due date)")
        counts = dated.groupby(["month", "status"]).size().reset_index(name="Builds")
        counts["Month"] = counts.month.map(lambda mo: f"{mo:%b %Y}")
        counts["rank"] = counts.status.map({s: i for i, s in enumerate(BUILD_STATUSES)})
        bars = alt.Chart(counts).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3).encode(
            x=alt.X("Month:N", sort=order, title=None, axis=alt.Axis(labelAngle=0)),
            y=alt.Y("sum(Builds):Q", title="Builds", axis=alt.Axis(tickMinStep=1, format="d")),
            color=alt.Color("status:N", title=None, scale=alt.Scale(domain=BUILD_STATUSES, range=STATUS_COLORS),
                            legend=alt.Legend(orient="bottom")),
            order=alt.Order("rank:Q"),
            tooltip=["Month", alt.Tooltip("status:N", title="Status"), "Builds"])
        totals = counts.groupby("Month", as_index=False)["Builds"].sum()
        labels = alt.Chart(totals).mark_text(dy=-8, fontWeight="bold").encode(
            x=alt.X("Month:N", sort=order), y="Builds:Q", text="Builds:Q")
        st.altair_chart((bars + labels).properties(height=320), width="stretch")
    with right:
        st.markdown("**Total builds to date** (running total)")
        per_month = dated.groupby("month").size().reindex(months).fillna(0).cumsum()
        cum = pd.DataFrame({"Month": order, "Total builds": per_month.astype(int).values})
        base = alt.Chart(cum).encode(x=alt.X("Month:N", sort=order, title=None, axis=alt.Axis(labelAngle=0)),
                                     y=alt.Y("Total builds:Q", title="Builds", axis=alt.Axis(tickMinStep=1, format="d")))
        line = base.mark_line(point=alt.OverlayMarkDef(size=70), color="#2F5D9A", strokeWidth=3)
        text = base.mark_text(dy=-12, fontWeight="bold", color="#2F5D9A").encode(text="Total builds:Q")
        st.altair_chart((line + text).properties(height=320), width="stretch")

    st.divider()
    st.markdown("**Month by month**")
    rows = []
    for mo in months:
        g = dated[dated.month == mo]
        rows.append({"Month": f"{mo:%B %Y}", "Total": len(g),
                     "Completed": int((g.status == "Completed").sum()),
                     "In Progress": int((g.status == "In Progress").sum()),
                     "Pending": int((g.status == "Pending").sum()),
                     "QC Completed": int((g.qc_status == "Completed").sum())})
    st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
    undated = total - len(dated)
    if undated:
        st.caption(f"{undated} build(s) have no due date yet, so they're counted in the totals but not the months.")


@st.fragment(run_every="30s")
def _live_builds_dashboard():
    """Re-reads the database and redraws the dashboard every 30 seconds."""
    try:
        builds = load_builds()
    except Exception as e:
        st.error("Could not refresh the dashboard. It will try again in 30 seconds.")
        with st.expander("Error details"):
            st.code(str(e))
        return
    _render_builds_dashboard(builds)
    st.caption(f"🟢 Live: updates automatically every 30 seconds. Last updated {now_local():%I:%M:%S %p} IST.")


def _start():
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
  <div class="lp-title" role="heading" aria-level="2">Build it.<br>Check it.<br>Hand it over.</div>
  <p class="lp-sub">Track every company build and its QC checklist in one place,
  from the first Salesforce update to the last GuideCX task.</p>
</div>
""", unsafe_allow_html=True)
        b, _ = st.columns([1, 1.4])
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


SECTIONS = {"📊 Dashboard": "Live totals and charts",
            "📋 Builds": "Tracker for every company build",
            "✅ QC": "Implementation checklist per build"}


def render_builds_page():
    try:
        builds = load_builds()
        checks = load_build_checks()
    except Exception as e:
        st.error("Company Builds could not load its data. Check that the builds and build_checks "
                 "tables exist in Supabase.")
        with st.expander("Error details"):
            st.code(str(e))
        st.stop()

    if "_flash" in st.session_state:
        st.success(st.session_state.pop("_flash"))

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
        st.divider()
        if view == "📊 Dashboard":
            st.caption("🟢 The dashboard refreshes on its own every 30 seconds.")
        else:
            if st.button("🔄 Refresh data", key="builds_refresh", width="stretch",
                         help="Load the latest changes saved by others. Save your own edits first, "
                              "because refreshing clears anything not yet saved."):
                st.rerun()
            st.caption(f"Data as of {now_local():%I:%M %p} IST. Save your edits before refreshing.")
        st.button("🏠 Home", key="go_home", width="stretch", on_click=_go_home)

    st.title(view)
    st.caption(SECTIONS[view])
    if view == "📊 Dashboard":
        _live_builds_dashboard()
    elif view == "📋 Builds":
        _render_build_tracker(builds, checks)
    else:
        _render_build_checklist(builds, checks)


# ─────────────────────────────────────────────────────────────
# App
# ─────────────────────────────────────────────────────────────
if st.session_state.get("started"):
    render_builds_page()
else:
    render_landing_page()
