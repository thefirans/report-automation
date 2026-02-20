import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import datetime
import json

# ──────────────────────────────────────────────
# Page config
# ──────────────────────────────────────────────
st.set_page_config(page_title="Report Automation", page_icon="📊", layout="centered")
st.title("📊 Report Automation")

# ──────────────────────────────────────────────
# Google Sheets auth (cached so it doesn't re-auth every interaction)
# ──────────────────────────────────────────────
@st.cache_resource
def get_gspread_client():
    """Authenticate with Google using secrets stored in Streamlit Cloud."""
    scope = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive",
    ]
    service_account_info = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT_JSON"])
    creds = Credentials.from_service_account_info(service_account_info, scopes=scope)
    return gspread.authorize(creds)


# ──────────────────────────────────────────────
# Shared helpers
# ──────────────────────────────────────────────
def money_to_float(x) -> float:
    """Convert '$170.63' / '170.63' / NaN → float. Returns 0.0 on failure."""
    if pd.isna(x):
        return 0.0
    s = str(x).strip().replace("$", "").replace(",", "")
    if not s:
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def get_yesterday_str() -> str:
    """Returns yesterday's date as DD.MM"""
    return (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%d.%m")


def apply_standard_formatting(sh, worksheet, num_cols):
    """Apply center alignment, bold blue header, and auto-resize columns."""
    last_col = chr(ord("A") + num_cols - 1)

    worksheet.format(f"A1:{last_col}", {"horizontalAlignment": "CENTER"})
    worksheet.format(f"A1:{last_col}1", {
        "backgroundColor": {"red": 0.78, "green": 0.87, "blue": 1},
        "textFormat": {"bold": True},
    })
    sh.batch_update({"requests": [{
        "autoResizeDimensions": {
            "dimensions": {
                "sheetId": worksheet._properties["sheetId"],
                "dimension": "COLUMNS",
                "startIndex": 0,
                "endIndex": num_cols,
            }
        }
    }]})


# ──────────────────────────────────────────────
# Shared constants
# ──────────────────────────────────────────────
FOLDER_ID = "1MHx8XyOxlj9UMbz9V2ZqxCi_UleizEwC"
SHARE_EMAIL = "yuskov.y@workflow.com.ua"


# ══════════════════════════════════════════════
# TAB 1 — Workflow Pros CRM
# ══════════════════════════════════════════════
def run_workflow_pros_crm(csv_file):
    """
    Original Ontario/Alberta pipeline — same logic as before.
    Reads Workflow Pros CRM CSV, filters, matches invoices against
    🟢GOOD REVIEWS sheet, colors yellow/orange, uploads to Google Sheets.
    """

    REQUIRED_COLS = [
        "Invoice ID", "Payment Date", "Payment Method", "Payment Status",
        "Payment Amount", "Outstanding Balance", "Client", "Email",
        "Phone Number", "Total", "Created By",
    ]

    status = st.status("Running Workflow Pros CRM report…", expanded=True)
    progress = st.progress(0)

    # ── 1. Read CSV ──────────────────────────────
    status.write("📂 Reading CSV…")
    df = pd.read_csv(csv_file)
    progress.progress(10)

    # ── 2. Validate columns ──────────────────────
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"❌ Missing columns in CSV: {missing}")
        return None

    df = df[REQUIRED_COLS].copy()
    status.write("✅ Columns validated.")
    progress.progress(20)

    # ── 3. Filter: Payment Amount > 0 ────────────
    df = df[df["Payment Amount"].apply(money_to_float) > 0].copy()
    df["Invoice ID"] = df["Invoice ID"].astype(str).str.strip()

    # Remove duplicate Invoice IDs (keep first occurrence only)
    before_dedup = len(df)
    df = df.drop_duplicates(subset=["Invoice ID"], keep="first")
    dupes_removed = before_dedup - len(df)
    status.write(f"✅ Filtered to {len(df)} rows (Payment Amount > 0, {dupes_removed} duplicate invoices removed).")
    progress.progress(30)

    # ── 4. Auth Google Sheets ────────────────────
    status.write("🔑 Connecting to Google Sheets…")
    try:
        client = get_gspread_client()
    except Exception as e:
        st.error(f"❌ Google auth failed: {e}")
        return None
    progress.progress(40)

    # ── 5. Fetch review invoices ─────────────────
    status.write("📋 Loading review invoices…")
    try:
        reviews_sheet = client.open("🟢GOOD REVIEWS")
        alex_invoices = {
            str(x).strip()
            for x in reviews_sheet.worksheet("Oleksandr Leoshko").col_values(3)
            if str(x).strip()
        }
        eugene_invoices = {
            str(x).strip()
            for x in reviews_sheet.worksheet("Eugene Yuskov").col_values(3)
            if str(x).strip()
        }
    except Exception as e:
        st.error(f"❌ Could not read GOOD REVIEWS sheet: {e}")
        return None
    progress.progress(50)

    # ── 6. Categorize rows ───────────────────────
    status.write("🔀 Categorizing rows…")
    regular_no_due, regular_due, yellow_rows, orange_rows = [], [], [], []

    for _, row in df.iterrows():
        inv = str(row["Invoice ID"]).strip()
        outstanding = money_to_float(row["Outstanding Balance"])

        if inv in alex_invoices:
            yellow_rows.append(row)
        elif inv in eugene_invoices:
            orange_rows.append(row)
        elif outstanding > 0:
            regular_due.append(row)
        else:
            regular_no_due.append(row)

    ordered_df = pd.DataFrame(regular_no_due + regular_due + yellow_rows + orange_rows)
    if not ordered_df.empty:
        ordered_df = ordered_df[df.columns]
    else:
        ordered_df = pd.DataFrame(columns=df.columns)
    ordered_df = ordered_df.fillna("")
    progress.progress(60)

    # ── 7. Create Google Sheet ───────────────────
    yesterday = get_yesterday_str()
    sheet_name = f"Workflow CRM {yesterday} automated"

    status.write(f"📝 Creating sheet: {sheet_name}")
    sh = client.create(sheet_name, folder_id=FOLDER_ID)
    worksheet = sh.get_worksheet(0)
    progress.progress(70)

    # ── 8. Upload data ───────────────────────────
    status.write("⬆️ Uploading data…")
    worksheet.update([ordered_df.columns.tolist()] + ordered_df.values.tolist())
    progress.progress(80)

    # ── 9. Formatting ────────────────────────────
    status.write("🎨 Applying formatting…")
    num_cols = len(REQUIRED_COLS)
    last_col = chr(ord("A") + num_cols - 1)  # K

    start_yellow = len(regular_no_due) + len(regular_due) + 2
    end_yellow = start_yellow + len(yellow_rows)
    start_orange = end_yellow
    end_orange = start_orange + len(orange_rows)

    formats = []
    if yellow_rows:
        formats.append({
            "range": f"A{start_yellow}:{last_col}{end_yellow - 1}",
            "format": {"backgroundColor": {"red": 1, "green": 1, "blue": 0}},
        })
    if orange_rows:
        formats.append({
            "range": f"A{start_orange}:{last_col}{end_orange - 1}",
            "format": {"backgroundColor": {"red": 1, "green": 0.6, "blue": 0}},
        })
    if formats:
        worksheet.batch_format(formats)

    apply_standard_formatting(sh, worksheet, num_cols)
    progress.progress(90)

    # ── 10. Share ─────────────────────────────────
    status.write(f"🔗 Sharing with {SHARE_EMAIL}…")
    sh.share(SHARE_EMAIL, perm_type="user", role="writer")
    progress.progress(100)

    url = f"https://docs.google.com/spreadsheets/d/{sh.id}"
    status.update(label="✅ Report complete!", state="complete", expanded=False)
    return url


# ══════════════════════════════════════════════
# TAB 2 — USA (Housecall)
# ══════════════════════════════════════════════
def run_usa_housecall(csv_file):
    """
    USA Housecall Pro CRM pipeline.
    Different CSV columns than Workflow Pros. Cleans employee names,
    matches by Invoice Number ONLY (not client name) against 🟢GOOD REVIEWS,
    colors yellow/orange, uploads to Google Sheets.
    """

    # Columns expected in the Housecall CSV
    NEEDED_COLS = [
        "Customer Name", "Invoice Number", "Invoice Status",
        "Assigned Employee Name", "Job Status", "Payment Amount",
    ]

    # Columns that go into the final Google Sheet
    OUTPUT_COLS = [
        "Customer Name", "Invoice Number", "Invoice Status",
        "Clean Employee Name", "Job Status",
    ]

    status = st.status("Running USA (Housecall) report…", expanded=True)
    progress = st.progress(0)

    # ── 1. Read CSV ──────────────────────────────
    status.write("📂 Reading CSV…")
    df = pd.read_csv(csv_file)
    progress.progress(10)

    # ── 2. Validate columns ──────────────────────
    missing = [c for c in NEEDED_COLS if c not in df.columns]
    if missing:
        st.error(f"❌ Missing columns in CSV: {missing}")
        return None
    status.write("✅ Columns validated.")
    progress.progress(20)

    # ── 3. Filter & clean ────────────────────────
    df_clean = df[df["Payment Amount"].apply(money_to_float) > 0].copy()

    # Clean employee name — remove brackets and quotes like ["John Smith"]
    df_clean.insert(
        loc=df_clean.columns.get_loc("Assigned Employee Name"),
        column="Clean Employee Name",
        value=df_clean["Assigned Employee Name"]
            .astype(str)
            .str.replace(r'[\[\]\"]', "", regex=True)
            .str.strip(),
    )

    final_df = df_clean[OUTPUT_COLS].copy()
    final_df["Invoice Number"] = final_df["Invoice Number"].astype(str).str.strip()

    # Remove duplicate Invoice Numbers (keep first occurrence only)
    before_dedup = len(final_df)
    final_df = final_df.drop_duplicates(subset=["Invoice Number"], keep="first")
    dupes_removed = before_dedup - len(final_df)
    status.write(f"✅ Filtered to {len(final_df)} rows (Payment Amount > 0, {dupes_removed} duplicate invoices removed).")
    progress.progress(30)

    # ── 4. Auth Google Sheets ────────────────────
    status.write("🔑 Connecting to Google Sheets…")
    try:
        client = get_gspread_client()
    except Exception as e:
        st.error(f"❌ Google auth failed: {e}")
        return None
    progress.progress(40)

    # ── 5. Fetch review invoices (Invoice Number ONLY — no name matching) ──
    status.write("📋 Loading review invoices…")
    try:
        reviews_sheet = client.open("🟢GOOD REVIEWS")
        alex_invoices = {
            str(x).strip()
            for x in reviews_sheet.worksheet("Oleksandr Leoshko").col_values(3)
            if str(x).strip()
        }
        eugene_invoices = {
            str(x).strip()
            for x in reviews_sheet.worksheet("Eugene Yuskov").col_values(3)
            if str(x).strip()
        }
    except Exception as e:
        st.error(f"❌ Could not read GOOD REVIEWS sheet: {e}")
        return None
    progress.progress(50)

    # ── 6. Categorize rows (by Invoice Number ONLY) ──
    status.write("🔀 Categorizing rows…")
    regular_rows, yellow_rows, orange_rows = [], [], []

    for _, row in final_df.iterrows():
        inv_num = str(row["Invoice Number"]).strip()

        if inv_num in alex_invoices:
            yellow_rows.append(row)
        elif inv_num in eugene_invoices:
            orange_rows.append(row)
        else:
            regular_rows.append(row)

    ordered_df = pd.DataFrame(regular_rows + yellow_rows + orange_rows)
    if not ordered_df.empty:
        ordered_df = ordered_df[OUTPUT_COLS]
    else:
        ordered_df = pd.DataFrame(columns=OUTPUT_COLS)
    ordered_df = ordered_df.fillna("")
    progress.progress(60)

    # ── 7. Create Google Sheet ───────────────────
    yesterday = get_yesterday_str()
    sheet_name = f"USA Housecall {yesterday} automated"

    status.write(f"📝 Creating sheet: {sheet_name}")
    sh = client.create(sheet_name, folder_id=FOLDER_ID)
    worksheet = sh.get_worksheet(0)
    progress.progress(70)

    # ── 8. Upload data ───────────────────────────
    status.write("⬆️ Uploading data…")
    worksheet.update([ordered_df.columns.tolist()] + ordered_df.values.tolist())
    progress.progress(80)

    # ── 9. Formatting ────────────────────────────
    status.write("🎨 Applying formatting…")
    num_cols = len(OUTPUT_COLS)
    last_col = chr(ord("A") + num_cols - 1)  # E

    start_yellow = len(regular_rows) + 2
    end_yellow = start_yellow + len(yellow_rows)
    start_orange = end_yellow
    end_orange = start_orange + len(orange_rows)

    formats = []
    if yellow_rows:
        formats.append({
            "range": f"A{start_yellow}:{last_col}{end_yellow - 1}",
            "format": {"backgroundColor": {"red": 1, "green": 1, "blue": 0}},
        })
    if orange_rows:
        formats.append({
            "range": f"A{start_orange}:{last_col}{end_orange - 1}",
            "format": {"backgroundColor": {"red": 1, "green": 0.6, "blue": 0}},
        })
    if formats:
        worksheet.batch_format(formats)

    apply_standard_formatting(sh, worksheet, num_cols)
    progress.progress(90)

    # ── 10. Share ─────────────────────────────────
    status.write(f"🔗 Sharing with {SHARE_EMAIL}…")
    sh.share(SHARE_EMAIL, perm_type="user", role="writer")
    progress.progress(100)

    url = f"https://docs.google.com/spreadsheets/d/{sh.id}"
    status.update(label="✅ Report complete!", state="complete", expanded=False)
    return url


# ══════════════════════════════════════════════
# TAB 3 — Plumbing
# ══════════════════════════════════════════════

# ⚠️ TODO: Replace with the actual Google Sheet name (or open by URL/ID instead)
PLUMBING_REVIEWS_SHEET_ID = "1HImgvjKQHGYMARHIpMJOf0961Urpbkq5zGrVpAkQAOU"

# ⚠️ TODO: Replace with real corporate emails to share the plumbing report with
PLUMBING_SHARE_EMAILS = [
    "yuskov.y@workflow.com.ua",
    #"alina.tryncha@workflow.com.ua",
    "oleksandr.leoshko@workflow.com.ua",
]

# Review tabs inside the plumbing reviews sheet + which column has Invoice Number
# Format: (tab_name, column_index)  — column 3 = column C
PLUMBING_REVIEW_TABS = [
    ("Alina", 1),       # column C
    ("Alex", 3),        # column C
    ("Eugene", 2) # column C
]


def run_plumbing(xlsx_file):
    """
    Plumbing pipeline:
    1. Read XLSX with plumbing-specific columns
    2. Upload to new Google Sheet named "Plumbing {yesterday} automated"
    3. Check 3 review tabs (Alina, Alex, Eugene) for duplicate Invoice Numbers
    4. Color duplicate rows dark red with white text
    5. Share with specified users and return URL
    """

    REQUIRED_COLS = [
        "Payment Type", "Amount", "Invoice Number", "Invoice Total",
        "Invoice Balance", "Payment Method", "Paid On",
        "Completion Date", "Assigned Technicians",
    ]

    # Columns to keep in the final output (drop Invoice Balance & Payment Method)
    OUTPUT_COLS = [
        "Payment Type", "Amount", "Invoice Number", "Invoice Total",
        "Paid On", "Completion Date", "Assigned Technicians",
    ]

    status = st.status("Running Plumbing report…", expanded=True)
    progress = st.progress(0)

    # ── 1. Read XLSX ─────────────────────────────
    status.write("📂 Reading XLSX…")
    df = pd.read_excel(xlsx_file)
    progress.progress(10)

    # ── 2. Validate columns ──────────────────────
    missing = [c for c in REQUIRED_COLS if c not in df.columns]
    if missing:
        st.error(f"❌ Missing columns in XLSX: {missing}")
        return None

    df = df[OUTPUT_COLS].copy()

    # Fix Invoice Number: remove trailing .0 (Excel reads integers as floats)
    df["Invoice Number"] = (
        df["Invoice Number"]
        .astype(str)
        .str.strip()
        .str.replace(r"\.0$", "", regex=True)
    )
    df = df.fillna("")

    # Remove duplicate Invoice Numbers (keep first occurrence only)
    before_dedup = len(df)
    df = df.drop_duplicates(subset=["Invoice Number"], keep="first")
    dupes_removed = before_dedup - len(df)

    # Convert ALL values to plain strings — XLSX often contains datetime/numpy
    # types that are not JSON-serializable and crash the Google Sheets API.
    for col in df.columns:
        df[col] = df[col].astype(str).replace("NaT", "").replace("nan", "")

    status.write(f"✅ Loaded {len(df)} rows ({dupes_removed} duplicate invoices removed).")
    progress.progress(20)

    # ── 3. Auth Google Sheets ────────────────────
    status.write("🔑 Connecting to Google Sheets…")
    try:
        client = get_gspread_client()
    except Exception as e:
        st.error(f"❌ Google auth failed: {e}")
        return None
    progress.progress(30)

    # ── 4. Create Google Sheet ───────────────────
    yesterday = get_yesterday_str()
    sheet_name = f"Plumbing {yesterday} automated"

    status.write(f"📝 Creating sheet: {sheet_name}")
    sh = client.create(sheet_name, folder_id=FOLDER_ID)
    worksheet = sh.get_worksheet(0)
    progress.progress(40)

    # ── 5. Upload data ───────────────────────────
    status.write("⬆️ Uploading data…")
    worksheet.update([df.columns.tolist()] + df.values.tolist())
    progress.progress(50)

    # ── 6. Fetch review invoices from all tabs ───
    status.write("📋 Loading plumbing review invoices…")
    try:
        reviews_sheet = client.open_by_key(PLUMBING_REVIEWS_SHEET_ID)

        # Map invoice number → tab name (first match wins)
        invoice_to_tab = {}
        for tab_name, col_index in PLUMBING_REVIEW_TABS:
            try:
                ws = reviews_sheet.worksheet(tab_name)
                tab_invoices = [
                    str(x).strip()
                    for x in ws.col_values(col_index)
                    if str(x).strip()
                ]
                for inv in tab_invoices:
                    if inv not in invoice_to_tab:
                        invoice_to_tab[inv] = tab_name
                status.write(f"   ✅ {tab_name}: {len(tab_invoices)} invoices loaded")
            except gspread.exceptions.WorksheetNotFound:
                status.write(f"   ⚠️ Tab '{tab_name}' not found — skipping")

    except Exception as e:
        st.error(f"❌ Could not read plumbing reviews sheet: {e}")
        return None

    progress.progress(55)

    # ── 7. Separate clean rows from duplicates, put dupes at the end ──
    status.write("🔀 Reordering rows (duplicates → bottom)…")
    clean_rows = []
    duplicate_rows = []

    for _, row in df.iterrows():
        inv_num = str(row["Invoice Number"]).strip()
        if inv_num and inv_num in invoice_to_tab:
            row_copy = row.copy()
            row_copy["Found On"] = invoice_to_tab[inv_num]
            duplicate_rows.append(row_copy)
        else:
            row_copy = row.copy()
            row_copy["Found On"] = ""
            clean_rows.append(row_copy)

    # Build final DataFrame with the extra "Found On" column
    FINAL_COLS = OUTPUT_COLS + ["Found On"]
    ordered_df = pd.DataFrame(clean_rows + duplicate_rows)
    if not ordered_df.empty:
        ordered_df = ordered_df[FINAL_COLS]
    else:
        ordered_df = pd.DataFrame(columns=FINAL_COLS)
    ordered_df = ordered_df.fillna("")

    # Convert ALL values to plain strings for Google Sheets API
    for col in ordered_df.columns:
        ordered_df[col] = ordered_df[col].astype(str).replace("NaT", "").replace("nan", "")

    status.write(f"   {len(clean_rows)} clean rows + {len(duplicate_rows)} duplicate(s).")
    progress.progress(65)

    # ── 8. Upload reordered data ─────────────────
    status.write("⬆️ Uploading data…")
    worksheet.update([ordered_df.columns.tolist()] + ordered_df.values.tolist())
    progress.progress(75)

    # ── 9. Color duplicate rows (now at the bottom) dark red ──
    status.write("🎨 Applying formatting…")
    num_cols = len(FINAL_COLS)  # 8 columns (G + Found On = H)
    last_col = chr(ord("A") + num_cols - 1)  # H

    if duplicate_rows:
        start_dupes = len(clean_rows) + 2  # +2: row 1 = header, data from row 2
        end_dupes = start_dupes + len(duplicate_rows) - 1

        worksheet.batch_format([{
            "range": f"A{start_dupes}:{last_col}{end_dupes}",
            "format": {
                "backgroundColor": {"red": 0.6, "green": 0, "blue": 0},
                "textFormat": {
                    "foregroundColor": {"red": 1, "green": 1, "blue": 1},
                },
            },
        }])

    apply_standard_formatting(sh, worksheet, num_cols)
    progress.progress(90)

    # ── 9. Share ─────────────────────────────────
    for email in PLUMBING_SHARE_EMAILS:
        status.write(f"🔗 Sharing with {email}…")
        sh.share(email, perm_type="user", role="writer")
    progress.progress(100)

    url = f"https://docs.google.com/spreadsheets/d/{sh.id}"
    status.update(label="✅ Plumbing report complete!", state="complete", expanded=False)
    return url


# ══════════════════════════════════════════════
# MAIN UI — Three Tabs
# ══════════════════════════════════════════════
tab1, tab2, tab3 = st.tabs([
    "🔧 Workflow Pros CRM",
    "🇺🇸 USA (Housecall)",
    "🚿 Plumbing",
])

# ── Tab 1: Workflow Pros CRM ─────────────────
with tab1:
    st.subheader("Workflow Pros CRM")
    st.caption("Ontario & Alberta — upload CSV from Workflow Pros CRM")
    csv_file_1 = st.file_uploader("Upload CSV file", type=["csv"], key="workflow_csv")

    if st.button("🚀 RUN", type="primary", use_container_width=True, key="btn_workflow"):
        if csv_file_1 is None:
            st.warning("Please upload a CSV file first.")
        else:
            url = run_workflow_pros_crm(csv_file_1)
            if url:
                st.success("Report generated successfully!")
                st.markdown(f"### [📄 Open Google Sheet]({url})")

# ── Tab 2: USA (Housecall) ───────────────────
with tab2:
    st.subheader("USA (Housecall)")
    st.caption("Housecall Pro CRM — upload CSV export")
    csv_file_2 = st.file_uploader("Upload CSV file", type=["csv"], key="usa_csv")

    if st.button("🚀 RUN", type="primary", use_container_width=True, key="btn_usa"):
        if csv_file_2 is None:
            st.warning("Please upload a CSV file first.")
        else:
            url = run_usa_housecall(csv_file_2)
            if url:
                st.success("Report generated successfully!")
                st.markdown(f"### [📄 Open Google Sheet]({url})")

# ── Tab 3: Plumbing ──────────────────────────
with tab3:
    st.subheader("Plumbing")
    st.caption("Upload XLSX export — duplicate invoices will be highlighted in dark red")
    xlsx_file = st.file_uploader("Upload XLSX file", type=["xlsx"], key="plumbing_xlsx")

    if st.button("🚀 RUN", type="primary", use_container_width=True, key="btn_plumbing"):
        if xlsx_file is None:
            st.warning("Please upload an XLSX file first.")
        else:
            url = run_plumbing(xlsx_file)
            if url:
                st.success("Report generated successfully!")
                st.markdown(f"### [📄 Open Google Sheet]({url})")

# ── Footer ────────────────────────────────────
st.divider()
st.caption("Report Automation v2.0")