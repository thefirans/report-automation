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
# Helpers
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


REQUIRED_COLS = [
    "Invoice ID",
    "Payment Date",
    "Payment Method",
    "Payment Status",
    "Payment Amount",
    "Outstanding Balance",
    "Client",
    "Email",
    "Phone Number",
    "Total",
    "Created By",
]

FOLDER_ID = "1MHx8XyOxlj9UMbz9V2ZqxCi_UleizEwC"
SHARE_EMAIL = "yuskov.y@workflow.com.ua"


def process_report(csv_file, province: str):
    """Main pipeline: read CSV → clean → upload to Google Sheets → return URL."""

    status = st.status("Running report…", expanded=True)
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
    status.write(f"✅ Filtered to {len(df)} rows with Payment Amount > 0.")
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
    yesterday = (datetime.datetime.now() - datetime.timedelta(days=1)).strftime("%d.%m")
    sheet_name = f"{province} {yesterday} workflow crm automated"

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

    # Color rows
    header_offset = 2  # row 1 = header, data starts at row 2
    start_yellow = len(regular_no_due) + len(regular_due) + header_offset
    end_yellow = start_yellow + len(yellow_rows)
    start_orange = end_yellow
    end_orange = start_orange + len(orange_rows)

    formats = []
    if yellow_rows:
        formats.append({
            "range": f"A{start_yellow}:K{end_yellow - 1}",
            "format": {"backgroundColor": {"red": 1, "green": 1, "blue": 0}},
        })
    if orange_rows:
        formats.append({
            "range": f"A{start_orange}:K{end_orange - 1}",
            "format": {"backgroundColor": {"red": 1, "green": 0.6, "blue": 0}},
        })
    if formats:
        worksheet.batch_format(formats)

    # Center all + bold header
    worksheet.format("A1:K", {"horizontalAlignment": "CENTER"})
    worksheet.format(
        "A1:K1",
        {
            "backgroundColor": {"red": 0.78, "green": 0.87, "blue": 1},
            "textFormat": {"bold": True},
        },
    )

    # Auto-resize columns
    sh.batch_update({
        "requests": [{
            "autoResizeDimensions": {
                "dimensions": {
                    "sheetId": worksheet._properties["sheetId"],
                    "dimension": "COLUMNS",
                    "startIndex": 0,
                    "endIndex": 11,
                }
            }
        }]
    })
    progress.progress(90)

    # ── 10. Share ─────────────────────────────────
    status.write(f"🔗 Sharing with {SHARE_EMAIL}…")
    sh.share(SHARE_EMAIL, perm_type="user", role="writer")
    progress.progress(100)

    url = f"https://docs.google.com/spreadsheets/d/{sh.id}"
    status.update(label="✅ Report complete!", state="complete", expanded=False)

    return url


# ──────────────────────────────────────────────
# UI
# ──────────────────────────────────────────────
province = st.radio("Choose Province:", ["Ontario", "Alberta"], horizontal=True)
csv_file = st.file_uploader("Upload CSV file", type=["csv"])

if st.button("🚀 RUN", type="primary", use_container_width=True):
    if csv_file is None:
        st.warning("Please upload a CSV file first.")
    else:
        url = process_report(csv_file, province)
        if url:
            st.success("Report generated successfully!")
            st.markdown(f"### [📄 Open Google Sheet]({url})")
