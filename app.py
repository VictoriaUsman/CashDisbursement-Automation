import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pandas as pd
import io
import os
import json
import warnings
import anthropic
from dotenv import load_dotenv
load_dotenv()

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(__file__), "service_account.json")
CONFIG_FILE          = os.path.join(os.path.dirname(__file__), "config.json")

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

AMOUNT_COLS = [
    "amount paid", "vatable amount", "vat amount", "non vat amount",
    "other  charges amount", "ewt amount (negative entry)", "tb amount",
]

# Hardcoded account classifications (account title → COS / OPEX / SALES)
ACCOUNT_TAGS: dict[str, str] = {
    # ── ECON ────────────────────────────────────────────────────────────────
    "13TH MONTH PAY":                                    "COS",
    "SALARIES AND WAGES EXPENSE-CHINESE":                "COS",
    "SALARIES AND WAGES EXPENSE-LOCAL":                  "COS",
    "LABOR FEE":                                         "COS",
    "SUBCONTRACTOR COST":                                "COS",
    "HDMF, SSS, PHIC EXPENSE":                          "COS",
    "BI WIRE (SUPERSONIC MFG INC)":                      "COS",
    "CEMENT- HOLCIM":                                    "COS",
    "CEMENT-EXCEM":                                      "COS",
    "CEMENT-ONE SAMEX":                                  "COS",
    "COAL":                                              "COS",
    "GRAVEL":                                            "COS",
    "SAND":                                              "COS",
    "PURCHASE DOLLAR":                                   "COS",
    "PURCHASE PESO":                                     "COS",
    "FREIGHT SERVICES/TRUCKING":                         "COS",
    "FREIGHT SERVICES/TRUCKING-KAIKES":                  "COS",
    "GASOLINE , OIL, LUBRICANTS-PRODUCTION":             "COS",
    "GASOLINE , OIL, LUBRICANTS-VEHICLE":                "COS",
    "IMPORT CHARGES/CONTAINER":                          "COS",
    "INSURANCE EXPENSE":                                 "COS",
    "RENTAL STOCKYARD":                                  "COS",
    "SECURITY SERVICES":                                 "COS",
    "TRAVEL, TRANSPORTATION AND ACCOMODATION EXPENSE":   "COS",
    "TOLLGATE/TERMINAL FEE":                             "COS",
    "MATERIALS AND SUPPLIES-OFFICE":                     "COS",
    "ADVERTISING EXPENSE":                               "OPEX",
    "BANK CHARGE/SERVICE CHARGE FEE":                    "OPEX",
    "COMMISSION":                                        "OPEX",
    "COMMUNICATION EXPENSE":                             "OPEX",
    "COMPANY EVENTS":                                    "OPEX",
    "COURIER EXPENSE":                                   "OPEX",
    "DEPRECIATION/AMORTIZATION":                         "OPEX",
    "EMPLOYEES BENEFITS":                                "OPEX",
    "MEALS AND PANTRY":                                  "OPEX",
    "MEDICAL EXPENSE":                                   "OPEX",
    "MISCELLANEOUS EXPENSE":                             "OPEX",
    "NOTARY FEE":                                        "OPEX",
    "OFFICE SUPPLIES AND PRINTING EXPENSE":              "OPEX",
    "OTHER OFFICE EXPENSES":                             "OPEX",
    "PARKING FEE":                                       "OPEX",
    "PERMITS, TAXES AND LICENSES":                       "OPEX",
    "POWER, LIGHT, WATER":                               "OPEX",
    "PROFESSIONAL FEES- FRANCIS":                        "OPEX",
    "PROFESSIONAL FEES-DEOGRACIAS":                      "OPEX",
    "PROFESSIONAL FEES-OTHERS":                          "OPEX",
    "PROFESSIONAL FEES-SUNWU":                           "OPEX",
    "PROPERTY AND EQUIPMENT EXPENSE":                    "OPEX",
    "RENT EXPENSE":                                      "OPEX",
    "REPAIRS AND MAINTENANCE":                           "OPEX",
    "REPRESENTATION AND ENTERTAINMENT":                  "OPEX",
    "SUBSCRIPTION":                                      "OPEX",
    "VISA 9G":                                           "OPEX",
    # ── ECOTEX ──────────────────────────────────────────────────────────────
    "MATERIALS, TOOLS AND SUPPLIES-OPERATIONS":          "COS",
    'PURCHASES-JA & A ENT ALLAN LAWRENCE':               "COS",
    "PURCHASES-FRONTIERRY CONST TRD.":                   "COS",
    "PURCHASE (IMPORT)-CONTAINER":                       "COS",
    "PURCHASES-ONE SAMEX CEMENT":                        "COS",
    "PURCHASES-HOLCIM":                                  "COS",
    "PURCHASES-CYCLETREND WHITE PAPER":                  "COS",
    "PURCHASES-C ONE C TWO WHITE PAPER":                 "COS",
    "PURCHASES-Coal":                                    "COS",
    "PURCHASES-Sand":                                    "COS",
    "PURCHASES-LOCAL":                                   "COS",
    "FREIGHT SERVICES- E TRANS":                         "COS",
    "FREIGHT SERVICES- TRUCKING":                        "COS",
    "GASOLINE , OIL, LUBRICANTS":                        "COS",
    "UTILITIES":                                         "OPEX",
    "PROFESSIONAL FEES":                                 "OPEX",
    "SALARIES AND WAGES EXPENSE":                        "OPEX",
    "TAXES AND LICENSES":                                "OPEX",
    "TRAVEL, TRANSPORTATION AND ACCOMODAITION EXPENSE":  "OPEX",
}

DEFAULT_SOURCES = [
    {
        "label": "DAVAO",
        "sheet_id": "1uGiFSN9L3bUl__o5fXQK9SNgv7l4fL7g",
        "tab": "DAVAO Cash Disbursement",
        "output_sheet_id": "178dBRrrsME9IB_kDtlG7QasGsz-A9x2H2hsygr0ZXZ8",
        "output_tab": "DAVAO",
    }
]

# ── Config persistence ───────────────────────────────────────────────────────
def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE) as f:
                data = json.load(f)
                return data.get("sources", DEFAULT_SOURCES), data.get("pl_classifications", {})
        except Exception:
            pass
    return DEFAULT_SOURCES, {}

def save_config(sources, pl_classifications=None):
    existing = {}
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE) as f:
                existing = json.load(f)
        except Exception:
            pass
    existing["sources"] = sources
    if pl_classifications is not None:
        existing["pl_classifications"] = pl_classifications
    with open(CONFIG_FILE, "w") as f:
        json.dump(existing, f, indent=2)

def build_gl_excel(df, acct_col, tb_col, display_cols):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = "General Ledger"

    navy_fill   = PatternFill("solid", fgColor="1E3A5F")
    navy_font   = Font(color="FFFFFF", bold=True)
    acct_fill   = PatternFill("solid", fgColor="E8EEF7")
    acct_font   = Font(bold=True)
    sub_fill    = PatternFill("solid", fgColor="D0E8FF")
    sub_font    = Font(bold=True)
    right_align = Alignment(horizontal="right")

    all_cols = [acct_col] + display_cols

    # Column headers
    for ci, h in enumerate(all_cols, 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = navy_font
        cell.fill = navy_fill
        cell.alignment = Alignment(horizontal="center")

    r = 2
    for acct_title, group in df.groupby(acct_col, sort=True):
        # Account title header row
        cell = ws.cell(row=r, column=1, value=acct_title)
        cell.font = acct_font
        cell.fill = acct_fill
        for ci in range(2, len(all_cols) + 1):
            ws.cell(row=r, column=ci).fill = acct_fill
        r += 1

        # Detail rows
        for _, data_row in group[display_cols].iterrows():
            ws.cell(row=r, column=1, value="")
            for ci, col in enumerate(display_cols, 2):
                val = data_row[col]
                try:
                    val = float(str(val).replace(",", "").strip())
                except (ValueError, AttributeError):
                    pass
                c = ws.cell(row=r, column=ci, value=val)
                if isinstance(val, float):
                    c.alignment = right_align
                    c.number_format = '#,##0.00'
            r += 1

        # Subtotal row
        subtotal = group[tb_col].sum()
        lbl = ws.cell(row=r, column=1, value="Subtotal")
        lbl.font = sub_font
        lbl.fill = sub_fill
        tb_ci = all_cols.index(tb_col) + 1
        for ci in range(1, len(all_cols) + 1):
            c = ws.cell(row=r, column=ci)
            c.fill = sub_fill
            if ci == tb_ci:
                c.value = subtotal
                c.font = sub_font
                c.alignment = right_align
                c.number_format = '#,##0.00'
        r += 2  # blank row between accounts

    for ci in range(1, len(all_cols) + 1):
        ws.column_dimensions[get_column_letter(ci)].width = 22

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def df_to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()


def generate_pl_summary(company_name: str, active_months: list, pl_rows: list) -> str:
    lines = [f"{company_name} — Income Statement"]
    for row in pl_rows:
        month_vals = {m: row.get(m, "-") for m in active_months}
        ye = row.get("As at Year-end", "-")
        lines.append(f"{row['label']}: {', '.join(f'{m[:3]}={v}' for m, v in month_vals.items())}  |  Year-end: {ye}")
    text = "\n".join(lines)

    client = anthropic.Anthropic(api_key=os.environ.get("ANTHROPIC_API_KEY"))
    message = client.messages.create(
        model="claude-opus-4-7",
        max_tokens=300,
        messages=[{
            "role": "user",
            "content": (
                f"You are a financial analyst. Based on the following P&L statement data, "
                f"write exactly 3 concise sentences summarizing the company's financial performance. "
                f"Focus on revenue, profitability, and any notable trends.\n\n{text}"
            ),
        }],
    )
    return message.content[0].text


# ── Page config ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Cash Disbursement", page_icon="💸", layout="wide")

# ── Session state ────────────────────────────────────────────────────────────
if "sources"            not in st.session_state:
    st.session_state.sources, st.session_state.pl_classifications = load_config()
if "pl_classifications" not in st.session_state: st.session_state.pl_classifications = {}
if "preview_df"         not in st.session_state: st.session_state.preview_df  = None
if "output_urls"        not in st.session_state: st.session_state.output_urls = []

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.title("💸 Cash Disbursement")
    page = st.radio("Navigation", ["📋 Consolidation", "📒 General Ledger", "📊 Disbursement Summary", "📈 P&L Statement"], label_visibility="collapsed")
    st.divider()

    # Credentials
    sa_valid  = False
    env_json  = os.environ.get("SERVICE_ACCOUNT_JSON")
    if env_json:
        try:
            sa = json.loads(env_json)
            st.success("Service account loaded (env)")
            st.caption(sa.get("client_email", ""))
            sa_valid = True
        except json.JSONDecodeError:
            st.error("SERVICE_ACCOUNT_JSON env var is invalid JSON")
    elif os.path.exists(SERVICE_ACCOUNT_FILE):
        try:
            with open(SERVICE_ACCOUNT_FILE) as f:
                sa = json.load(f)
            st.success("Service account loaded")
            st.caption(sa.get("client_email", ""))
            sa_valid = True
        except json.JSONDecodeError:
            st.error("service_account.json is empty or invalid JSON")
            st.info("Paste your Google service account credentials into service_account.json.")
    else:
        st.error("service_account.json not found")
        st.info("Place your Google service account JSON file in the same folder as this app.")

# ── Helpers ──────────────────────────────────────────────────────────────────
def get_creds():
    env_json = os.environ.get("SERVICE_ACCOUNT_JSON")
    if env_json:
        return Credentials.from_service_account_info(json.loads(env_json), scopes=SCOPES)
    return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

def col_letter(idx):
    result, idx = "", idx + 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result

def filter_rows(rows):
    if not rows:
        return rows
    header = rows[0]
    try:
        idx = [h.strip().lower() for h in header].index("amount paid")
    except ValueError:
        return rows
    return [header] + [r for r in rows[1:] if r[idx].strip()]

def read_sheet_rows(creds, sheet_id, tab_name):
    gc = gspread.authorize(creds)
    try:
        ws = gc.open_by_key(sheet_id).worksheet(tab_name)
        return ws.get_all_values()
    except gspread.exceptions.APIError as e:
        if "not be an Office file" not in str(e):
            raise
        drive = build("drive", "v3", credentials=creds)
        file_meta = drive.files().get(fileId=sheet_id, fields="mimeType").execute()
        mime = file_meta.get("mimeType", "")
        buffer = io.BytesIO()
        req = (drive.files().export_media(fileId=sheet_id,
               mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
               if "google-apps.spreadsheet" in mime
               else drive.files().get_media(fileId=sheet_id))
        dl = MediaIoBaseDownload(buffer, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buffer.seek(0)
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
            df = pd.read_excel(buffer, sheet_name=tab_name, dtype=str).fillna("")
        df = df.loc[:, ~df.columns.str.startswith("Unnamed:")]
        return [df.columns.tolist()] + df.values.tolist()

def write_to_sheet(gc, output_sheet_id, output_tab, rows):
    out_sheet = gc.open_by_key(output_sheet_id)
    try:
        ws_out = out_sheet.worksheet(output_tab)
    except gspread.exceptions.WorksheetNotFound:
        ws_out = out_sheet.add_worksheet(title=output_tab, rows=len(rows) + 10, cols=50)
    ws_out.clear()
    ws_out.update(range_name="A1", values=rows)
    ws_out.format("1:1", {"textFormat": {"bold": True}})
    return f"https://docs.google.com/spreadsheets/d/{output_sheet_id}"

# ════════════════════════════════════════════════════════════════════════════
# PAGE: CONSOLIDATION
# ════════════════════════════════════════════════════════════════════════════
if page == "📋 Consolidation":
    st.title("📋 Consolidation")
    st.caption("Extract cash disbursement entries from multiple Google Sheets into one consolidated sheet.")

    # ── Source Sheets Table ──────────────────────────────────────────────────
    st.subheader("Source Sheets")
    hcols = st.columns([1.5, 2.5, 2, 2.5, 1.5, 0.4])
    for col, label in zip(hcols, ["Label", "Source Sheet ID", "Source Tab", "Output Sheet ID", "Output Tab", ""]):
        col.markdown(f"**{label}**")

    updated_sources, to_delete = [], None
    for i, src in enumerate(st.session_state.sources):
        cols = st.columns([1.5, 2.5, 2, 2.5, 1.5, 0.4])
        label        = cols[0].text_input("Label",           value=src.get("label", ""),          key=f"label_{i}",  label_visibility="collapsed")
        sheet_id     = cols[1].text_input("Source Sheet ID", value=src.get("sheet_id", ""),        key=f"id_{i}",     label_visibility="collapsed")
        tab          = cols[2].text_input("Source Tab",      value=src.get("tab", ""),             key=f"tab_{i}",    label_visibility="collapsed")
        out_sheet_id = cols[3].text_input("Output Sheet ID", value=src.get("output_sheet_id", ""), key=f"outid_{i}",  label_visibility="collapsed")
        out_tab      = cols[4].text_input("Output Tab",      value=src.get("output_tab", ""),      key=f"outtab_{i}", label_visibility="collapsed")
        if cols[5].button("✕", key=f"del_{i}"):
            to_delete = i
        updated_sources.append({
            "label": label.strip(), "sheet_id": sheet_id.strip().rstrip("/"),
            "tab": tab.strip(), "output_sheet_id": out_sheet_id.strip().rstrip("/"),
            "output_tab": out_tab.strip(),
        })

    if to_delete is not None:
        updated_sources.pop(to_delete)
    st.session_state.sources = updated_sources

    col_add, col_save = st.columns(2)
    with col_add:
        if st.button("＋ Add Sheet", width="stretch"):
            st.session_state.sources.append({
                "label": "", "sheet_id": "", "tab": "",
                "output_sheet_id": "178dBRrrsME9IB_kDtlG7QasGsz-A9x2H2hsygr0ZXZ8",
                "output_tab": "",
            })
            st.rerun()
    with col_save:
        if st.button("💾 Save Configuration", width="stretch"):
            save_config(st.session_state.sources)
            st.toast("Configuration saved!", icon="✅")


    st.divider()

    # ── Actions ──────────────────────────────────────────────────────────────
    col_preview, col_run = st.columns(2)

    with col_preview:
        if st.button("🔍 Preview Data", width="stretch", disabled=not sa_valid):
            all_rows, header, errors = [], None, []
            with st.spinner("Fetching data..."):
                try:
                    creds = get_creds()
                    for src in st.session_state.sources:
                        if not src["sheet_id"] or not src["tab"]:
                            continue
                        try:
                            rows = read_sheet_rows(creds, src["sheet_id"], src["tab"])
                            if not rows:
                                errors.append(f"No data in '{src['tab']}'")
                                continue
                            if header is None:
                                header = rows[0]
                            all_rows.extend(rows[1:] if len(rows) > 1 else [])
                        except Exception as e:
                            errors.append(f"{src.get('label') or src['tab']}: {e}")
                    if all_rows and header:
                        filtered = filter_rows([header] + all_rows)
                        st.session_state.preview_df = pd.DataFrame(filtered[1:], columns=filtered[0])
                    else:
                        st.session_state.preview_df = None
                except Exception as e:
                    st.error(f"Auth error: {e}")
            for err in errors:
                st.warning(err)

    with col_run:
        has_valid = any(s["sheet_id"] and s["tab"] and s["output_sheet_id"] and s["output_tab"]
                        for s in st.session_state.sources)
        if st.button("🚀 Run Consolidation", width="stretch", type="primary",
                     disabled=not (sa_valid and has_valid)):
            errors, output_urls = [], []
            st.session_state.output_urls = []

            groups: dict[tuple, list] = {}
            for src in st.session_state.sources:
                if not (src["sheet_id"] and src["tab"] and src["output_sheet_id"] and src["output_tab"]):
                    continue
                groups.setdefault((src["output_sheet_id"], src["output_tab"]), []).append(src)

            total_steps = sum(len(v) for v in groups.values()) + len(groups)
            step = 0
            progress = st.progress(0, text="Starting...")

            try:
                creds = get_creds()
                gc    = gspread.authorize(creds)

                for (out_id, out_tab), group_sources in groups.items():
                    all_rows, header = [], None

                    for src in group_sources:
                        label = src.get("label") or src["tab"]
                        progress.progress(step / total_steps, text=f"Reading: {label}...")
                        try:
                            rows = read_sheet_rows(creds, src["sheet_id"], src["tab"])
                            if not rows:
                                errors.append(f"No data in '{label}'")
                            else:
                                if header is None:
                                    header = rows[0]
                                    all_rows.append(header)
                                all_rows.extend(rows[1:] if len(rows) > 1 else [])
                        except Exception as e:
                            errors.append(f"{label}: {e}")
                        step += 1

                    if all_rows:
                        all_rows = filter_rows(all_rows)
                        progress.progress(step / total_steps, text=f"Writing → {out_tab}...")
                        url = write_to_sheet(gc, out_id, out_tab, all_rows)
                        output_urls.append({"label": out_tab, "url": url})
                        st.session_state.preview_df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
                    step += 1

                st.session_state.output_urls = output_urls
                progress.progress(1.0, text="Done!")

            except Exception as e:
                st.error(f"Error: {e}")
            for err in errors:
                st.warning(err)

    if st.session_state.output_urls:
        st.success(f"Consolidation complete — {len(st.session_state.output_urls)} output(s) written.")
        for item in st.session_state.output_urls:
            st.link_button(f"📄 Open: {item['label']}", item["url"], width="stretch")

    if st.session_state.preview_df is not None:
        df = st.session_state.preview_df
        st.subheader(f"Preview — {len(df):,} rows")
        st.dataframe(df, width="stretch", height=400)


# ════════════════════════════════════════════════════════════════════════════
# PAGE: GENERAL LEDGER
# ════════════════════════════════════════════════════════════════════════════
elif page == "📒 General Ledger":
    st.title("📒 General Ledger")

    # Build list of unique output tabs from config
    output_options = []
    seen = set()
    for src in st.session_state.sources:
        key = (src.get("output_sheet_id", ""), src.get("output_tab", ""))
        if key[0] and key[1] and key not in seen:
            seen.add(key)
            label = src.get("label") or src.get("output_tab", "")
            output_options.append({"label": src.get("output_tab", ""), "sheet_id": key[0], "tab": key[1]})

    if not output_options:
        st.info("No output tabs configured. Set up sources in the Consolidation page first.")
        st.stop()

    if not sa_valid:
        st.warning("Service account not loaded. Check the sidebar.")
        st.stop()

    # ── Controls ─────────────────────────────────────────────────────────────
    col_tab, col_month = st.columns([2, 3])

    tab_labels = [o["label"] for o in output_options]
    selected_tab_label = col_tab.selectbox("Output Tab", tab_labels)
    selected_option = next(o for o in output_options if o["label"] == selected_tab_label)

    selected_months = col_month.multiselect(
        "Filter by Month", MONTHS,
        placeholder="All months (no filter)"
    )

    # ── Load data ─────────────────────────────────────────────────────────────
    @st.cache_data(ttl=300, show_spinner="Loading data...")
    def load_output_tab(sheet_id, tab):
        creds = get_creds()
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(sheet_id).worksheet(tab)
        return ws.get_all_values()

    try:
        raw = load_output_tab(selected_option["sheet_id"], selected_option["tab"])
    except Exception as e:
        st.error(f"Could not load tab: {e}")
        st.stop()

    if not raw or len(raw) < 2:
        st.warning("No data found in the selected tab.")
        st.stop()

    df = pd.DataFrame(raw[1:], columns=[h.strip() for h in raw[0]])

    # Normalise column names for lookup
    df.columns = [c.strip() for c in df.columns]
    col_lower   = {c.lower(): c for c in df.columns}

    month_col = col_lower.get("month")
    acct_col  = col_lower.get("account titles") or col_lower.get("account title")
    tb_col    = col_lower.get("tb amount")

    missing = [n for n, c in [("MONTH", month_col), ("ACCOUNT TITLES", acct_col), ("TB Amount", tb_col)] if not c]
    if missing:
        st.error(f"Missing columns in selected tab: {', '.join(missing)}")
        st.stop()

    # ── Apply month filter ────────────────────────────────────────────────────
    if selected_months:
        df = df[df[month_col].str.strip().isin(selected_months)]

    # ── Convert TB Amount to numeric ──────────────────────────────────────────
    df[tb_col] = pd.to_numeric(
        df[tb_col].astype(str).str.replace(",", "").str.strip(),
        errors="coerce"
    ).fillna(0)

    # Remove blank account title rows and zero TB rows
    df = df[df[acct_col].str.strip() != ""]
    df = df[df[tb_col] != 0]
    df[acct_col] = df[acct_col].str.replace(r"[\r\n]+", " ", regex=True).str.strip()

    # Sort by account title then by month order
    month_order = {m: i for i, m in enumerate(MONTHS)}
    df["_month_order"] = df[month_col].str.strip().map(month_order).fillna(99)
    df = df.sort_values([acct_col, "_month_order"]).drop(columns=["_month_order"])

    # ── Display ───────────────────────────────────────────────────────────────
    month_label = ", ".join(selected_months) if selected_months else "All Months"
    total_tb    = df[tb_col].sum()

    display_cols = [c for c in df.columns if c != acct_col]

    col_title, col_dl = st.columns([3, 1])
    col_title.subheader(f"{selected_tab_label}  ·  {month_label}  ·  {len(df):,} entries")
    col_dl.download_button(
        label="⬇️ Download Excel",
        data=build_gl_excel(df, acct_col, tb_col, display_cols),
        file_name=f"general_ledger_{selected_tab_label}_{month_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    for acct_title, group in df.groupby(acct_col, sort=True):
        st.markdown(f"**{acct_title}**")
        st.dataframe(group[display_cols].reset_index(drop=True),
                     use_container_width=True, hide_index=True)
        subtotal = group[tb_col].sum()
        st.markdown(
            f"<div style='text-align:right; font-size:0.95rem; font-weight:600; "
            f"padding:2px 4px; margin-bottom:16px;'>"
            f"Subtotal: &nbsp;&nbsp; {subtotal:,.2f}</div>",
            unsafe_allow_html=True,
        )

    st.divider()
    st.markdown(
        f"<div style='text-align:right; font-size:1.1rem; font-weight:700; padding:6px 4px;'>"
        f"Total TB Amount: &nbsp;&nbsp; {total_tb:,.2f}</div>",
        unsafe_allow_html=True,
    )


# ════════════════════════════════════════════════════════════════════════════
# PAGE: DISBURSEMENT SUMMARY
# ════════════════════════════════════════════════════════════════════════════
elif page == "📊 Disbursement Summary":
    st.title("📊 Disbursement Summary")

    # ── Output tab selector ───────────────────────────────────────────────────
    output_options = []
    seen = set()
    for src in st.session_state.sources:
        key = (src.get("output_sheet_id", ""), src.get("output_tab", ""))
        if key[0] and key[1] and key not in seen:
            seen.add(key)
            output_options.append({"label": src.get("output_tab", ""), "sheet_id": key[0], "tab": key[1]})

    if not output_options:
        st.info("No output tabs configured. Set up sources in the Consolidation page first.")
        st.stop()

    if not sa_valid:
        st.warning("Service account not loaded. Check the sidebar.")
        st.stop()

    col_sel, col_dl = st.columns([2, 1])
    tab_labels       = [o["label"] for o in output_options]
    selected_label   = col_sel.selectbox("Output Tab", tab_labels)
    selected_option  = next(o for o in output_options if o["label"] == selected_label)

    # ── Load data ─────────────────────────────────────────────────────────────
    @st.cache_data(ttl=300, show_spinner="Loading data...")
    def load_ds_tab(sheet_id, tab):
        creds = get_creds()
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(sheet_id).worksheet(tab)
        return ws.get_all_values()

    try:
        raw = load_ds_tab(selected_option["sheet_id"], selected_option["tab"])
    except Exception as e:
        st.error(f"Could not load tab: {e}")
        st.stop()

    if not raw or len(raw) < 2:
        st.warning("No data found in the selected tab.")
        st.stop()

    df = pd.DataFrame(raw[1:], columns=[h.strip() for h in raw[0]])
    df.columns   = [c.strip() for c in df.columns]
    col_lower    = {c.lower(): c for c in df.columns}

    month_col = col_lower.get("month")
    acct_col  = col_lower.get("account titles") or col_lower.get("account title")
    tb_col    = col_lower.get("tb amount")

    missing = [n for n, c in [("MONTH", month_col), ("ACCOUNT TITLES", acct_col), ("TB Amount", tb_col)] if not c]
    if missing:
        st.error(f"Missing columns: {', '.join(missing)}")
        st.stop()

    # ── Convert & pivot ───────────────────────────────────────────────────────
    df[tb_col] = pd.to_numeric(
        df[tb_col].astype(str).str.replace(",", "").str.strip(), errors="coerce"
    ).fillna(0)
    df = df[df[acct_col].str.strip() != ""]
    df[acct_col] = df[acct_col].str.replace(r"[\r\n]+", " ", regex=True).str.strip()

    active_months = [m for m in MONTHS if m in df[month_col].str.strip().unique()]
    pivot = (
        df.groupby([acct_col, month_col])[tb_col]
        .sum()
        .unstack(fill_value=0)
        .reindex(columns=active_months, fill_value=0)
    )
    pivot["As at Year-end"] = pivot.sum(axis=1)

    def fmt(v):
        return f"{v:,.2f}" if v != 0 else "-"

    def parse_prefix(title):
        t = title.strip()
        for p in ["COS", "OPEX", "RECON", "SALES"]:
            if t.upper().startswith(p):
                return p, t[len(p):].strip()
        return "OTHER", t

    # Categorise accounts
    sections = {"OTHER": [], "SALES": [], "COS": [], "OPEX": [], "RECON": []}
    for title in pivot.index:
        prefix, _ = parse_prefix(title)
        sections[prefix].append(title)

    SECTION_LABELS = {
        "OTHER":  "ASSETS / LIABILITIES & EQUITY",
        "SALES":  "SALES",
        "COS":    "COST OF SALES",
        "OPEX":   "OPERATING EXPENSES",
        "RECON":  "RECON",
    }

    # ── Build display dataframe ───────────────────────────────────────────────
    display_rows = []
    col_headers  = ["Account", "Prefix"] + active_months + ["As at Year-end"]

    def add_section_header(label):
        display_rows.append({"Account": label, "Prefix": "", "_is_header": True,
                              **{m: "" for m in active_months}, "As at Year-end": ""})

    def add_account_row(title, prefix_label, display_name):
        row = {"Account": display_name, "Prefix": prefix_label, "_is_header": False}
        for m in active_months:
            row[m] = fmt(pivot.loc[title, m]) if title in pivot.index else "-"
        row["As at Year-end"] = fmt(pivot.loc[title, "As at Year-end"]) if title in pivot.index else "-"
        display_rows.append(row)

    def add_subtotal_row(label, titles):
        row = {"Account": label, "Prefix": "", "_is_header": "subtotal"}
        for m in active_months:
            total = sum(pivot.loc[t, m] for t in titles if t in pivot.index)
            row[m] = fmt(total)
        total_ye = sum(pivot.loc[t, "As at Year-end"] for t in titles if t in pivot.index)
        row["As at Year-end"] = fmt(total_ye)
        display_rows.append(row)

    for section_key in ["OTHER", "SALES", "COS", "OPEX", "RECON"]:
        titles = sections[section_key]
        if not titles:
            continue
        add_section_header(SECTION_LABELS[section_key])
        for title in sorted(titles):
            prefix_label, display_name = parse_prefix(title)
            add_account_row(title, prefix_label if prefix_label != "OTHER" else "", display_name)
        if section_key in ("COS", "OPEX"):
            add_subtotal_row(f"Total {SECTION_LABELS[section_key]}", titles)

    result_df = pd.DataFrame(display_rows).drop(columns=["_is_header"])

    # ── Download ──────────────────────────────────────────────────────────────
    col_dl.download_button(
        "⬇️ Download Excel",
        data=df_to_excel_bytes(result_df),
        file_name=f"disbursement_summary_{selected_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # ── Render ────────────────────────────────────────────────────────────────
    st.markdown(f"### {selected_label} — DISBURSEMENT SUMMARY")

    header_html = "<tr><th>Prefix</th><th>Account</th>" + \
                  "".join(f"<th>{m[:3]}</th>" for m in active_months) + \
                  "<th>As at Year-end</th></tr>"

    rows_html = ""
    for row in display_rows:
        is_header = row["_is_header"]
        if is_header is True:
            style = "background:#1e3a5f;color:white;font-weight:700;"
            prefix_cell = f'<td style="{style}"></td>'
            acct_cell   = f'<td style="{style}" colspan="1">{row["Account"]}</td>'
        elif is_header == "subtotal":
            style = "background:#d0e8ff;font-weight:700;"
            prefix_cell = f'<td style="{style}"></td>'
            acct_cell   = f'<td style="{style}">{row["Account"]}</td>'
        else:
            style = ""
            prefix_cell = f'<td style="color:#555;font-size:0.8rem;">{row["Prefix"]}</td>'
            acct_cell   = f'<td>{row["Account"]}</td>'

        month_cells = "".join(
            f'<td style="{style}text-align:right;">{row.get(m, "-")}</td>'
            for m in active_months
        )
        ye_cell = f'<td style="{style}text-align:right;font-weight:600;">{row.get("As at Year-end", "-")}</td>'
        rows_html += f"<tr>{prefix_cell}{acct_cell}{month_cells}{ye_cell}</tr>"

    table_html = f"""
    <style>
      .ds-table {{ width:100%; border-collapse:collapse; font-size:0.85rem; }}
      .ds-table th {{ background:#1e3a5f; color:white; padding:6px 8px; text-align:center; }}
      .ds-table td {{ padding:4px 8px; border-bottom:1px solid #eee; }}
      .ds-table tr:hover td {{ background:#f5f9ff; }}
    </style>
    <table class="ds-table"><thead>{header_html}</thead><tbody>{rows_html}</tbody></table>
    """
    st.markdown(table_html, unsafe_allow_html=True)


# ════════════════════════════════════════════════════════════════════════════
# PAGE: P&L STATEMENT
# ════════════════════════════════════════════════════════════════════════════
elif page == "📈 P&L Statement":
    st.title("📈 P&L Statement")

    output_options = []
    seen = set()
    for src in st.session_state.sources:
        key = (src.get("output_sheet_id", ""), src.get("output_tab", ""))
        if key[0] and key[1] and key not in seen:
            seen.add(key)
            output_options.append({"label": src.get("output_tab", ""), "sheet_id": key[0], "tab": key[1]})

    if not output_options:
        st.info("No output tabs configured. Set up sources in the Consolidation page first.")
        st.stop()
    if not sa_valid:
        st.warning("Service account not loaded. Check the sidebar.")
        st.stop()

    col_sel, col_name, col_dl = st.columns([1.5, 2, 1])
    selected_label  = col_sel.selectbox("Output Tab", [o["label"] for o in output_options])
    selected_option = next(o for o in output_options if o["label"] == selected_label)
    company_name    = col_name.text_input("Company Name", value=selected_label)

    @st.cache_data(ttl=300, show_spinner="Loading data...")
    def load_pl_tab(sheet_id, tab):
        creds = get_creds()
        gc = gspread.authorize(creds)
        ws = gc.open_by_key(sheet_id).worksheet(tab)
        return ws.get_all_values()

    try:
        raw = load_pl_tab(selected_option["sheet_id"], selected_option["tab"])
    except Exception as e:
        st.error(f"Could not load tab: {e}")
        st.stop()

    if not raw or len(raw) < 2:
        st.warning("No data found in the selected tab.")
        st.stop()

    df = pd.DataFrame(raw[1:], columns=[h.strip() for h in raw[0]])
    col_lower = {c.strip().lower(): c.strip() for c in df.columns}

    month_col = col_lower.get("month")
    acct_col  = col_lower.get("account titles") or col_lower.get("account title")
    tb_col    = col_lower.get("tb amount")

    missing = [n for n, c in [("MONTH", month_col), ("ACCOUNT TITLES", acct_col), ("TB Amount", tb_col)] if not c]
    if missing:
        st.error(f"Missing columns: {', '.join(missing)}")
        st.stop()

    df[tb_col] = pd.to_numeric(
        df[tb_col].astype(str).str.replace(",", "").str.strip(), errors="coerce"
    ).fillna(0)
    df = df[df[acct_col].str.strip() != ""]
    df[acct_col] = df[acct_col].str.replace(r"[\r\n]+", " ", regex=True).str.strip()

    all_titles = sorted(df[acct_col].dropna().unique().tolist())

    # ── Account classification ────────────────────────────────────────────────
    # Use hardcoded ACCOUNT_TAGS first, then prefix detection, then manual overrides
    saved_cls    = st.session_state.pl_classifications.get(selected_label, {"cos": [], "opex": [], "sales": []})

    hardcoded_cos   = [t for t in all_titles if ACCOUNT_TAGS.get(t.strip()) == "COS"  or t.strip().upper().startswith("COS")]
    hardcoded_opex  = [t for t in all_titles if ACCOUNT_TAGS.get(t.strip()) == "OPEX" or t.strip().upper().startswith("OPEX")]
    hardcoded_sales = [t for t in all_titles if ACCOUNT_TAGS.get(t.strip()) == "SALES"or t.strip().upper().startswith("SALES")]

    with st.expander("⚙️ Account Classification Override", expanded=False):
        st.caption("Accounts are auto-classified from the hardcoded list. Use this to add or remove accounts.")
        cc1, cc2, cc3 = st.columns(3)
        manual_cos   = cc1.multiselect("Additional COS",   all_titles, default=saved_cls.get("cos", []),   key="cls_cos")
        manual_opex  = cc2.multiselect("Additional OPEX",  all_titles, default=saved_cls.get("opex", []),  key="cls_opex")
        manual_sales = cc3.multiselect("Additional SALES", all_titles, default=saved_cls.get("sales", []), key="cls_sales")
        if st.button("💾 Save Override"):
            st.session_state.pl_classifications[selected_label] = {
                "cos": manual_cos, "opex": manual_opex, "sales": manual_sales
            }
            save_config(st.session_state.sources, st.session_state.pl_classifications)
            st.toast("Override saved!", icon="✅")

    cos_accounts   = list(set(hardcoded_cos   + manual_cos))
    opex_accounts  = list(set(hardcoded_opex  + manual_opex))
    sales_accounts = list(set(hardcoded_sales + manual_sales))

    # Sum TB Amount by account and month
    monthly = (
        df.groupby([acct_col, month_col])[tb_col]
        .sum()
        .unstack(fill_value=0)
    )
    active_months = [m for m in MONTHS if m in monthly.columns]
    monthly = monthly.reindex(columns=active_months, fill_value=0)

    def section_sum(accounts):
        rows = [t for t in accounts if t in monthly.index]
        if not rows:
            return {m: 0.0 for m in active_months}
        return monthly.loc[rows].sum().to_dict()

    def fmt_pl(v):
        if v == 0:
            return "-"
        if v < 0:
            return f"({abs(v):,.0f})"
        return f"{v:,.0f}"

    # ── Calculate P&L lines ───────────────────────────────────────────────────
    revenues = section_sum(sales_accounts)
    cos      = section_sum(cos_accounts)
    opex     = section_sum(opex_accounts)

    gross_profit   = {m: revenues[m] - cos[m]                      for m in active_months}
    income_ops     = {m: gross_profit[m] - opex[m]                 for m in active_months}
    income_pretax  = {m: income_ops[m]                             for m in active_months}
    net_income     = {m: income_pretax[m]                          for m in active_months}

    def ye(d):
        return sum(d.values())

    # ── Build table rows ──────────────────────────────────────────────────────
    all_cols = active_months + ["As at Year-end"]

    def make_row(label, data: dict, is_header=False, is_calc=False, is_empty=False):
        if is_empty:
            return {"label": label, "is_header": is_header, "is_calc": is_calc,
                    **{m: "" for m in all_cols}}
        values = {m: fmt_pl(data.get(m, 0)) for m in active_months}
        values["As at Year-end"] = fmt_pl(ye(data))
        return {"label": label, "is_header": is_header, "is_calc": is_calc, **values}

    pl_rows = [
        make_row("Revenues",                                  revenues,      is_header=True),
        make_row("Cost of Sales",                             cos),
        make_row("Gross profit (loss)",                       gross_profit,  is_calc=True),
        make_row("Operating expenses",                        opex),
        make_row("Other operating income, net",               {},            is_empty=True),
        make_row("Income (loss) from operations",             income_ops,    is_calc=True),
        make_row("Interest expense",                          {},            is_empty=True),
        make_row("Income (loss) before income tax",           income_pretax, is_calc=True),
        make_row("Benefit from (provision for) income tax",   {},            is_empty=True),
        make_row("Net income (loss) for the year",            net_income,    is_calc=True),
    ]

    # ── Download ──────────────────────────────────────────────────────────────
    dl_df = pd.DataFrame([{k: v for k, v in r.items() if k not in ("is_header", "is_calc")}
                           for r in pl_rows])
    col_dl.download_button(
        "⬇️ Download Excel",
        data=df_to_excel_bytes(dl_df),
        file_name=f"pl_{selected_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    # ── Render HTML table ─────────────────────────────────────────────────────
    th = "".join(f"<th>{m[:3]}</th>" for m in active_months) + "<th>As at Year-end</th>"
    header_html = f"<tr><th style='text-align:left;'>Description</th>{th}</tr>"

    rows_html = ""
    for row in pl_rows:
        if row["is_calc"]:
            rs = "background:#e8f4e8;font-weight:700;"
        elif row["is_header"]:
            rs = "background:#1e3a5f;color:white;font-weight:700;"
        else:
            rs = ""

        label_cell = f'<td style="{rs}padding:5px 10px;">{row["label"]}</td>'
        val_cells  = "".join(
            f'<td style="{rs}text-align:right;padding:5px 8px;">{row.get(m, "")}</td>'
            for m in all_cols
        )
        rows_html += f"<tr>{label_cell}{val_cells}</tr>"

    pl_html = f"""
    <style>
      .pl-table {{ width:100%; border-collapse:collapse; font-size:0.88rem; margin-top:8px; }}
      .pl-table th {{ background:#1e3a5f; color:white; padding:6px 8px; text-align:center; }}
      .pl-table td {{ border-bottom:1px solid #eee; }}
      .pl-table tr:hover td {{ background:#f9f9f9; }}
    </style>
    <div style="font-size:1rem;font-weight:700;margin-bottom:2px;">{company_name}</div>
    <div style="font-size:0.9rem;color:#555;margin-bottom:12px;">INCOME STATEMENT</div>
    <table class="pl-table">
      <thead>{header_html}</thead>
      <tbody>{rows_html}</tbody>
    </table>
    """
    st.markdown(pl_html, unsafe_allow_html=True)

    # ── AI Summary ────────────────────────────────────────────────────────────
    st.divider()
    has_api_key = bool(os.environ.get("ANTHROPIC_API_KEY"))
    if not has_api_key:
        st.info("Set the `ANTHROPIC_API_KEY` environment variable to enable AI summaries.")
    else:
        if st.button("✨ Generate AI Summary", type="primary"):
            with st.spinner("Generating summary..."):
                try:
                    summary = generate_pl_summary(company_name, active_months, pl_rows)
                    st.info(summary)
                except Exception as e:
                    st.error(f"Could not generate summary: {e}")
