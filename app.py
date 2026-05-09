import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import pandas as pd
import io
import os
import json

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(__file__), "service_account.json")
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "config.json")

DEFAULT_SOURCES = [
    {
        "label": "DAVAO",
        "sheet_id": "1uGiFSN9L3bUl__o5fXQK9SNgv7l4fL7g",
        "tab": "DAVAO Cash Disbursement",
        "output_sheet_id": "178dBRrrsME9IB_kDtlG7QasGsz-A9x2H2hsygr0ZXZ8",
        "output_tab": "DAVAO",
    }
]

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE) as f:
                return json.load(f).get("sources", DEFAULT_SOURCES)
        except Exception:
            pass
    return DEFAULT_SOURCES

def save_config(sources):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"sources": sources}, f, indent=2)

st.set_page_config(page_title="Cash Disbursement Consolidator", page_icon="💸", layout="wide")
st.title("💸 Cash Disbursement Consolidator")
st.caption("Extract cash disbursement entries from multiple Google Sheets into one consolidated sheet.")

# ── Session state ────────────────────────────────────────────────────────────
if "sources" not in st.session_state:
    st.session_state.sources = load_config()
if "preview_df" not in st.session_state:
    st.session_state.preview_df = None
if "output_urls" not in st.session_state:
    st.session_state.output_urls = []

# ── Sidebar ──────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuration")
    sa_valid = False
    env_json = os.environ.get("SERVICE_ACCOUNT_JSON")
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
        info = json.loads(env_json)
        return Credentials.from_service_account_info(info, scopes=SCOPES)
    return Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

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
        if "google-apps.spreadsheet" in mime:
            req = drive.files().export_media(
                fileId=sheet_id,
                mimeType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            req = drive.files().get_media(fileId=sheet_id)
        dl = MediaIoBaseDownload(buffer, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        buffer.seek(0)
        import warnings
        with warnings.catch_warnings():
            warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")
            df = pd.read_excel(buffer, sheet_name=tab_name, dtype=str).fillna("")
        return [df.columns.tolist()] + df.values.tolist()

AMOUNT_COLS = [
    "amount paid", "vatable amount", "vat amount", "non vat amount",
    "other  charges amount", "ewt amount (negative entry)", "tb amount",
]

def col_letter(idx):
    """Convert 0-based column index to spreadsheet column letter."""
    result = ""
    idx += 1
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        result = chr(65 + rem) + result
    return result

def write_general_ledger(gc, output_sheet_id, tab_data: list[tuple[str, list]]):
    """
    tab_data: list of (tab_name, rows) for all consolidated tabs in this output sheet.
    Writes a General Ledger tab with unique account titles and SUMIF formulas
    referencing each consolidated tab — no backend computation.
    """
    # Collect unique account titles and resolve column positions from first valid tab
    all_account_titles: set[str] = set()
    acct_col_letter = None
    amt_col_info: list[tuple[str, str]] = []  # (display_name, col_letter)

    for tab_name, rows in tab_data:
        if not rows or len(rows) < 2:
            continue
        header = [h.strip() for h in rows[0]]
        header_lower = [h.lower() for h in header]

        try:
            acct_idx = header_lower.index("account titles")
        except ValueError:
            continue

        if acct_col_letter is None:
            acct_col_letter = col_letter(acct_idx)
            amt_col_info = [
                (header[i], col_letter(i))
                for i, h in enumerate(header_lower)
                if h in AMOUNT_COLS
            ]

        for row in rows[1:]:
            acct = row[acct_idx].strip() if acct_idx < len(row) else ""
            if acct:
                all_account_titles.add(acct)

    if not acct_col_letter or not all_account_titles:
        return

    sorted_titles = sorted(all_account_titles)

    # Build GL rows: header + one row per account title
    gl_header = ["Account Title"] + [name for name, _ in amt_col_info]
    gl_rows = [gl_header]

    for i, title in enumerate(sorted_titles):
        row_num = i + 2  # row 1 is header
        formula_cells = []
        for _, amt_letter in amt_col_info:
            # Sum across all consolidated tabs using SUMIF
            parts = [
                f"SUMIF('{tab}'!{acct_col_letter}:{acct_col_letter},A{row_num},'{tab}'!{amt_letter}:{amt_letter})"
                for tab, _ in tab_data
            ]
            formula_cells.append("=" + "+".join(parts))
        gl_rows.append([title] + formula_cells)

    out_sheet = gc.open_by_key(output_sheet_id)
    try:
        ws_gl = out_sheet.worksheet("General Ledger")
    except gspread.exceptions.WorksheetNotFound:
        ws_gl = out_sheet.add_worksheet(title="General Ledger", rows=len(gl_rows) + 10, cols=20)
    ws_gl.clear()
    ws_gl.update(range_name="A1", values=gl_rows)
    ws_gl.format("1:1", {"textFormat": {"bold": True}})

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

# ── Source Sheets Table ───────────────────────────────────────────────────────
st.subheader("📋 Source Sheets")

# Column headers
hcols = st.columns([1.5, 2.5, 2, 2.5, 1.5, 0.4])
hcols[0].markdown("**Label**")
hcols[1].markdown("**Source Sheet ID**")
hcols[2].markdown("**Source Tab**")
hcols[3].markdown("**Output Sheet ID**")
hcols[4].markdown("**Output Tab**")

updated_sources = []
to_delete = None

for i, src in enumerate(st.session_state.sources):
    cols = st.columns([1.5, 2.5, 2, 2.5, 1.5, 0.4])
    label        = cols[0].text_input("Label",           value=src.get("label", ""),           key=f"label_{i}",  label_visibility="collapsed")
    sheet_id     = cols[1].text_input("Source Sheet ID", value=src.get("sheet_id", ""),         key=f"id_{i}",     label_visibility="collapsed")
    tab          = cols[2].text_input("Source Tab",      value=src.get("tab", ""),              key=f"tab_{i}",    label_visibility="collapsed")
    out_sheet_id = cols[3].text_input("Output Sheet ID", value=src.get("output_sheet_id", ""),  key=f"outid_{i}",  label_visibility="collapsed")
    out_tab      = cols[4].text_input("Output Tab",      value=src.get("output_tab", ""),       key=f"outtab_{i}", label_visibility="collapsed")
    if cols[5].button("✕", key=f"del_{i}", help="Remove"):
        to_delete = i
    updated_sources.append({
        "label":           label.strip(),
        "sheet_id":        sheet_id.strip().rstrip("/"),
        "tab":             tab.strip(),
        "output_sheet_id": out_sheet_id.strip().rstrip("/"),
        "output_tab":      out_tab.strip(),
    })

if to_delete is not None:
    updated_sources.pop(to_delete)

st.session_state.sources = updated_sources

col_add, col_save = st.columns([1, 1])
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

# ── Actions ──────────────────────────────────────────────────────────────────
col_preview, col_run = st.columns(2)

with col_preview:
    if st.button("🔍 Preview Data", width="stretch", disabled=not sa_valid):
        all_rows = []
        header = None
        errors = []
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
    has_valid_sources = any(s["sheet_id"] and s["tab"] and s["output_sheet_id"] and s["output_tab"] for s in st.session_state.sources)
    if st.button("🚀 Run Consolidation", width="stretch", type="primary", disabled=not (sa_valid and has_valid_sources)):
        errors = []
        output_urls = []
        st.session_state.output_urls = []

        # Group sources by (output_sheet_id, output_tab)
        groups: dict[tuple, list] = {}
        for src in st.session_state.sources:
            if not (src["sheet_id"] and src["tab"] and src["output_sheet_id"] and src["output_tab"]):
                continue
            key = (src["output_sheet_id"], src["output_tab"])
            groups.setdefault(key, []).append(src)

        total_steps = sum(len(v) for v in groups.values()) + len(groups)
        step = 0
        progress = st.progress(0, text="Starting...")

        try:
            creds = get_creds()
            gc = gspread.authorize(creds)

            # Track tab data per output sheet for GL generation
            gl_tab_data: dict[str, list[tuple[str, list]]] = {}

            for (out_id, out_tab), group_sources in groups.items():
                all_rows = []
                header = None

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
                    gl_tab_data.setdefault(out_id, []).append((out_tab, all_rows))
                    output_urls.append({"label": out_tab, "url": url})
                    st.session_state.preview_df = pd.DataFrame(all_rows[1:], columns=all_rows[0])
                step += 1

            # Write General Ledger once per output sheet with SUMIF formulas
            for out_id, tab_data in gl_tab_data.items():
                progress.progress(0.95, text="Writing General Ledger...")
                write_general_ledger(gc, out_id, tab_data)

            st.session_state.output_urls = output_urls
            progress.progress(1.0, text="Done!")

        except Exception as e:
            st.error(f"Error: {e}")

        for err in errors:
            st.warning(err)

# ── Output links ──────────────────────────────────────────────────────────────
if st.session_state.output_urls:
    st.success(f"Consolidation complete — {len(st.session_state.output_urls)} output(s) written.")
    for item in st.session_state.output_urls:
        st.link_button(f"📄 Open: {item['label']}", item["url"], width="stretch")

# ── Preview table ─────────────────────────────────────────────────────────────
if st.session_state.preview_df is not None:
    df = st.session_state.preview_df
    st.subheader(f"Preview — {len(df):,} rows")
    st.dataframe(df, width="stretch", height=400)
