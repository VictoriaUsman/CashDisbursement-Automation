# Cash Disbursement Consolidator

Extracts cash disbursement entries from one or more Google Sheets and writes them into a consolidated Google Sheet. Supports multiple sources and output destinations, with persistent configuration.

---

## Requirements

- Python 3.8+
- A Google Cloud Service Account with Sheets and Drive APIs enabled

---

## Setup (One-Time)

### 1. Create a Google Cloud Service Account

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Create or select a project
3. Enable **Google Sheets API** and **Google Drive API**
4. Navigate to **IAM & Admin → Service Accounts → Create Service Account**
5. Download the JSON key and save it as `service_account.json` in this folder

### 2. Share Sheets with the Service Account

- Open `service_account.json` and copy the `client_email` value
- Share each **source sheet** with that email → **Viewer** (or Editor)
- Share each **output sheet** with that email → **Editor**

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

---

## Usage

```bash
streamlit run app.py
```

Opens in your browser at `http://localhost:8501`.

---

## UI Overview

### Source Sheets Table

Each row represents one source-to-output mapping:

| Field | Description |
|---|---|
| **Label** | A friendly name for this source (e.g. DAVAO, LOMA) |
| **Source Sheet ID** | ID of the Google Sheet to read from |
| **Source Tab** | Tab name inside the source sheet |
| **Output Sheet ID** | ID of the Google Sheet to write to |
| **Output Tab** | Tab name to write results into |

- Click **＋ Add Sheet** to add a new row
- Click **✕** to remove a row
- Click **💾 Save Configuration** to persist all entries to `config.json` — they will reload automatically on next launch

### Grouping Behavior

If multiple source rows share the same **Output Sheet ID + Output Tab**, their data is consolidated into that single tab (with one shared header row).

### Buttons

| Button | Action |
|---|---|
| **🔍 Preview Data** | Fetches all source data and displays it in a table — no writes |
| **🚀 Run Consolidation** | Reads all sources, filters rows, writes to each output tab |

### Filtering

Rows where **Amount Paid** is blank are automatically excluded from the output.

---

## Finding a Sheet ID

The Sheet ID is the long string in the Google Sheets URL:

```
https://docs.google.com/spreadsheets/d/<SHEET_ID>/edit
```

---

## File Structure

```
Accounting/
├── app.py                # Streamlit web UI
├── consolidate.py        # Headless CLI script (legacy)
├── config.json           # Saved sheet configuration (auto-generated)
├── requirements.txt      # Python dependencies
├── service_account.json  # Google credentials (DO NOT share or commit)
└── README.md
```

---

## Notes

- `service_account.json` contains sensitive credentials — keep it private and never commit it to git
- Each run **clears and rewrites** the output tab — old data is not preserved
- Trailing slashes in Sheet IDs are stripped automatically
- Office/Excel files stored in Google Drive are supported via the Drive API fallback
